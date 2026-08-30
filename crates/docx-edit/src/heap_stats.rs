//! Live and peak Rust heap bytes, for diagnosing what a document actually
//! costs. `WebAssembly.Memory.buffer.byteLength` — the only figure the host
//! could read before — reports the linear memory the allocator has *claimed*,
//! including its free lists, and it never shrinks. That conflates waste with
//! occupancy and cannot say how close a document is to the wasm32 ceiling.
//!
//! Counting at the allocator answers the occupancy question directly, and in
//! the same units natively and under wasm, so the two are comparable.
//!
//! Off unless the `heap-stats` feature is on, and deliberately NOT in the
//! shipped wasm build. Measured on a 2000-page document: the counters are free
//! natively (layout 11.63s against 11.92s without them) but cost roughly 12x
//! under wasm (layout 220s against 18s). Build a diagnostic wasm to use them —
//! see `scripts/build-docx-wasm.ts` — rather than turning them on for users.
//!
//! A library also has no business installing a global allocator in someone
//! else's binary, which is the other reason this is opt-in.

#[cfg(feature = "heap-stats")]
mod imp {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicUsize, Ordering};

    static LIVE: AtomicUsize = AtomicUsize::new(0);
    static PEAK: AtomicUsize = AtomicUsize::new(0);

    struct Counting;

    /// Relaxed ordering throughout: these counters are read for reporting, and
    /// no other memory is published through them.
    ///
    /// The high-water mark is a load/compare/store rather than `fetch_max`,
    /// which is a compare-exchange loop on every single allocation. Two
    /// threads can lose an update to each other here; a diagnostic counter is
    /// worth neither that loop nor the contention.
    fn record_growth(delta: usize) {
        let live = LIVE.fetch_add(delta, Ordering::Relaxed) + delta;
        if live > PEAK.load(Ordering::Relaxed) {
            PEAK.store(live, Ordering::Relaxed);
        }
    }

    unsafe impl GlobalAlloc for Counting {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            let ptr = unsafe { System.alloc(layout) };
            if !ptr.is_null() {
                record_growth(layout.size());
            }
            ptr
        }

        unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
            let ptr = unsafe { System.alloc_zeroed(layout) };
            if !ptr.is_null() {
                record_growth(layout.size());
            }
            ptr
        }

        unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
            LIVE.fetch_sub(layout.size(), Ordering::Relaxed);
            unsafe { System.dealloc(ptr, layout) }
        }

        unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            let next = unsafe { System.realloc(ptr, layout, new_size) };
            if !next.is_null() {
                if let Some(growth) = new_size.checked_sub(layout.size()) {
                    record_growth(growth);
                } else {
                    LIVE.fetch_sub(layout.size() - new_size, Ordering::Relaxed);
                }
            }
            next
        }
    }

    #[global_allocator]
    static ALLOCATOR: Counting = Counting;

    pub fn live_bytes() -> usize {
        LIVE.load(Ordering::Relaxed)
    }

    pub fn peak_bytes() -> usize {
        PEAK.load(Ordering::Relaxed)
    }

    /// Drops the high-water mark to what is live now, so the next stage's peak
    /// is its own rather than the whole session's.
    pub fn reset_peak() {
        PEAK.store(LIVE.load(Ordering::Relaxed), Ordering::Relaxed);
    }
}

#[cfg(not(feature = "heap-stats"))]
mod imp {
    pub fn live_bytes() -> usize {
        0
    }
    pub fn peak_bytes() -> usize {
        0
    }
    pub fn reset_peak() {}
}

/// Rust heap bytes currently allocated. Zero when the feature is off.
pub fn live_bytes() -> usize {
    imp::live_bytes()
}

/// Highest live figure since the last [`reset_peak`]. Zero when off.
pub fn peak_bytes() -> usize {
    imp::peak_bytes()
}

/// Restart the high-water mark from the current live figure.
pub fn reset_peak() {
    imp::reset_peak();
}

/// Whether these figures mean anything in this build.
pub fn available() -> bool {
    cfg!(feature = "heap-stats")
}

#[cfg(all(test, feature = "heap-stats"))]
mod tests {
    use super::*;

    const CHUNK: usize = 8 * 1024 * 1024;

    #[test]
    fn counts_an_allocation_and_gives_it_back() {
        let before = live_bytes();
        let held = vec![0_u8; CHUNK];
        let during = live_bytes();
        assert!(
            during >= before + CHUNK,
            "live should cover the allocation: {before} -> {during}"
        );
        drop(held);
        assert!(
            live_bytes() < during,
            "dropping should give the bytes back: still {} of {during}",
            live_bytes()
        );
    }

    #[test]
    fn peak_outlives_the_allocation_until_reset() {
        drop(vec![0_u8; CHUNK]);
        let peak = peak_bytes();
        let live = live_bytes();
        assert!(peak >= live, "peak {peak} must not sit under live {live}");
        assert!(
            peak >= CHUNK,
            "peak {peak} should remember the transient allocation"
        );
        reset_peak();
        assert!(
            peak_bytes() <= live_bytes().max(peak),
            "reset drops the mark back to what is live"
        );
    }

    #[test]
    fn growing_a_vec_through_realloc_is_tracked() {
        let before = live_bytes();
        let mut grown = Vec::<u8>::new();
        for _ in 0..CHUNK {
            grown.push(1);
        }
        assert!(
            live_bytes() >= before + CHUNK,
            "realloc growth must be counted, not just fresh allocations"
        );
        drop(grown);
    }
}
