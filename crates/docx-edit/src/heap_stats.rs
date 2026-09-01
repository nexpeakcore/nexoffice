//! Live and peak bytes *requested* from the Rust allocator, for diagnosing what
//! a document costs. This is the sum of `Layout::size()`, so it excludes
//! rounding, allocator metadata and fragmentation, and reads below true
//! occupancy. `WebAssembly.Memory.buffer.byteLength` — the only figure the host
//! could read before — is the opposite error: linear memory the allocator has
//! *claimed*, including free lists, and it never shrinks.
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
    /// `fetch_max` rather than load/compare/store: two threads racing there can
    /// store in the reverse order and leave the mark below a figure that was
    /// genuinely reached — even below `LIVE`, which makes the reading absurd
    /// rather than merely approximate.
    fn record_growth(delta: usize) {
        let live = LIVE.fetch_add(delta, Ordering::Relaxed) + delta;
        PEAK.fetch_max(live, Ordering::Relaxed);
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

    /// The counters are process-wide, so these tests perturb each other's
    /// readings when they run at the same time.
    static SERIAL: std::sync::Mutex<()> = std::sync::Mutex::new(());

    /// Assertions here compare against the chunk alone, never against a
    /// difference between two readings: the test harness allocates and frees on
    /// its own threads throughout, so a delta is not the test's to predict.
    #[test]
    fn counts_an_allocation_and_gives_it_back() {
        let _serial = SERIAL.lock().unwrap_or_else(|e| e.into_inner());
        let held = vec![0_u8; CHUNK];
        let during = live_bytes();
        assert!(
            during >= CHUNK,
            "live must cover an allocation this thread is still holding: {during}"
        );
        drop(held);
        let after = live_bytes();
        // Two readings of a process-wide counter cannot be compared exactly:
        // the harness allocates on its own threads between them. The slack
        // absorbs that while staying two orders of magnitude under the chunk,
        // so it cannot hide a chunk that was never returned.
        const SLACK: usize = 64 * 1024;
        assert!(
            after + CHUNK <= during + SLACK,
            "dropping must give the chunk back: {after} against {during}"
        );
    }

    #[test]
    fn peak_outlives_the_allocation_until_reset() {
        let _serial = SERIAL.lock().unwrap_or_else(|e| e.into_inner());
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
        let _serial = SERIAL.lock().unwrap_or_else(|e| e.into_inner());
        let mut grown = Vec::<u8>::new();
        for _ in 0..CHUNK {
            grown.push(1);
        }
        let live = live_bytes();
        assert!(
            live >= CHUNK,
            "realloc growth must be counted, not just fresh allocations: {live}"
        );
        drop(grown);
    }
}
