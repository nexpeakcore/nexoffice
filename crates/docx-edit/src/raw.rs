//! Raw story mutations using UTF-16 story indices.

use std::sync::Arc;

use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::{
    Any, Assoc, IndexedSequence, Map, MapPrelim, MapRef, Out, ReadTxn, Text, TextRef,
    TransactionMut,
};

use crate::op::{OpError, OpResult};
use crate::{COMMENTS, EditCtx, EditingDoc, KIND_KEY, anchor_value, out_len, story_ref};

/// One low-level story mutation. Indices are UTF-16 story units (every embed = 1).
#[derive(Clone, Debug, PartialEq)]
pub enum RawOp {
    /// Inserts text with explicit attributes.
    Insert {
        index: u32,
        text: String,
        attrs: Attrs,
    },
    /// Remove `len` units starting at `index` (text, embeds, or pilcrows alike).
    Delete { index: u32, len: u32 },
    /// Re-format `len` units at `index`; `Any::Null` values clear an attribute.
    Format { index: u32, len: u32, attrs: Attrs },
    /// Insert a map-backed embed at `index` with discriminator `kind` (`pilcrow`,
    /// `break`, `opaque`, …) and `payload` map entries. `attrs` are the embed's
    /// text-level formatting (usually just tracked-change stamps, if any).
    InsertEmbed {
        index: u32,
        kind: String,
        payload: Vec<(String, Any)>,
        attrs: Attrs,
    },
    /// Sets one key on the map-backed embed at the index.
    SetEmbedAttr { index: u32, key: String, value: Any },
    /// Upserts a side-map comment with sticky UTF-16 ranges.
    SetComment {
        id: String,
        ranges: Vec<(u32, u32)>,
        author: String,
        date: String,
        body: Any,
    },
    /// Remove the side-map comment keyed by `id`. Errors when it does not exist.
    RemoveComment { id: String },
}

/// Finds the map-backed embed sitting exactly at story `index`.
fn embed_at<T: ReadTxn>(story: &yrs::TextRef, txn: &T, index: u32) -> OpResult<MapRef> {
    let mut offset = 0u32;
    for diff in story.diff(txn, YChange::identity) {
        if offset == index {
            if let Out::YMap(map) = diff.insert {
                return Ok(map);
            }
            break;
        }
        offset += out_len(&diff.insert);
        if offset > index {
            break;
        }
    }
    Err(OpError::OutOfBounds {
        index,
        len: story.len(txn),
    })
}

impl EditingDoc {
    /// Applies raw story operations in one transaction.
    pub fn apply_raw_ops(&self, story_id: &str, ops: Vec<RawOp>, ctx: &EditCtx) -> OpResult<()> {
        let mut txn = self.transact_for(ctx);
        apply_raw_ops_to_story(&mut txn, story_id, ops)
    }

    pub(crate) fn apply_raw_story_batches(
        &self,
        batches: Vec<(String, Vec<RawOp>)>,
        ctx: &EditCtx,
    ) -> OpResult<()> {
        let mut txn = self.transact_for(ctx);
        for (story_id, ops) in batches {
            apply_raw_ops_to_story(&mut txn, &story_id, ops)?;
        }
        Ok(())
    }
}

fn apply_raw_ops_to_story(
    txn: &mut TransactionMut<'_>,
    story_id: &str,
    ops: Vec<RawOp>,
) -> OpResult<()> {
    let story = story_ref(txn, story_id).map_err(OpError::from)?;
    let mut writer = StoryWriter::new();
    for op in ops {
        match op {
            RawOp::Insert { index, text, attrs } => {
                guard_index(&story, txn, index)?;
                writer.insert_text(&story, txn, index, &text, attrs)?;
            }
            RawOp::Delete { index, len } => {
                writer.invalidate();
                guard_range(&story, txn, index, len)?;
                story.remove_range(txn, index, len);
            }
            RawOp::Format { index, len, attrs } => {
                writer.invalidate();
                guard_range(&story, txn, index, len)?;
                if len > 0 {
                    story.format(txn, index, len, attrs);
                }
            }
            RawOp::InsertEmbed {
                index,
                kind,
                payload,
                attrs,
            } => {
                guard_index(&story, txn, index)?;
                let embed = writer.insert_embed(&story, txn, index, attrs)?;
                embed.insert(txn, KIND_KEY, kind.as_str());
                for (key, value) in payload {
                    embed.insert(txn, key, value);
                }
            }
            RawOp::SetEmbedAttr { index, key, value } => {
                writer.invalidate();
                let embed = embed_at(&story, txn, index)?;
                embed.insert(txn, key, value);
            }
            RawOp::SetComment {
                id,
                ranges,
                author,
                date,
                body,
            } => {
                writer.invalidate();
                if ranges.is_empty() {
                    return Err(OpError::InvalidComment(
                        "at least one anchored range is required".into(),
                    ));
                }
                let mut anchors = Vec::with_capacity(ranges.len());
                for (start, end) in ranges {
                    let len = end
                        .checked_sub(start)
                        .filter(|len| *len > 0)
                        .ok_or(OpError::InvalidRange { start, end })?;
                    guard_range(&story, txn, start, len)?;
                    let start_anchor =
                        story
                            .sticky_index(txn, start, Assoc::After)
                            .ok_or_else(|| {
                                OpError::InvalidComment("start anchor could not be made".into())
                            })?;
                    let end_anchor =
                        story.sticky_index(txn, end, Assoc::Before).ok_or_else(|| {
                            OpError::InvalidComment("end anchor could not be made".into())
                        })?;
                    anchors.push(anchor_value(story_id, &start_anchor, &end_anchor));
                }
                let comments = txn
                    .get_map(COMMENTS)
                    .expect("comments root is declared by EditingDoc::new");
                let comment = comments.insert(txn, id.as_str(), MapPrelim::default());
                comment.insert(txn, "author", author.as_str());
                comment.insert(txn, "date", date.as_str());
                comment.insert(txn, "parentId", Any::Null);
                comment.insert(txn, "done", false);
                comment.insert(txn, "body", body);
                comment.insert(txn, "anchors", Any::Array(Arc::from(anchors)));
            }
            RawOp::RemoveComment { id } => {
                writer.invalidate();
                let comments = txn
                    .get_map(COMMENTS)
                    .expect("comments root is declared by EditingDoc::new");
                if comments.remove(txn, &id).is_none() {
                    return Err(OpError::UnknownComment(id));
                }
            }
        }
    }
    Ok(())
}

/// Where a run of ascending inserts is writing into a story.
///
/// Seeding a story is exactly such a run, and yrs reaches an insertion index
/// by walking the block list from the story's start — so N chunks cost
/// 0 + 1 + ... + N block visits. With the `yrs-cursor` feature this carries
/// the position the write already left behind (see vendor/README.md); without
/// it every insert walks, which is all upstream yrs 0.27 offers.
#[cfg(feature = "yrs-cursor")]
struct StoryWriter(Option<(yrs::TextCursor, u32)>);

#[cfg(feature = "yrs-cursor")]
impl StoryWriter {
    fn new() -> Self {
        Self(None)
    }

    /// Drop the cursor. Every op that is not an insert moves the story out
    /// from under it.
    fn invalidate(&mut self) {
        self.0 = None;
    }

    /// The cursor to write at `index`, reopening it when it sits anywhere
    /// else: a cursor only equals an index insert while it is at that index.
    fn at(
        &mut self,
        story: &TextRef,
        txn: &mut TransactionMut<'_>,
        index: u32,
    ) -> OpResult<&mut yrs::TextCursor> {
        if !self.0.as_ref().is_some_and(|(_, at)| *at == index) {
            let open = story
                .cursor(txn, index)
                .ok_or(OpError::OutOfBounds { index, len: index })?;
            self.0 = Some((open, index));
        }
        Ok(&mut self.0.as_mut().expect("cursor opened above").0)
    }

    /// The index is tracked here rather than read back from the cursor:
    /// `ItemPosition::forward` counts only `String` and `Embed` content, so it
    /// misses the map-backed embeds this crate uses for pilcrows, while
    /// `find_position` counts every non-format block. Ops already state
    /// exactly how many story units they write.
    fn insert_text(
        &mut self,
        story: &TextRef,
        txn: &mut TransactionMut<'_>,
        index: u32,
        text: &str,
        attrs: Attrs,
    ) -> OpResult<()> {
        let written = text.encode_utf16().count() as u32;
        let at = self.at(story, txn, index)?;
        story.insert_with_attributes_at(txn, at, text, attrs);
        if let Some((_, at)) = self.0.as_mut() {
            *at = index + written;
        }
        Ok(())
    }

    fn insert_embed(
        &mut self,
        story: &TextRef,
        txn: &mut TransactionMut<'_>,
        index: u32,
        attrs: Attrs,
    ) -> OpResult<MapRef> {
        let at = self.at(story, txn, index)?;
        let embed = story.insert_embed_with_attributes_at(txn, at, MapPrelim::default(), attrs);
        if let Some((_, at)) = self.0.as_mut() {
            *at = index + 1;
        }
        Ok(embed)
    }
}

#[cfg(not(feature = "yrs-cursor"))]
struct StoryWriter;

#[cfg(not(feature = "yrs-cursor"))]
impl StoryWriter {
    fn new() -> Self {
        Self
    }

    fn invalidate(&mut self) {}

    fn insert_text(
        &mut self,
        story: &TextRef,
        txn: &mut TransactionMut<'_>,
        index: u32,
        text: &str,
        attrs: Attrs,
    ) -> OpResult<()> {
        story.insert_with_attributes(txn, index, text, attrs);
        Ok(())
    }

    fn insert_embed(
        &mut self,
        story: &TextRef,
        txn: &mut TransactionMut<'_>,
        index: u32,
        attrs: Attrs,
    ) -> OpResult<MapRef> {
        Ok(story.insert_embed_with_attributes(txn, index, MapPrelim::default(), attrs))
    }
}

fn guard_index<T: ReadTxn>(story: &yrs::TextRef, txn: &T, index: u32) -> OpResult<()> {
    if index <= story.len(txn) {
        Ok(())
    } else {
        Err(OpError::OutOfBounds {
            index,
            len: story.len(txn),
        })
    }
}

fn guard_range<T: ReadTxn>(story: &yrs::TextRef, txn: &T, index: u32, len: u32) -> OpResult<()> {
    let story_len = story.len(txn);
    if index.checked_add(len).is_some_and(|end| end <= story_len) {
        Ok(())
    } else {
        Err(OpError::OutOfBounds {
            index: index.saturating_add(len),
            len: story_len,
        })
    }
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use super::*;
    use crate::{EditCtx, PILCROW_KIND};

    fn attrs(pairs: &[(&str, Any)]) -> Attrs {
        pairs
            .iter()
            .map(|(key, value)| (Arc::from(*key), value.clone()))
            .collect()
    }

    #[test]
    fn insert_format_and_delete_apply_in_one_batch() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        // story: A(0) B(1) pilcrow(2)
        let ctx = EditCtx::local(String::new(), String::new());
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::Insert {
                    index: 1,
                    text: "X".into(),
                    attrs: attrs(&[("bold", Any::Bool(true))]),
                },
                // story: A(0) X(1) B(2) pilcrow(3)
                RawOp::Format {
                    index: 0,
                    len: 1,
                    attrs: attrs(&[("italic", Any::Bool(true))]),
                },
                RawOp::Delete { index: 2, len: 1 }, // remove B
                                                    // story: A(0) X(1) pilcrow(2)
            ],
            &ctx,
        )
        .unwrap();

        let paras = doc.paragraphs("body").unwrap();
        assert_eq!(paras.len(), 1);
        assert_eq!(paras[0].text, "AX");
    }

    #[test]
    fn insert_embed_pilcrow_splits_a_paragraph() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        doc.apply_raw_ops(
            "body",
            vec![RawOp::InsertEmbed {
                index: 1,
                kind: PILCROW_KIND.to_owned(),
                payload: vec![
                    ("paraId".into(), Any::from("7:seed")),
                    ("pStyle".into(), Any::from("Normal")),
                    ("alignment".into(), Any::from("left")),
                ],
                attrs: Attrs::new(),
            }],
            &ctx,
        )
        .unwrap();

        let paras = doc.paragraphs("body").unwrap();
        assert_eq!(paras.len(), 2);
        assert_eq!(paras[0].text, "A");
        assert_eq!(paras[1].text, "B");
    }

    #[test]
    fn set_embed_attr_updates_a_pilcrow_property() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        // The pilcrow sits at index 2 (after "AB").
        doc.apply_raw_ops(
            "body",
            vec![RawOp::SetEmbedAttr {
                index: 2,
                key: "alignment".into(),
                value: Any::from("center"),
            }],
            &ctx,
        )
        .unwrap();

        let paras = doc.paragraphs("body").unwrap();
        assert_eq!(
            paras[0].properties.get("alignment"),
            Some(&Any::from("center"))
        );
    }

    #[test]
    fn set_comment_upserts_and_remove_comment_deletes() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "ABCDE", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "9".into(),
                ranges: vec![(1, 3)],
                author: "Ada".into(),
                date: "2026-07-13T12:00:00Z".into(),
                body: Any::from("hi"),
            }],
            &ctx,
        )
        .unwrap();
        let resolved = doc.resolve_comment("9").unwrap();
        assert_eq!((resolved[0].start, resolved[0].end), (1, 3));

        // Sticky anchors ride an insertion before the range.
        doc.apply_raw_ops(
            "body",
            vec![RawOp::Insert {
                index: 0,
                text: "xx".into(),
                attrs: Attrs::new(),
            }],
            &ctx,
        )
        .unwrap();
        let resolved = doc.resolve_comment("9").unwrap();
        assert_eq!((resolved[0].start, resolved[0].end), (3, 5));

        // Upsert on the same key re-anchors.
        doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "9".into(),
                ranges: vec![(0, 2), (4, 6)],
                author: String::new(),
                date: String::new(),
                body: Any::Null,
            }],
            &ctx,
        )
        .unwrap();
        let resolved = doc.resolve_comment("9").unwrap();
        assert_eq!(resolved.len(), 2);
        assert_eq!((resolved[0].start, resolved[0].end), (0, 2));
        assert_eq!((resolved[1].start, resolved[1].end), (4, 6));

        doc.apply_raw_ops("body", vec![RawOp::RemoveComment { id: "9".into() }], &ctx)
            .unwrap();
        assert!(doc.resolve_comment("9").is_err());
        // Removing an unknown comment is a typed error, not silence.
        let missing =
            doc.apply_raw_ops("body", vec![RawOp::RemoveComment { id: "9".into() }], &ctx);
        assert!(matches!(missing, Err(OpError::UnknownComment(_))));
    }

    #[test]
    fn set_comment_rejects_empty_and_out_of_bounds_ranges() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        let empty = doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "1".into(),
                ranges: vec![(1, 1)],
                author: String::new(),
                date: String::new(),
                body: Any::Null,
            }],
            &ctx,
        );
        assert!(matches!(empty, Err(OpError::InvalidRange { .. })));
        let outside = doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "1".into(),
                ranges: vec![(0, 99)],
                author: String::new(),
                date: String::new(),
                body: Any::Null,
            }],
            &ctx,
        );
        assert!(matches!(outside, Err(OpError::OutOfBounds { .. })));
    }

    #[test]
    fn out_of_bounds_index_is_a_typed_error() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        let result = doc.apply_raw_ops("body", vec![RawOp::Delete { index: 99, len: 1 }], &ctx);
        assert!(matches!(result, Err(OpError::OutOfBounds { .. })));
    }
}
