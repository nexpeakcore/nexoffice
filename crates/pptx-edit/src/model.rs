use std::collections::BTreeMap;

use ooxml_drawingml::{ShapeFill, ShapeOutline};
use pptx_parse::{GraphicFrameData, Placeholder};
use serde::{Deserialize, Serialize};
use thiserror::Error;

#[derive(Clone, Copy, Debug, Default, Eq, Hash, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum EditOrigin {
    #[default]
    Local,
    Agent,
    Remote,
    System,
}

#[derive(Clone, Debug, Default, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct EditCtx {
    pub origin: EditOrigin,
    pub author: String,
}

impl EditCtx {
    pub fn local(author: impl Into<String>) -> Self {
        Self {
            origin: EditOrigin::Local,
            author: author.into(),
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextStyle {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub font_size_pt: Option<f64>,
    pub color: Option<String>,
    pub font_family: Option<String>,
    pub underline: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextStylePatch {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub font_size_pt: Option<f64>,
    pub color: Option<String>,
    pub font_family: Option<String>,
    pub underline: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextRunSnapshot {
    pub text: String,
    pub style: TextStyle,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphSnapshot {
    pub id: String,
    pub alignment: Option<String>,
    pub level: u32,
    pub bullet_json: Option<String>,
    pub runs: Vec<TextRunSnapshot>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct StorySnapshot {
    pub id: String,
    pub length: u32,
    pub paragraphs: Vec<ParagraphSnapshot>,
}

impl StorySnapshot {
    pub fn plain_text(&self) -> String {
        self.paragraphs
            .iter()
            .map(|paragraph| {
                paragraph
                    .runs
                    .iter()
                    .map(|run| run.text.as_str())
                    .collect::<String>()
            })
            .collect::<Vec<_>>()
            .join("\n")
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeSnapshot {
    pub id: String,
    pub source_id: u32,
    pub kind: ShapeKind,
    pub name: String,
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
    pub rotation_deg: f64,
    pub flip_h: bool,
    pub flip_v: bool,
    pub geometry: String,
    pub adjust_values: BTreeMap<String, f64>,
    pub placeholder: Option<Placeholder>,
    pub fill: Option<ShapeFill>,
    pub resolved_fill_color: Option<String>,
    pub outline: Option<ShapeOutline>,
    pub resolved_outline_color: Option<String>,
    pub media_part_path: Option<String>,
    pub graphic: Option<GraphicFrameData>,
    pub text_stories: Vec<StorySnapshot>,
    pub children: Vec<ShapeSnapshot>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ShapeKind {
    Shape,
    Picture,
    GraphicFrame,
    Group,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideSnapshot {
    pub id: String,
    pub source_part_path: Option<String>,
    pub layout_part_path: Option<String>,
    pub name: Option<String>,
    pub shapes: Vec<ShapeSnapshot>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DeckSnapshot {
    pub width_emu: i64,
    pub height_emu: i64,
    pub slides: Vec<SlideSnapshot>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SlideReceipt {
    pub slide_id: String,
    pub from_index: Option<u32>,
    pub to_index: Option<u32>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeReceipt {
    pub slide_id: String,
    pub shape_id: String,
    pub index: u32,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TransformReceipt {
    pub slide_id: String,
    pub shape_id: String,
    pub before: ShapeRect,
    pub after: ShapeRect,
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct ShapeRect {
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextReceipt {
    pub story_id: String,
    pub start: u32,
    pub end: u32,
    pub text: String,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeDraft {
    pub name: String,
    pub rect: ShapeRect,
    pub text: String,
    pub style: TextStyle,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct PresetShapeDraft {
    pub name: String,
    pub geometry: String,
    pub rect: ShapeRect,
    pub fill: Option<String>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeStroke {
    pub color: Option<String>,
    pub width_pt: Option<f64>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeFillReceipt {
    pub slide_id: String,
    pub shape_id: String,
    pub before: Option<String>,
    pub after: Option<String>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeStrokeReceipt {
    pub slide_id: String,
    pub shape_id: String,
    pub before: Option<ShapeStroke>,
    pub after: Option<ShapeStroke>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ShapeAdjustReceipt {
    pub slide_id: String,
    pub shape_id: String,
    pub before: BTreeMap<String, f64>,
    pub after: BTreeMap<String, f64>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum UpdateOrigin {
    Local,
    Remote,
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct UpdateEvent {
    pub update: Vec<u8>,
    pub origin: UpdateOrigin,
}

#[derive(Debug, Error)]
pub enum EditError {
    #[error("invalid client ID {0}")]
    InvalidClientId(u64),
    #[error("could not parse PPTX: {0}")]
    Parse(String),
    #[error("invalid deck state: {0}")]
    InvalidState(String),
    #[error("invalid yrs update: {0}")]
    InvalidUpdate(String),
    #[error("invalid yrs state vector: {0}")]
    InvalidStateVector(String),
    #[error("slide {0:?} was not found")]
    SlideNotFound(String),
    #[error("shape {0:?} was not found")]
    ShapeNotFound(String),
    #[error("story {0:?} was not found")]
    StoryNotFound(String),
    #[error("index {index} is outside length {length}")]
    OutOfBounds { index: u32, length: u32 },
    #[error("text range {start}..{end} crosses a paragraph boundary")]
    ParagraphBoundary { start: u32, end: u32 },
    #[error("index {0} is not a paragraph break this story can join")]
    NotAParagraphBreak(u32),
    #[error("invalid shape geometry: {0}")]
    InvalidGeometry(String),
    #[error("invalid shape adjustment: {0}")]
    InvalidAdjustment(String),
    #[error("update observer failed: {0}")]
    Observer(String),
    #[error("JSON boundary error: {0}")]
    Json(String),
    #[error("this deck holds a change the PPTX writer cannot save yet: {0}")]
    Unprojectable(String),
    #[error("this deck cannot be saved as a file: {0}")]
    Unsavable(String),
    #[error("this save writes more than one save may write: {0}")]
    WriteLimit(String),
    #[error("writing the deck failed: {0}")]
    WriteFailed(String),
    #[error("the written deck did not read back as the deck it was planned from: {0}")]
    VerificationFailed(String),
}

/// What a failed save means for the work sitting in the deck.
///
/// The desktop offers to abandon edits on exactly one of these, so the
/// distinction is the whole point: a writer that cannot express a change is
/// not a disk that would not take the bytes, and only the first is fixed by
/// undoing something.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Hash)]
pub enum SaveFault {
    /// The writer cannot express a change this deck holds. Undoing the change
    /// the message names lets the same save succeed.
    Unprojectable,
    /// This deck can never be written as a file, whatever is undone. A replica
    /// opened from a collaborative update is the only one so far: it never
    /// carried the original package bytes to splice into.
    Unsavable,
    /// The edit is larger than one save may write. Saving less at a time works.
    Limit,
    /// Writing broke, or the writer reached a state it does not hold. The edit
    /// is intact and the same save may well succeed on the next attempt.
    WriteFailed,
    /// The bytes written did not read back as the deck they were planned from.
    /// A writer bug, never the user's change — so the edits must survive it.
    VerificationFailed,
}

impl SaveFault {
    /// A stable name for this fault, for hosts that must branch on it rather
    /// than read the message. Never change one of these: a host that does not
    /// recognise a code treats the save as [`SaveFault::WriteFailed`], and
    /// renaming an existing code silently moves users onto that fallback.
    pub fn code(self) -> &'static str {
        match self {
            Self::Unprojectable => "unprojectable",
            Self::Unsavable => "unsavable",
            Self::Limit => "limit",
            Self::WriteFailed => "write-failed",
            Self::VerificationFailed => "verification-failed",
        }
    }

    /// Whether undoing the change the message names lets the save through.
    ///
    /// This is what an offer to abandon edits hangs on, so it is false for
    /// every fault the user's own change did not cause.
    pub fn undoing_helps(self) -> bool {
        matches!(self, Self::Unprojectable)
    }
}

impl EditError {
    /// How a save that ended in this error should be read.
    ///
    /// Errors that are not about writing at all — a bad client ID, a broken
    /// update — reach a save through the snapshot it takes first, and they get
    /// the same fail-safe reading as a broken write: retryable, and never a
    /// reason to throw work away.
    pub fn save_fault(&self) -> SaveFault {
        match self {
            Self::Unprojectable(_) => SaveFault::Unprojectable,
            Self::Unsavable(_) => SaveFault::Unsavable,
            Self::WriteLimit(_) => SaveFault::Limit,
            Self::VerificationFailed(_) => SaveFault::VerificationFailed,
            _ => SaveFault::WriteFailed,
        }
    }
}

pub type EditResult<T> = Result<T, EditError>;
