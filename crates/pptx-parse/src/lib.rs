//! Bounded PresentationML parsing and part-preserving package writes.

mod drawing;
mod error;
mod model;
mod package;
mod relationships;
mod theme;
mod write;
mod xml;

pub use error::PptxError;
pub use model::*;
pub use package::{parse_pptx, parse_pptx_with_limits, write_pptx};
pub use relationships::{Relationship, TargetMode, relationship_types};
pub use write::{
    ParagraphRewrite, RunPiece, RunRef, RunStylePatch, ShapeInsertion, ShapeTransformRewrite,
    TextBodyLocation, adjust_value_to_val, font_size_to_sz, rewrite_slide_geometry,
    rewrite_slide_shape_insertions, rewrite_slide_shape_removals, rewrite_slide_text,
    serialize_shape,
};
pub use xml::{ParseLimits, is_legal_xml_character, sanitize_xml_text};
