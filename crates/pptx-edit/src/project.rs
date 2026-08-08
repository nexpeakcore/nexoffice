//! Projection of deck edits back onto the source package.
//!
//! Run text, paragraph splits, paragraph merges, soft line breaks, and moved
//! or resized shapes and pictures (their explicit `<a:xfrm>` rewritten in
//! place) are expressible. Everything the projection cannot write is refused
//! by name, so a save never quietly drops an edit.
//!
//! The projection is verified twice rather than assumed. Planning re-seeds a
//! story from the rewritten text bodies and requires the result to equal the
//! live deck exactly — bar the characters the edit added, which are read as
//! belonging to the run they were typed into, see [`adopt_run_style`] — so a
//! change the plan did not account for — a formatting patch, a paragraph break,
//! a run split — is refused. That check runs on the
//! model held in memory, so it alone would trust the byte splice to have landed
//! on the `<a:t>` the plan named. The bytes are therefore zipped and parsed
//! back, and the package that reads out of them must equal the package the plan
//! built.
//!
//! A splice that lands on the wrong run reads back as a different model, and
//! that is not a refusal: the edit was expressible and this writer mis-wrote
//! it. It carries [`crate::SaveFault::VerificationFailed`], as does every other
//! way a save can end that the user's own change did not cause, so a host is
//! never told to abandon work over a bug of ours.

use std::collections::{BTreeMap, BTreeSet};
use std::ops::Range;

use ooxml_drawingml::Theme;
use pptx_parse::{
    GraphicFrameData, ParagraphRewrite, PptxPackage, RunPiece, RunProperties, RunRef,
    RunStylePatch, ShapeNode, ShapeTransformRewrite, TextBody, TextBodyLocation, TextParagraph,
    TextRun, font_size_to_sz, rewrite_slide_shape_removals,
};
use yrs::{Map, TextRef, Transact, WriteTxn};

use crate::deck::{seed_baseline, snapshot_doc, theme_for_layout};
use crate::story::{seed_story, snapshot_story, style_from_run_properties};
use crate::{
    DeckSession, EditError, EditResult, ParagraphSnapshot, STORIES, ShapeSnapshot, SlideSnapshot,
    StorySnapshot, TextRunSnapshot, TextStyle,
};

/// Stands for a paragraph end while a group of paragraphs is diffed as one
/// string.
///
/// U+0000 can appear in neither side: XML cannot carry it, so no source run
/// holds one, and [`pptx_parse::sanitize_xml_text`] strips it from anything the
/// editor inserts.
const PARAGRAPH_MARK: char = '\u{0}';

/// Ceilings on what one save is allowed to rewrite.
///
/// [`pptx_parse::ParseLimits`] bounds what a deck may cost to read. These bound
/// what it may cost to write: a collaborative peer or a clipboard paste can put
/// an arbitrarily long string into a run, and every save after that re-escapes
/// it (up to five bytes out for each `&` in), re-zips the part and re-seeds a
/// document to verify it. The defaults sit far above what a real slide holds —
/// a slide's whole text is normally a couple of kilobytes — so honest content
/// never meets them.
#[derive(Clone, Debug)]
pub struct WriteLimits {
    /// Largest text one `<a:t>` may be rewritten to, in bytes.
    pub max_run_text_bytes: usize,
    /// Largest number of runs one save may rewrite.
    pub max_run_edits: usize,
    /// Largest total text one save may rewrite, in bytes.
    pub max_total_edit_bytes: usize,
}

impl Default for WriteLimits {
    fn default() -> Self {
        Self {
            max_run_text_bytes: 128 * 1024,
            max_run_edits: 20_000,
            max_total_edit_bytes: 8 * 1024 * 1024,
        }
    }
}

impl DeckSession {
    /// Projects the deck's text edits onto the parts the package was opened
    /// with, returning a package whose bytes are ready to re-zip.
    ///
    /// Untouched slides and every non-slide part keep their source bytes, and
    /// a touched slide keeps every byte outside the `<a:t>` elements that
    /// changed.
    ///
    /// Every way this can fail carries a [`crate::SaveFault`], because they do
    /// not mean the same thing to the work in the deck: only
    /// [`EditError::Unprojectable`] names a change the caller could undo to get
    /// the save through.
    pub fn project(&self) -> EditResult<PptxPackage> {
        self.project_with_limits(&WriteLimits::default())
    }

    /// [`DeckSession::project`] under explicit write budgets.
    pub fn project_with_limits(&self, limits: &WriteLimits) -> EditResult<PptxPackage> {
        Ok(self.verified_projection(limits)?.0)
    }

    /// Projects the deck and returns the re-zipped file.
    ///
    /// Prefer this over zipping the result of [`DeckSession::project`]: the
    /// bytes returned here are the ones the byte-level verification read back,
    /// and producing them costs one zip rather than two.
    pub fn save_bytes(&self) -> EditResult<Vec<u8>> {
        self.save_bytes_with_limits(&WriteLimits::default())
    }

    /// [`DeckSession::save_bytes`] under explicit write budgets.
    pub fn save_bytes_with_limits(&self, limits: &WriteLimits) -> EditResult<Vec<u8>> {
        Ok(self.verified_projection(limits)?.1)
    }

    /// Plans the projection, writes it, and requires the written bytes to read
    /// back as the package the plan built.
    ///
    /// A package with no rewritten part skips the read-back: nothing was
    /// spliced, so there is no address to have got wrong.
    fn verified_projection(&self, limits: &WriteLimits) -> EditResult<(PptxPackage, Vec<u8>)> {
        let (package, rewrote_parts) = self.planned_projection(limits)?;
        let bytes = pptx_parse::write_pptx(&package).map_err(|error| {
            EditError::WriteFailed(format!("the deck could not be re-zipped: {error}"))
        })?;
        if !rewrote_parts {
            return Ok((package, bytes));
        }
        let read_back = pptx_parse::parse_pptx(&bytes).map_err(|error| {
            EditError::VerificationFailed(format!("the rewritten deck did not parse: {error}"))
        })?;
        if read_back != package {
            return Err(EditError::VerificationFailed(
                "the rewritten part bytes read back as a different deck, so the text was written \
                 somewhere other than the runs the edit named"
                    .to_owned(),
            ));
        }
        Ok((package, bytes))
    }

    fn planned_projection(&self, limits: &WriteLimits) -> EditResult<(PptxPackage, bool)> {
        let mut package = (*self.package).clone();
        if package
            .part_bytes(&package.presentation.part_path)
            .is_none()
        {
            return Err(EditError::Unsavable(
                "this replica was opened from a collaborative update, which does not carry the \
                 original package bytes"
                    .to_owned(),
            ));
        }
        let current = self.snapshot()?;
        let baseline = snapshot_doc(&seed_baseline(&package)?, &package)?;
        if current.width_emu != baseline.width_emu || current.height_emu != baseline.height_emu {
            return Err(unprojectable("the slide size changed"));
        }
        if current.slides.len() != baseline.slides.len() {
            return Err(unprojectable("slides were added or removed"));
        }

        let mut plan = Plan::new(limits);
        for (index, (baseline_slide, current_slide)) in
            baseline.slides.iter().zip(&current.slides).enumerate()
        {
            require_same_slide(baseline_slide, current_slide)?;
            let source = package
                .slides
                .get(index)
                .ok_or_else(|| EditError::WriteFailed("a slide lost its source part".to_owned()))?;
            let context = SlideContext {
                index,
                part_path: source.part_path.clone(),
                theme: theme_for_layout(&package, source.layout_part_path.as_deref()),
            };
            let alignment = align_removed_shapes(&baseline_slide.shapes, &current_slide.shapes)
                .ok_or_else(|| {
                    unprojectable(format!("slide {} added or rearranged shapes", index + 1))
                })?;
            for (shape_index, source_shape) in source.shapes.iter().enumerate() {
                match alignment[shape_index] {
                    Some(current_index) => project_shape(
                        &context,
                        source_shape,
                        &baseline_slide.shapes[shape_index],
                        &current_slide.shapes[current_index],
                        &mut vec![shape_index],
                        &mut plan,
                    )?,
                    None => {
                        plan.removals
                            .entry(context.part_path.clone())
                            .or_default()
                            .push(shape_index);
                        plan.removed.push((index, shape_index));
                    }
                }
            }
        }

        let rewrote_parts =
            !plan.edits.is_empty() || !plan.transforms.is_empty() || !plan.removals.is_empty();
        for (part_path, edits) in &plan.edits {
            let bytes = package
                .part_bytes(part_path)
                .ok_or_else(|| EditError::WriteFailed(format!("part {part_path} is missing")))?;
            let rewritten = pptx_parse::rewrite_slide_text(part_path, bytes, edits)
                .map_err(|error| EditError::WriteFailed(error.to_string()))?;
            package.replace_part(part_path, rewritten);
        }
        for (part_path, transforms) in &plan.transforms {
            let bytes = package
                .part_bytes(part_path)
                .ok_or_else(|| EditError::WriteFailed(format!("part {part_path} is missing")))?;
            let rewritten = pptx_parse::rewrite_slide_geometry(part_path, bytes, transforms)
                .map_err(|error| {
                    let message = error.to_string();
                    // A shape whose placement lives in its layout has no
                    // <a:xfrm> to rewrite — that is the user's move to undo,
                    // not a writer bug.
                    if message.contains("no explicit") {
                        unprojectable(
                            "a moved shape takes its position from the slide layout, which the \
                             writer cannot rewrite yet",
                        )
                    } else {
                        EditError::WriteFailed(message)
                    }
                })?;
            package.replace_part(part_path, rewritten);
        }
        for (part_path, removals) in &plan.removals {
            let bytes = package
                .part_bytes(part_path)
                .ok_or_else(|| EditError::WriteFailed(format!("part {part_path} is missing")))?;
            let rewritten = rewrite_slide_shape_removals(part_path, bytes, removals)
                .map_err(|error| EditError::WriteFailed(error.to_string()))?;
            package.replace_part(part_path, rewritten);
        }
        for moved in &plan.moved {
            let transform = package
                .slides
                .get_mut(moved.slide_index)
                .and_then(|slide| shape_transform_mut(&mut slide.shapes, &moved.shape_path))
                .ok_or_else(|| {
                    EditError::WriteFailed("a moved shape left the shape tree".to_owned())
                })?;
            transform.x = moved.x;
            transform.y = moved.y;
            transform.width = moved.width;
            transform.height = moved.height;
        }
        for body in plan.bodies {
            let target = package
                .slides
                .get_mut(body.slide_index)
                .and_then(|slide| {
                    text_body_mut(&mut slide.shapes, &body.shape_path, &body.location)
                })
                .ok_or_else(|| {
                    EditError::WriteFailed("a rewritten text body left the shape tree".to_owned())
                })?;
            *target = body.body;
        }
        let mut removed = plan.removed.clone();
        removed.sort_by(|a, b| b.cmp(a));
        for (slide_index, shape_index) in removed {
            let slide = package.slides.get_mut(slide_index).ok_or_else(|| {
                EditError::WriteFailed("a removed shape left the slide list".to_owned())
            })?;
            if shape_index >= slide.shapes.len() {
                return Err(EditError::WriteFailed(
                    "a removed shape left the shape tree".to_owned(),
                ));
            }
            slide.shapes.remove(shape_index);
        }
        Ok((package, rewrote_parts))
    }
}

struct SlideContext<'a> {
    index: usize,
    part_path: String,
    theme: Option<&'a Theme>,
}

struct Plan<'a> {
    limits: &'a WriteLimits,
    edits: BTreeMap<String, Vec<ParagraphRewrite>>,
    bodies: Vec<ProjectedBody>,
    transforms: BTreeMap<String, Vec<ShapeTransformRewrite>>,
    moved: Vec<MovedShape>,
    removals: BTreeMap<String, Vec<usize>>,
    removed: Vec<(usize, usize)>,
    charged_edits: usize,
    charged_bytes: usize,
}

impl<'a> Plan<'a> {
    fn new(limits: &'a WriteLimits) -> Self {
        Self {
            limits,
            edits: BTreeMap::new(),
            bodies: Vec::new(),
            transforms: BTreeMap::new(),
            moved: Vec::new(),
            removals: BTreeMap::new(),
            removed: Vec::new(),
            charged_edits: 0,
            charged_bytes: 0,
        }
    }

    fn charge(&mut self, text: &str, origin: &str) -> EditResult<()> {
        if text.len() > self.limits.max_run_text_bytes {
            return Err(EditError::WriteLimit(format!(
                "{origin}: the edit writes {} bytes into one run, over the {} bytes one run may \
                 hold in a save",
                text.len(),
                self.limits.max_run_text_bytes
            )));
        }
        self.charged_edits += 1;
        if self.charged_edits > self.limits.max_run_edits {
            return Err(EditError::WriteLimit(format!(
                "the save rewrites more than the {} runs one save may rewrite",
                self.limits.max_run_edits
            )));
        }
        self.charged_bytes += text.len();
        if self.charged_bytes > self.limits.max_total_edit_bytes {
            return Err(EditError::WriteLimit(format!(
                "the save rewrites more than the {} bytes of text one save may write",
                self.limits.max_total_edit_bytes
            )));
        }
        Ok(())
    }
}

struct ProjectedBody {
    slide_index: usize,
    shape_path: Vec<usize>,
    location: TextBodyLocation,
    body: TextBody,
}

struct MovedShape {
    slide_index: usize,
    shape_path: Vec<usize>,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
}

fn project_shape(
    context: &SlideContext<'_>,
    source: &ShapeNode,
    baseline: &ShapeSnapshot,
    current: &ShapeSnapshot,
    shape_path: &mut Vec<usize>,
    plan: &mut Plan<'_>,
) -> EditResult<()> {
    // A pure move/resize is expressible for plain shapes and pictures: their
    // <a:off>/<a:ext> are rewritten in place. Everything else about the shape
    // must still match, so the placement is neutralised before the comparison
    // rather than skipped by it.
    let mut current = current.clone();
    if baseline.x != current.x
        || baseline.y != current.y
        || baseline.width != current.width
        || baseline.height != current.height
    {
        if !matches!(source, ShapeNode::Shape(_) | ShapeNode::Picture(_)) {
            return Err(unprojectable(format!(
                "shape {:?} was moved or resized, which the writer can only express for plain \
                 shapes and pictures",
                current.name
            )));
        }
        plan.transforms
            .entry(context.part_path.clone())
            .or_default()
            .push(ShapeTransformRewrite {
                shape_path: shape_path.clone(),
                x: current.x,
                y: current.y,
                width: current.width,
                height: current.height,
            });
        plan.moved.push(MovedShape {
            slide_index: context.index,
            shape_path: shape_path.clone(),
            x: current.x,
            y: current.y,
            width: current.width,
            height: current.height,
        });
        current.x = baseline.x;
        current.y = baseline.y;
        current.width = baseline.width;
        current.height = baseline.height;
    }
    let current = &current;
    require_same_shape(baseline, current)?;
    match source {
        ShapeNode::Shape(shape) => {
            if let Some(body) = &shape.text {
                project_body(
                    context,
                    body,
                    baseline,
                    current,
                    0,
                    TextBodyLocation::Shape,
                    shape_path,
                    plan,
                )?;
            }
        }
        ShapeNode::GraphicFrame(frame) => {
            if let GraphicFrameData::Table { rows } = &frame.data {
                let mut story_index = 0;
                for (row, cells) in rows.iter().enumerate() {
                    for (cell, body) in cells.iter().enumerate() {
                        project_body(
                            context,
                            body,
                            baseline,
                            current,
                            story_index,
                            TextBodyLocation::TableCell { row, cell },
                            shape_path,
                            plan,
                        )?;
                        story_index += 1;
                    }
                }
            }
        }
        ShapeNode::Group(group) => {
            if group.children.len() != current.children.len() {
                return Err(unprojectable(format!(
                    "group {:?} added or removed children",
                    current.name
                )));
            }
            for (index, child) in group.children.iter().enumerate() {
                shape_path.push(index);
                project_shape(
                    context,
                    child,
                    &baseline.children[index],
                    &current.children[index],
                    shape_path,
                    plan,
                )?;
                shape_path.pop();
            }
        }
        ShapeNode::Picture(_) => {}
    }
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn project_body(
    context: &SlideContext<'_>,
    source: &TextBody,
    baseline: &ShapeSnapshot,
    current: &ShapeSnapshot,
    story_index: usize,
    location: TextBodyLocation,
    shape_path: &[usize],
    plan: &mut Plan<'_>,
) -> EditResult<()> {
    let (Some(baseline_story), Some(current_story)) = (
        baseline.text_stories.get(story_index),
        current.text_stories.get(story_index),
    ) else {
        return Err(unprojectable(format!(
            "{} lost a text body",
            describe(context, current)
        )));
    };
    if baseline_story == current_story {
        return Ok(());
    }
    let place = describe(context, current);
    // The baseline is seeded from the package this projection splices into, so
    // the two disagreeing is the writer contradicting itself rather than
    // anything the user typed.
    if source.paragraphs.len() != baseline_story.paragraphs.len() {
        return Err(EditError::WriteFailed(format!(
            "{place} was seeded from paragraphs the source body does not have"
        )));
    }
    let groups = align_paragraphs(baseline_story, current_story)
        .map_err(|reason| unprojectable(format!("{place}: {reason}")))?;

    let mut predicted = source.clone();
    predicted.paragraphs.clear();
    let mut typed = current_story.clone();
    // Every output paragraph carries its pieces — an untouched paragraph as
    // all-Keep — so the style pass below can turn any of them into a rewrite.
    let mut outputs: Vec<Vec<RunPiece>> = Vec::new();
    let mut spans: Vec<GroupSpan> = Vec::new();
    for group in &groups {
        let sources = &source.paragraphs[group.source.clone()];
        let currents = &current_story.paragraphs[group.current.clone()];
        if sources.len() == 1
            && currents.len() == 1
            && paragraph_text(&sources[0]) == snapshot_text(&currents[0])
        {
            spans.push(GroupSpan {
                source_start: group.source.start,
                source_len: 1,
                output: outputs.len()..outputs.len() + 1,
                text_changed: false,
            });
            outputs.push(
                (0..sources[0].runs.len())
                    .map(|run| {
                        RunPiece::Keep(RunRef {
                            paragraph: group.source.start,
                            run,
                        })
                    })
                    .collect(),
            );
            predicted.paragraphs.push(sources[0].clone());
            continue;
        }
        let origin = format!("{place} paragraph {}", group.source.start + 1);
        let planned = plan_group(sources, currents, group.source.start)
            .map_err(|reason| unprojectable(format!("{origin}: {reason}")))?;
        if let Some(insertion) = &planned.insertion {
            adopt_run_style(
                &mut typed.paragraphs[group.current.clone()],
                insertion,
                &style_from_run_properties(&insertion.properties, context.theme),
            );
        }
        spans.push(GroupSpan {
            source_start: group.source.start,
            source_len: group.source.len(),
            output: outputs.len()..outputs.len() + planned.pieces.len(),
            text_changed: true,
        });
        outputs.extend(planned.pieces);
        predicted.paragraphs.extend(planned.paragraphs);
    }

    // Formatting pass: where a paragraph's characters agree but their styles
    // do not, split its pieces along the typed run boundaries and patch each
    // one's <a:rPr>. What this pass cannot express it leaves alone for the
    // comparison below to refuse by name.
    for (index, typed_paragraph) in typed.paragraphs.iter().enumerate() {
        let Some(predicted_paragraph) = predicted.paragraphs.get(index) else {
            break;
        };
        if paragraph_styles_match(predicted_paragraph, typed_paragraph, context.theme) {
            continue;
        }
        let _ = predicted_paragraph;
        if let Some((pieces, runs)) =
            reconcile_paragraph(&outputs[index], source, typed_paragraph, context.theme)
        {
            outputs[index] = pieces;
            predicted.paragraphs[index].runs = runs;
            if let Some(span) = spans.iter_mut().find(|span| span.output.contains(&index)) {
                span.text_changed = true;
            }
        }
    }

    let mut rewrites = Vec::new();
    for span in &spans {
        if !span.text_changed {
            continue;
        }
        let origin = format!("{place} paragraph {}", span.source_start + 1);
        let pieces: Vec<Vec<RunPiece>> = outputs[span.output.clone()].to_vec();
        for piece in pieces.iter().flatten() {
            if let RunPiece::Text(_, text) | RunPiece::Styled(_, text, _) = piece {
                plan.charge(text, &origin)?;
            }
        }
        rewrites.push(ParagraphRewrite {
            shape_path: shape_path.to_vec(),
            location: location.clone(),
            first_paragraph: span.source_start,
            source_paragraphs: span.source_len,
            paragraphs: pieces,
        });
    }

    if content(&story_snapshot(
        &predicted,
        &current_story.id,
        context.theme,
    )?) != content(&typed)
    {
        return Err(unprojectable(format!(
            "{place} changed in a way this writer cannot express, such as formatting, a bullet, or \
             a run split",
        )));
    }
    if rewrites.is_empty() {
        return Ok(());
    }
    plan.edits
        .entry(context.part_path.clone())
        .or_default()
        .extend(rewrites);
    plan.bodies.push(ProjectedBody {
        slide_index: context.index,
        shape_path: shape_path.to_vec(),
        location,
        body: predicted,
    });
    Ok(())
}

/// A run of source paragraphs and the current paragraphs that replaced them.
struct Group {
    source: Range<usize>,
    current: Range<usize>,
}

/// For each baseline shape, its slot in the current slide — `None` when the
/// deck deleted it. `None` for the whole slide when the current shapes are not
/// an in-order subset of the baseline's (something was added or rearranged).
fn align_removed_shapes(
    baseline: &[ShapeSnapshot],
    current: &[ShapeSnapshot],
) -> Option<Vec<Option<usize>>> {
    let mut result = Vec::with_capacity(baseline.len());
    let mut cursor = 0;
    for shape in baseline {
        if current
            .get(cursor)
            .is_some_and(|candidate| candidate.id == shape.id)
        {
            result.push(Some(cursor));
            cursor += 1;
        } else {
            result.push(None);
        }
    }
    (cursor == current.len()).then_some(result)
}

/// One planned group's place in the output, kept so the formatting pass can
/// promote an untouched group to a rewrite when it patches a style.
struct GroupSpan {
    source_start: usize,
    source_len: usize,
    output: Range<usize>,
    text_changed: bool,
}

/// `#RRGGBB` uppercase, so two spellings of one colour compare equal.
fn normalize_color(color: &str) -> String {
    let bare = color.strip_prefix('#').unwrap_or(color);
    if bare.len() == 6 && bare.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        format!("#{}", bare.to_ascii_uppercase())
    } else {
        color.to_owned()
    }
}

fn normalized_style(style: &TextStyle) -> TextStyle {
    TextStyle {
        color: style.color.as_deref().map(normalize_color),
        ..style.clone()
    }
}

/// Resolved, merged runs of both sides compare equal — colours normalised.
fn paragraph_styles_match(
    predicted: &TextParagraph,
    typed: &ParagraphSnapshot,
    theme: Option<&Theme>,
) -> bool {
    let merge = |runs: Vec<(String, TextStyle)>| -> Vec<(String, TextStyle)> {
        let mut merged: Vec<(String, TextStyle)> = Vec::new();
        for (text, style) in runs {
            if text.is_empty() {
                continue;
            }
            match merged.last_mut() {
                Some(last) if last.1 == style => last.0.push_str(&text),
                _ => merged.push((text, style)),
            }
        }
        merged
    };
    let left = merge(
        predicted
            .runs
            .iter()
            .map(|run| {
                (
                    run.text.clone(),
                    normalized_style(&style_from_run_properties(&run.properties, theme)),
                )
            })
            .collect(),
    );
    let right = merge(
        typed
            .runs
            .iter()
            .map(|run| (run.text.clone(), normalized_style(&run.style)))
            .collect(),
    );
    left == right
}

/// The `RunStylePatch` that takes `source` to `target`, or `None` for a
/// difference the writer cannot spell (a size that does not round-trip).
fn style_patch(source: &TextStyle, target: &TextStyle) -> Option<RunStylePatch> {
    let mut patch = RunStylePatch::default();
    if source.bold != target.bold {
        patch.bold = Some(target.bold);
    }
    if source.italic != target.italic {
        patch.italic = Some(target.italic);
    }
    if source.underline != target.underline {
        patch.underline = Some(target.underline.clone());
    }
    if source.font_size_pt != target.font_size_pt {
        if let Some(points) = target.font_size_pt
            && font_size_to_sz(points).is_none()
        {
            return None;
        }
        patch.font_size_pt = Some(target.font_size_pt);
    }
    let source_color = source.color.as_deref().map(normalize_color);
    let target_color = target.color.as_deref().map(normalize_color);
    if source_color != target_color {
        patch.color_rgb = Some(
            target_color
                .as_deref()
                .map(|color| color.trim_start_matches('#').to_owned()),
        );
    }
    if source.font_family != target.font_family {
        patch.font_family = Some(target.font_family.clone());
    }
    Some(patch)
}

fn apply_patch_to_properties(properties: &RunProperties, patch: &RunStylePatch) -> RunProperties {
    let mut out = properties.clone();
    if let Some(bold) = &patch.bold {
        out.bold = *bold;
    }
    if let Some(italic) = &patch.italic {
        out.italic = *italic;
    }
    if let Some(underline) = &patch.underline {
        out.underline = underline.clone();
    }
    if let Some(size) = &patch.font_size_pt {
        out.font_size_pt =
            size.and_then(|points| font_size_to_sz(points).map(|sz| sz as f64 / 100.0));
    }
    if let Some(color) = &patch.color_rgb {
        out.color = color.as_ref().map(|hex| ooxml_drawingml::ColorValue {
            rgb: Some(hex.clone()),
            ..ooxml_drawingml::ColorValue::default()
        });
    }
    if let Some(family) = &patch.font_family {
        out.font_family = family.clone();
    }
    out
}

/// Splits one output paragraph's pieces along the typed run boundaries and
/// patches each segment's style. `None` leaves the paragraph as planned, so
/// the final comparison refuses instead of this pass guessing.
fn reconcile_paragraph(
    pieces: &[RunPiece],
    source: &TextBody,
    typed: &ParagraphSnapshot,
    theme: Option<&Theme>,
) -> Option<(Vec<RunPiece>, Vec<TextRun>)> {
    // Segment list from the planned pieces: which source run each character
    // is written from, and what it says.
    let mut segments: Vec<(RunRef, String)> = Vec::new();
    for piece in pieces {
        match piece {
            RunPiece::Keep(target) => {
                let run = source
                    .paragraphs
                    .get(target.paragraph)?
                    .runs
                    .get(target.run)?;
                if run.line_break || run.field_id.is_some() || run.text.is_empty() {
                    return None;
                }
                segments.push((*target, run.text.clone()));
            }
            RunPiece::Text(target, text) => segments.push((*target, text.clone())),
            RunPiece::Break(_) | RunPiece::Styled(_, _, _) => return None,
        }
    }
    let source_run = |target: &RunRef| -> Option<&TextRun> {
        source
            .paragraphs
            .get(target.paragraph)?
            .runs
            .get(target.run)
    };
    for (target, _) in &segments {
        let run = source_run(target)?;
        if run.line_break || run.field_id.is_some() {
            return None;
        }
    }
    let piece_text: String = segments.iter().map(|(_, text)| text.as_str()).collect();
    let typed_text: String = typed.runs.iter().map(|run| run.text.as_str()).collect();
    if piece_text != typed_text {
        return None;
    }

    // Character-aligned walk over both segmentations.
    let mut out_pieces: Vec<RunPiece> = Vec::new();
    let mut out_runs: Vec<TextRun> = Vec::new();
    let mut piece_iter = segments.iter();
    let mut piece_current = piece_iter.next()?.clone();
    let mut piece_chars: Vec<char> = piece_current.1.chars().collect();
    let mut piece_offset = 0usize;
    let mut typed_iter = typed.runs.iter();
    let mut typed_current = typed_iter.next()?;
    let mut typed_remaining = typed_current.text.chars().count();

    loop {
        // Skip empty carriers on either side.
        while piece_offset >= piece_chars.len() {
            match piece_iter.next() {
                Some(next) => {
                    piece_current = next.clone();
                    piece_chars = piece_current.1.chars().collect();
                    piece_offset = 0;
                }
                None => break,
            }
        }
        while typed_remaining == 0 {
            match typed_iter.next() {
                Some(next) => {
                    typed_current = next;
                    typed_remaining = typed_current.text.chars().count();
                }
                None => break,
            }
        }
        if piece_offset >= piece_chars.len() {
            break;
        }
        if typed_remaining == 0 {
            return None;
        }
        let take = (piece_chars.len() - piece_offset).min(typed_remaining);
        let text: String = piece_chars[piece_offset..piece_offset + take]
            .iter()
            .collect();
        piece_offset += take;
        typed_remaining -= take;

        let target = piece_current.0;
        let run = source_run(&target)?;
        let source_style = style_from_run_properties(&run.properties, theme);
        let patch = style_patch(&source_style, &typed_current.style)?;

        let piece = if patch.is_empty() {
            if text == run.text {
                RunPiece::Keep(target)
            } else {
                RunPiece::Text(target, text.clone())
            }
        } else {
            RunPiece::Styled(target, text.clone(), patch.clone())
        };
        let properties = if patch.is_empty() {
            run.properties.clone()
        } else {
            apply_patch_to_properties(&run.properties, &patch)
        };

        // Merge with the previous piece when it continues the same run with
        // the same treatment, so an unsplit run stays one run.
        let merged = match (out_pieces.last_mut(), out_runs.last_mut(), &piece) {
            (
                Some(RunPiece::Styled(last_target, last_text, last_patch)),
                Some(last_run),
                RunPiece::Styled(next_target, _, next_patch),
            ) if *last_target == *next_target && *last_patch == *next_patch => {
                last_text.push_str(&text);
                last_run.text.push_str(&text);
                true
            }
            (
                Some(RunPiece::Text(last_target, last_text)),
                Some(last_run),
                RunPiece::Text(next_target, _),
            ) if *last_target == *next_target => {
                last_text.push_str(&text);
                last_run.text.push_str(&text);
                true
            }
            _ => false,
        };
        if !merged {
            out_pieces.push(piece);
            out_runs.push(TextRun {
                text,
                properties,
                field_id: None,
                field_type: None,
                line_break: false,
            });
        }
    }

    Some((out_pieces, out_runs))
}

/// Matches current paragraphs to the source paragraphs they came from.
///
/// A paragraph mark keeps its identity when the story around it is edited, and
/// [`DeckSession::insert_paragraph_break`] adds a mark ahead of the one it
/// splits, so a current paragraph carrying a source id is the end of that source
/// paragraph: the current paragraphs before it are what a split made, and the
/// source paragraphs before it are what a merge swallowed.
fn align_paragraphs(
    baseline: &StorySnapshot,
    current: &StorySnapshot,
) -> Result<Vec<Group>, String> {
    let mut index_of = BTreeMap::new();
    for (index, paragraph) in baseline.paragraphs.iter().enumerate() {
        if index_of.insert(paragraph.id.as_str(), index).is_some() {
            return Err("two paragraphs of the body share a paragraph id".to_owned());
        }
    }
    let mut groups = Vec::new();
    let mut source = 0;
    let mut start = 0;
    for (index, paragraph) in current.paragraphs.iter().enumerate() {
        let Some(&ends) = index_of.get(paragraph.id.as_str()) else {
            continue;
        };
        if ends < source {
            return Err("the paragraphs were reordered".to_owned());
        }
        groups.push(Group {
            source: source..ends + 1,
            current: start..index + 1,
        });
        source = ends + 1;
        start = index + 1;
    }
    if source != baseline.paragraphs.len() || start != current.paragraphs.len() {
        return Err("a new paragraph has no source paragraph to end it".to_owned());
    }
    Ok(groups)
}

/// What one group of source paragraphs becomes, as both a model and a plan the
/// part writer can splice.
struct GroupPlan {
    paragraphs: Vec<TextParagraph>,
    pieces: Vec<Vec<RunPiece>>,
    insertion: Option<Insertion>,
}

/// The characters a group's change adds, and the source run whose `<a:rPr>`
/// will carry them.
///
/// Added text is written inside that run's `<a:r>`, so it reads back with the
/// run's own style. See [`adopt_run_style`] for why the deck is compared as if
/// it already had.
///
/// A diff of characters alone cannot always say which copy of a repeated
/// character was typed: adding an `X` to `DOCX` before its last letter and
/// after it produce the same string. `window` is therefore every position the
/// added characters could occupy — the diff's own placement widened left for as
/// long as sliding the addition one character left would read the same — and
/// `length` is how many of them were added.
struct Insertion {
    /// Character offsets into the group's current text, paragraph marks
    /// included, matching what [`current_characters`] lays out.
    window: Range<usize>,
    length: usize,
    properties: RunProperties,
}

/// One position in the source text of a group.
#[derive(Clone, Copy)]
struct Slot {
    character: char,
    /// The paragraph, run and character offset the slot sits at, or `None` for
    /// a paragraph mark, which belongs to no run.
    place: Option<(usize, usize, usize)>,
    /// For a paragraph mark, the source paragraph it ends.
    mark: Option<usize>,
}

/// Where in the source runs a change sits.
#[derive(Clone, Copy)]
struct Anchor {
    paragraph: usize,
    run: usize,
    start: usize,
    end: usize,
}

/// Turns the difference between a group of source paragraphs and what the deck
/// now shows into paragraphs to predict and pieces to write.
///
/// The group's paragraphs are diffed as one string with [`PARAGRAPH_MARK`]
/// standing for each paragraph end, so a split reads as an inserted mark, a
/// merge as a deleted one, and a soft line break as an inserted `\n`. The common
/// prefix and suffix bound the change; the text it removes must lie inside one
/// run, and the text it adds is written into that run, divided into `<a:r>`,
/// `<a:br>` and paragraph pieces. Anything wider — a change straddling two runs,
/// a field's text, or two edits far apart in one body — has no faithful rewrite
/// and is refused.
fn plan_group(
    source: &[TextParagraph],
    current: &[ParagraphSnapshot],
    first: usize,
) -> Result<GroupPlan, String> {
    let slots = source_slots(source);
    let before: Vec<char> = slots.iter().map(|slot| slot.character).collect();
    let after: Vec<char> = current_characters(current);
    let mut prefix = 0;
    while prefix < before.len() && prefix < after.len() && before[prefix] == after[prefix] {
        prefix += 1;
    }
    let mut suffix = 0;
    while prefix + suffix < before.len()
        && prefix + suffix < after.len()
        && before[before.len() - 1 - suffix] == after[after.len() - 1 - suffix]
    {
        suffix += 1;
    }
    let changed_end = before.len() - suffix;
    let replacement: String = after[prefix..after.len() - suffix].iter().collect();
    let removed: BTreeSet<usize> = slots[prefix..changed_end]
        .iter()
        .filter_map(|slot| slot.mark)
        .collect();
    let anchor = locate_anchor(&slots, source, prefix, changed_end, &replacement)?;
    let mut plan = build_group(source, first, &removed, anchor, &replacement)?;
    plan.insertion = anchor
        .filter(|_| !replacement.is_empty())
        .map(|anchor| Insertion {
            window: insertion_window(
                &after,
                prefix,
                changed_end == prefix,
                replacement.chars().count(),
            ),
            length: replacement.chars().count(),
            properties: source[anchor.paragraph].runs[anchor.run].properties.clone(),
        });
    Ok(plan)
}

fn source_slots(source: &[TextParagraph]) -> Vec<Slot> {
    let mut slots = Vec::new();
    for (paragraph_index, paragraph) in source.iter().enumerate() {
        if paragraph_index > 0 {
            slots.push(Slot {
                character: PARAGRAPH_MARK,
                place: None,
                mark: Some(paragraph_index - 1),
            });
        }
        for (run_index, run) in paragraph.runs.iter().enumerate() {
            for (offset, character) in run.text.chars().enumerate() {
                slots.push(Slot {
                    character,
                    place: Some((paragraph_index, run_index, offset)),
                    mark: None,
                });
            }
        }
    }
    slots
}

/// Every position the added characters could sit at, as offsets into the
/// current text.
///
/// The diff places an addition as far right as it can, so the window only ever
/// widens left, and only for an addition that removed nothing: where characters
/// were also taken out, the text around the change pins it down.
fn insertion_window(
    after: &[char],
    prefix: usize,
    added_only: bool,
    length: usize,
) -> Range<usize> {
    let mut start = prefix;
    while added_only && start > 0 && after[start - 1] == after[start - 1 + length] {
        start -= 1;
    }
    start..prefix + length
}

fn current_characters(current: &[ParagraphSnapshot]) -> Vec<char> {
    let mut characters = Vec::new();
    for (index, paragraph) in current.iter().enumerate() {
        if index > 0 {
            characters.push(PARAGRAPH_MARK);
        }
        characters.extend(snapshot_text(paragraph).chars());
    }
    characters
}

/// Finds the run the change is confined to, or `None` when the group holds no
/// run and the change is only paragraph marks.
fn locate_anchor(
    slots: &[Slot],
    source: &[TextParagraph],
    prefix: usize,
    changed_end: usize,
    replacement: &str,
) -> Result<Option<Anchor>, String> {
    if let Some(anchor) = deleted_anchor(&slots[prefix..changed_end])? {
        if let Some(holder) = holder_name(&source[anchor.paragraph].runs[anchor.run], anchor.run) {
            return Err(format!("the change lands on {holder}"));
        }
        return Ok(Some(anchor));
    }
    inserted_anchor(slots, source, prefix, changed_end, replacement)
}

/// The run the change takes its characters out of, if it takes any.
fn deleted_anchor(deleted: &[Slot]) -> Result<Option<Anchor>, String> {
    let mut anchor: Option<Anchor> = None;
    for slot in deleted {
        let Some((paragraph, run, offset)) = slot.place else {
            continue;
        };
        match &mut anchor {
            None => {
                anchor = Some(Anchor {
                    paragraph,
                    run,
                    start: offset,
                    end: offset + 1,
                })
            }
            Some(open) if open.paragraph == paragraph && open.run == run && open.end == offset => {
                open.end = offset + 1;
            }
            Some(_) => return Err("the change spans more than one run".to_owned()),
        }
    }
    Ok(anchor)
}

/// The run a change that only adds characters is written into.
///
/// A caret sits between two runs, and either can hold what is typed. The run
/// before it is preferred, which is where a caret's own style comes from, unless
/// the change starts by ending the paragraph, in which case what follows the new
/// paragraph mark belongs to the run after it. A run that cannot hold text — an
/// `<a:br>` or an `<a:fld>` — is passed over when there is text to place, and
/// still anchors a change that only ends a paragraph. `None` means the group has
/// no run at all, which only an empty paragraph can be.
fn inserted_anchor(
    slots: &[Slot],
    source: &[TextParagraph],
    prefix: usize,
    changed_end: usize,
    replacement: &str,
) -> Result<Option<Anchor>, String> {
    let preceding = slots[..prefix]
        .iter()
        .rev()
        .take_while(|slot| slot.mark.is_none())
        .find_map(|slot| slot.place)
        .map(|(paragraph, run, offset)| (paragraph, run, offset + 1));
    let following = slots[changed_end..]
        .iter()
        .take_while(|slot| slot.mark.is_none())
        .find_map(|slot| slot.place);
    let order = if replacement.starts_with(PARAGRAPH_MARK) {
        [following, preceding]
    } else {
        [preceding, following]
    };
    let holds_text = replacement
        .chars()
        .any(|character| character != PARAGRAPH_MARK);
    for (paragraph, run, offset) in order.into_iter().flatten() {
        if holds_text && holder_name(&source[paragraph].runs[run], run).is_some() {
            continue;
        }
        return Ok(Some(Anchor {
            paragraph,
            run,
            start: offset,
            end: offset,
        }));
    }
    let Some((paragraph, run, offset)) = order.into_iter().flatten().next() else {
        if holds_text {
            return Err("the paragraph has no run to hold text".to_owned());
        }
        if source.iter().any(|paragraph| !paragraph.runs.is_empty()) {
            return Err("the change has no run to sit beside".to_owned());
        }
        return Ok(None);
    };
    let holder = holder_name(&source[paragraph].runs[run], run)
        .unwrap_or_else(|| format!("run {}", run + 1));
    if offset > 0 && offset < source[paragraph].runs[run].text.chars().count() {
        return Err(format!("the change lands on {holder}"));
    }
    Err(format!(
        "the change lands past {holder}, which has no run after it to hold the text"
    ))
}

/// What a run is called in a refusal when it cannot hold run text.
fn holder_name(run: &TextRun, index: usize) -> Option<String> {
    if run.line_break {
        return Some(format!("line break {}", index + 1));
    }
    if run.field_id.is_some() || run.field_type.is_some() {
        return Some(format!("field {}", index + 1));
    }
    None
}

/// Walks the group's source runs once, writing every run through untouched
/// except the one the change sits in, which is divided around the paragraph
/// marks and line breaks the change adds.
fn build_group(
    source: &[TextParagraph],
    first: usize,
    removed: &BTreeSet<usize>,
    anchor: Option<Anchor>,
    replacement: &str,
) -> Result<GroupPlan, String> {
    let mut plan = GroupPlan {
        paragraphs: Vec::new(),
        pieces: Vec::new(),
        insertion: None,
    };
    let mut pieces: Vec<RunPiece> = Vec::new();
    let mut runs: Vec<TextRun> = Vec::new();
    let mut owner = 0;

    for (local, paragraph) in source.iter().enumerate() {
        if local > 0 && !removed.contains(&(local - 1)) {
            plan.paragraphs.push(TextParagraph {
                properties: source[owner].properties.clone(),
                runs: std::mem::take(&mut runs),
                end_properties: source[local - 1].end_properties.clone(),
            });
            plan.pieces.push(std::mem::take(&mut pieces));
            owner = local;
        }
        for (index, run) in paragraph.runs.iter().enumerate() {
            let reference = RunRef {
                paragraph: first + local,
                run: index,
            };
            let Some(anchor) =
                anchor.filter(|anchor| anchor.paragraph == local && anchor.run == index)
            else {
                pieces.push(RunPiece::Keep(reference));
                runs.push(run.clone());
                continue;
            };
            let divided = divide(run, anchor, replacement);
            let mut wrote = false;
            for (paragraph_index, lines) in divided.iter().enumerate() {
                if paragraph_index > 0 {
                    plan.paragraphs.push(TextParagraph {
                        properties: source[owner].properties.clone(),
                        runs: std::mem::take(&mut runs),
                        end_properties: None,
                    });
                    plan.pieces.push(std::mem::take(&mut pieces));
                    owner = local;
                }
                for (line_index, line) in lines.iter().enumerate() {
                    if line_index > 0 {
                        if let Some(holder) = holder_name(run, index) {
                            return Err(format!("the change lands on {holder}"));
                        }
                        pieces.push(RunPiece::Break(reference));
                        runs.push(TextRun {
                            text: "\n".to_owned(),
                            properties: run.properties.clone(),
                            field_id: None,
                            field_type: None,
                            line_break: true,
                        });
                    }
                    if line.is_empty() {
                        continue;
                    }
                    wrote = true;
                    if *line == run.text {
                        pieces.push(RunPiece::Keep(reference));
                        runs.push(run.clone());
                    } else {
                        if let Some(holder) = holder_name(run, index) {
                            return Err(format!("the change lands on {holder}"));
                        }
                        pieces.push(RunPiece::Text(reference, line.clone()));
                        runs.push(TextRun {
                            text: line.clone(),
                            ..run.clone()
                        });
                    }
                }
            }
            if !wrote {
                if let Some(holder) = holder_name(run, index) {
                    return Err(format!("the change lands on {holder}"));
                }
                pieces.push(RunPiece::Text(reference, String::new()));
                runs.push(TextRun {
                    text: String::new(),
                    ..run.clone()
                });
            }
        }
        if anchor.is_none() && local + 1 == source.len() {
            for _ in replacement.chars().filter(|c| *c == PARAGRAPH_MARK) {
                plan.paragraphs.push(TextParagraph {
                    properties: source[owner].properties.clone(),
                    runs: std::mem::take(&mut runs),
                    end_properties: None,
                });
                plan.pieces.push(std::mem::take(&mut pieces));
                owner = local;
            }
        }
    }
    plan.paragraphs.push(TextParagraph {
        properties: source[owner].properties.clone(),
        runs,
        end_properties: source[source.len() - 1].end_properties.clone(),
    });
    plan.pieces.push(pieces);
    Ok(plan)
}

/// The text of one run after the change, split into paragraphs and, inside
/// each, into the lines a `<a:br>` separates.
///
/// Only the characters the change adds are split on: a newline the source run
/// already holds inside its own `<a:t>` stays where it is.
fn divide(run: &TextRun, anchor: Anchor, replacement: &str) -> Vec<Vec<String>> {
    let text: Vec<char> = run.text.chars().collect();
    let mut divided = vec![vec![String::new()]];
    push_line(&mut divided, &text[..anchor.start.min(text.len())]);
    for character in replacement.chars() {
        match character {
            PARAGRAPH_MARK => divided.push(vec![String::new()]),
            '\n' => {
                if let Some(lines) = divided.last_mut() {
                    lines.push(String::new());
                }
            }
            character => {
                if let Some(line) = divided.last_mut().and_then(|lines| lines.last_mut()) {
                    line.push(character);
                }
            }
        }
    }
    push_line(&mut divided, &text[anchor.end.min(text.len())..]);
    divided
}

fn push_line(divided: &mut [Vec<String>], characters: &[char]) {
    if let Some(line) = divided.last_mut().and_then(|lines| lines.last_mut()) {
        line.extend(characters);
    }
}

/// A story stripped of what a part cannot carry, so the projection is compared
/// on content alone.
///
/// A paragraph id lives only in the editor — no PresentationML element holds one
/// — and a split gives the halves ids the source paragraph never had, so ids are
/// replaced by position. Runs that share a style are one run to a reader and to
/// the layout, so adjacent equal-styled runs are joined: that is what lets a
/// merge whose two paragraphs end and start in the same style compare equal to
/// the two `<a:r>` elements it writes.
fn content(story: &StorySnapshot) -> StorySnapshot {
    let mut story = story.clone();
    for (index, paragraph) in story.paragraphs.iter_mut().enumerate() {
        paragraph.id = index.to_string();
        let mut runs: Vec<TextRunSnapshot> = Vec::new();
        for mut run in std::mem::take(&mut paragraph.runs) {
            if run.text.is_empty() {
                continue;
            }
            run.style = normalized_style(&run.style);
            match runs.last_mut() {
                Some(last) if last.style == run.style => last.text.push_str(&run.text),
                _ => runs.push(run),
            }
        }
        paragraph.runs = runs;
    }
    story
}

/// Reads the characters an edit added as the anchor run's own, where the style
/// they were typed with says nothing that run's `<a:rPr>` denies.
///
/// A caret carries a whole style, not the half of one an `<a:rPr>` spells out:
/// an editor materialises what the text under the caret resolves to and inserts
/// with that, so text typed into a run that leaves `i` and `u` to be inherited
/// arrives carrying `italic: Some(false)` and `underline: Some("none")`. Nothing
/// about the run changed — the same characters, in the same run, resolve to the
/// same appearance — but the deck now holds two runs where the file holds one,
/// and comparing the two literally would refuse every keystroke of a normal
/// deck.
///
/// A field the run does spell out is a different matter. Text typed as
/// `bold: Some(false)` into a run whose `<a:rPr>` carries `b="1"` is a run
/// split, and text typed with no value at all where the run has one would lose
/// that value; both leave the characters as they are, so the comparison sees
/// them and refuses. Formatting applied to text the edit did not add falls
/// outside the insertion and is likewise untouched, which is what keeps a bold
/// toggle over half a run a refusal rather than a silent no-op.
///
/// The cost is that where the run is silent, the value the caret carried is
/// dropped and the character inherits instead — it takes the appearance of the
/// run it was typed into rather than the one the editor guessed for it.
fn adopt_run_style(paragraphs: &mut [ParagraphSnapshot], insertion: &Insertion, style: &TextStyle) {
    let Some(range) = typed_range(&run_spans(paragraphs), insertion, style) else {
        return;
    };
    let mut offset = 0;
    for (index, paragraph) in paragraphs.iter_mut().enumerate() {
        if index > 0 {
            offset += 1;
        }
        let mut runs = Vec::new();
        for run in std::mem::take(&mut paragraph.runs) {
            let start = offset;
            offset += run.text.chars().count();
            let from = range.start.max(start);
            let to = range.end.min(offset);
            if from >= to {
                runs.push(run);
                continue;
            }
            let mut characters = run.text.chars();
            let head: String = characters.by_ref().take(from - start).collect();
            let typed: String = characters.by_ref().take(to - from).collect();
            let tail: String = characters.collect();
            for (text, style) in [(head, &run.style), (typed, style), (tail, &run.style)] {
                if !text.is_empty() {
                    runs.push(TextRunSnapshot {
                        text,
                        style: style.clone(),
                    });
                }
            }
        }
        paragraph.runs = runs;
    }
}

/// Which characters of [`Insertion::window`] are the ones that were typed.
///
/// They are the characters in the window that are not already the anchor run's
/// own: one unbroken stretch of them, no longer than the edit added, and every
/// one of them a style the anchor run's `<a:rPr>` does not contradict. Anything
/// else — a stretch too long, two stretches with untouched text between them, a
/// style the run denies — is a formatting change this reading cannot account
/// for, and leaving it alone is what makes the comparison refuse it.
fn typed_range(
    spans: &[(Range<usize>, TextStyle)],
    insertion: &Insertion,
    style: &TextStyle,
) -> Option<Range<usize>> {
    let mut typed: Option<Range<usize>> = None;
    for (span, run) in spans {
        let from = span.start.max(insertion.window.start);
        let to = span.end.min(insertion.window.end);
        if from >= to || run == style {
            continue;
        }
        if !specialises(run, style) {
            return None;
        }
        match &mut typed {
            None => typed = Some(from..to),
            Some(open) if open.end == from => open.end = to,
            Some(_) => return None,
        }
    }
    typed.filter(|range| range.len() <= insertion.length)
}

/// Where each run of the current paragraphs sits, as character offsets counting
/// one for each paragraph mark, the way [`current_characters`] lays them out.
fn run_spans(paragraphs: &[ParagraphSnapshot]) -> Vec<(Range<usize>, TextStyle)> {
    let mut spans = Vec::new();
    let mut offset = 0;
    for (index, paragraph) in paragraphs.iter().enumerate() {
        if index > 0 {
            offset += 1;
        }
        for run in &paragraph.runs {
            let start = offset;
            offset += run.text.chars().count();
            spans.push((start..offset, run.style.clone()));
        }
    }
    spans
}

/// Whether `run` carries every value `anchor` carries, and differs only by
/// naming values `anchor` leaves to be inherited.
///
/// The six fields are the whole of [`TextStyle`], and each is compared: bold,
/// italic and underline because a run split turns on them, the size, the colour
/// and the font family because they are as much a formatting change as the
/// other three. The colour is compared as the hex the deck resolved it to on
/// the way in, so a run wearing a theme colour and one wearing that colour
/// literally are the same run to this test.
fn specialises(run: &TextStyle, anchor: &TextStyle) -> bool {
    agrees(run.bold, anchor.bold)
        && agrees(run.italic, anchor.italic)
        && agrees(run.font_size_pt, anchor.font_size_pt)
        && agrees(run.color.as_deref(), anchor.color.as_deref())
        && agrees(run.font_family.as_deref(), anchor.font_family.as_deref())
        && agrees(run.underline.as_deref(), anchor.underline.as_deref())
}

fn agrees<T: PartialEq>(run: Option<T>, anchor: Option<T>) -> bool {
    anchor.is_none_or(|anchor| run == Some(anchor))
}

fn story_snapshot(
    body: &TextBody,
    story_id: &str,
    theme: Option<&Theme>,
) -> EditResult<StorySnapshot> {
    let doc = crate::doc_with_client_id(crate::BOOTSTRAP_CLIENT_ID);
    let mut txn = doc.transact_mut();
    let stories = txn.get_or_insert_map(STORIES);
    seed_story(&stories, &mut txn, story_id, body, theme)?;
    let story = stories
        .get(&txn, story_id)
        .and_then(|value| value.cast::<TextRef>().ok())
        .ok_or_else(|| EditError::InvalidState("projected story is missing".to_owned()))?;
    snapshot_story(&story, &txn, story_id)
}

/// Compares every slide field except the shapes, so a field added to the
/// snapshot later is refused by default rather than silently dropped.
fn require_same_slide(baseline: &SlideSnapshot, current: &SlideSnapshot) -> EditResult<()> {
    let mut left = baseline.clone();
    let mut right = current.clone();
    left.shapes.clear();
    right.shapes.clear();
    if left != right {
        return Err(unprojectable(format!(
            "slide {:?} was replaced, renamed or reordered",
            current.id
        )));
    }
    Ok(())
}

/// Compares every shape field except its stories and children, which the
/// caller walks. Comparing whole values rather than named fields keeps an
/// unhandled future field a refusal instead of a silent loss.
fn require_same_shape(baseline: &ShapeSnapshot, current: &ShapeSnapshot) -> EditResult<()> {
    let mut left = baseline.clone();
    let mut right = current.clone();
    let baseline_stories = std::mem::take(&mut left.text_stories).len();
    let current_stories = std::mem::take(&mut right.text_stories).len();
    left.children.clear();
    right.children.clear();
    if left != right {
        return Err(unprojectable(format!(
            "shape {:?} was moved, resized, restyled or replaced",
            current.name
        )));
    }
    if baseline_stories != current_stories {
        return Err(unprojectable(format!(
            "shape {:?} added or removed a text body",
            current.name
        )));
    }
    Ok(())
}

fn shape_transform_mut<'a>(
    shapes: &'a mut [ShapeNode],
    shape_path: &[usize],
) -> Option<&'a mut pptx_parse::ShapeTransform> {
    let (index, rest) = shape_path.split_first()?;
    let shape = shapes.get_mut(*index)?;
    if let Some(_next) = rest.first() {
        match shape {
            ShapeNode::Group(group) => shape_transform_mut(&mut group.children, rest),
            _ => None,
        }
    } else {
        match shape {
            ShapeNode::Shape(shape) => Some(&mut shape.base.transform),
            ShapeNode::Picture(picture) => Some(&mut picture.base.transform),
            _ => None,
        }
    }
}

fn text_body_mut<'a>(
    shapes: &'a mut [ShapeNode],
    shape_path: &[usize],
    location: &TextBodyLocation,
) -> Option<&'a mut TextBody> {
    let (index, rest) = shape_path.split_first()?;
    let shape = shapes.get_mut(*index)?;
    if !rest.is_empty() {
        let ShapeNode::Group(group) = shape else {
            return None;
        };
        return text_body_mut(&mut group.children, rest, location);
    }
    match (shape, location) {
        (ShapeNode::Shape(shape), TextBodyLocation::Shape) => shape.text.as_mut(),
        (ShapeNode::GraphicFrame(frame), TextBodyLocation::TableCell { row, cell }) => {
            let GraphicFrameData::Table { rows } = &mut frame.data else {
                return None;
            };
            rows.get_mut(*row)?.get_mut(*cell)
        }
        _ => None,
    }
}

fn paragraph_text(paragraph: &TextParagraph) -> String {
    paragraph.runs.iter().map(|run| run.text.as_str()).collect()
}

fn snapshot_text(paragraph: &ParagraphSnapshot) -> String {
    paragraph.runs.iter().map(|run| run.text.as_str()).collect()
}

fn describe(context: &SlideContext<'_>, shape: &ShapeSnapshot) -> String {
    format!("slide {} shape {:?}", context.index + 1, shape.name)
}

/// Names a change this writer cannot express yet.
///
/// Only reach for this when undoing the change would let the same save
/// through: it is the one error the desktop answers by offering to abandon
/// edits. A broken write, a limit, or the writer contradicting itself are
/// [`EditError::WriteFailed`], [`EditError::WriteLimit`] and
/// [`EditError::VerificationFailed`].
fn unprojectable(reason: impl Into<String>) -> EditError {
    EditError::Unprojectable(reason.into())
}
