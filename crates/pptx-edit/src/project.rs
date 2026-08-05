//! Projection of deck edits back onto the source package.
//!
//! Only run text is expressible in this slice. Everything the projection
//! cannot write is refused by name, so a save never quietly drops an edit.
//!
//! The projection is verified twice rather than assumed. Planning re-seeds a
//! story from the rewritten text bodies and requires the result to equal the
//! live deck exactly, so a change the plan did not account for — a formatting
//! patch, a paragraph break, a run split — is refused. That check runs on the
//! model held in memory, so it alone would trust the byte splice to have landed
//! on the `<a:t>` the plan named. The bytes are therefore zipped and parsed
//! back, and the package that reads out of them must equal the package the plan
//! built. A splice that lands on the wrong run reads back as a different model
//! and the save is refused.

use std::collections::BTreeMap;

use ooxml_drawingml::Theme;
use pptx_parse::{
    GraphicFrameData, PptxPackage, RunTextEdit, ShapeNode, TextBody, TextBodyLocation,
    TextParagraph, TextRun,
};
use yrs::{Map, TextRef, Transact, WriteTxn};

use crate::deck::{seed_baseline, snapshot_doc, theme_for_layout};
use crate::story::{seed_story, snapshot_story};
use crate::{
    DeckSession, EditError, EditResult, ParagraphSnapshot, STORIES, ShapeSnapshot, SlideSnapshot,
    StorySnapshot,
};

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
    /// changed. Returns [`EditError::Unprojectable`] when the deck holds a
    /// change this writer cannot express.
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
        let bytes = pptx_parse::write_pptx(&package)
            .map_err(|error| unprojectable(format!("the deck could not be re-zipped: {error}")))?;
        if !rewrote_parts {
            return Ok((package, bytes));
        }
        let read_back = pptx_parse::parse_pptx(&bytes).map_err(|error| {
            unprojectable(format!("the rewritten deck did not read back: {error}"))
        })?;
        if read_back != package {
            return Err(unprojectable(
                "the rewritten part bytes read back as a different deck, so the text was written \
                 somewhere other than the runs the edit named",
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
            return Err(EditError::Unprojectable(
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
                .ok_or_else(|| unprojectable("a slide lost its source part"))?;
            let context = SlideContext {
                index,
                part_path: source.part_path.clone(),
                theme: theme_for_layout(&package, source.layout_part_path.as_deref()),
            };
            if source.shapes.len() != current_slide.shapes.len() {
                return Err(unprojectable(format!(
                    "slide {} added or removed shapes",
                    index + 1
                )));
            }
            for (shape_index, source_shape) in source.shapes.iter().enumerate() {
                project_shape(
                    &context,
                    source_shape,
                    &baseline_slide.shapes[shape_index],
                    &current_slide.shapes[shape_index],
                    &mut vec![shape_index],
                    &mut plan,
                )?;
            }
        }

        let rewrote_parts = !plan.edits.is_empty();
        for (part_path, edits) in &plan.edits {
            let bytes = package
                .part_bytes(part_path)
                .ok_or_else(|| EditError::Unprojectable(format!("part {part_path} is missing")))?;
            let rewritten = pptx_parse::rewrite_slide_run_text(part_path, bytes, edits)
                .map_err(|error| EditError::Unprojectable(error.to_string()))?;
            package.replace_part(part_path, rewritten);
        }
        for body in plan.bodies {
            let target = package
                .slides
                .get_mut(body.slide_index)
                .and_then(|slide| {
                    text_body_mut(&mut slide.shapes, &body.shape_path, &body.location)
                })
                .ok_or_else(|| unprojectable("a rewritten text body left the shape tree"))?;
            *target = body.body;
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
    edits: BTreeMap<String, Vec<RunTextEdit>>,
    bodies: Vec<ProjectedBody>,
    charged_edits: usize,
    charged_bytes: usize,
}

impl<'a> Plan<'a> {
    fn new(limits: &'a WriteLimits) -> Self {
        Self {
            limits,
            edits: BTreeMap::new(),
            bodies: Vec::new(),
            charged_edits: 0,
            charged_bytes: 0,
        }
    }

    fn charge(&mut self, text: &str, origin: &str) -> EditResult<()> {
        if text.len() > self.limits.max_run_text_bytes {
            return Err(unprojectable(format!(
                "{origin}: the edit writes {} bytes into one run, over the {} bytes one run may \
                 hold in a save",
                text.len(),
                self.limits.max_run_text_bytes
            )));
        }
        self.charged_edits += 1;
        if self.charged_edits > self.limits.max_run_edits {
            return Err(unprojectable(format!(
                "the save rewrites more than the {} runs one save may rewrite",
                self.limits.max_run_edits
            )));
        }
        self.charged_bytes += text.len();
        if self.charged_bytes > self.limits.max_total_edit_bytes {
            return Err(unprojectable(format!(
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

fn project_shape(
    context: &SlideContext<'_>,
    source: &ShapeNode,
    baseline: &ShapeSnapshot,
    current: &ShapeSnapshot,
    shape_path: &mut Vec<usize>,
    plan: &mut Plan<'_>,
) -> EditResult<()> {
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
    if source.paragraphs.len() != current_story.paragraphs.len() {
        return Err(unprojectable(format!(
            "{} added or removed a paragraph",
            describe(context, current)
        )));
    }

    let mut predicted = source.clone();
    let mut edits = Vec::new();
    for (index, paragraph) in source.paragraphs.iter().enumerate() {
        let before = paragraph_text(paragraph);
        let after = snapshot_text(&current_story.paragraphs[index]);
        if before == after {
            continue;
        }
        let origin = format!("{} paragraph {}", describe(context, current), index + 1);
        let (run_index, text) = locate_run_edit(&paragraph.runs, &before, &after)
            .map_err(|reason| unprojectable(format!("{origin}: {reason}")))?;
        plan.charge(&text, &origin)?;
        predicted.paragraphs[index].runs[run_index].text = text.clone();
        edits.push(RunTextEdit {
            shape_path: shape_path.to_vec(),
            location: location.clone(),
            paragraph_index: index,
            run_index,
            text,
        });
    }

    if story_snapshot(&predicted, &current_story.id, context.theme)? != *current_story {
        return Err(unprojectable(format!(
            "{} changed in a way only run text cannot express, such as formatting, a bullet, or a \
             run split",
            describe(context, current)
        )));
    }
    if edits.is_empty() {
        return Ok(());
    }
    plan.edits
        .entry(context.part_path.clone())
        .or_default()
        .extend(edits);
    plan.bodies.push(ProjectedBody {
        slide_index: context.index,
        shape_path: shape_path.to_vec(),
        location,
        body: predicted,
    });
    Ok(())
}

/// Finds the single `<a:r>` whose text the change is confined to.
///
/// The common prefix and suffix of the paragraph's old and new text bound the
/// change; anything wider than one run — a run merge, a split, or an edit that
/// straddles a field or a line break — has no faithful single-run rewrite and
/// is refused.
///
/// An insertion sitting exactly on the trailing edge of an `<a:br>` or `<a:fld>`
/// is the one case where the containing run is not the right target: the break
/// or the field cannot hold text, but the run that follows it can, and the text
/// lands where the caret was either way. The insertion is handed to that run;
/// its style is then the following run's, which the story comparison in
/// `project_body` refuses if that is not what the deck shows.
fn locate_run_edit(runs: &[TextRun], before: &str, after: &str) -> Result<(usize, String), String> {
    if runs.is_empty() {
        return Err("the paragraph has no run to hold text".to_owned());
    }
    let before: Vec<char> = before.chars().collect();
    let after: Vec<char> = after.chars().collect();
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
    if replacement.contains(['\n', '\r']) {
        return Err("line breaks are not written yet".to_owned());
    }

    let mut offset = 0;
    let mut deferred = None;
    for (index, run) in runs.iter().enumerate() {
        let end = offset + run.text.chars().count();
        if offset <= prefix && changed_end <= end {
            let holder = if run.line_break {
                Some(format!("line break {}", index + 1))
            } else if run.field_id.is_some() || run.field_type.is_some() {
                Some(format!("field {}", index + 1))
            } else {
                None
            };
            if let Some(holder) = holder {
                if prefix == changed_end && changed_end == end {
                    deferred = Some(format!(
                        "the change lands past {holder}, which has no run after it to hold the text"
                    ));
                    offset = end;
                    continue;
                }
                return Err(format!("the change lands on {holder}"));
            }
            let mut text: Vec<char> = run.text.chars().collect();
            text.splice(prefix - offset..changed_end - offset, replacement.chars());
            return Ok((index, text.into_iter().collect()));
        }
        offset = end;
    }
    Err(deferred.unwrap_or_else(|| "the change spans more than one run".to_owned()))
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

fn unprojectable(reason: impl Into<String>) -> EditError {
    EditError::Unprojectable(reason.into())
}
