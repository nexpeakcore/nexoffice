//! Text projection into slide parts.
//!
//! The source XML is the template. A rewrite only ever re-emits bytes the part
//! already holds: the character data of the `<a:t>` elements an edit names, and
//! — where a paragraph is split, merged, or given a line break — copies of the
//! `<a:p>` tag, the `<a:pPr>` and the `<a:rPr>` the source already spells.
//! Every other byte of the part is copied through untouched.
//!
//! # Paragraph structure
//!
//! [`ParagraphRewrite`] describes what one contiguous group of source
//! paragraphs becomes. One source paragraph and one output paragraph is a
//! plain text edit; one source paragraph and two output paragraphs is a split;
//! two source paragraphs and one output paragraph is a merge.
//!
//! Splitting inserts `</a:p>` and the source `<a:p>` open tag followed by a
//! copy of the source `<a:pPr>`, so the new paragraph carries the level,
//! alignment, bullet, indents and default run properties the split paragraph
//! had. Nothing else is synthesised: an `<a:endParaRPr>` is never invented and
//! never duplicated, so it stays at the end of the group, on the last output
//! paragraph — which is where PowerPoint keeps it, and which makes a merge the
//! exact inverse of a split.

use std::collections::{BTreeMap, BTreeSet};

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::PptxError;
use crate::xml::is_legal_xml_character;

/// Which text body of a shape an edit addresses.
#[derive(Clone, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub enum TextBodyLocation {
    /// The `<p:txBody>` of a `<p:sp>`.
    Shape,
    /// The `<a:txBody>` of the table cell at `row`/`cell` in a `<p:graphicFrame>`.
    TableCell { row: usize, cell: usize },
}

/// One run of one source paragraph.
///
/// `paragraph` counts `<a:p>` children of the text body, `run` counts `<a:r>`,
/// `<a:fld>` and `<a:br>` children of that paragraph together, matching the
/// parsed run list.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub struct RunRef {
    pub paragraph: usize,
    pub run: usize,
}

/// One piece of an output paragraph.
///
/// Every piece names the source run it is made of, so a rewrite can only
/// re-emit bytes the part already holds.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum RunPiece {
    /// The source run, byte for byte.
    Keep(RunRef),
    /// The source `<a:r>` re-emitted around different character data.
    Text(RunRef, String),
    /// A new `<a:br>` carrying a copy of the source run's `<a:rPr>`.
    Break(RunRef),
}

impl RunPiece {
    fn target(&self) -> RunRef {
        match self {
            Self::Keep(target) | Self::Text(target, _) | Self::Break(target) => *target,
        }
    }

    fn holds_text(&self) -> bool {
        matches!(self, Self::Keep(_) | Self::Text(_, _))
    }
}

/// What one contiguous group of source paragraphs becomes.
///
/// `shape_path` indexes the shape tree the way the parser walks it: each entry
/// is the position of a `<p:sp>`/`<p:pic>`/`<p:graphicFrame>`/`<p:grpSp>` among
/// its container's shape children, descending through groups.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ParagraphRewrite {
    pub shape_path: Vec<usize>,
    pub location: TextBodyLocation,
    /// Index of the first source `<a:p>` the rewrite covers.
    pub first_paragraph: usize,
    /// How many source paragraphs it covers; more than one is a merge.
    pub source_paragraphs: usize,
    /// The paragraphs written in their place, each naming the source runs it
    /// keeps in source order.
    pub paragraphs: Vec<Vec<RunPiece>>,
}

impl ParagraphRewrite {
    /// A rewrite that only replaces the character data of one run.
    pub fn run_text(
        shape_path: Vec<usize>,
        location: TextBodyLocation,
        paragraph: usize,
        runs: usize,
        run: usize,
        text: impl Into<String>,
    ) -> Self {
        let text = text.into();
        let pieces = (0..runs)
            .map(|index| {
                let target = RunRef {
                    paragraph,
                    run: index,
                };
                if index == run {
                    RunPiece::Text(target, text.clone())
                } else {
                    RunPiece::Keep(target)
                }
            })
            .collect();
        Self {
            shape_path,
            location,
            first_paragraph: paragraph,
            source_paragraphs: 1,
            paragraphs: vec![pieces],
        }
    }

    fn body(&self) -> BodyKey {
        (self.shape_path.clone(), self.location.clone())
    }

    fn last_paragraph(&self) -> usize {
        self.first_paragraph + self.source_paragraphs - 1
    }
}

type BodyKey = (Vec<usize>, TextBodyLocation);

/// Rewrites the named paragraphs of a slide part, leaving every other byte in
/// place.
///
/// Fails when a rewrite names a paragraph or run the part does not contain, so
/// a projection can never be silently dropped; when a piece asks an `<a:br>` or
/// an `<a:fld>` to hold run text; and when replacement text holds a character
/// XML cannot represent.
///
/// # Splice bounds
///
/// Every span is derived from [`Reader::buffer_position`] read immediately
/// before an event (the first byte of that event) and immediately after it (one
/// past its last byte). An element's bytes therefore span its `Start`'s
/// `opens_at` to its `End`'s `ends_at`, and a `<a:t>` run's character data spans
/// the `ends_at` of its `Start` to the `opens_at` of its `End`. quick-xml
/// documents that position as consumed input rather than as element bounds, so
/// every span is checked against the source bytes before use — see
/// [`check_element_span`], [`check_element_bounds`] and
/// [`check_character_data_span`].
pub fn rewrite_slide_text(
    part: &str,
    bytes: &[u8],
    rewrites: &[ParagraphRewrite],
) -> Result<Vec<u8>, PptxError> {
    if rewrites.is_empty() {
        return Ok(bytes.to_vec());
    }
    let wanted = coverage(part, rewrites)?;
    let collected = collect_spans(part, bytes, &wanted)?;

    let mut splices: Vec<(usize, usize, Vec<u8>)> = Vec::new();
    for rewrite in rewrites {
        let mut spans = BTreeMap::new();
        for index in rewrite.first_paragraph..=rewrite.last_paragraph() {
            let key = (rewrite.body(), index);
            let paragraph = collected
                .get(&key)
                .ok_or_else(|| missing(part, &key.0, index, None))?;
            spans.insert(index, paragraph);
        }
        splices.extend(plan_splices(part, bytes, rewrite, &spans)?);
    }

    splices.sort_by_key(|(start, _, _)| *start);
    let mut output = Vec::with_capacity(bytes.len());
    let mut cursor = 0;
    for (start, end, replacement) in splices {
        if start < cursor || end < start || end > bytes.len() {
            return Err(malformed(part, start, "two rewrites overlap in the part"));
        }
        output.extend_from_slice(&bytes[cursor..start]);
        output.extend_from_slice(&replacement);
        cursor = end;
    }
    output.extend_from_slice(&bytes[cursor..]);
    Ok(output)
}

/// The paragraphs each rewrite covers, refusing overlapping rewrites so no
/// paragraph is written twice.
fn coverage(
    part: &str,
    rewrites: &[ParagraphRewrite],
) -> Result<BTreeMap<BodyKey, BTreeSet<usize>>, PptxError> {
    let mut wanted: BTreeMap<BodyKey, BTreeSet<usize>> = BTreeMap::new();
    for rewrite in rewrites {
        if rewrite.source_paragraphs == 0 || rewrite.paragraphs.is_empty() {
            return Err(missing(
                part,
                &rewrite.body(),
                rewrite.first_paragraph,
                Some("a rewrite covers no paragraph"),
            ));
        }
        let covered = wanted.entry(rewrite.body()).or_default();
        for index in rewrite.first_paragraph..=rewrite.last_paragraph() {
            if !covered.insert(index) {
                return Err(missing(
                    part,
                    &rewrite.body(),
                    index,
                    Some("two rewrites cover the same paragraph"),
                ));
            }
        }
    }
    Ok(wanted)
}

#[derive(Default)]
struct ParagraphSpans {
    open: (usize, usize),
    properties: Option<(usize, usize)>,
    runs: Vec<RunSpans>,
    end_properties: Option<(usize, usize)>,
    /// `None` for a self-closing `<a:p/>`, which has no separate close tag.
    close: Option<(usize, usize)>,
}

impl ParagraphSpans {
    fn content_start(&self) -> usize {
        self.properties.map_or(self.open.1, |(_, end)| end)
    }

    fn content_end(&self) -> usize {
        self.end_properties
            .map(|(start, _)| start)
            .or_else(|| self.close.map(|(start, _)| start))
            .unwrap_or(self.open.1)
    }
}

struct RunSpans {
    regular: bool,
    span: (usize, usize),
    properties: Option<(usize, usize)>,
    text: Option<TextSpans>,
}

/// Where a run's `<a:t>` keeps its character data.
///
/// `data` is the character data of a `<a:t>text</a:t>`, or the whole element of
/// a self-closing `<a:t/>`, in which case `empty_tag` holds the source tag's own
/// bytes so a rewrite can rebuild the element around them.
struct TextSpans {
    data: (usize, usize),
    empty_tag: Option<Vec<u8>>,
}

fn collect_spans(
    part: &str,
    bytes: &[u8],
    wanted: &BTreeMap<BodyKey, BTreeSet<usize>>,
) -> Result<BTreeMap<(BodyKey, usize), ParagraphSpans>, PptxError> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut frames: Vec<Frame> = Vec::new();
    let mut shape_path: Vec<usize> = Vec::new();
    let mut collected: BTreeMap<(BodyKey, usize), ParagraphSpans> = BTreeMap::new();
    let mut current: Option<(BodyKey, usize, ParagraphSpans)> = None;

    loop {
        let opens_at = reader.buffer_position() as usize;
        let event = reader
            .read_event()
            .map_err(|error| malformed(part, opens_at, error.to_string()))?;
        let ends_at = reader.buffer_position() as usize;
        match event {
            Event::Start(start) => {
                let name = local_name(start.name().into_inner()).to_vec();
                let mut frame = open_frame(&name, &mut frames, &mut shape_path);
                frame.start = opens_at;
                frame.role = span_role(&name, &frame, &frames, current.is_some());
                if let SpanRole::Paragraph = frame.role
                    && let Some((key, index)) = paragraph_key(&frames, &shape_path)
                    && wanted
                        .get(&key)
                        .is_some_and(|indexes| indexes.contains(&index))
                {
                    let mut spans = ParagraphSpans::default();
                    check_element_bounds(part, bytes, opens_at, ends_at)?;
                    spans.open = (opens_at, ends_at);
                    current = Some((key, index, spans));
                } else if let SpanRole::Paragraph = frame.role {
                    frame.role = SpanRole::None;
                }
                if let Some((_, _, spans)) = current.as_mut() {
                    match frame.role {
                        SpanRole::Run => spans.runs.push(RunSpans {
                            regular: name == b"r",
                            span: (opens_at, opens_at),
                            properties: None,
                            text: None,
                        }),
                        SpanRole::RunText => frame.text_start = ends_at,
                        _ => {}
                    }
                }
                frames.push(frame);
            }
            Event::Empty(tag) => {
                let name = local_name(tag.name().into_inner()).to_vec();
                let mut frame = open_frame(&name, &mut frames, &mut shape_path);
                frame.start = opens_at;
                frame.role = span_role(&name, &frame, &frames, current.is_some());
                if let SpanRole::Paragraph = frame.role
                    && let Some((key, index)) = paragraph_key(&frames, &shape_path)
                    && wanted
                        .get(&key)
                        .is_some_and(|indexes| indexes.contains(&index))
                {
                    check_element_span(part, bytes, opens_at, ends_at)?;
                    collected.insert(
                        (key, index),
                        ParagraphSpans {
                            open: (opens_at, ends_at),
                            ..ParagraphSpans::default()
                        },
                    );
                } else if let Some((_, _, spans)) = current.as_mut() {
                    check_element_span(part, bytes, opens_at, ends_at)?;
                    match frame.role {
                        SpanRole::Run => spans.runs.push(RunSpans {
                            regular: name == b"r",
                            span: (opens_at, ends_at),
                            properties: None,
                            text: None,
                        }),
                        SpanRole::RunText => {
                            if let Some(run) = spans.runs.last_mut() {
                                run.text = Some(TextSpans {
                                    data: (opens_at, ends_at),
                                    empty_tag: Some(tag.to_vec()),
                                });
                            }
                        }
                        role => record_span(spans, role, (opens_at, ends_at)),
                    }
                }
                close_frame(frame, &mut shape_path);
            }
            Event::End(_) => {
                let Some(frame) = frames.pop() else {
                    return Err(malformed(part, opens_at, "unexpected closing element"));
                };
                match frame.role {
                    SpanRole::Paragraph => {
                        if let Some((key, index, mut spans)) = current.take() {
                            check_element_bounds(part, bytes, opens_at, ends_at)?;
                            spans.close = Some((opens_at, ends_at));
                            collected.insert((key, index), spans);
                        }
                    }
                    SpanRole::RunText => {
                        if let Some((_, _, spans)) = current.as_mut()
                            && let Some(run) = spans.runs.last_mut()
                        {
                            check_character_data_span(part, bytes, frame.text_start, opens_at)?;
                            run.text = Some(TextSpans {
                                data: (frame.text_start, opens_at),
                                empty_tag: None,
                            });
                        }
                    }
                    role => {
                        if let Some((_, _, spans)) = current.as_mut() {
                            check_element_bounds(part, bytes, frame.start, ends_at)?;
                            record_span(spans, role, (frame.start, ends_at));
                        }
                    }
                }
                close_frame(frame, &mut shape_path);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(collected)
}

fn record_span(spans: &mut ParagraphSpans, role: SpanRole, span: (usize, usize)) {
    match role {
        SpanRole::ParagraphProperties => spans.properties = Some(span),
        SpanRole::EndProperties => spans.end_properties = Some(span),
        SpanRole::Run => {
            if let Some(run) = spans.runs.last_mut() {
                run.span = span;
            }
        }
        SpanRole::RunProperties => {
            if let Some(run) = spans.runs.last_mut() {
                run.properties = Some(span);
            }
        }
        _ => {}
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum SpanRole {
    None,
    Paragraph,
    ParagraphProperties,
    Run,
    RunProperties,
    RunText,
    EndProperties,
}

fn span_role(name: &[u8], frame: &Frame, frames: &[Frame], collecting: bool) -> SpanRole {
    if let FrameKind::Paragraph { .. } = frame.kind {
        return SpanRole::Paragraph;
    }
    if !collecting {
        return SpanRole::None;
    }
    if let FrameKind::Run { .. } = frame.kind {
        return SpanRole::Run;
    }
    match (name, frames.last().map(|frame| &frame.kind)) {
        (b"pPr", Some(FrameKind::Paragraph { .. })) => SpanRole::ParagraphProperties,
        (b"endParaRPr", Some(FrameKind::Paragraph { .. })) => SpanRole::EndProperties,
        (b"rPr", Some(FrameKind::Run { .. })) => SpanRole::RunProperties,
        (b"t", Some(FrameKind::Run { regular: true, .. })) => SpanRole::RunText,
        _ => SpanRole::None,
    }
}

struct Frame {
    kind: FrameKind,
    pushed_shape: bool,
    role: SpanRole,
    start: usize,
    text_start: usize,
}

enum FrameKind {
    Plain,
    ShapeTree {
        next_shape: usize,
    },
    Shape,
    GraphicFrame,
    Graphic,
    GraphicData,
    Table {
        next_row: usize,
    },
    Row {
        row: usize,
        next_cell: usize,
    },
    Cell {
        row: usize,
        cell: usize,
    },
    TextBody {
        location: Option<TextBodyLocation>,
        next_paragraph: usize,
    },
    Paragraph {
        next_run: usize,
    },
    Run {
        regular: bool,
    },
}

fn open_frame(name: &[u8], frames: &mut [Frame], shape_path: &mut Vec<usize>) -> Frame {
    let parent = frames.last_mut().map(|frame| &mut frame.kind);
    match (name, parent) {
        (b"spTree", _) => plain(FrameKind::ShapeTree { next_shape: 0 }),
        (
            b"sp" | b"pic" | b"graphicFrame" | b"grpSp",
            Some(FrameKind::ShapeTree { next_shape }),
        ) => {
            shape_path.push(*next_shape);
            *next_shape += 1;
            Frame {
                kind: match name {
                    b"sp" => FrameKind::Shape,
                    b"graphicFrame" => FrameKind::GraphicFrame,
                    b"grpSp" => FrameKind::ShapeTree { next_shape: 0 },
                    _ => FrameKind::Plain,
                },
                pushed_shape: true,
                role: SpanRole::None,
                start: 0,
                text_start: 0,
            }
        }
        (b"graphic", Some(FrameKind::GraphicFrame)) => plain(FrameKind::Graphic),
        (b"graphicData", Some(FrameKind::Graphic)) => plain(FrameKind::GraphicData),
        (b"tbl", Some(FrameKind::GraphicData)) => plain(FrameKind::Table { next_row: 0 }),
        (b"tr", Some(FrameKind::Table { next_row })) => {
            let row = *next_row;
            *next_row += 1;
            plain(FrameKind::Row { row, next_cell: 0 })
        }
        (b"tc", Some(FrameKind::Row { row, next_cell })) => {
            let cell = *next_cell;
            *next_cell += 1;
            plain(FrameKind::Cell { row: *row, cell })
        }
        (b"txBody", Some(FrameKind::Shape)) => plain(FrameKind::TextBody {
            location: Some(TextBodyLocation::Shape),
            next_paragraph: 0,
        }),
        (b"txBody", Some(FrameKind::Cell { row, cell })) => plain(FrameKind::TextBody {
            location: Some(TextBodyLocation::TableCell {
                row: *row,
                cell: *cell,
            }),
            next_paragraph: 0,
        }),
        (
            b"p",
            Some(FrameKind::TextBody {
                location: Some(_),
                next_paragraph,
            }),
        ) => {
            *next_paragraph += 1;
            plain(FrameKind::Paragraph { next_run: 0 })
        }
        (b"r" | b"fld" | b"br", Some(FrameKind::Paragraph { next_run })) => {
            *next_run += 1;
            plain(FrameKind::Run {
                regular: name == b"r",
            })
        }
        _ => plain(FrameKind::Plain),
    }
}

fn plain(kind: FrameKind) -> Frame {
    Frame {
        kind,
        pushed_shape: false,
        role: SpanRole::None,
        start: 0,
        text_start: 0,
    }
}

fn close_frame(frame: Frame, shape_path: &mut Vec<usize>) {
    if frame.pushed_shape {
        shape_path.pop();
    }
}

/// The body and paragraph index of the paragraph frame just opened, whose own
/// frame is not on the stack yet.
fn paragraph_key(frames: &[Frame], shape_path: &[usize]) -> Option<(BodyKey, usize)> {
    let FrameKind::TextBody {
        location: Some(location),
        next_paragraph,
    } = &frames.last()?.kind
    else {
        return None;
    };
    Some((
        (shape_path.to_vec(), location.clone()),
        next_paragraph.checked_sub(1)?,
    ))
}

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|byte| *byte == b':') {
        Some(index) => &name[index + 1..],
        None => name,
    }
}

/// Turns one rewrite into the splices that write it.
///
/// The walk carries a cursor through the source bytes: content that stays is
/// stepped over rather than re-emitted, so every byte between two pieces —
/// whitespace, comments, elements this crate does not model — survives.
fn plan_splices(
    part: &str,
    bytes: &[u8],
    rewrite: &ParagraphRewrite,
    spans: &BTreeMap<usize, &ParagraphSpans>,
) -> Result<Vec<(usize, usize, Vec<u8>)>, PptxError> {
    let first = rewrite.first_paragraph;
    let last = rewrite.last_paragraph();
    let stream = flatten(part, rewrite, spans)?;
    let mut splices = Splices::new();
    let mut cursor = spans[&first].content_start();
    let mut source = first;
    let mut boundaries = 0;
    let mut index = 0;

    while index < stream.len() {
        let Item::Piece(piece) = &stream[index] else {
            boundaries += 1;
            index += 1;
            continue;
        };
        let target = piece.target();
        if target.paragraph != source {
            cross(
                part,
                bytes,
                spans,
                &mut splices,
                cursor,
                source,
                target.paragraph,
                &mut boundaries,
            )?;
            source = target.paragraph;
        }
        let mut pending = boundary_bytes(bytes, spans[&source], std::mem::take(&mut boundaries));
        let group_end = group_end(&stream, index, target);
        let run = &spans[&source].runs[target.run];
        let group = &stream[index..group_end];
        if let [Item::Piece(single @ (RunPiece::Keep(_) | RunPiece::Text(_, _)))] = group {
            if !pending.is_empty() {
                splices.push(part, run.span.0, run.span.0, std::mem::take(&mut pending))?;
            }
            if let RunPiece::Text(_, text) = single {
                let target_text = run
                    .text
                    .as_ref()
                    .ok_or_else(|| missing_run(part, rewrite, target, "has no <a:t> to write"))?;
                splices.push(
                    part,
                    target_text.data.0,
                    target_text.data.1,
                    render_text(part, rewrite, target, target_text, text)?,
                )?;
            }
        } else {
            let mut rebuilt = std::mem::take(&mut pending);
            for item in group {
                match item {
                    Item::Boundary => rebuilt.extend(boundary_bytes(bytes, spans[&source], 1)),
                    Item::Piece(RunPiece::Keep(_)) => {
                        rebuilt.extend_from_slice(&bytes[run.span.0..run.span.1]);
                    }
                    Item::Piece(RunPiece::Text(_, text)) => {
                        let target_text = run.text.as_ref().ok_or_else(|| {
                            missing_run(part, rewrite, target, "has no <a:t> to write")
                        })?;
                        rebuilt.extend_from_slice(&bytes[run.span.0..target_text.data.0]);
                        rebuilt.extend(render_text(part, rewrite, target, target_text, text)?);
                        rebuilt.extend_from_slice(&bytes[target_text.data.1..run.span.1]);
                    }
                    Item::Piece(RunPiece::Break(_)) => {
                        rebuilt.extend(break_bytes(bytes, run));
                    }
                }
            }
            splices.push(part, run.span.0, run.span.1, rebuilt)?;
        }
        cursor = run.span.1;
        index = group_end;
    }

    if source != last {
        cross(
            part,
            bytes,
            spans,
            &mut splices,
            cursor,
            source,
            last,
            &mut boundaries,
        )?;
    }
    if boundaries > 0 {
        let at = spans[&last].content_end();
        splices.push(
            part,
            at,
            at,
            boundary_bytes(bytes, spans[&last], boundaries),
        )?;
    }
    Ok(splices.items)
}

/// Steps from one source paragraph to a later one.
///
/// A paragraph end the output still wants is stepped over rather than rewritten,
/// so an untouched paragraph boundary keeps its own bytes; one the output has no
/// paragraph left for is cut away, which is what merges two paragraphs into one.
#[allow(clippy::too_many_arguments)]
fn cross(
    part: &str,
    bytes: &[u8],
    spans: &BTreeMap<usize, &ParagraphSpans>,
    splices: &mut Splices,
    cursor: usize,
    source: usize,
    target: usize,
    boundaries: &mut usize,
) -> Result<usize, PptxError> {
    let start = spans[&target].content_start();
    let crossings = target - source;
    if *boundaries >= crossings {
        *boundaries -= crossings;
    } else {
        let written = boundary_bytes(bytes, spans[&source], *boundaries);
        *boundaries = 0;
        splices.push(part, cursor, start, written)?;
    }
    Ok(start)
}

enum Item<'a> {
    Boundary,
    Piece(&'a RunPiece),
}

/// Flattens the output paragraphs into one stream of pieces separated by
/// paragraph boundaries, checking that the pieces name every source run of the
/// group exactly once, in order, and that only a regular `<a:r>` is asked to
/// hold text.
fn flatten<'a>(
    part: &str,
    rewrite: &'a ParagraphRewrite,
    spans: &BTreeMap<usize, &ParagraphSpans>,
) -> Result<Vec<Item<'a>>, PptxError> {
    let mut stream = Vec::new();
    let mut expected = Vec::new();
    for paragraph in rewrite.first_paragraph..=rewrite.last_paragraph() {
        for run in 0..spans[&paragraph].runs.len() {
            expected.push(RunRef { paragraph, run });
        }
    }
    let pieces: Vec<&RunPiece> = rewrite.paragraphs.iter().flatten().collect();
    let mut seen: Vec<RunRef> = Vec::new();
    for (index, piece) in pieces.iter().enumerate() {
        let target = piece.target();
        let run = spans
            .get(&target.paragraph)
            .and_then(|spans| spans.runs.get(target.run))
            .ok_or_else(|| missing_run(part, rewrite, target, "is not in the paragraph"))?;
        if !run.regular && !matches!(piece, RunPiece::Keep(_)) {
            return Err(missing_run(
                part,
                rewrite,
                target,
                "is a line break or a field, which cannot hold run text",
            ));
        }
        if let Some(previous) = seen.last()
            && *previous > target
        {
            return Err(missing_run(
                part,
                rewrite,
                target,
                "is written out of source order",
            ));
        }
        if piece.holds_text() {
            if seen.last() != Some(&target) {
                seen.push(target);
            }
        } else {
            let neighbours = [index.checked_sub(1), Some(index + 1)];
            if !neighbours.iter().flatten().any(|neighbour| {
                pieces
                    .get(*neighbour)
                    .is_some_and(|piece| piece.holds_text() && piece.target() == target)
            }) {
                return Err(missing_run(
                    part,
                    rewrite,
                    target,
                    "would take its line break properties from a run it does not touch",
                ));
            }
        }
    }
    if seen != expected {
        return Err(missing(
            part,
            &rewrite.body(),
            rewrite.first_paragraph,
            Some("the rewrite does not keep every source run exactly once"),
        ));
    }

    for (index, paragraph) in rewrite.paragraphs.iter().enumerate() {
        if index > 0 {
            stream.push(Item::Boundary);
        }
        stream.extend(paragraph.iter().map(Item::Piece));
    }
    Ok(stream)
}

/// The end of the run of consecutive items writing one source run, taking in
/// the paragraph boundaries between them but not those that follow it.
fn group_end(stream: &[Item<'_>], start: usize, target: RunRef) -> usize {
    let mut end = start;
    let mut index = start;
    while index < stream.len() {
        match &stream[index] {
            Item::Boundary => index += 1,
            Item::Piece(piece) if piece.target() == target => {
                index += 1;
                end = index;
            }
            Item::Piece(_) => break,
        }
    }
    end
}

/// `count` copies of `</a:p>` followed by the source paragraph's own open tag
/// and a copy of its `<a:pPr>`, which is what starts the next paragraph of a
/// split.
fn boundary_bytes(bytes: &[u8], spans: &ParagraphSpans, count: usize) -> Vec<u8> {
    let mut output = Vec::new();
    for _ in 0..count {
        let Some(close) = spans.close else {
            output.extend_from_slice(&bytes[spans.open.0..spans.open.1]);
            continue;
        };
        output.extend_from_slice(&bytes[close.0..close.1]);
        output.extend_from_slice(&bytes[spans.open.0..spans.open.1]);
        if let Some(properties) = spans.properties {
            output.extend_from_slice(&bytes[properties.0..properties.1]);
        }
    }
    output
}

/// A `<a:br>` carrying a copy of the source run's `<a:rPr>`, spelt with the
/// run's own namespace prefix.
fn break_bytes(bytes: &[u8], run: &RunSpans) -> Vec<u8> {
    let name = element_name(&bytes[run.span.0..run.span.1]);
    let prefix = match name.iter().position(|byte| *byte == b':') {
        Some(index) => &name[..=index],
        None => b"".as_slice(),
    };
    let Some(properties) = run.properties else {
        let mut output = b"<".to_vec();
        output.extend_from_slice(prefix);
        output.extend_from_slice(b"br/>");
        return output;
    };
    let mut output = b"<".to_vec();
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"br>");
    output.extend_from_slice(&bytes[properties.0..properties.1]);
    output.extend_from_slice(b"</");
    output.extend_from_slice(prefix);
    output.extend_from_slice(b"br>");
    output
}

fn element_name(tag: &[u8]) -> &[u8] {
    let rest = tag.strip_prefix(b"<").unwrap_or(tag);
    let end = rest
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || *byte == b'>' || *byte == b'/')
        .unwrap_or(rest.len());
    &rest[..end]
}

fn render_text(
    part: &str,
    rewrite: &ParagraphRewrite,
    target: RunRef,
    spans: &TextSpans,
    text: &str,
) -> Result<Vec<u8>, PptxError> {
    let escaped = escape_text(part, rewrite, target, text)?;
    let Some(tag) = &spans.empty_tag else {
        return Ok(escaped);
    };
    let mut output = b"<".to_vec();
    output.extend_from_slice(tag);
    output.push(b'>');
    output.extend_from_slice(&escaped);
    output.extend_from_slice(b"</");
    output.extend_from_slice(element_name(tag));
    output.push(b'>');
    Ok(output)
}

struct Splices {
    items: Vec<(usize, usize, Vec<u8>)>,
    cursor: usize,
}

impl Splices {
    fn new() -> Self {
        Self {
            items: Vec::new(),
            cursor: 0,
        }
    }

    fn push(
        &mut self,
        part: &str,
        start: usize,
        end: usize,
        bytes: Vec<u8>,
    ) -> Result<(), PptxError> {
        if end < start || start < self.cursor {
            return Err(malformed(
                part,
                start,
                "a rewrite planned splices that run backwards",
            ));
        }
        if start == end && bytes.is_empty() {
            return Ok(());
        }
        match self.items.last_mut() {
            Some(last) if last.1 == start => {
                last.1 = end;
                last.2.extend(bytes);
            }
            _ => self.items.push((start, end, bytes)),
        }
        self.cursor = end;
        Ok(())
    }
}

/// Encodes replacement text as `<a:t>` character data, refusing anything XML
/// cannot represent.
///
/// A carriage return is written as a character reference because an XML parser
/// normalises a literal one to a line feed, which would silently change the
/// text the next time the deck is opened.
fn escape_text(
    part: &str,
    rewrite: &ParagraphRewrite,
    target: RunRef,
    text: &str,
) -> Result<Vec<u8>, PptxError> {
    if let Some((index, character)) = text
        .chars()
        .enumerate()
        .find(|(_, character)| !is_legal_xml_character(*character))
    {
        return Err(PptxError::UnwritableText {
            part: part.to_owned(),
            target: describe(&rewrite.body(), target),
            reason: format!(
                "character {} of the run is U+{:04X}, which XML cannot store; delete that \
                 character — retyping the run's text replaces it — and save again",
                index + 1,
                character as u32
            ),
        });
    }
    let escaped = quick_xml::escape::escape(text);
    if escaped.contains('\r') {
        return Ok(escaped.replace('\r', "&#13;").into_bytes());
    }
    Ok(escaped.into_owned().into_bytes())
}

/// Asserts that `start..end` covers exactly one self-closing element.
///
/// The splice bounds come from [`Reader::buffer_position`] before and after an
/// event, which quick-xml documents as a byte count of consumed input rather
/// than as element bounds. Checking the slice turns a change in that meaning
/// into a loud refusal instead of a silently corrupted part.
fn check_element_span(part: &str, bytes: &[u8], start: usize, end: usize) -> Result<(), PptxError> {
    let span = bytes
        .get(start..end)
        .ok_or_else(|| malformed(part, start, "element bounds ran past the part"))?;
    if !span.starts_with(b"<") || !span.ends_with(b"/>") {
        return Err(malformed(
            part,
            start,
            "element bounds did not cover a self-closing tag",
        ));
    }
    Ok(())
}

/// Asserts that `start..end` covers a whole tag or element.
///
/// See [`check_element_span`] for why the bounds are checked at all.
fn check_element_bounds(
    part: &str,
    bytes: &[u8],
    start: usize,
    end: usize,
) -> Result<(), PptxError> {
    let span = bytes
        .get(start..end)
        .ok_or_else(|| malformed(part, start, "element bounds ran past the part"))?;
    if !span.starts_with(b"<") || !span.ends_with(b">") {
        return Err(malformed(part, start, "element bounds did not cover a tag"));
    }
    Ok(())
}

/// Asserts that `start..end` covers exactly the character data of an element.
///
/// See [`check_element_span`] for why the bounds are checked at all.
fn check_character_data_span(
    part: &str,
    bytes: &[u8],
    start: usize,
    end: usize,
) -> Result<(), PptxError> {
    let span = bytes
        .get(start..end)
        .ok_or_else(|| malformed(part, start, "character data bounds ran past the part"))?;
    if !bytes[..start].ends_with(b">") || !bytes[end..].starts_with(b"</") || span.contains(&b'<') {
        return Err(malformed(
            part,
            start,
            "character data bounds did not cover the inside of an element",
        ));
    }
    Ok(())
}

fn missing(part: &str, body: &BodyKey, paragraph: usize, reason: Option<&str>) -> PptxError {
    let target = describe_body(body, paragraph);
    PptxError::MissingTextTarget {
        part: part.to_owned(),
        target: match reason {
            Some(reason) => format!("{target}: {reason}"),
            None => target,
        },
    }
}

fn missing_run(part: &str, rewrite: &ParagraphRewrite, target: RunRef, reason: &str) -> PptxError {
    PptxError::MissingTextTarget {
        part: part.to_owned(),
        target: format!("{} {reason}", describe(&rewrite.body(), target)),
    }
}

fn describe(body: &BodyKey, target: RunRef) -> String {
    format!(
        "{} run {}",
        describe_body(body, target.paragraph),
        target.run
    )
}

fn describe_body(body: &BodyKey, paragraph: usize) -> String {
    let path = body
        .0
        .iter()
        .map(usize::to_string)
        .collect::<Vec<_>>()
        .join(".");
    let location = match &body.1 {
        TextBodyLocation::Shape => "text body".to_owned(),
        TextBodyLocation::TableCell { row, cell } => format!("table cell {row}/{cell}"),
    };
    format!("shape {path} {location} paragraph {paragraph}")
}

fn malformed(part: &str, offset: usize, message: impl Into<String>) -> PptxError {
    PptxError::MalformedXml {
        part: part.to_owned(),
        offset: offset as u64,
        message: message.into(),
    }
}

/// One shape's rewritten placement, in EMU.
///
/// The values land in the `<a:off>`/`<a:ext>` of the `<a:xfrm>` the shape
/// already spells inside its `<p:spPr>`; a shape whose position lives in a
/// layout (no explicit `<a:xfrm>`) is refused rather than given one, so the
/// writer keeps re-emitting only structure the part already holds.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ShapeTransformRewrite {
    /// Child indexes into the shape tree, counting `sp`/`pic`/`graphicFrame`/
    /// `grpSp` the way the parser does.
    pub shape_path: Vec<usize>,
    pub x: i64,
    pub y: i64,
    pub width: i64,
    pub height: i64,
}

/// Rewrites the named shapes' `<a:off>`/`<a:ext>`, leaving every other byte in
/// place. Only `sp` and `pic` shapes qualify: groups and graphic frames keep
/// child geometry that plain offset/extent splices cannot express.
pub fn rewrite_slide_geometry(
    part: &str,
    bytes: &[u8],
    rewrites: &[ShapeTransformRewrite],
) -> Result<Vec<u8>, PptxError> {
    if rewrites.is_empty() {
        return Ok(bytes.to_vec());
    }
    let mut wanted: BTreeMap<Vec<usize>, &ShapeTransformRewrite> = BTreeMap::new();
    for rewrite in rewrites {
        if wanted.insert(rewrite.shape_path.clone(), rewrite).is_some() {
            return Err(malformed(part, 0, "two rewrites cover the same shape"));
        }
    }

    struct GeoFrame {
        name: Vec<u8>,
        counts_shapes: bool,
        next_shape: usize,
        pushed_shape: bool,
    }

    /// The full byte span and qualified tag name of one `<a:off>`/`<a:ext>`.
    type TagSpan = (usize, usize, Vec<u8>);

    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut frames: Vec<GeoFrame> = Vec::new();
    let mut shape_path: Vec<usize> = Vec::new();
    // The target's spPr>xfrm child spans, keyed by the covered shape path.
    let mut found: BTreeMap<Vec<usize>, (Option<TagSpan>, Option<TagSpan>)> = BTreeMap::new();

    let target_state = |frames: &[GeoFrame],
                        shape_path: &[usize],
                        wanted: &BTreeMap<Vec<usize>, &ShapeTransformRewrite>|
     -> bool {
        // Directly inside the wanted shape's own spPr>xfrm: the frame under
        // the spPr must be the shape frame itself, and its path the target's.
        wanted.contains_key(shape_path) && frames.last().is_some_and(|frame| frame.pushed_shape)
    };

    loop {
        let opens_at = reader.buffer_position() as usize;
        let event = reader
            .read_event()
            .map_err(|error| malformed(part, opens_at, error.to_string()))?;
        let ends_at = reader.buffer_position() as usize;
        match event {
            Event::Start(ref start) | Event::Empty(ref start) => {
                let raw_name = start.name().into_inner().to_vec();
                let name = local_name(&raw_name).to_vec();
                let parent_counts = frames.last().is_some_and(|frame| frame.counts_shapes);
                let mut pushed_shape = false;
                if matches!(name.as_slice(), b"sp" | b"pic" | b"graphicFrame" | b"grpSp")
                    && parent_counts
                {
                    let parent = frames.last_mut().expect("counting parent");
                    shape_path.push(parent.next_shape);
                    parent.next_shape += 1;
                    pushed_shape = true;
                }
                let in_xfrm = frames.last().is_some_and(|frame| frame.name == b"xfrm")
                    && frames.len() >= 2
                    && frames[frames.len() - 2].name == b"spPr"
                    && target_state(&frames[..frames.len() - 2], &shape_path, &wanted);
                if in_xfrm && matches!(name.as_slice(), b"off" | b"ext") {
                    if !matches!(event, Event::Empty(_)) {
                        return Err(malformed(
                            part,
                            opens_at,
                            "an <a:off>/<a:ext> with children cannot be rewritten",
                        ));
                    }
                    let slot = found.entry(shape_path.clone()).or_default();
                    let record = (opens_at, ends_at, raw_name.clone());
                    if name == b"off" {
                        slot.0.get_or_insert(record);
                    } else {
                        slot.1.get_or_insert(record);
                    }
                }
                match event {
                    Event::Start(_) => frames.push(GeoFrame {
                        counts_shapes: name == b"spTree" || (name == b"grpSp" && pushed_shape),
                        name,
                        next_shape: 0,
                        pushed_shape,
                    }),
                    _ => {
                        // An empty element opens no frame; a shape spelled
                        // `<p:sp/>` still consumed its index.
                        if pushed_shape {
                            shape_path.pop();
                        }
                    }
                }
            }
            Event::End(_) => {
                if let Some(frame) = frames.pop()
                    && frame.pushed_shape
                {
                    shape_path.pop();
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    let mut splices: Vec<(usize, usize, Vec<u8>)> = Vec::new();
    for (path, rewrite) in &wanted {
        let Some((off, ext)) = found.get(path).map(|(off, ext)| (off.clone(), ext.clone())) else {
            return Err(malformed(
                part,
                0,
                format!("shape {path:?} spells no explicit <a:xfrm> to rewrite"),
            ));
        };
        let (Some(off), Some(ext)) = (off, ext) else {
            return Err(malformed(
                part,
                0,
                format!("shape {path:?} spells no explicit <a:off>/<a:ext> to rewrite"),
            ));
        };
        let off_name = String::from_utf8_lossy(&off.2).into_owned();
        let ext_name = String::from_utf8_lossy(&ext.2).into_owned();
        splices.push((
            off.0,
            off.1,
            format!(r#"<{off_name} x="{}" y="{}"/>"#, rewrite.x, rewrite.y).into_bytes(),
        ));
        splices.push((
            ext.0,
            ext.1,
            format!(
                r#"<{ext_name} cx="{}" cy="{}"/>"#,
                rewrite.width, rewrite.height
            )
            .into_bytes(),
        ));
    }

    splices.sort_by_key(|(start, _, _)| *start);
    let mut output = Vec::with_capacity(bytes.len());
    let mut cursor = 0;
    for (start, end, replacement) in splices {
        if start < cursor || end < start || end > bytes.len() {
            return Err(malformed(part, start, "two geometry rewrites overlap"));
        }
        output.extend_from_slice(&bytes[cursor..start]);
        output.extend_from_slice(&replacement);
        cursor = end;
    }
    output.extend_from_slice(&bytes[cursor..]);
    Ok(output)
}

#[cfg(test)]
mod tests {
    use super::*;

    const SLIDE: &str = concat!(
        r#"<p:sld xmlns:a="a" xmlns:p="p"><p:cSld><p:spTree>"#,
        r#"<p:sp><p:txBody><a:bodyPr/><a:p><a:pPr algn="ctr"/>"#,
        r#"<a:r><a:rPr b="1"/><a:t>one</a:t></a:r>"#,
        r#"<a:br/><a:fld id="x"><a:t>3</a:t></a:fld>"#,
        r#"<a:r><a:rPr i="1"/><a:t>two</a:t></a:r></a:p></p:txBody></p:sp>"#,
        r#"<p:grpSp><p:pic/><p:sp><p:txBody><a:p><a:r><a:t/></a:r></a:p></p:txBody></p:sp></p:grpSp>"#,
        r#"<p:graphicFrame><a:graphic><a:graphicData><a:tbl>"#,
        r#"<a:tr><a:tc><a:txBody><a:p><a:r><a:t>cell</a:t></a:r></a:p></a:txBody></a:tc></a:tr>"#,
        r#"</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#,
        r#"</p:spTree></p:cSld></p:sld>"#,
    );

    /// Two paragraphs of their own, so a split and a merge can be spelt out
    /// against source bytes a test can quote in full.
    const PARAGRAPHS: &str = concat!(
        r#"<p:sld xmlns:a="a" xmlns:p="p"><p:cSld><p:spTree><p:sp><p:txBody>"#,
        r#"<a:p><a:pPr lvl="1"><a:buChar char="•"/></a:pPr>"#,
        r#"<a:r><a:rPr b="1"/><a:t>alpha</a:t></a:r>"#,
        r#"<a:r><a:rPr i="1"/><a:t>beta</a:t></a:r>"#,
        r#"<a:endParaRPr sz="1800"/></a:p>"#,
        r#"<a:p><a:pPr algn="r"/><a:r><a:t>gamma</a:t></a:r></a:p>"#,
        r#"</p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#,
    );

    fn shape_body(paragraph: usize, runs: usize, run: usize, text: &str) -> ParagraphRewrite {
        ParagraphRewrite::run_text(vec![0], TextBodyLocation::Shape, paragraph, runs, run, text)
    }

    fn reference(paragraph: usize, run: usize) -> RunRef {
        RunRef { paragraph, run }
    }

    fn rewrite(
        shape_path: &[usize],
        location: TextBodyLocation,
        first_paragraph: usize,
        source_paragraphs: usize,
        paragraphs: Vec<Vec<RunPiece>>,
    ) -> ParagraphRewrite {
        ParagraphRewrite {
            shape_path: shape_path.to_vec(),
            location,
            first_paragraph,
            source_paragraphs,
            paragraphs,
        }
    }

    #[test]
    fn rewrites_only_the_named_run_text() {
        let output = rewrite_slide_text(
            "slide1.xml",
            SLIDE.as_bytes(),
            &[shape_body(0, 4, 3, "TWO & <more>")],
        )
        .unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains(r#"<a:r><a:rPr i="1"/><a:t>TWO &amp; &lt;more&gt;</a:t></a:r>"#));
        assert_eq!(
            output.replace("TWO &amp; &lt;more&gt;", "two"),
            SLIDE,
            "every other byte survives the rewrite"
        );
    }

    #[test]
    fn addresses_group_children_table_cells_and_empty_runs() {
        let output = rewrite_slide_text(
            "slide1.xml",
            SLIDE.as_bytes(),
            &[
                ParagraphRewrite::run_text(vec![1, 1], TextBodyLocation::Shape, 0, 1, 0, "grouped"),
                ParagraphRewrite::run_text(
                    vec![2],
                    TextBodyLocation::TableCell { row: 0, cell: 0 },
                    0,
                    1,
                    0,
                    "filled",
                ),
            ],
        )
        .unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains("<a:t>grouped</a:t>"));
        assert!(output.contains("<a:t>filled</a:t>"));
        assert!(!output.contains("<a:t/>"));
    }

    #[test]
    fn refuses_edits_the_part_cannot_hold() {
        for target in [
            shape_body(0, 4, 1, "into a break"),
            shape_body(0, 4, 2, "into a field"),
            shape_body(0, 10, 9, "past the end"),
            ParagraphRewrite::run_text(vec![7], TextBodyLocation::Shape, 0, 1, 0, "missing shape"),
            ParagraphRewrite::run_text(
                vec![2],
                TextBodyLocation::TableCell { row: 4, cell: 0 },
                0,
                1,
                0,
                "missing row",
            ),
        ] {
            assert!(matches!(
                rewrite_slide_text("slide1.xml", SLIDE.as_bytes(), &[target]),
                Err(PptxError::MissingTextTarget { .. })
            ));
        }
    }

    #[test]
    fn two_rewrites_of_the_same_paragraph_are_refused() {
        let result = rewrite_slide_text(
            "slide1.xml",
            SLIDE.as_bytes(),
            &[shape_body(0, 4, 0, "first"), shape_body(0, 4, 0, "second")],
        );
        let Err(error @ PptxError::MissingTextTarget { .. }) = result else {
            panic!("two rewrites of one paragraph were accepted: {result:?}");
        };
        let message = error.to_string();
        assert!(
            message.contains("two rewrites cover the same paragraph"),
            "{message}"
        );
        assert!(
            message.contains("shape 0 text body paragraph 0"),
            "{message}"
        );
    }

    #[test]
    fn a_rewrite_that_drops_a_run_is_refused() {
        let result = rewrite_slide_text(
            "slide1.xml",
            SLIDE.as_bytes(),
            &[rewrite(
                &[0],
                TextBodyLocation::Shape,
                0,
                1,
                vec![vec![
                    RunPiece::Keep(reference(0, 0)),
                    RunPiece::Keep(reference(0, 1)),
                    RunPiece::Keep(reference(0, 2)),
                ]],
            )],
        );
        let Err(error @ PptxError::MissingTextTarget { .. }) = result else {
            panic!("a dropped run was accepted: {result:?}");
        };
        assert!(
            error.to_string().contains("every source run exactly once"),
            "{error}"
        );
    }

    #[test]
    fn a_rewrite_that_reorders_runs_is_refused() {
        let result = rewrite_slide_text(
            "slide1.xml",
            SLIDE.as_bytes(),
            &[rewrite(
                &[0],
                TextBodyLocation::Shape,
                0,
                1,
                vec![vec![
                    RunPiece::Keep(reference(0, 1)),
                    RunPiece::Keep(reference(0, 0)),
                    RunPiece::Keep(reference(0, 2)),
                    RunPiece::Keep(reference(0, 3)),
                ]],
            )],
        );
        let Err(error @ PptxError::MissingTextTarget { .. }) = result else {
            panic!("reordered runs were accepted: {result:?}");
        };
        assert!(error.to_string().contains("out of source order"), "{error}");
    }

    #[test]
    fn an_empty_edit_list_returns_the_source_bytes() {
        assert_eq!(
            rewrite_slide_text("slide1.xml", SLIDE.as_bytes(), &[]).unwrap(),
            SLIDE.as_bytes()
        );
    }

    #[test]
    fn splitting_at_a_run_boundary_moves_no_other_byte() {
        let output = rewrite_slide_text(
            "slide2.xml",
            PARAGRAPHS.as_bytes(),
            &[rewrite(
                &[0],
                TextBodyLocation::Shape,
                0,
                1,
                vec![
                    vec![RunPiece::Keep(reference(0, 0))],
                    vec![RunPiece::Keep(reference(0, 1))],
                ],
            )],
        )
        .unwrap();
        assert_eq!(
            String::from_utf8(output).unwrap(),
            PARAGRAPHS.replace(
                r#"</a:r><a:r><a:rPr i="1"/>"#,
                concat!(
                    r#"</a:r></a:p><a:p><a:pPr lvl="1"><a:buChar char="•"/></a:pPr>"#,
                    r#"<a:r><a:rPr i="1"/>"#,
                )
            ),
            "the new paragraph copies the source pPr and nothing else moves"
        );
    }

    #[test]
    fn splitting_inside_a_run_duplicates_its_run_properties() {
        let output = rewrite_slide_text(
            "slide2.xml",
            PARAGRAPHS.as_bytes(),
            &[rewrite(
                &[0],
                TextBodyLocation::Shape,
                0,
                1,
                vec![
                    vec![RunPiece::Text(reference(0, 0), "al".to_owned())],
                    vec![
                        RunPiece::Text(reference(0, 0), "pha".to_owned()),
                        RunPiece::Keep(reference(0, 1)),
                    ],
                ],
            )],
        )
        .unwrap();
        assert_eq!(
            String::from_utf8(output).unwrap(),
            PARAGRAPHS.replace(
                r#"<a:r><a:rPr b="1"/><a:t>alpha</a:t></a:r>"#,
                concat!(
                    r#"<a:r><a:rPr b="1"/><a:t>al</a:t></a:r>"#,
                    r#"</a:p><a:p><a:pPr lvl="1"><a:buChar char="•"/></a:pPr>"#,
                    r#"<a:r><a:rPr b="1"/><a:t>pha</a:t></a:r>"#,
                )
            ),
            "both halves of the divided run keep its <a:rPr> byte for byte"
        );
    }

    #[test]
    fn splitting_at_the_end_leaves_the_end_properties_on_the_new_paragraph() {
        let output = rewrite_slide_text(
            "slide2.xml",
            PARAGRAPHS.as_bytes(),
            &[rewrite(
                &[0],
                TextBodyLocation::Shape,
                0,
                1,
                vec![
                    vec![
                        RunPiece::Keep(reference(0, 0)),
                        RunPiece::Keep(reference(0, 1)),
                    ],
                    vec![],
                ],
            )],
        )
        .unwrap();
        assert_eq!(
            String::from_utf8(output).unwrap(),
            PARAGRAPHS.replace(
                r#"<a:endParaRPr sz="1800"/>"#,
                concat!(
                    r#"</a:p><a:p><a:pPr lvl="1"><a:buChar char="•"/></a:pPr>"#,
                    r#"<a:endParaRPr sz="1800"/>"#,
                )
            ),
            "the end paragraph mark stays at the end of the split"
        );
    }

    #[test]
    fn splitting_at_the_start_leaves_an_empty_paragraph_before_it() {
        let output = rewrite_slide_text(
            "slide2.xml",
            PARAGRAPHS.as_bytes(),
            &[rewrite(
                &[0],
                TextBodyLocation::Shape,
                0,
                1,
                vec![
                    vec![],
                    vec![
                        RunPiece::Keep(reference(0, 0)),
                        RunPiece::Keep(reference(0, 1)),
                    ],
                ],
            )],
        )
        .unwrap();
        assert_eq!(
            String::from_utf8(output).unwrap(),
            PARAGRAPHS.replace(
                r#"<a:r><a:rPr b="1"/><a:t>alpha</a:t></a:r>"#,
                concat!(
                    r#"</a:p><a:p><a:pPr lvl="1"><a:buChar char="•"/></a:pPr>"#,
                    r#"<a:r><a:rPr b="1"/><a:t>alpha</a:t></a:r>"#,
                )
            )
        );
    }

    #[test]
    fn merging_two_paragraphs_deletes_one_span_and_keeps_the_first_ppr() {
        let output = rewrite_slide_text(
            "slide2.xml",
            PARAGRAPHS.as_bytes(),
            &[rewrite(
                &[0],
                TextBodyLocation::Shape,
                0,
                2,
                vec![vec![
                    RunPiece::Keep(reference(0, 0)),
                    RunPiece::Keep(reference(0, 1)),
                    RunPiece::Keep(reference(1, 0)),
                ]],
            )],
        )
        .unwrap();
        assert_eq!(
            String::from_utf8(output).unwrap(),
            PARAGRAPHS.replace(
                concat!(
                    r#"<a:endParaRPr sz="1800"/></a:p>"#,
                    r#"<a:p><a:pPr algn="r"/><a:r><a:t>gamma</a:t></a:r>"#,
                ),
                r#"<a:r><a:t>gamma</a:t></a:r>"#
            ),
            "the merged paragraph keeps the first pPr and the last end properties"
        );
    }

    #[test]
    fn a_paragraph_boundary_the_output_keeps_is_stepped_over_not_rewritten() {
        let output = rewrite_slide_text(
            "slide2.xml",
            PARAGRAPHS.as_bytes(),
            &[rewrite(
                &[0],
                TextBodyLocation::Shape,
                0,
                2,
                vec![
                    vec![
                        RunPiece::Keep(reference(0, 0)),
                        RunPiece::Text(reference(0, 1), "BETA".to_owned()),
                    ],
                    vec![RunPiece::Keep(reference(1, 0))],
                ],
            )],
        )
        .unwrap();
        assert_eq!(
            String::from_utf8(output).unwrap(),
            PARAGRAPHS.replace("<a:t>beta</a:t>", "<a:t>BETA</a:t>"),
            "the boundary between the two paragraphs keeps its own bytes"
        );
    }

    #[test]
    fn a_line_break_copies_the_run_properties_it_divides() {
        let output = rewrite_slide_text(
            "slide2.xml",
            PARAGRAPHS.as_bytes(),
            &[rewrite(
                &[0],
                TextBodyLocation::Shape,
                0,
                1,
                vec![vec![
                    RunPiece::Text(reference(0, 0), "al".to_owned()),
                    RunPiece::Break(reference(0, 0)),
                    RunPiece::Text(reference(0, 0), "pha".to_owned()),
                    RunPiece::Keep(reference(0, 1)),
                ]],
            )],
        )
        .unwrap();
        assert_eq!(
            String::from_utf8(output).unwrap(),
            PARAGRAPHS.replace(
                r#"<a:r><a:rPr b="1"/><a:t>alpha</a:t></a:r>"#,
                concat!(
                    r#"<a:r><a:rPr b="1"/><a:t>al</a:t></a:r>"#,
                    r#"<a:br><a:rPr b="1"/></a:br>"#,
                    r#"<a:r><a:rPr b="1"/><a:t>pha</a:t></a:r>"#,
                )
            )
        );
    }

    /// A part whose `<a:t>` runs use a non-`a` prefix, carry attributes, and
    /// spell their tags with the whitespace XML permits.
    const PREFIXED: &str = concat!(
        r#"<pres:sld xmlns:dml="a" xmlns:pres="p"><pres:cSld><pres:spTree>"#,
        r#"<pres:sp><pres:txBody><dml:p>"#,
        r#"<dml:r><dml:t >first</dml:t ></dml:r>"#,
        r#"<dml:r><dml:t xml:space="preserve" /></dml:r>"#,
        r#"<dml:r><dml:t></dml:t></dml:r>"#,
        r#"</dml:p></pres:txBody></pres:sp>"#,
        r#"</pres:spTree></pres:cSld></pres:sld>"#,
    );

    fn every_prefix_resolves(xml: &str) -> bool {
        use quick_xml::NsReader;
        use quick_xml::name::ResolveResult;

        let mut reader = NsReader::from_str(xml);
        loop {
            match reader.read_resolved_event() {
                Err(_) => return false,
                Ok((_, Event::Eof)) => return true,
                Ok((ResolveResult::Unknown(_), _)) => return false,
                Ok(_) => {}
            }
        }
    }

    #[test]
    fn refuses_text_xml_cannot_represent() {
        for text in [
            "a\u{0}b",
            "a\u{8}b",
            "a\u{b}b",
            "a\u{c}b",
            "a\u{e}b",
            "a\u{1f}b",
            "a\u{fffe}b",
            "a\u{ffff}b",
        ] {
            for target in [
                shape_body(0, 4, 0, text),
                ParagraphRewrite::run_text(vec![1, 1], TextBodyLocation::Shape, 0, 1, 0, text),
                ParagraphRewrite::run_text(
                    vec![2],
                    TextBodyLocation::TableCell { row: 0, cell: 0 },
                    0,
                    1,
                    0,
                    text,
                ),
            ] {
                let result = rewrite_slide_text("slide1.xml", SLIDE.as_bytes(), &[target]);
                assert!(
                    matches!(result, Err(PptxError::UnwritableText { .. })),
                    "{text:?} was accepted: {result:?}"
                );
            }
        }
    }

    #[test]
    fn the_refusal_names_the_run_and_the_character() {
        let error = rewrite_slide_text(
            "slide1.xml",
            SLIDE.as_bytes(),
            &[shape_body(0, 4, 0, "ab\u{7}c")],
        )
        .unwrap_err();
        let message = error.to_string();
        assert!(message.contains("character 3"), "{message}");
        assert!(message.contains("U+0007"), "{message}");
        assert!(message.contains("paragraph 0 run 0"), "{message}");
        assert!(message.contains("retyping"), "{message}");
    }

    #[test]
    fn writes_the_control_characters_xml_allows() {
        let output = rewrite_slide_text(
            "slide1.xml",
            SLIDE.as_bytes(),
            &[shape_body(0, 4, 0, "a\tb\nc\rd")],
        )
        .unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains("<a:t>a\tb\nc&#13;d</a:t>"), "{output}");
        assert!(every_prefix_resolves(&output));
    }

    #[test]
    fn an_empty_run_keeps_its_prefix_and_attributes() {
        let output = rewrite_slide_text(
            "slide2.xml",
            PREFIXED.as_bytes(),
            &[
                ParagraphRewrite::run_text(vec![0], TextBodyLocation::Shape, 0, 3, 1, "preserved"),
                ParagraphRewrite::run_text(vec![0], TextBodyLocation::Shape, 0, 3, 2, "closed"),
            ],
        );
        let Err(error) = output else {
            panic!("two rewrites of one paragraph must be refused");
        };
        assert!(error.to_string().contains("two rewrites cover the same"));

        let output = rewrite_slide_text(
            "slide2.xml",
            PREFIXED.as_bytes(),
            &[rewrite(
                &[0],
                TextBodyLocation::Shape,
                0,
                1,
                vec![vec![
                    RunPiece::Keep(reference(0, 0)),
                    RunPiece::Text(reference(0, 1), "preserved".to_owned()),
                    RunPiece::Text(reference(0, 2), "closed".to_owned()),
                ]],
            )],
        )
        .unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(
            output.contains(r#"<dml:t xml:space="preserve" >preserved</dml:t>"#),
            "{output}"
        );
        assert!(output.contains("<dml:t>closed</dml:t>"), "{output}");
        assert!(!output.contains("a:t"), "{output}");
        assert!(every_prefix_resolves(&output), "{output}");
    }

    #[test]
    fn a_line_break_takes_the_namespace_prefix_of_the_run_it_divides() {
        let output = rewrite_slide_text(
            "slide2.xml",
            PREFIXED.as_bytes(),
            &[rewrite(
                &[0],
                TextBodyLocation::Shape,
                0,
                1,
                vec![vec![
                    RunPiece::Text(reference(0, 0), "fi".to_owned()),
                    RunPiece::Break(reference(0, 0)),
                    RunPiece::Text(reference(0, 0), "rst".to_owned()),
                    RunPiece::Keep(reference(0, 1)),
                    RunPiece::Keep(reference(0, 2)),
                ]],
            )],
        )
        .unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains("<dml:br/>"), "{output}");
        assert!(every_prefix_resolves(&output), "{output}");
    }

    #[test]
    fn splices_land_on_the_element_and_its_character_data() {
        let output = rewrite_slide_text(
            "slide2.xml",
            PREFIXED.as_bytes(),
            &[rewrite(
                &[0],
                TextBodyLocation::Shape,
                0,
                1,
                vec![vec![
                    RunPiece::Text(reference(0, 0), "ONE".to_owned()),
                    RunPiece::Text(reference(0, 1), "TWO".to_owned()),
                    RunPiece::Text(reference(0, 2), "THREE".to_owned()),
                ]],
            )],
        )
        .unwrap();
        let expected = PREFIXED
            .replace("<dml:t >first</dml:t >", "<dml:t >ONE</dml:t >")
            .replace(
                r#"<dml:t xml:space="preserve" />"#,
                r#"<dml:t xml:space="preserve" >TWO</dml:t>"#,
            )
            .replace("<dml:t></dml:t>", "<dml:t>THREE</dml:t>");
        assert_eq!(String::from_utf8(output).unwrap(), expected);
    }

    #[test]
    fn splice_bounds_that_drift_off_an_element_are_refused() {
        let bytes = PREFIXED.as_bytes();
        let data_start = PREFIXED.find("first").unwrap();
        let data_end = data_start + "first".len();
        assert!(check_character_data_span("s.xml", bytes, data_start, data_end).is_ok());
        for (start, end) in [
            (data_start - 1, data_end),
            (data_start, data_end + 1),
            (data_start, bytes.len() + 1),
        ] {
            assert!(
                check_character_data_span("s.xml", bytes, start, end).is_err(),
                "{start}..{end} was accepted as character data"
            );
        }

        let tag_start = PREFIXED.find("<dml:t xml:space").unwrap();
        let tag_end = PREFIXED[tag_start..].find("/>").unwrap() + tag_start + 2;
        assert!(check_element_span("s.xml", bytes, tag_start, tag_end).is_ok());
        for (start, end) in [
            (tag_start + 1, tag_end),
            (tag_start, tag_end - 1),
            (tag_start, bytes.len() + 1),
        ] {
            assert!(
                check_element_span("s.xml", bytes, start, end).is_err(),
                "{start}..{end} was accepted as an element"
            );
        }
    }

    const GEOMETRY: &str = concat!(
        r#"<p:sld xmlns:a="a" xmlns:p="p"><p:cSld><p:spTree>"#,
        r#"<p:sp><p:spPr><a:xfrm rot="0"><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm>"#,
        r#"<a:prstGeom prst="rect"/></p:spPr><p:txBody><a:p><a:r><a:t>t</a:t></a:r></a:p></p:txBody></p:sp>"#,
        r#"<p:pic><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr></p:pic>"#,
        r#"<p:sp><p:spPr><a:prstGeom prst="rect"/></p:spPr></p:sp>"#,
        r#"<p:grpSp><p:grpSpPr><a:xfrm><a:off x="5" y="6"/><a:ext cx="7" cy="8"/>"#,
        r#"<a:chOff x="0" y="0"/><a:chExt cx="7" cy="8"/></a:xfrm></p:grpSpPr>"#,
        r#"<p:sp><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr></p:sp></p:grpSp>"#,
        r#"</p:spTree></p:cSld></p:sld>"#,
    );

    fn transform(
        shape_path: Vec<usize>,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> ShapeTransformRewrite {
        ShapeTransformRewrite {
            shape_path,
            x,
            y,
            width,
            height,
        }
    }

    #[test]
    fn geometry_rewrites_off_and_ext_of_the_named_shape() {
        let out = rewrite_slide_geometry(
            "s.xml",
            GEOMETRY.as_bytes(),
            &[transform(vec![0], 111, 222, 333, 444)],
        )
        .unwrap();
        let out = String::from_utf8(out).unwrap();
        assert!(out.contains(
            r#"<a:xfrm rot="0"><a:off x="111" y="222"/><a:ext cx="333" cy="444"/></a:xfrm>"#
        ));
        assert!(
            out.contains(r#"<a:off x="1" y="2"/>"#),
            "the pic was untouched"
        );
        assert!(
            out.contains(r#"<a:off x="100" y="200"/>"#),
            "the group child was untouched"
        );
    }

    #[test]
    fn geometry_rewrites_a_picture_and_a_nested_child_independently() {
        let out = rewrite_slide_geometry(
            "s.xml",
            GEOMETRY.as_bytes(),
            &[
                transform(vec![1], 9, 8, 7, 6),
                transform(vec![3, 0], 990, 991, 992, 993),
            ],
        )
        .unwrap();
        let out = String::from_utf8(out).unwrap();
        assert!(out.contains(r#"<a:off x="9" y="8"/><a:ext cx="7" cy="6"/>"#));
        assert!(out.contains(r#"<a:off x="990" y="991"/><a:ext cx="992" cy="993"/>"#));
        assert!(
            out.contains(r#"<a:off x="10" y="20"/>"#),
            "shape 0 was untouched"
        );
        assert!(
            out.contains(r#"<a:off x="5" y="6"/>"#),
            "the group's own xfrm was untouched"
        );
    }

    #[test]
    fn geometry_refuses_a_shape_without_an_explicit_xfrm() {
        let error = rewrite_slide_geometry(
            "s.xml",
            GEOMETRY.as_bytes(),
            &[transform(vec![2], 1, 2, 3, 4)],
        )
        .unwrap_err();
        assert!(error.to_string().contains("no explicit"), "{error}");
    }

    #[test]
    fn geometry_never_touches_a_group_xfrm_via_the_group_path() {
        // The group's own grpSpPr xfrm is not inside an spPr, so the walk
        // finds nothing for path [3] and refuses instead of guessing.
        let error = rewrite_slide_geometry(
            "s.xml",
            GEOMETRY.as_bytes(),
            &[transform(vec![3], 1, 2, 3, 4)],
        )
        .unwrap_err();
        assert!(error.to_string().contains("no explicit"), "{error}");
    }
}
