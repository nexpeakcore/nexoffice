//! Text projection into slide parts.
//!
//! The source XML is the template. Only the character data of the `<a:t>`
//! elements an edit names is replaced; every other byte of the part — run
//! properties, paragraph properties, bullets, fields, namespace prefixes,
//! whitespace and unmodelled elements — is copied through untouched.

use std::collections::BTreeMap;

use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};

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

/// Replacement character data for the `<a:t>` of one run.
///
/// `shape_path` indexes the shape tree the way the parser walks it: each entry
/// is the position of a `<p:sp>`/`<p:pic>`/`<p:graphicFrame>`/`<p:grpSp>` among
/// its container's shape children, descending through groups. `run_index`
/// counts `<a:r>`, `<a:fld>` and `<a:br>` children of the paragraph together,
/// matching the parsed run list, but only an `<a:r>` may be edited.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct RunTextEdit {
    pub shape_path: Vec<usize>,
    pub location: TextBodyLocation,
    pub paragraph_index: usize,
    pub run_index: usize,
    pub text: String,
}

type RunKey = (Vec<usize>, TextBodyLocation, usize, usize);

/// Rewrites the named runs of a slide part, leaving every other byte in place.
///
/// Fails when an edit names a run the part does not contain, so a projection
/// can never be silently dropped, and when the replacement text holds a
/// character XML cannot represent.
///
/// # Splice bounds
///
/// Every splice is derived from [`Reader::buffer_position`] read immediately
/// before an event (`opens_at`, the first byte of that event) and immediately
/// after it (`ends_at`, one past its last byte). A `<a:t>` run's character data
/// therefore spans the `ends_at` of its `Start` to the `opens_at` of its `End`,
/// while an `<a:t/>` run spans its own `opens_at..ends_at`. quick-xml documents
/// that position as consumed input rather than as element bounds, so both
/// spans are checked against the source bytes before use — see
/// [`check_element_span`] and [`check_character_data_span`].
pub fn rewrite_slide_run_text(
    part: &str,
    bytes: &[u8],
    edits: &[RunTextEdit],
) -> Result<Vec<u8>, PptxError> {
    if edits.is_empty() {
        return Ok(bytes.to_vec());
    }
    let mut pending = BTreeMap::new();
    for edit in edits {
        let key = (
            edit.shape_path.clone(),
            edit.location.clone(),
            edit.paragraph_index,
            edit.run_index,
        );
        if pending.insert(key, edit.text.as_str()).is_some() {
            return Err(PptxError::MissingTextTarget {
                part: part.to_owned(),
                target: format!("duplicate edit for {}", describe(edit)),
            });
        }
    }

    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut frames: Vec<Frame> = Vec::new();
    let mut shape_path: Vec<usize> = Vec::new();
    let mut splices: Vec<(usize, usize, Vec<u8>)> = Vec::new();
    let mut open_text: Option<(usize, RunKey, usize)> = None;

    loop {
        let opens_at = reader.buffer_position() as usize;
        let event = reader
            .read_event()
            .map_err(|error| malformed(part, opens_at, error.to_string()))?;
        let ends_at = reader.buffer_position() as usize;
        match event {
            Event::Start(start) => {
                let name = local_name(start.name().into_inner()).to_vec();
                let frame = open_frame(&name, &mut frames, &mut shape_path);
                if name == b"t"
                    && let Some(key) = run_text_key(&frames, &shape_path)
                    && pending.contains_key(&key)
                {
                    open_text = Some((ends_at, key, frames.len()));
                }
                frames.push(frame);
            }
            Event::Empty(tag) => {
                let name = local_name(tag.name().into_inner()).to_vec();
                let frame = open_frame(&name, &mut frames, &mut shape_path);
                if name == b"t"
                    && let Some(key) = run_text_key(&frames, &shape_path)
                    && let Some(text) = pending.remove(&key)
                {
                    check_element_span(part, bytes, opens_at, ends_at)?;
                    splices.push((opens_at, ends_at, empty_run_text(part, &key, &tag, text)?));
                }
                close_frame(frame, &mut shape_path);
            }
            Event::End(_) => {
                let Some(frame) = frames.pop() else {
                    return Err(malformed(part, opens_at, "unexpected closing element"));
                };
                if let Some((start, key, depth)) = open_text.take() {
                    if depth == frames.len() {
                        if let Some(text) = pending.remove(&key) {
                            check_character_data_span(part, bytes, start, opens_at)?;
                            splices.push((start, opens_at, escape_text(part, &key, text)?));
                        }
                    } else {
                        open_text = Some((start, key, depth));
                    }
                }
                close_frame(frame, &mut shape_path);
            }
            Event::Eof => break,
            _ => {}
        }
    }

    if let Some((key, _)) = pending.into_iter().next() {
        return Err(PptxError::MissingTextTarget {
            part: part.to_owned(),
            target: describe_key(&key),
        });
    }

    let mut output = Vec::with_capacity(bytes.len());
    let mut cursor = 0;
    for (start, end, replacement) in splices {
        output.extend_from_slice(&bytes[cursor..start]);
        output.extend_from_slice(&replacement);
        cursor = end;
    }
    output.extend_from_slice(&bytes[cursor..]);
    Ok(output)
}

struct Frame {
    kind: FrameKind,
    pushed_shape: bool,
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
        index: usize,
        next_run: usize,
    },
    Run {
        regular: bool,
        index: usize,
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
            let index = *next_paragraph;
            *next_paragraph += 1;
            plain(FrameKind::Paragraph { index, next_run: 0 })
        }
        (b"r" | b"fld" | b"br", Some(FrameKind::Paragraph { next_run, .. })) => {
            let index = *next_run;
            *next_run += 1;
            plain(FrameKind::Run {
                regular: name == b"r",
                index,
            })
        }
        _ => plain(FrameKind::Plain),
    }
}

fn plain(kind: FrameKind) -> Frame {
    Frame {
        kind,
        pushed_shape: false,
    }
}

fn close_frame(frame: Frame, shape_path: &mut Vec<usize>) {
    if frame.pushed_shape {
        shape_path.pop();
    }
}

fn run_text_key(frames: &[Frame], shape_path: &[usize]) -> Option<RunKey> {
    let mut tail = frames.iter().rev();
    let &FrameKind::Run {
        regular: true,
        index: run,
    } = &tail.next()?.kind
    else {
        return None;
    };
    let &FrameKind::Paragraph {
        index: paragraph, ..
    } = &tail.next()?.kind
    else {
        return None;
    };
    let FrameKind::TextBody {
        location: Some(location),
        ..
    } = &tail.next()?.kind
    else {
        return None;
    };
    Some((shape_path.to_vec(), location.clone(), paragraph, run))
}

fn local_name(name: &[u8]) -> &[u8] {
    match name.iter().rposition(|byte| *byte == b':') {
        Some(index) => &name[index + 1..],
        None => name,
    }
}

/// Encodes replacement text as `<a:t>` character data, refusing anything XML
/// cannot represent.
///
/// A carriage return is written as a character reference because an XML parser
/// normalises a literal one to a line feed, which would silently change the
/// text the next time the deck is opened.
fn escape_text(part: &str, key: &RunKey, text: &str) -> Result<Vec<u8>, PptxError> {
    if let Some((index, character)) = text
        .chars()
        .enumerate()
        .find(|(_, character)| !is_legal_xml_character(*character))
    {
        return Err(PptxError::UnwritableText {
            part: part.to_owned(),
            target: describe_key(key),
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

/// Rebuilds an empty `<a:t/>` around its replacement text.
///
/// The source tag's own bytes are re-emitted verbatim, so a part using a
/// different namespace prefix or carrying attributes such as
/// `xml:space="preserve"` keeps both.
fn empty_run_text(
    part: &str,
    key: &RunKey,
    tag: &BytesStart<'_>,
    text: &str,
) -> Result<Vec<u8>, PptxError> {
    let escaped = escape_text(part, key, text)?;
    let mut output = b"<".to_vec();
    output.extend_from_slice(tag);
    output.push(b'>');
    output.extend_from_slice(&escaped);
    output.extend_from_slice(b"</");
    output.extend_from_slice(tag.name().into_inner());
    output.push(b'>');
    Ok(output)
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

fn describe(edit: &RunTextEdit) -> String {
    describe_key(&(
        edit.shape_path.clone(),
        edit.location.clone(),
        edit.paragraph_index,
        edit.run_index,
    ))
}

fn describe_key(key: &RunKey) -> String {
    let path = key
        .0
        .iter()
        .map(usize::to_string)
        .collect::<Vec<_>>()
        .join(".");
    let body = match &key.1 {
        TextBodyLocation::Shape => "text body".to_owned(),
        TextBodyLocation::TableCell { row, cell } => format!("table cell {row}/{cell}"),
    };
    format!("shape {path} {body} paragraph {} run {}", key.2, key.3)
}

fn malformed(part: &str, offset: usize, message: impl Into<String>) -> PptxError {
    PptxError::MalformedXml {
        part: part.to_owned(),
        offset: offset as u64,
        message: message.into(),
    }
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

    fn edit(
        shape_path: &[usize],
        location: TextBodyLocation,
        run_index: usize,
        text: &str,
    ) -> RunTextEdit {
        RunTextEdit {
            shape_path: shape_path.to_vec(),
            location,
            paragraph_index: 0,
            run_index,
            text: text.to_owned(),
        }
    }

    #[test]
    fn rewrites_only_the_named_run_text() {
        let output = rewrite_slide_run_text(
            "slide1.xml",
            SLIDE.as_bytes(),
            &[edit(&[0], TextBodyLocation::Shape, 3, "TWO & <more>")],
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
        let output = rewrite_slide_run_text(
            "slide1.xml",
            SLIDE.as_bytes(),
            &[
                edit(&[1, 1], TextBodyLocation::Shape, 0, "grouped"),
                edit(
                    &[2],
                    TextBodyLocation::TableCell { row: 0, cell: 0 },
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
            edit(&[0], TextBodyLocation::Shape, 1, "into a break"),
            edit(&[0], TextBodyLocation::Shape, 2, "into a field"),
            edit(&[0], TextBodyLocation::Shape, 9, "past the end"),
            edit(&[7], TextBodyLocation::Shape, 0, "missing shape"),
            edit(
                &[2],
                TextBodyLocation::TableCell { row: 4, cell: 0 },
                0,
                "missing row",
            ),
        ] {
            assert!(matches!(
                rewrite_slide_run_text("slide1.xml", SLIDE.as_bytes(), &[target]),
                Err(PptxError::MissingTextTarget { .. })
            ));
        }
    }

    #[test]
    fn two_edits_naming_the_same_run_are_refused() {
        for (label, second) in [("different text", "second"), ("the same text", "first")] {
            let result = rewrite_slide_run_text(
                "slide1.xml",
                SLIDE.as_bytes(),
                &[
                    edit(&[0], TextBodyLocation::Shape, 0, "first"),
                    edit(&[0], TextBodyLocation::Shape, 0, second),
                ],
            );
            let Err(error @ PptxError::MissingTextTarget { .. }) = result else {
                panic!("{label} was accepted: {result:?}");
            };
            let message = error.to_string();
            assert!(message.contains("duplicate edit for"), "{message}");
            assert!(
                message.contains("shape 0 text body paragraph 0 run 0"),
                "{message}"
            );
        }
    }

    #[test]
    fn an_empty_edit_list_returns_the_source_bytes() {
        assert_eq!(
            rewrite_slide_run_text("slide1.xml", SLIDE.as_bytes(), &[]).unwrap(),
            SLIDE.as_bytes()
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
                edit(&[0], TextBodyLocation::Shape, 0, text),
                edit(&[1, 1], TextBodyLocation::Shape, 0, text),
                edit(
                    &[2],
                    TextBodyLocation::TableCell { row: 0, cell: 0 },
                    0,
                    text,
                ),
            ] {
                let result = rewrite_slide_run_text("slide1.xml", SLIDE.as_bytes(), &[target]);
                assert!(
                    matches!(result, Err(PptxError::UnwritableText { .. })),
                    "{text:?} was accepted: {result:?}"
                );
            }
        }
    }

    #[test]
    fn the_refusal_names_the_run_and_the_character() {
        let error = rewrite_slide_run_text(
            "slide1.xml",
            SLIDE.as_bytes(),
            &[edit(&[0], TextBodyLocation::Shape, 0, "ab\u{7}c")],
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
        let output = rewrite_slide_run_text(
            "slide1.xml",
            SLIDE.as_bytes(),
            &[edit(&[0], TextBodyLocation::Shape, 0, "a\tb\nc\rd")],
        )
        .unwrap();
        let output = String::from_utf8(output).unwrap();
        assert!(output.contains("<a:t>a\tb\nc&#13;d</a:t>"), "{output}");
        assert!(every_prefix_resolves(&output));
    }

    #[test]
    fn an_empty_run_keeps_its_prefix_and_attributes() {
        let output = rewrite_slide_run_text(
            "slide2.xml",
            PREFIXED.as_bytes(),
            &[
                edit(&[0], TextBodyLocation::Shape, 1, "preserved"),
                edit(&[0], TextBodyLocation::Shape, 2, "closed"),
            ],
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
    fn splices_land_on_the_element_and_its_character_data() {
        let output = rewrite_slide_run_text(
            "slide2.xml",
            PREFIXED.as_bytes(),
            &[
                edit(&[0], TextBodyLocation::Shape, 0, "ONE"),
                edit(&[0], TextBodyLocation::Shape, 1, "TWO"),
                edit(&[0], TextBodyLocation::Shape, 2, "THREE"),
            ],
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
}
