//! Text projection into slide parts.
//!
//! The source XML is the template. Only the character data of the `<a:t>`
//! elements an edit names is replaced; every other byte of the part — run
//! properties, paragraph properties, bullets, fields, namespace prefixes,
//! whitespace and unmodelled elements — is copied through untouched.

use std::collections::BTreeMap;

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::PptxError;

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
/// can never be silently dropped.
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
        let position = reader.buffer_position() as usize;
        let event = reader
            .read_event()
            .map_err(|error| malformed(part, position, error.to_string()))?;
        match event {
            Event::Start(start) => {
                let name = local_name(start.name().into_inner()).to_vec();
                let frame = open_frame(&name, &mut frames, &mut shape_path);
                if name == b"t"
                    && let Some(key) = run_text_key(&frames, &shape_path)
                    && pending.contains_key(&key)
                {
                    open_text = Some((reader.buffer_position() as usize, key, frames.len()));
                }
                frames.push(frame);
            }
            Event::Empty(start) => {
                let name = local_name(start.name().into_inner()).to_vec();
                let frame = open_frame(&name, &mut frames, &mut shape_path);
                if name == b"t"
                    && let Some(key) = run_text_key(&frames, &shape_path)
                    && let Some(text) = pending.remove(&key)
                {
                    splices.push((
                        position,
                        reader.buffer_position() as usize,
                        empty_run_text(text),
                    ));
                }
                close_frame(frame, &mut shape_path);
            }
            Event::End(_) => {
                let Some(frame) = frames.pop() else {
                    return Err(malformed(part, position, "unexpected closing element"));
                };
                if let Some((start, key, depth)) = open_text.take() {
                    if depth == frames.len() {
                        if let Some(text) = pending.remove(&key) {
                            splices.push((start, position, escape_text(text)));
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

fn escape_text(text: &str) -> Vec<u8> {
    quick_xml::escape::escape(text).into_owned().into_bytes()
}

fn empty_run_text(text: &str) -> Vec<u8> {
    let mut output = b"<a:t>".to_vec();
    output.extend_from_slice(&escape_text(text));
    output.extend_from_slice(b"</a:t>");
    output
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
    fn an_empty_edit_list_returns_the_source_bytes() {
        assert_eq!(
            rewrite_slide_run_text("slide1.xml", SLIDE.as_bytes(), &[]).unwrap(),
            SLIDE.as_bytes()
        );
    }
}
