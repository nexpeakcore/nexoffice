//! Sheet drawings: cell anchors and the charts they frame.
//!
//! Read-only. The drawing and chart parts stay owned by the preserved
//! package on save; this parse exists so the canvas can paint them.

use xlsx_model::{AnchorCell, DrawingAnchor, SheetDrawing};

use crate::ParseError;
use crate::dom::{XmlElement, parse_dom};
use crate::xml::{find_part, resolve_part_path};

use ooxml_drawingml::chart::{ChartXml, parse_chart_space};
use ooxml_drawingml::{parse_color_value, resolve_color_value_to_hex};

impl ChartXml for XmlElement {
    fn local_name(&self) -> &str {
        XmlElement::local_name(self)
    }

    fn attribute(&self, prefix: Option<&str>, name: &str) -> Option<&str> {
        prefix
            .and_then(|prefix| XmlElement::attribute(self, &format!("{prefix}:{name}")))
            .or_else(|| XmlElement::attribute(self, name))
    }

    fn child_elements(&self) -> impl Iterator<Item = &Self> {
        XmlElement::child_elements(self)
    }

    fn descendant_text(&self) -> String {
        self.text_content()
    }

    fn solid_fill_hex(&self) -> Option<String> {
        let rgb = self
            .descendants_named("srgbClr")
            .first()
            .and_then(|element| element.attribute("val"))
            .map(str::to_owned);
        resolve_color_value_to_hex(Some(&parse_color_value(rgb.as_deref(), None, None, None)))
    }
}

/// The `<drawing r:id>` of a worksheet, scanned separately so the main
/// worksheet pass stays untouched.
pub(crate) fn worksheet_drawing_rid(data: &[u8]) -> Option<String> {
    let mut reader = quick_xml::Reader::from_reader(data);
    reader.config_mut().trim_text(true);
    loop {
        match reader.read_event().ok()? {
            quick_xml::events::Event::Empty(start) | quick_xml::events::Event::Start(start) => {
                let name = start.name();
                let local = name.into_inner().rsplit(|byte| *byte == b':').next();
                if local == Some(b"drawing") {
                    for attribute in start.attributes().flatten() {
                        let key = attribute.key.into_inner();
                        if key.ends_with(b"id") {
                            return String::from_utf8(attribute.value.into_owned()).ok();
                        }
                    }
                }
            }
            quick_xml::events::Event::Eof => return None,
            _ => {}
        }
    }
}

/// Parses one sheet's drawing part into anchored charts. Anchors without a
/// chart (pictures, plain shapes) are skipped — they are not modeled yet.
pub(crate) fn parse_sheet_drawings(
    parts: &[(String, Vec<u8>)],
    drawing_path: &str,
) -> Result<Vec<SheetDrawing>, ParseError> {
    let Some(bytes) = find_part(parts, drawing_path) else {
        return Ok(Vec::new());
    };
    let root = parse_dom(bytes, drawing_path)?;
    let drawing_rels = {
        let rels_path = relationship_part_path(drawing_path);
        find_part(parts, &rels_path)
            .map(|bytes| parse_dom(bytes, &rels_path))
            .transpose()?
    };

    let mut drawings = Vec::new();
    for anchor_element in root.child_elements() {
        let anchor = match anchor_element.local_name() {
            "twoCellAnchor" => {
                let (Some(from), Some(to)) = (
                    anchor_element.child("from").and_then(anchor_cell),
                    anchor_element.child("to").and_then(anchor_cell),
                ) else {
                    continue;
                };
                DrawingAnchor {
                    from,
                    to: Some(to),
                    extent_emu: None,
                }
            }
            "oneCellAnchor" => {
                let Some(from) = anchor_element.child("from").and_then(anchor_cell) else {
                    continue;
                };
                DrawingAnchor {
                    from,
                    to: None,
                    extent_emu: extent(anchor_element),
                }
            }
            _ => continue,
        };
        let Some(chart_rid) = anchor_element
            .descendants_named("chart")
            .first()
            .and_then(|element| element.attribute_local("id"))
            .map(str::to_owned)
        else {
            continue;
        };
        let Some(chart_target) = drawing_rels.as_ref().and_then(|rels| {
            rels.children_named("Relationship")
                .find(|relationship| relationship.attribute("Id") == Some(chart_rid.as_str()))
                .and_then(|relationship| relationship.attribute("Target"))
                .map(str::to_owned)
        }) else {
            continue;
        };
        let base = drawing_path.rsplit_once('/').map_or("", |(dir, _)| dir);
        let chart_path = resolve_part_path(base, &chart_target);
        let Some(chart_bytes) = find_part(parts, &chart_path) else {
            continue;
        };
        let chart_root = parse_dom(chart_bytes, &chart_path)?;
        let Some(chart) = parse_chart_space(&chart_root) else {
            continue;
        };
        drawings.push(SheetDrawing { anchor, chart });
    }
    Ok(drawings)
}

fn anchor_cell(element: &XmlElement) -> Option<AnchorCell> {
    let number = |name: &str| -> Option<i64> {
        element
            .child(name)
            .map(|child| child.text_content())
            .and_then(|text| text.trim().parse().ok())
    };
    Some(AnchorCell {
        col: u32::try_from(number("col")?).ok()?,
        col_offset_emu: number("colOff").unwrap_or(0),
        row: u32::try_from(number("row")?).ok()?,
        row_offset_emu: number("rowOff").unwrap_or(0),
    })
}

fn extent(anchor: &XmlElement) -> Option<(i64, i64)> {
    let ext = anchor.child("ext")?;
    let parse = |name: &str| ext.attribute(name).and_then(|value| value.parse().ok());
    Some((parse("cx")?, parse("cy")?))
}

fn relationship_part_path(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((directory, filename)) => format!("{directory}/_rels/{filename}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}
