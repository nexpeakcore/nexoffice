//! Sheet drawings: cell anchors and the charts they frame.
//!
//! Read-only. The drawing and chart parts stay owned by the preserved
//! package on save; this parse exists so the canvas can paint them.

use xlsx_model::styles::Theme;
use xlsx_model::{AnchorCell, DrawingAnchor, SheetDrawing};

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
/// Drawings are display-only, so parts that fail to parse (oversized,
/// malformed) drop their charts rather than failing the workbook open.
pub(crate) fn parse_sheet_drawings(
    parts: &[(String, Vec<u8>)],
    drawing_path: &str,
    theme: &Theme,
) -> Vec<SheetDrawing> {
    let Some(bytes) = find_part(parts, drawing_path) else {
        return Vec::new();
    };
    let Ok(root) = parse_dom(bytes, drawing_path) else {
        return Vec::new();
    };
    let drawing_rels = {
        let rels_path = relationship_part_path(drawing_path);
        find_part(parts, &rels_path).and_then(|bytes| parse_dom(bytes, &rels_path).ok())
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
                DrawingAnchor::Cell {
                    from,
                    to: Some(to),
                    extent_emu: None,
                }
            }
            "oneCellAnchor" => {
                let Some(from) = anchor_element.child("from").and_then(anchor_cell) else {
                    continue;
                };
                DrawingAnchor::Cell {
                    from,
                    to: None,
                    extent_emu: extent(anchor_element),
                }
            }
            "absoluteAnchor" => {
                let position = |name: &str| {
                    anchor_element
                        .child("pos")
                        .and_then(|pos| pos.attribute(name))
                        .and_then(|value| value.parse().ok())
                };
                let (Some(x), Some(y), Some(ext)) =
                    (position("x"), position("y"), extent(anchor_element))
                else {
                    continue;
                };
                DrawingAnchor::Absolute {
                    pos_emu: (x, y),
                    extent_emu: ext,
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
        let Ok(mut chart_root) = parse_dom(chart_bytes, &chart_path) else {
            continue;
        };
        resolve_scheme_colors(&mut chart_root, theme);
        let Some(chart) = parse_chart_space(&chart_root) else {
            continue;
        };
        drawings.push(SheetDrawing { anchor, chart });
    }
    drawings
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

/// Rewrites every `srgbClr`/`schemeClr`/`sysClr` in place into a plain
/// `srgbClr` resolved against the workbook theme with its lum/tint/shade
/// modifiers applied, so the theme-blind chart parser reads final colors.
fn resolve_scheme_colors(element: &mut crate::dom::XmlElement, theme: &Theme) {
    use crate::dom::XmlNode;

    for child in &mut element.children {
        if let XmlNode::Element(child) = child {
            resolve_scheme_colors(child, theme);
        }
    }
    let base = match element.local_name() {
        "srgbClr" => element.attribute("val").and_then(parse_hex_rgb),
        "schemeClr" => element
            .attribute("val")
            .and_then(|name| scheme_color(theme, name))
            .and_then(parse_hex_rgb),
        "sysClr" => element
            .attribute("lastClr")
            .and_then(parse_hex_rgb)
            .or_else(|| match element.attribute("val") {
                Some("windowText") => parse_hex_rgb(&theme.colors[0]),
                Some("window") => parse_hex_rgb(&theme.colors[1]),
                _ => None,
            }),
        _ => return,
    };
    let Some(rgb) = base else {
        return;
    };
    let resolved = apply_color_modifiers(rgb, element);
    element.name = "a:srgbClr".to_owned();
    element.attributes.clear();
    element.attributes.insert(
        "val".to_owned(),
        format!(
            "{:02X}{:02X}{:02X}",
            (resolved[0] * 255.0).round() as u8,
            (resolved[1] * 255.0).round() as u8,
            (resolved[2] * 255.0).round() as u8
        ),
    );
    element.children.clear();
}

/// Scheme name to declaration-order theme slot; unknown names resolve to none
/// rather than a wrong palette entry.
fn scheme_color<'a>(theme: &'a Theme, name: &str) -> Option<&'a str> {
    let slot = match name {
        "dk1" | "tx1" => 0,
        "lt1" | "bg1" => 1,
        "dk2" | "tx2" => 2,
        "lt2" | "bg2" => 3,
        "accent1" => 4,
        "accent2" => 5,
        "accent3" => 6,
        "accent4" => 7,
        "accent5" => 8,
        "accent6" => 9,
        "hlink" => 10,
        "folHlink" => 11,
        _ => return None,
    };
    Some(theme.colors[slot].as_str())
}

fn parse_hex_rgb(value: &str) -> Option<[f64; 3]> {
    let value = value.strip_prefix('#').unwrap_or(value);
    if value.len() != 6 {
        return None;
    }
    let packed = u32::from_str_radix(value, 16).ok()?;
    Some([
        f64::from((packed >> 16) & 0xff) / 255.0,
        f64::from((packed >> 8) & 0xff) / 255.0,
        f64::from(packed & 0xff) / 255.0,
    ])
}

/// DrawingML color modifier children, values in thousandths of a percent.
/// `lumMod`/`lumOff` scale and offset HSL luminance (Office's
/// "Lighter/Darker N%" variants); `shade` darkens toward black, `tint`
/// lightens toward white.
fn apply_color_modifiers(mut rgb: [f64; 3], element: &crate::dom::XmlElement) -> [f64; 3] {
    for child in element.child_elements() {
        let Some(value) = child
            .attribute("val")
            .and_then(|value| value.parse::<f64>().ok())
            .map(|value| (value / 100_000.0).clamp(0.0, 1.0))
        else {
            continue;
        };
        match child.local_name() {
            "shade" => rgb = rgb.map(|channel| channel * value),
            "tint" => rgb = rgb.map(|channel| channel * value + (1.0 - value)),
            "lumMod" => rgb = adjust_luminance(rgb, |luminance| luminance * value),
            "lumOff" => rgb = adjust_luminance(rgb, |luminance| luminance + value),
            _ => {}
        }
    }
    rgb
}

fn adjust_luminance(rgb: [f64; 3], adjust: impl Fn(f64) -> f64) -> [f64; 3] {
    let (hue, saturation, luminance) = rgb_to_hsl(rgb);
    hsl_to_rgb(hue, saturation, adjust(luminance).clamp(0.0, 1.0))
}

fn rgb_to_hsl([r, g, b]: [f64; 3]) -> (f64, f64, f64) {
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let luminance = (max + min) / 2.0;
    if max == min {
        return (0.0, 0.0, luminance);
    }
    let delta = max - min;
    let saturation = if luminance > 0.5 {
        delta / (2.0 - max - min)
    } else {
        delta / (max + min)
    };
    let hue = if max == r {
        ((g - b) / delta).rem_euclid(6.0)
    } else if max == g {
        (b - r) / delta + 2.0
    } else {
        (r - g) / delta + 4.0
    } / 6.0;
    (hue, saturation, luminance)
}

fn hsl_to_rgb(hue: f64, saturation: f64, luminance: f64) -> [f64; 3] {
    if saturation == 0.0 {
        return [luminance; 3];
    }
    let q = if luminance < 0.5 {
        luminance * (1.0 + saturation)
    } else {
        luminance + saturation - luminance * saturation
    };
    let p = 2.0 * luminance - q;
    let channel = |mut t: f64| {
        t = t.rem_euclid(1.0);
        if t < 1.0 / 6.0 {
            p + (q - p) * 6.0 * t
        } else if t < 0.5 {
            q
        } else if t < 2.0 / 3.0 {
            p + (q - p) * (2.0 / 3.0 - t) * 6.0
        } else {
            p
        }
    };
    [
        channel(hue + 1.0 / 3.0),
        channel(hue),
        channel(hue - 1.0 / 3.0),
    ]
}
