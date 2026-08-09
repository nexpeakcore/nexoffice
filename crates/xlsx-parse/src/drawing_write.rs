//! Serializes in-session-created charts into new drawing and chart parts.
//! Preserved drawings never pass through here — their source bytes ride.

use std::collections::HashSet;

use ooxml_drawingml::chart::{ChartSeries, ChartSpace};
use quick_xml::escape::escape;
use xlsx_model::workbook::{AnchorCell, DrawingAnchor, SheetDrawing};

use crate::ParseError;

pub(crate) const REL_DRAWING: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
pub(crate) const CT_DRAWING: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
pub(crate) const CT_CHART: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

const NS_PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const XML_DECL: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;

/// The namespace and relationship-type family the emitted parts join:
/// Transitional for ordinary packages, ISO Strict when the source is strict.
#[derive(Clone, Copy)]
pub(crate) struct DrawingNamespaces {
    xdr: &'static str,
    a: &'static str,
    c: &'static str,
    r: &'static str,
    chart_relationship: &'static str,
}

pub(crate) const TRANSITIONAL: DrawingNamespaces = DrawingNamespaces {
    xdr: "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    a: "http://schemas.openxmlformats.org/drawingml/2006/main",
    c: "http://schemas.openxmlformats.org/drawingml/2006/chart",
    r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    chart_relationship: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
};

pub(crate) const STRICT: DrawingNamespaces = DrawingNamespaces {
    xdr: "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing",
    a: "http://purl.oclc.org/ooxml/drawingml/main",
    c: "http://purl.oclc.org/ooxml/drawingml/chart",
    r: "http://purl.oclc.org/ooxml/officeDocument/relationships",
    chart_relationship: "http://purl.oclc.org/ooxml/officeDocument/relationships/chart",
};

const CAT_AXIS_ID: &str = "100000001";
const VAL_AXIS_ID: &str = "100000002";

/// New parts for one sheet's created drawings, plus what the caller must
/// register: the worksheet-level relationship target and content types.
pub(crate) struct EmittedDrawing {
    pub drawing_path: String,
    pub parts: Vec<(String, Vec<u8>)>,
    pub overrides: Vec<(String, &'static str)>,
}

/// Serializes `drawings` into a fresh drawing part with one chart part per
/// drawing. `None` when there is nothing to emit.
pub(crate) fn emit_sheet_drawings(
    drawings: &[SheetDrawing],
    used_paths: &mut HashSet<String>,
    namespaces: DrawingNamespaces,
) -> Result<Option<EmittedDrawing>, ParseError> {
    if drawings.is_empty() {
        return Ok(None);
    }
    let drawing_path = next_free_path(used_paths, |index| {
        format!("xl/drawings/drawing{index}.xml")
    });
    let chart_paths = drawings
        .iter()
        .map(|_| next_free_path(used_paths, |index| format!("xl/charts/chart{index}.xml")))
        .collect::<Vec<_>>();

    let mut anchors = String::new();
    let mut chart_relationships = String::new();
    let mut parts = Vec::new();
    let mut overrides = vec![(drawing_path.clone(), CT_DRAWING)];
    for (index, (drawing, chart_path)) in drawings.iter().zip(&chart_paths).enumerate() {
        let rel_id = format!("rId{}", index + 1);
        anchors.push_str(&anchor_xml(drawing, index, &rel_id, namespaces)?);
        let target = chart_path
            .strip_prefix("xl/")
            .map(|path| format!("../{path}"))
            .unwrap_or_else(|| chart_path.clone());
        chart_relationships.push_str(&format!(
            r#"<Relationship Id="{rel_id}" Type="{}" Target="{}"/>"#,
            namespaces.chart_relationship,
            escape(&target)
        ));
        parts.push((chart_path.clone(), chart_xml(&drawing.chart, namespaces)?));
        overrides.push((chart_path.clone(), CT_CHART));
    }

    parts.insert(
        0,
        (
            drawing_path.clone(),
            format!(
                r#"{XML_DECL}<xdr:wsDr xmlns:xdr="{}" xmlns:a="{}">{anchors}</xdr:wsDr>"#,
                namespaces.xdr, namespaces.a
            )
            .into_bytes(),
        ),
    );
    parts.insert(
        1,
        (
            relationship_part_path(&drawing_path),
            format!(
                r#"{XML_DECL}<Relationships xmlns="{NS_PKG_REL}">{chart_relationships}</Relationships>"#
            )
            .into_bytes(),
        ),
    );
    Ok(Some(EmittedDrawing {
        drawing_path,
        parts,
        overrides,
    }))
}

fn anchor_xml(
    drawing: &SheetDrawing,
    index: usize,
    rel_id: &str,
    namespaces: DrawingNamespaces,
) -> Result<String, ParseError> {
    let shape_id = index + 2;
    let name = escape(drawing.chart.title.as_deref().unwrap_or("Chart")).into_owned();
    let frame = format!(
        r#"<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{shape_id}" name="{name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="{}"><c:chart xmlns:c="{}" xmlns:r="{}" r:id="{rel_id}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/>"#,
        namespaces.c, namespaces.c, namespaces.r
    );
    let cell = |tag: &str, at: &AnchorCell| {
        format!(
            "<xdr:{tag}><xdr:col>{}</xdr:col><xdr:colOff>{}</xdr:colOff><xdr:row>{}</xdr:row><xdr:rowOff>{}</xdr:rowOff></xdr:{tag}>",
            at.col, at.col_offset_emu, at.row, at.row_offset_emu
        )
    };
    Ok(match &drawing.anchor {
        DrawingAnchor::Cell {
            from, to: Some(to), ..
        } => format!(
            "<xdr:twoCellAnchor>{}{}{frame}</xdr:twoCellAnchor>",
            cell("from", from),
            cell("to", to)
        ),
        DrawingAnchor::Cell {
            from,
            to: None,
            extent_emu: Some((cx, cy)),
        } => format!(
            r#"<xdr:oneCellAnchor>{}<xdr:ext cx="{cx}" cy="{cy}"/>{frame}</xdr:oneCellAnchor>"#,
            cell("from", from)
        ),
        DrawingAnchor::Cell {
            to: None,
            extent_emu: None,
            ..
        } => {
            return Err(ParseError::Malformed(
                "chart anchor has neither a to-cell nor an extent".to_owned(),
            ));
        }
        DrawingAnchor::Absolute {
            pos_emu,
            extent_emu,
        } => format!(
            r#"<xdr:absoluteAnchor><xdr:pos x="{}" y="{}"/><xdr:ext cx="{}" cy="{}"/>{frame}</xdr:absoluteAnchor>"#,
            pos_emu.0, pos_emu.1, extent_emu.0, extent_emu.1
        ),
    })
}

fn chart_xml(chart: &ChartSpace, namespaces: DrawingNamespaces) -> Result<Vec<u8>, ParseError> {
    let plot = plot_element(chart)?;
    let title = chart
        .title
        .as_deref()
        .map(|title| {
            format!(
                r#"<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>{}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>"#,
                escape(title)
            )
        })
        .unwrap_or_default();
    let legend = chart
        .legend
        .as_ref()
        .filter(|legend| legend.visible)
        .map(|legend| {
            format!(
                r#"<c:legend><c:legendPos val="{}"/><c:overlay val="0"/></c:legend>"#,
                escape(legend.position.as_deref().unwrap_or("b"))
            )
        })
        .unwrap_or_default();
    Ok(format!(
        r#"{XML_DECL}<c:chartSpace xmlns:c="{}" xmlns:a="{}" xmlns:r="{}"><c:chart>{title}<c:plotArea><c:layout/>{plot}</c:plotArea>{legend}<c:plotVisOnly val="1"/></c:chart></c:chartSpace>"#,
        namespaces.c, namespaces.a, namespaces.r
    )
    .into_bytes())
}

fn plot_element(chart: &ChartSpace) -> Result<String, ParseError> {
    let series = chart
        .series
        .iter()
        .enumerate()
        .map(|(index, series)| series_xml(series, index, &chart.chart_type))
        .collect::<Result<String, ParseError>>()?;
    Ok(match chart.chart_type.as_str() {
        "column" | "bar" => {
            let direction = if chart.chart_type == "bar" {
                "bar"
            } else {
                "col"
            };
            format!(
                r#"<c:barChart><c:barDir val="{direction}"/><c:grouping val="clustered"/><c:varyColors val="0"/>{series}{}</c:barChart>{}"#,
                axis_ids(),
                axes()
            )
        }
        "line" => format!(
            r#"<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>{series}<c:marker val="1"/>{}</c:lineChart>{}"#,
            axis_ids(),
            axes()
        ),
        "area" => format!(
            r#"<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>{series}{}</c:areaChart>{}"#,
            axis_ids(),
            axes()
        ),
        "pie" => format!(r#"<c:pieChart><c:varyColors val="1"/>{series}</c:pieChart>"#),
        "doughnut" => format!(
            r#"<c:doughnutChart><c:varyColors val="1"/>{series}<c:holeSize val="50"/></c:doughnutChart>"#
        ),
        other => {
            return Err(ParseError::Malformed(format!(
                "unsupported chart type for creation: {other}"
            )));
        }
    })
}

fn axis_ids() -> String {
    format!(r#"<c:axId val="{CAT_AXIS_ID}"/><c:axId val="{VAL_AXIS_ID}"/>"#)
}

fn axes() -> String {
    format!(
        r#"<c:catAx><c:axId val="{CAT_AXIS_ID}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="{VAL_AXIS_ID}"/></c:catAx><c:valAx><c:axId val="{VAL_AXIS_ID}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="{CAT_AXIS_ID}"/></c:valAx>"#
    )
}

fn series_xml(series: &ChartSeries, index: usize, chart_type: &str) -> Result<String, ParseError> {
    let name = series
        .name
        .as_deref()
        .map(|name| format!("<c:tx><c:v>{}</c:v></c:tx>", escape(name)))
        .unwrap_or_default();
    let color = series
        .color
        .strip_prefix('#')
        .unwrap_or(&series.color)
        .to_uppercase();
    if color.len() != 6 || !color.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(ParseError::Malformed(format!(
            "chart series color must be #RRGGBB, got {:?}",
            series.color
        )));
    }
    let properties = if chart_type == "line" {
        format!(
            r#"<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></c:spPr>"#
        )
    } else {
        format!(r#"<c:spPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></c:spPr>"#)
    };
    let categories = series
        .category_formula
        .as_deref()
        .map(|formula| {
            let points: String = series
                .categories
                .iter()
                .enumerate()
                .map(|(i, value)| format!("<c:pt idx=\"{i}\"><c:v>{}</c:v></c:pt>", escape(value)))
                .collect();
            format!(
                r#"<c:cat><c:strRef><c:f>{}</c:f><c:strCache><c:ptCount val="{}"/>{points}</c:strCache></c:strRef></c:cat>"#,
                escape(formula),
                series.categories.len()
            )
        })
        .unwrap_or_default();
    let values = series
        .value_formula
        .as_deref()
        .map(|formula| {
            let points: String = series
                .values
                .iter()
                .enumerate()
                .map(|(i, value)| format!("<c:pt idx=\"{i}\"><c:v>{value}</c:v></c:pt>"))
                .collect();
            format!(
                r#"<c:val><c:numRef><c:f>{}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="{}"/>{points}</c:numCache></c:numRef></c:val>"#,
                escape(formula),
                series.values.len()
            )
        })
        .unwrap_or_default();
    Ok(format!(
        r#"<c:ser><c:idx val="{index}"/><c:order val="{index}"/>{name}{properties}{categories}{values}</c:ser>"#
    ))
}

fn next_free_path(used_paths: &mut HashSet<String>, build: impl Fn(usize) -> String) -> String {
    for index in 1.. {
        let path = build(index);
        if used_paths.insert(path.clone()) {
            return path;
        }
    }
    unreachable!()
}

fn relationship_part_path(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((directory, filename)) => format!("{directory}/_rels/{filename}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}
