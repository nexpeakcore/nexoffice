//! An inline image wider than its line's available width is placed once.
//!
//! Word keeps an oversized inline object on a single line and lets it overflow
//! the indent and the margin. The engine used to wrap it off an empty line,
//! which left an empty leading row whose run span still covered the image, so
//! the display list painted the picture twice at two different x positions.
//!
//! The geometry here is synthesized from a real report: a `wp:inline` picture
//! 686 px wide inside a 686 px text column with `left=2097 right=1982
//! hanging=1955` twips and centered justification.

use docx_layout::display_list::{DisplayList, Primitive, build_display_list_json};
use docx_layout::{clear_measure_fonts, measure_paragraph_json, register_measure_font};
use serde_json::{Value, json};

const LIBERATION: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

/// Twips to CSS pixels.
fn tw(twips: f64) -> f64 {
    twips / 15.0
}

const COLUMN_WIDTH: f64 = 686.0;
const IMAGE_WIDTH: f64 = 686.0;
const IMAGE_HEIGHT: f64 = 91.13;
const PAGE_LEFT_MARGIN: f64 = 60.0;

fn image_block(indent: Value, alignment: Option<&str>) -> Value {
    let mut attrs = json!({ "indent": indent });
    if let Some(alignment) = alignment {
        attrs["alignment"] = json!(alignment);
    }
    json!({
        "kind": "paragraph",
        "id": 0,
        "runs": [{
            "kind": "image",
            "src": "rId9",
            "width": IMAGE_WIDTH,
            "height": IMAGE_HEIGHT,
            "pmStart": 1,
            "pmEnd": 2
        }],
        "attrs": attrs,
        "pmStart": 1,
        "pmEnd": 2
    })
}

/// Measures the block through the production measurement surface.
fn measure(block: &Value) -> Value {
    clear_measure_fonts();
    let id = register_measure_font(LIBERATION).expect("fixture font registers");
    let input = json!({
        "block": block,
        "maxWidth": COLUMN_WIDTH,
        "fontChains": { "liberation sans|0|0": [id] },
        "defaults": { "fontSize": 12.0, "fontFamily": "Liberation Sans" }
    });
    let out = measure_paragraph_json(&input.to_string()).expect("paragraph measures");
    serde_json::from_str(&out).expect("measure output is JSON")
}

/// Runs the measured paragraph through the display-list builder.
fn display_list(block: Value) -> DisplayList {
    let measure = measure(&block);
    let line_count = measure["lines"].as_array().expect("lines").len();
    let input = json!({
        "measured": [{ "block": block, "measure": measure }],
        "options": {},
        "layout": {
            "pages": [{
                "size": { "w": 794.0, "h": 1123.0 },
                "margins": { "top": 60.0, "right": 48.0, "bottom": 60.0, "left": PAGE_LEFT_MARGIN },
                "number": 1,
                "fragments": [{
                    "kind": "paragraph",
                    "blockId": 0,
                    "x": PAGE_LEFT_MARGIN,
                    "y": 60.0,
                    "width": COLUMN_WIDTH,
                    "height": measure["totalHeight"].as_f64().unwrap_or(0.0),
                    "fromLine": 0,
                    "toLine": line_count,
                    "pmStart": 1,
                    "pmEnd": 2
                }]
            }]
        }
    });
    let out = build_display_list_json(&input.to_string()).expect("display list builds");
    serde_json::from_str(&out).expect("display list is JSON")
}

fn images(list: &DisplayList) -> Vec<(String, f64, f64)> {
    list.pages
        .iter()
        .flat_map(|page| page.primitives.iter())
        .filter_map(|prim| match prim {
            Primitive::Image(image) => Some((
                image.rel_id.clone(),
                image.x.as_f64().unwrap_or(f64::NAN),
                image.w.as_f64().unwrap_or(f64::NAN),
            )),
            _ => None,
        })
        .collect()
}

#[test]
fn oversized_inline_image_with_hanging_indent_paints_once() {
    let block = image_block(
        json!({ "left": tw(2097.0), "right": tw(1982.0), "hanging": tw(1955.0) }),
        Some("center"),
    );
    let measure = measure(&block);
    let lines = measure["lines"].as_array().expect("lines");
    assert_eq!(
        lines.len(),
        1,
        "the picture owns exactly one line, with no empty row before it: {measure}"
    );

    let drawn = images(&display_list(block));
    assert_eq!(
        drawn.len(),
        1,
        "the picture is painted once, not once per line: {drawn:?}"
    );
    let (rel_id, x, width) = &drawn[0];
    assert_eq!(rel_id, "rId9");
    assert!((width - IMAGE_WIDTH).abs() < 0.5, "authored width is kept");
    // The first line hangs back to left − hanging, and centering cannot shift
    // an object wider than the line, so the picture starts at the hang.
    let expected_x = PAGE_LEFT_MARGIN + tw(2097.0) - tw(1955.0);
    assert!(
        (x - expected_x).abs() < 0.5,
        "picture starts at the hanging-indent x: expected {expected_x}, got {x}"
    );
}

/// The painter's own guard, pinned without the measurer: a line that ends
/// exactly where an image run starts spans that run but paints nothing.
#[test]
fn a_line_ending_at_an_image_run_does_not_paint_it() {
    let block = image_block(json!({ "left": 200.0, "right": 200.0 }), None);
    let input = json!({
        "measured": [{
            "block": block,
            "measure": {
                "kind": "paragraph",
                "totalHeight": 110.0,
                "lines": [
                    { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 0,
                      "width": 0.0, "ascent": 12.8, "descent": 3.2, "lineHeight": 18.4 },
                    { "headRun": 0, "headChar": 0, "tailRun": 0, "tailChar": 1,
                      "width": IMAGE_WIDTH, "ascent": 88.0, "descent": 3.2, "lineHeight": 91.2 }
                ]
            }
        }],
        "options": {},
        "layout": {
            "pages": [{
                "size": { "w": 794.0, "h": 1123.0 },
                "margins": { "top": 60.0, "right": 48.0, "bottom": 60.0, "left": PAGE_LEFT_MARGIN },
                "number": 1,
                "fragments": [{
                    "kind": "paragraph",
                    "blockId": 0,
                    "x": PAGE_LEFT_MARGIN,
                    "y": 60.0,
                    "width": COLUMN_WIDTH,
                    "height": 110.0,
                    "fromLine": 0,
                    "toLine": 2,
                    "pmStart": 1,
                    "pmEnd": 2
                }]
            }]
        }
    });
    let out = build_display_list_json(&input.to_string()).expect("display list builds");
    let list: DisplayList = serde_json::from_str(&out).expect("display list is JSON");
    assert_eq!(
        images(&list).len(),
        1,
        "the empty leading row must not repeat the picture"
    );
}

#[test]
fn oversized_inline_image_in_a_narrow_indent_paints_once() {
    let block = image_block(json!({ "left": 200.0, "right": 200.0 }), None);
    let measure = measure(&block);
    assert_eq!(
        measure["lines"].as_array().expect("lines").len(),
        1,
        "no empty leading row: {measure}"
    );

    let drawn = images(&display_list(block));
    assert_eq!(drawn.len(), 1, "the picture is painted once: {drawn:?}");
    let (_, x, width) = &drawn[0];
    assert!((width - IMAGE_WIDTH).abs() < 0.5, "authored width is kept");
    assert!(
        (x - (PAGE_LEFT_MARGIN + 200.0)).abs() < 0.5,
        "picture starts at the left indent and overflows it: {x}"
    );
}
