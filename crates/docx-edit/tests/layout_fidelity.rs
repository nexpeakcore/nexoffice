//! End-to-end layout-fidelity pins for run patterns that real documents use:
//! whitespace-only `xml:space="preserve"` runs carrying `w:spacing` tracking,
//! and long table-cell/tabbed paragraphs that must wrap. Both are exercised
//! through the product pipeline (parse → yrs seed → render bridge → measure)
//! with and without registered measurement fonts — the fontless path is the
//! deterministic synthetic fallback hosts hit when no font provider is wired.

use docx_edit::EditingDoc;
use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::seed_from_docx;
use docx_layout::measure_blocks::{MeasurementConfig, collect_font_requirements, measure_blocks};
use docx_layout::types::{BlockExtent, LayoutBlock, ParagraphExtent, Run};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn docx_from_body(body_xml: &str) -> Vec<u8> {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>{body_xml}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#
    );
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        ("word/document.xml".to_owned(), document.into_bytes()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn lowered_blocks(body_xml: &str) -> Vec<LayoutBlock> {
    let bytes = docx_from_body(body_xml);
    let doc = EditingDoc::new(1);
    seed_from_docx(&doc, &bytes).expect("seed");
    yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default()).expect("lower")
}

fn paragraph_text(paragraph: &docx_layout::types::ParagraphBlock) -> String {
    let mut text = String::new();
    for run in &paragraph.runs {
        match run {
            Run::Text(text_run) => text.push_str(&text_run.text),
            Run::Tab(_) => text.push('\t'),
            _ => {}
        }
    }
    text
}

fn fontless_config() -> MeasurementConfig {
    docx_layout::clear_measure_fonts();
    MeasurementConfig {
        authoritative_shaping: true,
        ..Default::default()
    }
}

fn liberation_config(blocks: &[LayoutBlock]) -> MeasurementConfig {
    docx_layout::clear_measure_fonts();
    let font_id = docx_layout::register_measure_font(FONT).expect("font registers");
    let mut config = MeasurementConfig {
        defaults: serde_json::json!({ "fontSize": 11.0, "fontFamily": "Liberation Sans" }),
        authoritative_shaping: true,
        ..Default::default()
    };
    for requirement in collect_font_requirements(blocks) {
        config.font_chains.insert(requirement.key, vec![font_id]);
    }
    config
}

const HEADING_BODY: &str = r#"<w:p w14:paraId="AAAA0001"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr><w:t>AAAA</w:t></w:r><w:r><w:rPr><w:b/><w:spacing w:val="-6"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr><w:t>BBB</w:t></w:r><w:r><w:rPr><w:b/><w:spacing w:val="-2"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:b/><w:sz w:val="20"/></w:rPr><w:t>CC</w:t></w:r></w:p>"#;

fn assert_space_runs_have_positive_advance(extent: &ParagraphExtent) {
    assert_eq!(extent.lines.len(), 1);
    let runs = extent.lines[0].run_advances.as_ref().expect("run advances");
    let mut space_runs = 0;
    for (index, run) in runs.iter().enumerate() {
        if index % 2 == 1 {
            space_runs += 1;
            assert!(
                run.advance.unwrap_or(0.0) > 0.0,
                "space-only run {index} lost its advance"
            );
        }
    }
    assert_eq!(space_runs, 2);
}

#[test]
fn space_only_runs_with_tracking_survive_to_layout() {
    let mut blocks = lowered_blocks(HEADING_BODY);
    let [LayoutBlock::Paragraph(paragraph)] = blocks.as_slice() else {
        panic!("one paragraph expected");
    };
    assert_eq!(paragraph_text(paragraph), "AAAA BBB CC");
    assert_eq!(paragraph.runs.len(), 5, "space-only runs stay distinct");

    let config = liberation_config(&blocks);
    let extents = measure_blocks(&mut blocks, 500.0, &config).expect("measure with fonts");
    let BlockExtent::Paragraph(extent) = &extents[0] else {
        panic!("paragraph extent expected");
    };
    assert_space_runs_have_positive_advance(extent);
}

#[test]
fn space_only_runs_survive_fontless_fallback() {
    let mut blocks = lowered_blocks(HEADING_BODY);
    let config = fontless_config();
    let extents = measure_blocks(&mut blocks, 500.0, &config).expect("measure without fonts");
    let BlockExtent::Paragraph(extent) = &extents[0] else {
        panic!("paragraph extent expected");
    };
    assert_space_runs_have_positive_advance(extent);
}

const CELL_BODY: &str = r#"<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1705"/><w:gridCol w:w="7928"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="1705" w:type="dxa"/></w:tcPr><w:p w14:paraId="BBBB0001"><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>Label:</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="7928" w:type="dxa"/></w:tcPr><w:p w14:paraId="BBBB0002"><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>123456789 - Aaaaa bbbbb Ccccccc ddddd Eeeee ffffff G</w:t></w:r><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>h</w:t></w:r><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>Iiiiiii jjjj Kkkk (Llllllllll) - MM Nnn Oooooooo pppp qqqq</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="BBBB0003"/>"#;

fn cell_paragraph_extent(blocks: &[LayoutBlock], extents: &[BlockExtent]) -> ParagraphExtent {
    for (block, extent) in blocks.iter().zip(extents) {
        let (LayoutBlock::Table(table), BlockExtent::Table(table_extent)) = (block, extent) else {
            continue;
        };
        let cell = &table.rows[0].cells[1];
        let cell_extent = &table_extent.rows[0].cells[1];
        for (cell_block, cell_block_extent) in cell.blocks.iter().zip(&cell_extent.blocks) {
            let LayoutBlock::Paragraph(_) = cell_block else {
                continue;
            };
            let BlockExtent::Paragraph(paragraph_extent) = cell_block_extent else {
                panic!("paragraph extent expected");
            };
            return paragraph_extent.clone();
        }
    }
    panic!("cell paragraph not measured");
}

fn assert_cell_paragraph_wraps(config: MeasurementConfig) {
    let mut blocks = lowered_blocks(CELL_BODY);
    let extents = measure_blocks(&mut blocks, 643.0, &config).expect("measure");
    let extent = cell_paragraph_extent(&blocks, &extents);
    // 7928 twips ≈ 528.5px.
    let cell_text_width = 7928.0 / 15.0;
    assert!(
        extent.lines.len() >= 2,
        "long cell paragraph must wrap, got {} line(s)",
        extent.lines.len()
    );
    for line in &extent.lines {
        assert!(
            line.width <= cell_text_width + 8.0,
            "line width {} exceeds the cell text width {cell_text_width}",
            line.width
        );
    }
}

#[test]
fn table_cell_paragraph_wraps_with_fonts() {
    let blocks = lowered_blocks(CELL_BODY);
    assert_cell_paragraph_wraps(liberation_config(&blocks));
}

#[test]
fn table_cell_paragraph_wraps_in_fontless_fallback() {
    assert_cell_paragraph_wraps(fontless_config());
}

const TABBED_BODY: &str = r#"<w:p w14:paraId="CCCC0001"><w:pPr><w:tabs><w:tab w:val="left" w:pos="743"/></w:tabs></w:pPr><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:tab/><w:t xml:space="preserve">Wwwww wwwww wwwww wwwww wwwww wwwww wwwww wwwww wwwww wwwww wwwww wwwww wwwww wwwww</w:t></w:r></w:p>"#;

fn assert_tabbed_paragraph_wraps(config: MeasurementConfig) {
    let mut blocks = lowered_blocks(TABBED_BODY);
    let width = 300.0;
    let extents = measure_blocks(&mut blocks, width, &config).expect("measure");
    let BlockExtent::Paragraph(extent) = &extents[0] else {
        panic!("paragraph extent expected");
    };
    assert!(
        extent.lines.len() >= 2,
        "tabbed paragraph must wrap, got {} line(s)",
        extent.lines.len()
    );
    for line in &extent.lines {
        assert!(
            line.width <= width + 8.0,
            "line width {} exceeds the wrap width {width}",
            line.width
        );
    }
}

#[test]
fn tabbed_paragraph_wraps_with_fonts() {
    let blocks = lowered_blocks(TABBED_BODY);
    assert_tabbed_paragraph_wraps(liberation_config(&blocks));
}

#[test]
fn tabbed_paragraph_wraps_in_fontless_fallback() {
    assert_tabbed_paragraph_wraps(fontless_config());
}
