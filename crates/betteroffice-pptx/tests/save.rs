use betteroffice_pptx::{
    DeckSnapshot, EditCtx, Error, Presentation, SaveFault, ShapeSnapshot, StorySnapshot, TextStyle,
    TextStylePatch,
};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn parts(bytes: &[u8]) -> Vec<(String, Vec<u8>)> {
    ooxml_opc::unzip_parts(bytes).unwrap()
}

fn context() -> EditCtx {
    EditCtx::local("save-test")
}

fn stories(snapshot: &DeckSnapshot) -> Vec<StorySnapshot> {
    fn walk(shape: &ShapeSnapshot, output: &mut Vec<StorySnapshot>) {
        output.extend(shape.text_stories.iter().cloned());
        for child in &shape.children {
            walk(child, output);
        }
    }
    let mut output = Vec::new();
    for slide in &snapshot.slides {
        for shape in &slide.shapes {
            walk(shape, &mut output);
        }
    }
    output
}

/// The story holding `needle`, with the style of the run that holds it, which
/// is what an editor carries at the caret.
fn story_holding(presentation: &Presentation, needle: &str) -> (String, TextStyle) {
    let snapshot = presentation.snapshot().unwrap();
    for story in stories(&snapshot) {
        for paragraph in &story.paragraphs {
            for run in &paragraph.runs {
                if run.text.contains(needle) {
                    return (story.id.clone(), run.style.clone());
                }
            }
        }
    }
    panic!("no story holds {needle:?}");
}

fn plain_text(presentation: &Presentation) -> String {
    stories(&presentation.snapshot().unwrap())
        .iter()
        .map(StorySnapshot::plain_text)
        .collect()
}

#[test]
fn an_untouched_save_reproduces_every_source_part() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    assert_eq!(parts(&presentation.save().unwrap()), parts(FIXTURE));
}

#[test]
fn a_text_edit_rewrites_only_its_own_run_and_leaves_every_other_part_alone() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let (story_id, style) = story_holding(&presentation, "A Rust-native");
    presentation
        .insert_text(&context(), &story_id, 0, "Now: ", &style)
        .unwrap();

    let saved = presentation.save().unwrap();
    let before = parts(FIXTURE);
    let after = parts(&saved);
    assert_eq!(before.len(), after.len());
    for ((source_path, source_bytes), (saved_path, saved_bytes)) in before.iter().zip(&after) {
        assert_eq!(source_path, saved_path);
        if source_path == "ppt/slides/slide1.xml" {
            let source = String::from_utf8(source_bytes.clone()).unwrap();
            let saved = String::from_utf8(saved_bytes.clone()).unwrap();
            assert_eq!(
                saved,
                source.replace("<a:t>A Rust-native", "<a:t>Now: A Rust-native"),
                "the edited slide differs only inside the edited <a:t>"
            );
        } else {
            assert_eq!(source_bytes, saved_bytes, "{source_path} was rewritten");
        }
    }

    let reopened = Presentation::open(&saved).unwrap();
    assert!(plain_text(&reopened).contains("Now: A Rust-native editing engine"));
    assert_eq!(reopened.slides().len(), 3);
    assert_eq!(
        stories(&reopened.snapshot().unwrap()).len(),
        stories(&presentation.snapshot().unwrap()).len()
    );
}

#[test]
fn a_table_cell_edit_reaches_the_graphic_frame() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let table_story = stories(&presentation.snapshot().unwrap())
        .into_iter()
        .find(|story| story.id.contains(":table:"))
        .expect("the fixture has a table");
    let style = table_story.paragraphs[0].runs[0].style.clone();
    let before = table_story.plain_text();

    presentation
        .insert_text(&context(), &table_story.id, 0, "> ", &style)
        .unwrap();
    let saved = presentation.save().unwrap();

    let reopened = Presentation::open(&saved).unwrap();
    let reparsed = stories(&reopened.snapshot().unwrap())
        .into_iter()
        .find(|story| story.id == table_story.id)
        .expect("the cell survived the save");
    assert_eq!(reparsed.plain_text(), format!("> {before}"));
}

#[test]
fn an_edit_inside_one_run_of_a_multi_run_paragraph_is_expressible() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let (story_id, _) = story_holding(&presentation, "01");
    let paragraph = presentation.story(&story_id).unwrap().paragraphs.remove(0);
    assert!(
        paragraph.runs.len() > 1,
        "the fixture paragraph has several runs"
    );
    let head = paragraph.runs[0].text.chars().count() as u32;

    presentation
        .insert_text(
            &context(),
            &story_id,
            head + 1,
            "!",
            &paragraph.runs[1].style,
        )
        .unwrap();

    let saved = presentation.save().unwrap();
    let reopened = Presentation::open(&saved).unwrap();
    let reparsed = presentation.story(&story_id).unwrap();
    assert_eq!(
        stories(&reopened.snapshot().unwrap())
            .into_iter()
            .find(|story| story.id == story_id)
            .unwrap()
            .plain_text(),
        reparsed.plain_text()
    );
    assert_eq!(reparsed.paragraphs[0].runs.len(), paragraph.runs.len());
}

#[test]
fn one_save_carries_edits_from_two_paragraphs_and_two_slides() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let (title, first_style) = story_holding(&presentation, "Office files,");
    let (_, second_style) = story_holding(&presentation, "without the office.");
    let (heading, heading_style) = story_holding(&presentation, "ONE NATIVE STACK");
    assert_ne!(title, heading, "the two edits land on different slides");
    let second_line = presentation
        .story(&title)
        .unwrap()
        .plain_text()
        .find("without the office.")
        .expect("both paragraphs live in one body") as u32;

    presentation
        .insert_text(&context(), &title, 0, "Now: ", &first_style)
        .unwrap();
    presentation
        .insert_text(
            &context(),
            &title,
            second_line + "Now: ".len() as u32,
            "just ",
            &second_style,
        )
        .unwrap();
    presentation
        .insert_text(&context(), &heading, 0, "TOP ", &heading_style)
        .unwrap();

    let saved = presentation.save().unwrap();
    let before = parts(FIXTURE);
    let after = parts(&saved);
    assert_eq!(before.len(), after.len());
    for ((source_path, source_bytes), (saved_path, saved_bytes)) in before.iter().zip(&after) {
        assert_eq!(source_path, saved_path);
        let edited = match source_path.as_str() {
            "ppt/slides/slide1.xml" => Some(
                String::from_utf8(source_bytes.clone())
                    .unwrap()
                    .replace("<a:t>Office files,", "<a:t>Now: Office files,")
                    .replace("<a:t>without the office.", "<a:t>just without the office."),
            ),
            "ppt/slides/slide2.xml" => Some(
                String::from_utf8(source_bytes.clone())
                    .unwrap()
                    .replace("<a:t>ONE NATIVE STACK", "<a:t>TOP ONE NATIVE STACK"),
            ),
            _ => None,
        };
        match edited {
            Some(expected) => assert_eq!(
                String::from_utf8(saved_bytes.clone()).unwrap(),
                expected,
                "{source_path} differs outside the runs the edits named"
            ),
            None => assert_eq!(source_bytes, saved_bytes, "{source_path} was rewritten"),
        }
    }

    let reopened = Presentation::open(&saved).unwrap();
    let text = plain_text(&reopened);
    assert!(text.contains("Now: Office files,\njust without the office."));
    assert!(text.contains("TOP ONE NATIVE STACK"));
}

/// The `<a:pPr>` and `<a:rPr>` of the demo deck's title paragraphs, which a
/// split has to copy verbatim.
const TITLE_PPR: &str = r#"<a:pPr algn="l" lvl="0"></a:pPr>"#;
const TITLE_RPR: &str = concat!(
    r#"<a:rPr lang="en-US" sz="4400" b="1">"#,
    r#"<a:solidFill><a:srgbClr val="101828"/></a:solidFill>"#,
    r#"<a:latin typeface="Arial"/></a:rPr>"#,
);

/// Every part of `saved` against `FIXTURE`, with `slide1` compared to what the
/// edit should have made of it and every other part to its own source bytes.
fn assert_only_slide_one_changed(saved: &[u8], expected_slide_one: &str) {
    let before = parts(FIXTURE);
    let after = parts(saved);
    assert_eq!(before.len(), after.len());
    for ((source_path, source_bytes), (saved_path, saved_bytes)) in before.iter().zip(&after) {
        assert_eq!(source_path, saved_path);
        if source_path == "ppt/slides/slide1.xml" {
            assert_eq!(
                String::from_utf8(saved_bytes.clone()).unwrap(),
                expected_slide_one,
                "the edited slide differs outside the structure the edit changed"
            );
        } else {
            assert_eq!(source_bytes, saved_bytes, "{source_path} was rewritten");
        }
    }
}

fn slide_one() -> String {
    String::from_utf8(
        parts(FIXTURE)
            .into_iter()
            .find(|(path, _)| path == "ppt/slides/slide1.xml")
            .expect("the demo deck has a first slide")
            .1,
    )
    .unwrap()
}

#[test]
fn a_paragraph_split_rewrites_only_the_slide_it_happened_on() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let (story_id, _) = story_holding(&presentation, "Office files,");
    presentation
        .insert_paragraph_break(&context(), &story_id, 7)
        .unwrap();

    let saved = presentation.save().unwrap();
    assert_only_slide_one_changed(
        &saved,
        &slide_one().replace(
            &format!("<a:r>{TITLE_RPR}<a:t>Office files,</a:t></a:r>"),
            &format!(
                "<a:r>{TITLE_RPR}<a:t>Office </a:t></a:r></a:p><a:p>{TITLE_PPR}\
                 <a:r>{TITLE_RPR}<a:t>files,</a:t></a:r>"
            ),
        ),
    );

    let reopened = Presentation::open(&saved).unwrap();
    assert!(plain_text(&reopened).contains("Office \nfiles,\nwithout the office."));
    assert_eq!(reopened.slides().len(), 3);
    assert_eq!(
        reopened.story(&story_id).unwrap().paragraphs.len(),
        3,
        "the title body now holds three paragraphs"
    );
}

#[test]
fn a_paragraph_merge_rewrites_only_the_slide_it_happened_on() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let (story_id, _) = story_holding(&presentation, "Office files,");
    let break_index = presentation
        .story(&story_id)
        .unwrap()
        .plain_text()
        .find('\n')
        .expect("both title lines live in one body") as u32;
    presentation
        .delete_paragraph_break(&context(), &story_id, break_index)
        .unwrap();

    let saved = presentation.save().unwrap();
    assert_only_slide_one_changed(
        &saved,
        &slide_one().replace(
            &format!(
                r#"<a:endParaRPr lang="en-US" sz="1800"/></a:p><a:p>{TITLE_PPR}<a:r><a:rPr lang="en-US" sz="4400" b="1"><a:solidFill><a:srgbClr val="315EFB"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>without the office.</a:t></a:r>"#
            ),
            r#"<a:r><a:rPr lang="en-US" sz="4400" b="1"><a:solidFill><a:srgbClr val="315EFB"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>without the office.</a:t></a:r>"#,
        ),
    );

    let reopened = Presentation::open(&saved).unwrap();
    assert!(plain_text(&reopened).contains("Office files,without the office."));
    let paragraphs = reopened.story(&story_id).unwrap().paragraphs;
    assert_eq!(paragraphs.len(), 1);
    assert_eq!(paragraphs[0].runs.len(), 2, "each line keeps its own run");
}

#[test]
fn a_line_break_rewrites_only_the_run_it_divides() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let (story_id, style) = story_holding(&presentation, "Office files,");
    presentation
        .insert_text(&context(), &story_id, 7, "\n", &style)
        .unwrap();

    let saved = presentation.save().unwrap();
    assert_only_slide_one_changed(
        &saved,
        &slide_one().replace(
            &format!("<a:r>{TITLE_RPR}<a:t>Office files,</a:t></a:r>"),
            &format!(
                "<a:r>{TITLE_RPR}<a:t>Office </a:t></a:r><a:br>{}</a:br>\
                 <a:r>{TITLE_RPR}<a:t>files,</a:t></a:r>",
                TITLE_RPR
            ),
        ),
    );

    let reopened = Presentation::open(&saved).unwrap();
    assert!(plain_text(&reopened).contains("Office \nfiles,\nwithout the office."));
    assert_eq!(
        reopened.story(&story_id).unwrap().paragraphs.len(),
        2,
        "a soft break keeps the two paragraphs the deck had"
    );
}

type Refusal = (&'static str, Box<dyn Fn(&Presentation)>);

#[test]
fn a_change_this_slice_cannot_write_refuses_instead_of_dropping_it() {
    let refusals: Vec<Refusal> = vec![
        (
            "two paragraph breaks in different runs",
            Box::new(|presentation: &Presentation| {
                let (story_id, _) = story_holding(presentation, "01");
                presentation
                    .insert_paragraph_break(&context(), &story_id, 1)
                    .unwrap();
                presentation
                    .insert_paragraph_break(&context(), &story_id, 6)
                    .unwrap();
            }),
        ),
        (
            "a new slide",
            Box::new(|presentation: &Presentation| {
                presentation.insert_slide(&context(), 0, None).unwrap();
            }),
        ),
    ];

    for (label, edit) in refusals {
        let presentation = Presentation::open(FIXTURE).unwrap();
        edit(&presentation);
        match presentation.save() {
            // The fault, not the wording, is what a caller acts on: a broken
            // write and a blown budget also arrive as `Error::Edit`, and only
            // this one means undoing the change gets the save through.
            Err(Error::Edit(error)) => {
                assert_eq!(
                    error.save_fault(),
                    SaveFault::Unprojectable,
                    "{label} ended with {error}"
                );
                assert!(error.save_fault().undoing_helps(), "{label}");
            }
            Err(other) => panic!("{label} failed with {other}"),
            Ok(_) => panic!("{label} was saved instead of refused"),
        }
    }
}

#[test]
fn taking_back_an_unwritable_change_restores_a_savable_deck() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let inserted = presentation.insert_slide(&context(), 0, None).unwrap();
    assert!(presentation.save().is_err());

    presentation
        .delete_slide(&context(), &inserted.slide_id)
        .unwrap();
    assert_eq!(parts(&presentation.save().unwrap()), parts(FIXTURE));
}

/// A formatting patch over part of a run is spelled into the file: the run is
/// split, the styled half carries the patch, and only that slide changes.
#[test]
fn a_mid_run_formatting_patch_saves_and_reads_back() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let (story_id, _) = story_holding(&presentation, "A Rust-native");
    presentation
        .format_text(
            &context(),
            &story_id,
            2,
            9,
            &TextStylePatch {
                bold: Some(true),
                color: Some("#325ee6".to_owned()),
                ..TextStylePatch::default()
            },
        )
        .unwrap();

    let saved = presentation.save().unwrap();
    let (before, after) = (parts(FIXTURE), parts(&saved));
    let changed: Vec<&str> = before
        .iter()
        .zip(&after)
        .filter(|(source, written)| source != written)
        .map(|(source, _)| source.0.as_str())
        .collect();
    assert_eq!(
        changed.len(),
        1,
        "exactly one slide part changed: {changed:?}"
    );

    let reopened = Presentation::open(&saved).unwrap();
    let reopened_snapshot = reopened.snapshot().unwrap();
    let styled: Vec<_> = reopened_snapshot
        .slides
        .iter()
        .flat_map(|slide| &slide.shapes)
        .flat_map(|shape| &shape.text_stories)
        .flat_map(|story| &story.paragraphs)
        .flat_map(|paragraph| &paragraph.runs)
        .filter(|run| run.style.bold == Some(true) && run.style.color.as_deref() == Some("#325EE6"))
        .collect();
    assert_eq!(styled.len(), 1, "one styled split run: {styled:?}");
    assert_eq!(styled[0].text.chars().count(), 7);
}

/// A pure move is no longer a refusal: the shape's own `<a:xfrm>` is spliced
/// and every untouched part keeps its source bytes.
#[test]
fn a_moved_shape_saves_with_every_other_part_untouched() {
    let presentation = Presentation::open(FIXTURE).unwrap();
    let snapshot = presentation.snapshot().unwrap();
    presentation
        .move_shape(
            &context(),
            &snapshot.slides[0].id,
            &snapshot.slides[0].shapes[0].id,
            10,
            20,
        )
        .unwrap();

    let saved = presentation.save().unwrap();
    let (before, after) = (parts(FIXTURE), parts(&saved));
    assert_eq!(before.len(), after.len());
    let changed: Vec<&str> = before
        .iter()
        .zip(&after)
        .filter(|(source, written)| source != written)
        .map(|(source, _)| source.0.as_str())
        .collect();
    assert_eq!(changed, ["ppt/slides/slide1.xml"]);

    let reopened = Presentation::open(&saved).unwrap();
    let shape = &reopened.snapshot().unwrap().slides[0].shapes[0];
    assert_eq!((shape.x, shape.y), (10, 20));
}
