use betteroffice_pptx::{
    DeckSnapshot, EditCtx, Error, Presentation, ShapeDraft, ShapeRect, ShapeSnapshot,
    StorySnapshot, TextStyle, TextStylePatch,
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

type Refusal = (&'static str, Box<dyn Fn(&Presentation)>);

#[test]
fn a_change_this_slice_cannot_write_refuses_instead_of_dropping_it() {
    let refusals: Vec<Refusal> = vec![
        (
            "formatting",
            Box::new(|presentation: &Presentation| {
                let (story_id, _) = story_holding(presentation, "A Rust-native");
                presentation
                    .format_text(
                        &context(),
                        &story_id,
                        0,
                        4,
                        &TextStylePatch {
                            bold: Some(true),
                            ..TextStylePatch::default()
                        },
                    )
                    .unwrap();
            }),
        ),
        (
            "a new paragraph",
            Box::new(|presentation: &Presentation| {
                let (story_id, _) = story_holding(presentation, "A Rust-native");
                presentation
                    .insert_paragraph_break(&context(), &story_id, 4)
                    .unwrap();
            }),
        ),
        (
            "unstyled text inside a styled run",
            Box::new(|presentation: &Presentation| {
                let (story_id, _) = story_holding(presentation, "A Rust-native");
                presentation
                    .insert_text(&context(), &story_id, 0, "plain", &TextStyle::default())
                    .unwrap();
            }),
        ),
        (
            "a moved shape",
            Box::new(|presentation: &Presentation| {
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
            }),
        ),
        (
            "a new shape",
            Box::new(|presentation: &Presentation| {
                let snapshot = presentation.snapshot().unwrap();
                presentation
                    .add_text_box(
                        &context(),
                        &snapshot.slides[0].id,
                        &ShapeDraft {
                            name: "Added".to_owned(),
                            rect: ShapeRect {
                                x: 0,
                                y: 0,
                                width: 100_000,
                                height: 100_000,
                            },
                            text: "hello".to_owned(),
                            style: TextStyle::default(),
                        },
                    )
                    .unwrap();
            }),
        ),
        (
            "a new slide",
            Box::new(|presentation: &Presentation| {
                presentation.insert_slide(&context(), 0, None).unwrap();
            }),
        ),
        (
            "a deleted slide",
            Box::new(|presentation: &Presentation| {
                let snapshot = presentation.snapshot().unwrap();
                presentation
                    .delete_slide(&context(), &snapshot.slides[1].id)
                    .unwrap();
            }),
        ),
        (
            "a reordered deck",
            Box::new(|presentation: &Presentation| {
                let snapshot = presentation.snapshot().unwrap();
                presentation
                    .move_slide(&context(), &snapshot.slides[0].id, 2)
                    .unwrap();
            }),
        ),
    ];

    for (label, edit) in refusals {
        let presentation = Presentation::open(FIXTURE).unwrap();
        edit(&presentation);
        match presentation.save() {
            Err(Error::Edit(error)) => {
                let message = error.to_string();
                assert!(
                    message.starts_with("this deck holds a change the PPTX writer cannot save yet"),
                    "{label} refused with {message:?}"
                );
            }
            Err(other) => panic!("{label} failed with {other}"),
            Ok(_) => panic!("{label} was saved instead of refused"),
        }
    }
}

#[test]
fn undoing_an_unwritable_edit_restores_a_savable_deck() {
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
    assert!(presentation.save().is_err());

    assert!(presentation.undo());
    assert_eq!(parts(&presentation.save().unwrap()), parts(FIXTURE));
}
