use pptx_edit::{DeckSession, EditCtx, EditError, StorySnapshot};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

fn first_story(session: &DeckSession) -> StorySnapshot {
    session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .flat_map(|slide| slide.shapes.iter())
        .find_map(|shape| shape.text_stories.first().cloned())
        .expect("the fixture has a text story")
}

#[test]
fn a_replica_opened_from_an_update_refuses_to_save() {
    let source = DeckSession::open(FIXTURE, 11).unwrap();
    let replica = DeckSession::open_from_update(&source.encode_state_as_update_v1(), 22).unwrap();
    assert_eq!(
        replica.snapshot().unwrap().slides.len(),
        source.snapshot().unwrap().slides.len(),
        "the replica still holds the whole deck"
    );

    let error = replica.project().unwrap_err();
    assert!(
        matches!(&error, EditError::Unprojectable(reason)
            if reason.contains("collaborative update")),
        "{error}"
    );
}

#[test]
fn a_projection_carries_the_edit_into_the_model_and_the_part_bytes() {
    let session = DeckSession::open(FIXTURE, 33).unwrap();
    let story = first_story(&session);
    let style = story.paragraphs[0].runs[0].style.clone();
    session
        .insert_text(&EditCtx::local("test"), &story.id, 0, "AA", &style)
        .unwrap();

    let projected = session.project().unwrap();
    let slide = projected
        .slides
        .iter()
        .find(|slide| {
            slide.shapes.iter().any(|shape| {
                matches!(shape, pptx_parse::ShapeNode::Shape(shape)
                    if shape.text.as_ref().is_some_and(|body| body
                        .paragraphs
                        .iter()
                        .flat_map(|paragraph| &paragraph.runs)
                        .any(|run| run.text.starts_with("AA"))))
            })
        })
        .expect("the projected model holds the edit");
    let bytes = projected
        .part_bytes(&slide.part_path)
        .expect("the projected package keeps its part bytes");
    assert!(String::from_utf8_lossy(bytes).contains("<a:t>AA"));
}

#[test]
fn a_text_deletion_that_empties_a_run_still_projects() {
    let session = DeckSession::open(FIXTURE, 44).unwrap();
    let story = first_story(&session);
    let length = story.paragraphs[0].runs[0].text.chars().count() as u32;
    session
        .delete_text(&EditCtx::local("test"), &story.id, 0, length)
        .unwrap();

    let projected = session.project().unwrap();
    let reopened = DeckSession::open(&pptx_parse::write_pptx(&projected).unwrap(), 55).unwrap();
    assert_eq!(
        reopened.story(&story.id).unwrap().plain_text(),
        session.story(&story.id).unwrap().plain_text()
    );
}
