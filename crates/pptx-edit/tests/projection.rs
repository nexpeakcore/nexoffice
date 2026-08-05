use pptx_edit::{DeckSession, EditCtx, EditError, StorySnapshot, TextStyle, WriteLimits};

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

const SLIDE_PART: &str = "ppt/slides/slide1.xml";
const SECOND_SLIDE_PART: &str = "ppt/slides/slide2.xml";
const THIRD_SLIDE_PART: &str = "ppt/slides/slide3.xml";

/// A slide whose only shape holds the paragraphs a test asks for, so a case can
/// be built out of exact `<a:r>`/`<a:br>`/`<a:fld>` sequences instead of
/// whatever a real deck happens to contain.
const SLIDE_TEMPLATE: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
    r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#,
    r#" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
    r#" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
    r#"<p:cSld><p:spTree>"#,
    r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
    r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Body"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>"#,
    r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6096000" cy="1981200"/></a:xfrm></p:spPr>"#,
    r#"<p:txBody><a:bodyPr/><a:lstStyle/>PARAGRAPHS</p:txBody>"#,
    r#"</p:sp></p:spTree></p:cSld></p:sld>"#,
);

fn slide_xml(paragraphs: &str) -> String {
    SLIDE_TEMPLATE.replace("PARAGRAPHS", paragraphs)
}

/// Re-zips the fixture with its first slide replaced, so opening the result
/// parses model and bytes from the same source.
fn deck_with(paragraphs: &str) -> Vec<u8> {
    let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
    assert!(package.replace_part(SLIDE_PART, slide_xml(paragraphs).into_bytes()));
    pptx_parse::write_pptx(&package).unwrap()
}

/// [`deck_with`] for the first two slides, so a save can be asked to rewrite
/// more than one part while the third slide stays as the fixture wrote it.
fn deck_with_slides(first: &str, second: &str) -> Vec<u8> {
    let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
    assert!(package.replace_part(SLIDE_PART, slide_xml(first).into_bytes()));
    assert!(package.replace_part(SECOND_SLIDE_PART, slide_xml(second).into_bytes()));
    pptx_parse::write_pptx(&package).unwrap()
}

fn session_with(paragraphs: &str) -> DeckSession {
    DeckSession::open(&deck_with(paragraphs), 77).unwrap()
}

fn part_text(deck: &[u8], part: &str) -> String {
    let package = pptx_parse::parse_pptx(deck).unwrap();
    String::from_utf8(package.part_bytes(part).unwrap().to_vec()).unwrap()
}

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

/// Every story of the deck in slide order, so a test can address the body of a
/// slide other than the first.
fn stories(session: &DeckSession) -> Vec<StorySnapshot> {
    session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .flat_map(|slide| slide.shapes.clone())
        .flat_map(|shape| shape.text_stories)
        .collect()
}

fn type_text(session: &DeckSession, story_id: &str, index: u32, text: &str) {
    session
        .insert_text(
            &EditCtx::local("test"),
            story_id,
            index,
            text,
            &TextStyle::default(),
        )
        .unwrap();
}

fn refusal(error: &EditError) -> String {
    match error {
        EditError::Unprojectable(reason) => reason.clone(),
        other => panic!("expected a refusal, got {other}"),
    }
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

#[test]
fn a_splice_that_lands_on_the_wrong_run_is_refused() {
    let deck = deck_with(r#"<a:p><a:r><a:t>AAA</a:t></a:r><a:r><a:t>BBB</a:t></a:r></a:p>"#);
    let mut package = pptx_parse::parse_pptx(&deck).unwrap();
    let bytes = String::from_utf8(package.part_bytes(SLIDE_PART).unwrap().to_vec()).unwrap();
    let swapped = bytes
        .replace("<a:t>AAA</a:t>", "<a:t>ZZZ</a:t>")
        .replace("<a:t>BBB</a:t>", "<a:t>AAA</a:t>")
        .replace("<a:t>ZZZ</a:t>", "<a:t>BBB</a:t>");
    assert_ne!(swapped, bytes);
    assert!(package.replace_part(SLIDE_PART, swapped.into_bytes()));

    let session = DeckSession::from_package(package, 88).unwrap();
    let story = first_story(&session);
    assert_eq!(story.plain_text(), "AAABBB");
    type_text(&session, &story.id, 0, "X");

    let reason = refusal(&session.project().unwrap_err());
    assert!(
        reason.contains("read back as a different deck"),
        "the wrong run was rewritten and the save must say so: {reason}"
    );
}

#[test]
fn an_edit_after_a_line_break_lands_in_the_run_that_follows_it() {
    let session =
        session_with(r#"<a:p><a:r><a:t>ab</a:t></a:r><a:br/><a:r><a:t>cd</a:t></a:r></a:p>"#);
    let story = first_story(&session);
    assert_eq!(story.plain_text(), "ab\ncd");
    type_text(&session, &story.id, 3, "X");

    let bytes = session.save_bytes().unwrap();
    let package = pptx_parse::parse_pptx(&bytes).unwrap();
    let part = String::from_utf8(package.part_bytes(SLIDE_PART).unwrap().to_vec()).unwrap();
    assert!(part.contains("<a:t>ab</a:t>"), "{part}");
    assert!(part.contains("<a:t>Xcd</a:t>"), "{part}");
    assert_eq!(
        DeckSession::open(&bytes, 99)
            .unwrap()
            .story(&story.id)
            .unwrap()
            .plain_text(),
        "ab\nXcd"
    );
}

#[test]
fn an_edit_past_a_trailing_line_break_is_refused_by_name() {
    let session = session_with(r#"<a:p><a:r><a:t>ab</a:t></a:r><a:br/></a:p>"#);
    let story = first_story(&session);
    assert_eq!(story.plain_text(), "ab\n");
    type_text(&session, &story.id, 3, "X");

    let reason = refusal(&session.project().unwrap_err());
    assert!(
        reason.contains("no run after it to hold the text"),
        "a break with nothing after it needs a new run this writer cannot add: {reason}"
    );
}

#[test]
fn an_edit_inside_a_line_break_is_still_refused() {
    let session =
        session_with(r#"<a:p><a:r><a:t>ab</a:t></a:r><a:br/><a:r><a:t>cd</a:t></a:r></a:p>"#);
    let story = first_story(&session);
    session
        .delete_text(&EditCtx::local("test"), &story.id, 2, 3)
        .unwrap();

    let reason = refusal(&session.project().unwrap_err());
    assert!(
        reason.contains("the change lands on line break"),
        "removing the break itself is still unprojectable: {reason}"
    );
}

#[test]
fn an_edit_taking_text_from_both_sides_of_a_line_break_is_refused_by_name() {
    let session =
        session_with(r#"<a:p><a:r><a:t>ab</a:t></a:r><a:br/><a:r><a:t>cd</a:t></a:r></a:p>"#);
    let story = first_story(&session);
    assert_eq!(story.plain_text(), "ab\ncd");
    session
        .delete_text(&EditCtx::local("test"), &story.id, 1, 4)
        .unwrap();
    assert_eq!(session.story(&story.id).unwrap().plain_text(), "ad");

    let reason = refusal(&session.project().unwrap_err());
    assert!(
        reason.contains("the change spans more than one run"),
        "text taken from both sides of a break has no single-run rewrite: {reason}"
    );
}

#[test]
fn an_edit_that_lands_on_a_field_is_refused_by_name() {
    let session = session_with(concat!(
        r#"<a:p><a:r><a:t>page </a:t></a:r>"#,
        r#"<a:fld id="{4B0A}" type="slidenum"><a:t>3</a:t></a:fld>"#,
        r#"<a:r><a:t>/9</a:t></a:r></a:p>"#,
    ));
    let story = first_story(&session);
    assert_eq!(story.plain_text(), "page 3/9");
    session
        .delete_text(&EditCtx::local("test"), &story.id, 5, 6)
        .unwrap();

    let reason = refusal(&session.project().unwrap_err());
    assert!(
        reason.contains("the change lands on field 2"),
        "a field's text belongs to PowerPoint, not to the writer: {reason}"
    );
}

#[test]
fn an_edit_that_spans_two_runs_is_refused_by_name() {
    let session = session_with(r#"<a:p><a:r><a:t>ab</a:t></a:r><a:r><a:t>cd</a:t></a:r></a:p>"#);
    let story = first_story(&session);
    assert_eq!(story.plain_text(), "abcd");
    session
        .delete_text(&EditCtx::local("test"), &story.id, 1, 3)
        .unwrap();

    let reason = refusal(&session.project().unwrap_err());
    assert!(
        reason.contains("the change spans more than one run"),
        "a change wider than one run has no single-run rewrite: {reason}"
    );
}

#[test]
fn a_deletion_inside_a_run_rewrites_only_that_runs_text() {
    let paragraphs = r#"<a:p><a:r><a:rPr b="1"/><a:t>abcdef</a:t></a:r></a:p>"#;
    let session = session_with(paragraphs);
    let story = first_story(&session);
    session
        .delete_text(&EditCtx::local("test"), &story.id, 2, 4)
        .unwrap();

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(paragraphs).replace("<a:t>abcdef</a:t>", "<a:t>abef</a:t>"),
        "a partial deletion moves no byte outside its own <a:t>"
    );
    assert_eq!(
        DeckSession::open(&bytes, 131)
            .unwrap()
            .story(&story.id)
            .unwrap()
            .plain_text(),
        "abef"
    );
}

#[test]
fn the_runs_around_an_edit_keep_their_entities_and_their_whitespace() {
    let paragraphs = concat!(
        r#"<a:p><a:r><a:t>Ben &amp; Co</a:t></a:r>"#,
        r#"<a:r><a:rPr b="1"/><a:t>target</a:t></a:r>"#,
        r#"<a:r><a:t xml:space="preserve">   </a:t></a:r></a:p>"#,
    );
    let session = session_with(paragraphs);
    let story = first_story(&session);
    assert_eq!(story.plain_text(), "Ben & Cotarget   ");
    let style = story.paragraphs[0]
        .runs
        .iter()
        .find(|run| run.text == "target")
        .expect("the edited run")
        .style
        .clone();
    session
        .insert_text(
            &EditCtx::local("test"),
            &story.id,
            "Ben & Cotar".chars().count() as u32,
            "X",
            &style,
        )
        .unwrap();

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(paragraphs).replace("<a:t>target</a:t>", "<a:t>tarXget</a:t>"),
        "the runs on either side keep their entity and their spaces byte for byte"
    );
    assert_eq!(
        DeckSession::open(&bytes, 141)
            .unwrap()
            .story(&story.id)
            .unwrap()
            .plain_text(),
        "Ben & CotarXget   "
    );
}

#[test]
fn one_save_rewrites_several_paragraphs_across_several_slides() {
    let first = concat!(
        r#"<a:p><a:r><a:t>alpha</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:t>beta</a:t></a:r></a:p>"#,
    );
    let second = r#"<a:p><a:r><a:t>gamma</a:t></a:r></a:p>"#;
    let deck = deck_with_slides(first, second);
    let session = DeckSession::open(&deck, 151).unwrap();
    let stories = stories(&session);
    let (front, back) = (stories[0].clone(), stories[1].clone());
    assert_eq!(front.plain_text(), "alpha\nbeta");
    assert_eq!(back.plain_text(), "gamma");

    type_text(&session, &front.id, 0, "1");
    type_text(&session, &front.id, 7, "2");
    type_text(&session, &back.id, 0, "3");

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(first)
            .replace("<a:t>alpha</a:t>", "<a:t>1alpha</a:t>")
            .replace("<a:t>beta</a:t>", "<a:t>2beta</a:t>"),
        "both paragraphs of one body are rewritten in place"
    );
    assert_eq!(
        part_text(&bytes, SECOND_SLIDE_PART),
        slide_xml(second).replace("<a:t>gamma</a:t>", "<a:t>3gamma</a:t>")
    );
    assert_eq!(
        part_text(&bytes, THIRD_SLIDE_PART),
        part_text(&deck, THIRD_SLIDE_PART),
        "the slide no edit named keeps its source bytes"
    );

    let reopened = DeckSession::open(&bytes, 161).unwrap();
    assert_eq!(
        reopened.story(&front.id).unwrap().plain_text(),
        "1alpha\n2beta"
    );
    assert_eq!(reopened.story(&back.id).unwrap().plain_text(), "3gamma");
}

#[test]
fn a_run_larger_than_the_write_budget_refuses_the_save() {
    let session = session_with(r#"<a:p><a:r><a:t>ab</a:t></a:r></a:p>"#);
    let story = first_story(&session);
    let limits = WriteLimits::default();
    let paste = "z".repeat(limits.max_run_text_bytes + 1);
    type_text(&session, &story.id, 0, &paste);

    let reason = refusal(&session.project().unwrap_err());
    assert!(
        reason.contains("bytes into one run"),
        "an oversized paste must be refused by name: {reason}"
    );
}

#[test]
fn the_write_budgets_bound_the_runs_and_the_bytes_one_save_rewrites() {
    let session = session_with(concat!(
        r#"<a:p><a:r><a:t>one</a:t></a:r></a:p>"#,
        r#"<a:p><a:r><a:t>two</a:t></a:r></a:p>"#,
    ));
    let story = first_story(&session);
    type_text(&session, &story.id, 0, "X");
    type_text(&session, &story.id, 5, "Y");

    let reason = refusal(
        &session
            .project_with_limits(&WriteLimits {
                max_run_edits: 1,
                ..WriteLimits::default()
            })
            .unwrap_err(),
    );
    assert!(reason.contains("runs one save may rewrite"), "{reason}");

    let reason = refusal(
        &session
            .project_with_limits(&WriteLimits {
                max_total_edit_bytes: 5,
                ..WriteLimits::default()
            })
            .unwrap_err(),
    );
    assert!(
        reason.contains("bytes of text one save may write"),
        "{reason}"
    );

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        DeckSession::open(&bytes, 111)
            .unwrap()
            .story(&story.id)
            .unwrap()
            .plain_text(),
        "Xone\nYtwo",
        "an ordinary edit passes the default budgets untouched"
    );
}
