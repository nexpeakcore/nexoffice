use pptx_edit::{
    DeckSession, EditCtx, EditError, SaveFault, ShapeSnapshot, StorySnapshot, TextStyle,
    TextStylePatch, WriteLimits,
};

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

/// What the editor falls back to for a value nothing under the caret states:
/// `initialStyle` in `packages/pptx-react/src/PptxEditor.tsx`.
fn editor_fallback() -> TextStyle {
    TextStyle {
        bold: Some(false),
        italic: Some(false),
        font_size_pt: Some(24.0),
        color: Some("#111827".to_owned()),
        font_family: Some("Arial".to_owned()),
        underline: Some("none".to_owned()),
    }
}

/// The style the editor inserts typed text with.
///
/// A port of `effectiveStyleFromSelection` in
/// `packages/pptx-react/src/textFormatting.ts` for a collapsed caret: the run
/// the caret sits in, or the last run before it, resolved against
/// [`editor_fallback`] so that every field arrives spelled out — including the
/// `b`, `i` and `u` the run's own `<a:rPr>` usually leaves to be inherited.
/// This, not the run's own style, is what a keystroke carries into the model.
fn effective_style(story: &StorySnapshot, index: u32) -> TextStyle {
    let fallback = editor_fallback();
    let mut spans: Vec<(u32, u32, TextStyle)> = Vec::new();
    let mut position = 0;
    for (index, paragraph) in story.paragraphs.iter().enumerate() {
        for run in &paragraph.runs {
            let start = position;
            position += run.text.encode_utf16().count() as u32;
            if position > start {
                spans.push((
                    start,
                    position,
                    TextStyle {
                        bold: run.style.bold.or(fallback.bold),
                        italic: run.style.italic.or(fallback.italic),
                        font_size_pt: run.style.font_size_pt.or(fallback.font_size_pt),
                        color: run.style.color.clone().or_else(|| fallback.color.clone()),
                        font_family: run
                            .style
                            .font_family
                            .clone()
                            .or_else(|| fallback.font_family.clone()),
                        underline: run
                            .style
                            .underline
                            .clone()
                            .or_else(|| fallback.underline.clone()),
                    },
                ));
            }
        }
        if index + 1 < story.paragraphs.len() {
            position += 1;
        }
    }
    spans
        .iter()
        .find(|(start, end, _)| index >= *start && index < *end)
        .or_else(|| spans.iter().rev().find(|(_, end, _)| *end <= index))
        .or_else(|| spans.first())
        .map(|(_, _, style)| style.clone())
        .unwrap_or(fallback)
}

/// Types `text` the way the editor does: at `index`, carrying the whole style
/// the caret resolves to.
fn type_as_the_editor_does(session: &DeckSession, story_id: &str, index: u32, text: &str) {
    let style = effective_style(&session.story(story_id).unwrap(), index);
    session
        .insert_text(&EditCtx::local("test"), story_id, index, text, &style)
        .unwrap();
}

/// The writer's account of a change it cannot express.
///
/// Asserting the fault, not just the wording, is the point: a broken write and
/// a blown budget also arrive with a reason, and reading either as a refusal
/// would tell the caller that undoing something fixes a save that undoing
/// cannot reach.
fn refusal(error: &EditError) -> String {
    fault_reason(error, SaveFault::Unprojectable)
}

fn fault_reason(error: &EditError, expected: SaveFault) -> String {
    assert_eq!(
        error.save_fault(),
        expected,
        "expected a {expected:?} save, got: {error}"
    );
    match error {
        EditError::Unprojectable(reason)
        | EditError::Unsavable(reason)
        | EditError::WriteLimit(reason)
        | EditError::WriteFailed(reason)
        | EditError::VerificationFailed(reason) => reason.clone(),
        other => other.to_string(),
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
    let reason = fault_reason(&error, SaveFault::Unsavable);
    assert!(reason.contains("collaborative update"), "{reason}");
    // No edit put the replica here and no undo takes it back out, so a host
    // that offers to abandon edits over this offers a way out that does not
    // exist.
    assert!(!error.save_fault().undoing_helps());
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
fn a_splice_that_lands_on_the_wrong_run_fails_verification() {
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

    // The writer put the text somewhere the edit did not name. Nothing the
    // user did causes this and no undo clears it, so it must never be read as
    // a refusal — the edit is sound and has to survive.
    let error = session.project().unwrap_err();
    let reason = fault_reason(&error, SaveFault::VerificationFailed);
    assert!(
        reason.contains("read back as a different deck"),
        "the wrong run was rewritten and the save must say so: {reason}"
    );
    assert!(!error.save_fault().undoing_helps());
}

#[test]
fn a_split_whose_bytes_land_on_the_wrong_run_fails_verification() {
    let deck = deck_with(r#"<a:p><a:r><a:t>AAA</a:t></a:r><a:r><a:t>BBB</a:t></a:r></a:p>"#);
    let mut package = pptx_parse::parse_pptx(&deck).unwrap();
    let bytes = String::from_utf8(package.part_bytes(SLIDE_PART).unwrap().to_vec()).unwrap();
    let swapped = bytes
        .replace("<a:t>AAA</a:t>", "<a:t>ZZZ</a:t>")
        .replace("<a:t>BBB</a:t>", "<a:t>AAA</a:t>")
        .replace("<a:t>ZZZ</a:t>", "<a:t>BBB</a:t>");
    assert!(package.replace_part(SLIDE_PART, swapped.into_bytes()));

    let session = DeckSession::from_package(package, 217).unwrap();
    let story = first_story(&session);
    assert_eq!(story.plain_text(), "AAABBB");
    press_enter(&session, &story.id, 1);

    let error = session.project().unwrap_err();
    let reason = fault_reason(&error, SaveFault::VerificationFailed);
    assert!(
        reason.contains("read back as a different deck"),
        "a split spliced into the wrong run must be caught, not shipped: {reason}"
    );
    assert!(!error.save_fault().undoing_helps());
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

/// A paragraph with properties worth carrying, formatting worth keeping and an
/// end paragraph mark worth placing, so a split can be checked byte for byte.
const BULLETED: &str = concat!(
    r#"<a:p><a:pPr lvl="2" marL="457200" indent="-228600"><a:buChar char="•"/></a:pPr>"#,
    r#"<a:r><a:rPr b="1" sz="1400"/><a:t>alpha</a:t></a:r>"#,
    r#"<a:r><a:rPr i="1"/><a:t>beta</a:t></a:r>"#,
    r#"<a:endParaRPr sz="1800"/></a:p>"#,
);

const BULLET_PPR: &str =
    r#"<a:pPr lvl="2" marL="457200" indent="-228600"><a:buChar char="•"/></a:pPr>"#;

fn press_enter(session: &DeckSession, story_id: &str, index: u32) {
    session
        .insert_paragraph_break(&EditCtx::local("test"), story_id, index)
        .unwrap();
}

#[test]
fn a_split_at_the_start_of_a_paragraph_leaves_an_empty_paragraph_above_it() {
    let session = session_with(BULLETED);
    let story = first_story(&session);
    press_enter(&session, &story.id, 0);

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(BULLETED).replace(
            r#"<a:r><a:rPr b="1" sz="1400"/>"#,
            &format!(r#"</a:p><a:p>{BULLET_PPR}<a:r><a:rPr b="1" sz="1400"/>"#),
        ),
        "the empty paragraph copies the pPr and every run keeps its own bytes"
    );

    let reopened = DeckSession::open(&bytes, 201).unwrap();
    let story = reopened
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .flat_map(|slide| slide.shapes.clone())
        .find_map(|shape| shape.text_stories.first().cloned())
        .unwrap();
    assert_eq!(story.plain_text(), "\nalphabeta");
    assert_eq!(story.paragraphs[0].level, 2);
    assert_eq!(story.paragraphs[1].level, 2);
    assert_eq!(
        story.paragraphs[0].bullet_json,
        story.paragraphs[1].bullet_json
    );
}

#[test]
fn a_split_at_the_end_of_a_paragraph_moves_the_end_mark_onto_the_new_one() {
    let session = session_with(BULLETED);
    let story = first_story(&session);
    press_enter(&session, &story.id, "alphabeta".chars().count() as u32);

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(BULLETED).replace(
            r#"<a:endParaRPr sz="1800"/>"#,
            &format!(r#"</a:p><a:p>{BULLET_PPR}<a:endParaRPr sz="1800"/>"#),
        ),
        "the paragraph mark of the source paragraph ends the new empty one"
    );
    assert_eq!(
        DeckSession::open(&bytes, 202)
            .unwrap()
            .story(&story.id)
            .unwrap()
            .plain_text(),
        "alphabeta\n"
    );
}

#[test]
fn a_split_inside_a_formatted_run_gives_both_halves_its_run_properties() {
    let session = session_with(BULLETED);
    let story = first_story(&session);
    press_enter(&session, &story.id, 2);

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(BULLETED).replace(
            r#"<a:r><a:rPr b="1" sz="1400"/><a:t>alpha</a:t></a:r>"#,
            &format!(
                concat!(
                    r#"<a:r><a:rPr b="1" sz="1400"/><a:t>al</a:t></a:r>"#,
                    r#"</a:p><a:p>{}"#,
                    r#"<a:r><a:rPr b="1" sz="1400"/><a:t>pha</a:t></a:r>"#,
                ),
                BULLET_PPR
            ),
        ),
        "the divided run's <a:rPr> is copied, not rebuilt"
    );

    let reopened = DeckSession::open(&bytes, 203).unwrap();
    let reparsed = reopened.story(&story.id).unwrap();
    assert_eq!(reparsed.plain_text(), "al\nphabeta");
    assert_eq!(reparsed.paragraphs[0].runs[0].style.bold, Some(true));
    assert_eq!(reparsed.paragraphs[1].runs[0].style.bold, Some(true));
    assert_eq!(
        reparsed.paragraphs[1].runs[0].style.font_size_pt,
        Some(14.0)
    );
}

#[test]
fn a_split_leaves_the_slides_it_did_not_touch_byte_for_byte() {
    let deck = deck_with_slides(BULLETED, r#"<a:p><a:r><a:t>gamma</a:t></a:r></a:p>"#);
    let session = DeckSession::open(&deck, 204).unwrap();
    let story = stories(&session)[0].clone();
    press_enter(&session, &story.id, 5);

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SECOND_SLIDE_PART),
        part_text(&deck, SECOND_SLIDE_PART)
    );
    assert_eq!(
        part_text(&bytes, THIRD_SLIDE_PART),
        part_text(&deck, THIRD_SLIDE_PART)
    );
    let before = ooxml_opc::unzip_parts(&deck).unwrap();
    let after = ooxml_opc::unzip_parts(&bytes).unwrap();
    for ((path, source), (saved_path, saved)) in before.iter().zip(&after) {
        assert_eq!(path, saved_path);
        if path != SLIDE_PART {
            assert_eq!(source, saved, "{path} was rewritten");
        }
    }
}

#[test]
fn a_merge_keeps_the_first_paragraphs_properties_and_the_last_end_mark() {
    let paragraphs = concat!(
        r#"<a:p><a:pPr algn="ctr"/><a:r><a:rPr b="1"/><a:t>alpha</a:t></a:r>"#,
        r#"<a:endParaRPr sz="900"/></a:p>"#,
        r#"<a:p><a:pPr lvl="3" algn="r"><a:buNone/></a:pPr>"#,
        r#"<a:r><a:rPr i="1" sz="2400"/><a:t>beta</a:t></a:r>"#,
        r#"<a:endParaRPr sz="1800"/></a:p>"#,
    );
    let session = session_with(paragraphs);
    let story = first_story(&session);
    assert_eq!(story.plain_text(), "alpha\nbeta");
    session
        .delete_paragraph_break(&EditCtx::local("test"), &story.id, 5)
        .unwrap();
    assert_eq!(session.story(&story.id).unwrap().plain_text(), "alphabeta");

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(paragraphs).replace(
            concat!(
                r#"<a:endParaRPr sz="900"/></a:p>"#,
                r#"<a:p><a:pPr lvl="3" algn="r"><a:buNone/></a:pPr>"#,
            ),
            "",
        ),
        "one deletion joins the paragraphs; both runs keep their own bytes"
    );

    let reparsed = DeckSession::open(&bytes, 205)
        .unwrap()
        .story(&story.id)
        .unwrap();
    assert_eq!(reparsed.plain_text(), "alphabeta");
    assert_eq!(reparsed.paragraphs.len(), 1);
    assert_eq!(reparsed.paragraphs[0].alignment.as_deref(), Some("ctr"));
    assert_eq!(reparsed.paragraphs[0].level, 0);
    assert_eq!(reparsed.paragraphs[0].runs[0].style.bold, Some(true));
    assert_eq!(reparsed.paragraphs[0].runs[1].style.italic, Some(true));
}

#[test]
fn a_split_and_a_merge_are_inverses_down_to_the_byte() {
    let deck = deck_with(BULLETED);
    let session = DeckSession::open(&deck, 206).unwrap();
    let story = first_story(&session);
    press_enter(&session, &story.id, 5);
    let split = session.save_bytes().unwrap();
    assert_ne!(part_text(&split, SLIDE_PART), part_text(&deck, SLIDE_PART));

    let session = DeckSession::open(&split, 207).unwrap();
    let story = first_story(&session);
    session
        .delete_paragraph_break(&EditCtx::local("test"), &story.id, 5)
        .unwrap();
    assert_eq!(
        part_text(&session.save_bytes().unwrap(), SLIDE_PART),
        part_text(&deck, SLIDE_PART),
        "merging what a split produced restores the source part"
    );
}

#[test]
fn merging_a_split_run_keeps_the_two_runs_it_was_divided_into() {
    let deck = deck_with(BULLETED);
    let session = DeckSession::open(&deck, 213).unwrap();
    let story = first_story(&session);
    press_enter(&session, &story.id, 3);
    let split = session.save_bytes().unwrap();

    let session = DeckSession::open(&split, 214).unwrap();
    let story = first_story(&session);
    session
        .delete_paragraph_break(&EditCtx::local("test"), &story.id, 3)
        .unwrap();
    let merged = session.save_bytes().unwrap();
    let reparsed = DeckSession::open(&merged, 215)
        .unwrap()
        .story(&story.id)
        .unwrap();
    assert_eq!(reparsed.plain_text(), "alphabeta");
    assert_eq!(reparsed.paragraphs.len(), 1);
    assert_eq!(
        part_text(&merged, SLIDE_PART),
        slide_xml(BULLETED).replace(
            r#"<a:r><a:rPr b="1" sz="1400"/><a:t>alpha</a:t></a:r>"#,
            concat!(
                r#"<a:r><a:rPr b="1" sz="1400"/><a:t>alp</a:t></a:r>"#,
                r#"<a:r><a:rPr b="1" sz="1400"/><a:t>ha</a:t></a:r>"#,
            ),
        ),
        "joining runs is not this writer's to do, so the divided run stays divided"
    );
}

#[test]
fn a_line_break_typed_inside_a_run_writes_a_br_carrying_its_properties() {
    let session = session_with(BULLETED);
    let story = first_story(&session);
    let style = story.paragraphs[0].runs[0].style.clone();
    session
        .insert_text(&EditCtx::local("test"), &story.id, 2, "\n", &style)
        .unwrap();

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(BULLETED).replace(
            r#"<a:r><a:rPr b="1" sz="1400"/><a:t>alpha</a:t></a:r>"#,
            concat!(
                r#"<a:r><a:rPr b="1" sz="1400"/><a:t>al</a:t></a:r>"#,
                r#"<a:br><a:rPr b="1" sz="1400"/></a:br>"#,
                r#"<a:r><a:rPr b="1" sz="1400"/><a:t>pha</a:t></a:r>"#,
            ),
        ),
        "the break carries the formatting of the run it divides, as PowerPoint writes it"
    );

    let reparsed = DeckSession::open(&bytes, 208)
        .unwrap()
        .story(&story.id)
        .unwrap();
    assert_eq!(reparsed.plain_text(), "al\nphabeta");
    assert_eq!(reparsed.paragraphs.len(), 1);
}

#[test]
fn a_line_break_typed_next_to_an_unstyled_run_needs_no_run_properties() {
    let paragraphs = r#"<a:p><a:r><a:t>abcd</a:t></a:r></a:p>"#;
    let session = session_with(paragraphs);
    let story = first_story(&session);
    type_text(&session, &story.id, 2, "\n");

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(paragraphs).replace(
            r#"<a:r><a:t>abcd</a:t></a:r>"#,
            r#"<a:r><a:t>ab</a:t></a:r><a:br/><a:r><a:t>cd</a:t></a:r>"#,
        )
    );
    assert_eq!(
        DeckSession::open(&bytes, 209)
            .unwrap()
            .story(&story.id)
            .unwrap()
            .plain_text(),
        "ab\ncd"
    );
}

#[test]
fn typing_and_then_splitting_in_one_place_saves_as_one_change() {
    let paragraphs = r#"<a:p><a:pPr algn="ctr"/><a:r><a:t>abcd</a:t></a:r></a:p>"#;
    let session = session_with(paragraphs);
    let story = first_story(&session);
    type_text(&session, &story.id, 2, "XY");
    press_enter(&session, &story.id, 4);

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(paragraphs).replace(
            r#"<a:r><a:t>abcd</a:t></a:r>"#,
            concat!(
                r#"<a:r><a:t>abXY</a:t></a:r>"#,
                r#"</a:p><a:p><a:pPr algn="ctr"/>"#,
                r#"<a:r><a:t>cd</a:t></a:r>"#,
            ),
        )
    );
    assert_eq!(
        DeckSession::open(&bytes, 210)
            .unwrap()
            .story(&story.id)
            .unwrap()
            .plain_text(),
        "abXY\ncd"
    );
}

#[test]
fn a_split_of_an_empty_paragraph_writes_a_second_empty_paragraph() {
    let paragraphs = concat!(
        r#"<a:p><a:pPr algn="ctr"/><a:endParaRPr sz="1800"/></a:p>"#,
        r#"<a:p><a:r><a:t>after</a:t></a:r></a:p>"#,
    );
    let session = session_with(paragraphs);
    let story = first_story(&session);
    press_enter(&session, &story.id, 0);

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(paragraphs).replace(
            r#"<a:p><a:pPr algn="ctr"/><a:endParaRPr sz="1800"/></a:p>"#,
            r#"<a:p><a:pPr algn="ctr"/></a:p><a:p><a:pPr algn="ctr"/><a:endParaRPr sz="1800"/></a:p>"#,
        )
    );
    assert_eq!(
        DeckSession::open(&bytes, 211)
            .unwrap()
            .story(&story.id)
            .unwrap()
            .plain_text(),
        "\n\nafter"
    );
}

#[test]
fn two_splits_inside_one_run_write_three_paragraphs() {
    let paragraphs = r#"<a:p><a:pPr algn="ctr"/><a:r><a:t>abcdefgh</a:t></a:r></a:p>"#;
    let session = session_with(paragraphs);
    let story = first_story(&session);
    press_enter(&session, &story.id, 6);
    press_enter(&session, &story.id, 2);

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(paragraphs).replace(
            r#"<a:r><a:t>abcdefgh</a:t></a:r>"#,
            concat!(
                r#"<a:r><a:t>ab</a:t></a:r>"#,
                r#"</a:p><a:p><a:pPr algn="ctr"/><a:r><a:t>cdef</a:t></a:r>"#,
                r#"</a:p><a:p><a:pPr algn="ctr"/><a:r><a:t>gh</a:t></a:r>"#,
            ),
        )
    );
    assert_eq!(
        DeckSession::open(&bytes, 216)
            .unwrap()
            .story(&story.id)
            .unwrap()
            .plain_text(),
        "ab\ncdef\ngh"
    );
}

#[test]
fn two_edits_in_different_runs_of_one_paragraph_are_still_refused_by_name() {
    let session = session_with(concat!(
        r#"<a:p><a:r><a:t>abcd</a:t></a:r>"#,
        r#"<a:r><a:rPr b="1"/><a:t>efgh</a:t></a:r></a:p>"#,
    ));
    let story = first_story(&session);
    press_enter(&session, &story.id, 2);
    press_enter(&session, &story.id, 7);

    let reason = refusal(&session.project().unwrap_err());
    assert!(
        reason.contains("the change spans more than one run"),
        "two changes in one body that are not one change have no faithful rewrite: {reason}"
    );
}

#[test]
fn a_split_inside_a_field_is_refused_by_name() {
    let session = session_with(concat!(
        r#"<a:p><a:r><a:t>page </a:t></a:r>"#,
        r#"<a:fld id="{4B0A}" type="slidenum"><a:t>12</a:t></a:fld></a:p>"#,
    ));
    let story = first_story(&session);
    assert_eq!(story.plain_text(), "page 12");
    press_enter(&session, &story.id, 6);

    let reason = refusal(&session.project().unwrap_err());
    assert!(
        reason.contains("the change lands on field 2"),
        "a field's text belongs to PowerPoint, not to the writer: {reason}"
    );
}

#[test]
fn typing_into_a_paragraph_with_no_run_is_still_refused_by_name() {
    let session = session_with(r#"<a:p><a:endParaRPr sz="1800"/></a:p>"#);
    let story = first_story(&session);
    type_text(&session, &story.id, 0, "X");

    let reason = refusal(&session.project().unwrap_err());
    assert!(
        reason.contains("no run to hold text"),
        "an empty paragraph has no run to carry a style: {reason}"
    );
}

#[test]
fn a_split_keeps_the_paragraph_level_on_both_halves() {
    let session = session_with(BULLETED);
    let story = first_story(&session);
    press_enter(&session, &story.id, 5);
    let split = session.snapshot().unwrap();
    assert_eq!(
        split.slides[0].shapes[0].text_stories[0].paragraphs.len(),
        2
    );

    let bytes = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&bytes, 212).unwrap();
    let levels: Vec<u32> = reopened
        .story(&story.id)
        .unwrap()
        .paragraphs
        .iter()
        .map(|paragraph| paragraph.level)
        .collect();
    assert_eq!(levels, vec![2, 2], "a split keeps the level on both halves");
}

#[test]
fn a_run_larger_than_the_write_budget_stops_the_save() {
    let session = session_with(r#"<a:p><a:r><a:t>ab</a:t></a:r></a:p>"#);
    let story = first_story(&session);
    let limits = WriteLimits::default();
    let paste = "z".repeat(limits.max_run_text_bytes + 1);
    type_text(&session, &story.id, 0, &paste);

    // A budget is not a capability gap: the writer can express this paste, it
    // is simply too much for one save. Saying so is what lets a host offer to
    // save less rather than to throw the paste away.
    let error = session.project().unwrap_err();
    let reason = fault_reason(&error, SaveFault::Limit);
    assert!(
        reason.contains("bytes into one run"),
        "an oversized paste must be stopped by name: {reason}"
    );
    assert!(!error.save_fault().undoing_helps());
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

    let reason = fault_reason(
        &session
            .project_with_limits(&WriteLimits {
                max_run_edits: 1,
                ..WriteLimits::default()
            })
            .unwrap_err(),
        SaveFault::Limit,
    );
    assert!(reason.contains("runs one save may rewrite"), "{reason}");

    let reason = fault_reason(
        &session
            .project_with_limits(&WriteLimits {
                max_total_edit_bytes: 5,
                ..WriteLimits::default()
            })
            .unwrap_err(),
        SaveFault::Limit,
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

#[test]
fn a_keystroke_into_a_run_with_no_run_properties_rewrites_only_its_text() {
    let paragraphs = r#"<a:p><a:r><a:t>abcd</a:t></a:r></a:p>"#;
    let session = session_with(paragraphs);
    let story = first_story(&session);
    assert_eq!(
        story.paragraphs[0].runs[0].style,
        TextStyle::default(),
        "the run states nothing, so every field of the caret's style is the editor's own"
    );
    type_as_the_editor_does(&session, &story.id, 2, "X");

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(paragraphs).replace("<a:t>abcd</a:t>", "<a:t>abXcd</a:t>"),
        "a keystroke moves no byte outside the <a:t> it landed in"
    );
    assert_eq!(
        DeckSession::open(&bytes, 301)
            .unwrap()
            .story(&story.id)
            .unwrap()
            .plain_text(),
        "abXcd"
    );
}

#[test]
fn a_keystroke_into_a_partly_stated_run_saves_at_its_start_middle_and_end() {
    let paragraphs = r#"<a:p><a:r><a:rPr b="1" sz="1400"/><a:t>abcd</a:t></a:r></a:p>"#;
    for (index, (caret, expected)) in [(0, "Xabcd"), (2, "abXcd"), (4, "abcdX")]
        .into_iter()
        .enumerate()
    {
        let session = session_with(paragraphs);
        let story = first_story(&session);
        let style = effective_style(&story, caret);
        assert_eq!(
            (style.bold, style.italic, style.underline.as_deref()),
            (Some(true), Some(false), Some("none")),
            "the caret keeps the b the run states and spells out the i and u it does not"
        );
        type_as_the_editor_does(&session, &story.id, caret, "X");

        let bytes = session.save_bytes().unwrap();
        assert_eq!(
            part_text(&bytes, SLIDE_PART),
            slide_xml(paragraphs).replace("<a:t>abcd</a:t>", &format!("<a:t>{expected}</a:t>")),
            "the run keeps its own <a:rPr>, whatever the caret spelled out"
        );
        assert_eq!(
            DeckSession::open(&bytes, 310 + index as u64)
                .unwrap()
                .story(&story.id)
                .unwrap()
                .plain_text(),
            expected
        );
    }
}

#[test]
fn a_keystroke_leaves_the_other_runs_of_its_paragraph_byte_for_byte() {
    let paragraphs = concat!(
        r#"<a:p><a:r><a:rPr b="1"/><a:t>alpha</a:t></a:r>"#,
        r#"<a:r><a:rPr i="1" sz="1200"/><a:t>beta</a:t></a:r>"#,
        r#"<a:r><a:t>gamma</a:t></a:r></a:p>"#,
    );
    let session = session_with(paragraphs);
    let story = first_story(&session);
    assert_eq!(story.plain_text(), "alphabetagamma");
    type_as_the_editor_does(&session, &story.id, 7, "X");

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(paragraphs).replace("<a:t>beta</a:t>", "<a:t>beXta</a:t>"),
        "only the run the caret sat in is rewritten"
    );
    assert_eq!(
        DeckSession::open(&bytes, 320)
            .unwrap()
            .story(&story.id)
            .unwrap()
            .plain_text(),
        "alphabeXtagamma"
    );
}

#[test]
fn a_keystroke_repeating_the_letter_beside_it_still_lands_in_its_own_run() {
    let paragraphs = concat!(
        r#"<a:p><a:r><a:rPr b="1" sz="1000"/><a:t>DOCX</a:t></a:r>"#,
        r#"<a:r><a:rPr b="1" sz="1000"/><a:t> tab</a:t></a:r></a:p>"#,
    );
    let session = session_with(paragraphs);
    let story = first_story(&session);
    type_as_the_editor_does(&session, &story.id, 3, "X");

    let bytes = session.save_bytes().unwrap();
    assert_eq!(
        part_text(&bytes, SLIDE_PART),
        slide_xml(paragraphs).replace("<a:t>DOCX</a:t>", "<a:t>DOCXX</a:t>"),
        "which of the two Xs was typed cannot be read off the text, and either reads the same"
    );
    assert_eq!(
        DeckSession::open(&bytes, 330)
            .unwrap()
            .story(&story.id)
            .unwrap()
            .plain_text(),
        "DOCXX tab"
    );
}

/// Bolding half a run splits it in the file: the untouched half keeps its
/// bytes and the styled half gets a synthesised `<a:rPr b="1"/>`.
#[test]
fn a_bold_toggle_over_part_of_a_run_splits_it_in_the_saved_file() {
    let session = session_with(r#"<a:p><a:r><a:t>abcd</a:t></a:r></a:p>"#);
    let story = first_story(&session);
    session
        .format_text(
            &EditCtx::local("test"),
            &story.id,
            2,
            4,
            &TextStylePatch {
                bold: Some(true),
                ..TextStylePatch::default()
            },
        )
        .unwrap();

    let saved = session.save_bytes().unwrap();
    let slide_xml = part_text(&saved, SLIDE_PART);
    assert!(slide_xml.contains(r#"<a:t>ab</a:t>"#), "{slide_xml}");
    assert!(
        slide_xml.contains(r#"<a:rPr b="1"/><a:t>cd</a:t>"#),
        "{slide_xml}"
    );

    let reopened = DeckSession::open(&saved, 91).unwrap();
    let story = first_story(&reopened);
    let runs: Vec<_> = story.paragraphs[0]
        .runs
        .iter()
        .map(|run| (run.text.as_str(), run.style.bold))
        .collect();
    assert_eq!(runs, [("ab", None), ("cd", Some(true))]);
}

#[test]
fn a_bold_toggle_beside_a_keystroke_both_save() {
    let session = session_with(r#"<a:p><a:r><a:t>abcd</a:t></a:r></a:p>"#);
    let story = first_story(&session);
    type_as_the_editor_does(&session, &story.id, 2, "X");
    session
        .format_text(
            &EditCtx::local("test"),
            &story.id,
            0,
            2,
            &TextStylePatch {
                bold: Some(true),
                ..TextStylePatch::default()
            },
        )
        .unwrap();

    let saved = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&saved, 92).unwrap();
    let story = first_story(&reopened);
    let runs: Vec<_> = story.paragraphs[0]
        .runs
        .iter()
        .map(|run| (run.text.as_str(), run.style.bold))
        .collect();
    assert_eq!(runs, [("ab", Some(true)), ("Xcd", None)]);
}

#[test]
fn a_keystroke_contradicting_the_run_it_lands_in_splits_the_run() {
    let paragraphs = r#"<a:p><a:r><a:rPr b="1"/><a:t>abcd</a:t></a:r></a:p>"#;
    let session = session_with(paragraphs);
    let story = first_story(&session);
    let mut style = effective_style(&story, 2);
    style.bold = Some(false);
    session
        .insert_text(&EditCtx::local("test"), &story.id, 2, "X", &style)
        .unwrap();

    let saved = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&saved, 93).unwrap();
    let story = first_story(&reopened);
    let runs: Vec<_> = story.paragraphs[0]
        .runs
        .iter()
        .map(|run| (run.text.as_str(), run.style.bold))
        .collect();
    assert_eq!(
        runs,
        [("ab", Some(true)), ("X", Some(false)), ("cd", Some(true))]
    );
}

/// Every text story of the demo deck, groups and table cells included.
fn demo_stories(session: &DeckSession) -> Vec<StorySnapshot> {
    fn walk(shape: &ShapeSnapshot, stories: &mut Vec<StorySnapshot>) {
        stories.extend(shape.text_stories.iter().cloned());
        for child in &shape.children {
            walk(child, stories);
        }
    }
    let mut stories = Vec::new();
    for slide in &session.snapshot().unwrap().slides {
        for shape in &slide.shapes {
            walk(shape, &mut stories);
        }
    }
    stories
}

/// The case the suite used to miss: a keystroke as the editor makes it, in
/// every text story of a real deck.
///
/// Typing with the run's own style saved before this test existed; typing with
/// the style the editor actually sends refused in every one of these stories,
/// because a real deck's runs state `b` and `sz` and leave `i` and `u` to be
/// inherited.
#[test]
fn every_text_story_of_the_demo_deck_saves_after_a_keystroke() {
    let package = pptx_parse::parse_pptx(FIXTURE).unwrap();
    let session = DeckSession::from_package(package.clone(), 340).unwrap();
    let stories = demo_stories(&session);
    assert!(
        stories.len() > 40,
        "the demo deck holds a deck's worth of text"
    );

    let mut refused = Vec::new();
    for (index, story) in stories.iter().enumerate() {
        for caret in [0, 3, 5, 20, 50] {
            let session =
                DeckSession::from_package(package.clone(), 1_000 + index as u64 * 8 + caret)
                    .unwrap();
            let live = session.story(&story.id).unwrap();
            if caret as u32 >= live.length {
                continue;
            }
            type_as_the_editor_does(&session, &story.id, caret as u32, "X");
            if let Err(error) = session.save_bytes() {
                refused.push(format!("{} at {caret}: {error}", story.id));
            }
        }
    }
    assert!(
        refused.is_empty(),
        "a keystroke must be savable everywhere:\n{}",
        refused.join("\n")
    );
}

/// The one reading a host may take from a failed save.
///
/// A code that drifts, or a new fault that quietly lands on the one the
/// desktop answers with an offer to abandon edits, is what this pins.
#[test]
fn only_a_change_the_writer_cannot_express_is_undone_away() {
    let faults = [
        (SaveFault::Unprojectable, "unprojectable", true),
        (SaveFault::Unsavable, "unsavable", false),
        (SaveFault::Limit, "limit", false),
        (SaveFault::WriteFailed, "write-failed", false),
        (SaveFault::VerificationFailed, "verification-failed", false),
    ];
    for (fault, code, undoing_helps) in faults {
        assert_eq!(fault.code(), code);
        assert_eq!(fault.undoing_helps(), undoing_helps, "{code}");
    }
    let codes: std::collections::BTreeSet<_> =
        faults.iter().map(|(fault, ..)| fault.code()).collect();
    assert_eq!(codes.len(), faults.len(), "two faults answer to one code");
}

/// Errors that are not about writing still reach a save, through the snapshot
/// the projection takes before it writes anything. Reading one of those as a
/// refusal would offer to throw away work over a bad client ID.
#[test]
fn an_error_the_writer_did_not_raise_is_never_a_refusal() {
    for error in [
        EditError::InvalidClientId(0),
        EditError::Parse("truncated".to_owned()),
        EditError::InvalidState("no such story".to_owned()),
        EditError::InvalidUpdate("short".to_owned()),
        EditError::Observer("listener".to_owned()),
        EditError::Json("boundary".to_owned()),
    ] {
        assert_eq!(error.save_fault(), SaveFault::WriteFailed, "{error}");
        assert!(!error.save_fault().undoing_helps(), "{error}");
    }
}

/// `undoing_helps` is a promise, not a label: taking the named change back out
/// has to leave a deck that saves.
#[test]
fn taking_back_a_refused_change_lets_the_same_save_through() {
    let session = session_with(r#"<a:p><a:r><a:t>ab</a:t></a:r><a:br/></a:p>"#);
    let story = first_story(&session);
    type_text(&session, &story.id, 3, "X");

    let error = session.save_bytes().unwrap_err();
    assert_eq!(error.save_fault(), SaveFault::Unprojectable, "{error}");
    assert!(error.save_fault().undoing_helps());

    session
        .delete_text(&EditCtx::local("test"), &story.id, 3, 4)
        .unwrap();
    session
        .save_bytes()
        .expect("the deck saves once the change the refusal named is gone");
}

/// The slide's only sp spells an explicit `<a:xfrm>`, so a move and a resize
/// splice its `<a:off>`/`<a:ext>` in place and everything round-trips.
#[test]
fn a_moved_and_resized_shape_saves_and_reads_back() {
    let session = session_with(r#"<a:p><a:r><a:t>anchor</a:t></a:r></a:p>"#);
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = &slide.shapes[0];
    let context = EditCtx::local("test");
    session
        .move_shape(&context, &slide.id, &shape.id, 914_400, 457_200)
        .unwrap();
    session
        .resize_shape(&context, &slide.id, &shape.id, 5_000_000, 2_000_000)
        .unwrap();

    let saved = session.save_bytes().unwrap();
    let slide_xml = part_text(&saved, SLIDE_PART);
    assert!(
        slide_xml.contains(r#"<a:off x="914400" y="457200"/>"#),
        "{slide_xml}"
    );
    assert!(
        slide_xml.contains(r#"<a:ext cx="5000000" cy="2000000"/>"#),
        "{slide_xml}"
    );

    let reopened = DeckSession::open(&saved, 78).unwrap();
    let shape = reopened.snapshot().unwrap().slides[0].shapes[0].clone();
    assert_eq!(
        (shape.x, shape.y, shape.width, shape.height),
        (914_400, 457_200, 5_000_000, 2_000_000)
    );
}

/// A move combined with a text edit lands both in the same part.
#[test]
fn a_move_and_a_text_edit_share_one_save() {
    let session = session_with(r#"<a:p><a:r><a:t>ab</a:t></a:r></a:p>"#);
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = &slide.shapes[0];
    let story = first_story(&session);
    type_text(&session, &story.id, 2, "c");
    session
        .move_shape(&EditCtx::local("test"), &slide.id, &shape.id, 111, 222)
        .unwrap();

    let saved = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&saved, 79).unwrap();
    let shape = reopened.snapshot().unwrap().slides[0].shapes[0].clone();
    assert_eq!((shape.x, shape.y), (111, 222));
    let story = first_story(&reopened);
    assert_eq!(story.paragraphs[0].runs[0].text, "abc");
}

/// A placeholder that inherits its placement from the layout has no
/// `<a:xfrm>` to rewrite; moving it is refused as the user's change.
#[test]
fn moving_a_shape_without_an_explicit_transform_is_refused() {
    let deck = {
        let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let xml = slide_xml(r#"<a:p><a:r><a:t>x</a:t></a:r></a:p>"#).replace(
            r#"<a:xfrm><a:off x="0" y="0"/><a:ext cx="6096000" cy="1981200"/></a:xfrm>"#,
            "",
        );
        assert!(package.replace_part(SLIDE_PART, xml.into_bytes()));
        pptx_parse::write_pptx(&package).unwrap()
    };
    let session = DeckSession::open(&deck, 80).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = &slide.shapes[0];
    session
        .move_shape(&EditCtx::local("test"), &slide.id, &shape.id, 5, 6)
        .unwrap();

    let error = session.save_bytes().unwrap_err();
    assert_eq!(error.save_fault(), SaveFault::Unprojectable, "{error}");
    assert!(error.to_string().contains("layout"), "{error}");
}

/// A removed shape's element is cut from the slide; the rest keeps its bytes.
#[test]
fn a_removed_shape_saves_and_stays_gone() {
    let session = DeckSession::open(FIXTURE, 95).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let victim = slide.shapes[0].clone();
    session
        .remove_shape(&EditCtx::local("test"), &slide.id, &victim.id)
        .unwrap();

    let saved = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&saved, 96).unwrap();
    let reopened_slide = &reopened.snapshot().unwrap().slides[0];
    assert_eq!(reopened_slide.shapes.len(), slide.shapes.len() - 1);
    assert!(
        reopened_slide
            .shapes
            .iter()
            .all(|shape| shape.name != victim.name),
        "the removed shape does not come back"
    );
}

/// A removal combined with a text edit in a surviving shape lands both.
#[test]
fn a_removal_and_a_text_edit_share_one_save() {
    let session = session_with(r#"<a:p><a:r><a:t>keep</a:t></a:r></a:p>"#);
    let snapshot = session.snapshot().unwrap();
    let other_slide = &snapshot.slides[1];
    let victim = other_slide.shapes[0].clone();
    session
        .remove_shape(&EditCtx::local("test"), &other_slide.id, &victim.id)
        .unwrap();
    let story = first_story(&session);
    type_text(&session, &story.id, 4, "!");

    let saved = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&saved, 97).unwrap();
    let story = first_story(&reopened);
    assert_eq!(story.paragraphs[0].runs[0].text, "keep!");
    assert_eq!(
        reopened.snapshot().unwrap().slides[1].shapes.len(),
        other_slide.shapes.len() - 1
    );
}

/// A text box added this session is synthesised into the slide and survives
/// the round trip with its text and style.
#[test]
fn an_added_text_box_saves_and_reads_back() {
    let session = DeckSession::open(FIXTURE, 98).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    session
        .add_text_box(
            &EditCtx::local("test"),
            &slide.id,
            &pptx_edit::ShapeDraft {
                name: "Note".to_owned(),
                rect: pptx_edit::ShapeRect {
                    x: 914_400,
                    y: 914_400,
                    width: 2_000_000,
                    height: 500_000,
                },
                text: "hello box".to_owned(),
                style: TextStyle {
                    bold: Some(true),
                    color: Some("#325EE6".to_owned()),
                    ..TextStyle::default()
                },
            },
        )
        .unwrap();

    let saved = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&saved, 99).unwrap();
    let reopened_slide = &reopened.snapshot().unwrap().slides[0];
    assert_eq!(reopened_slide.shapes.len(), slide.shapes.len() + 1);
    let added = reopened_slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Note")
        .expect("the new text box is in the reopened deck");
    assert_eq!(
        (added.x, added.y, added.width, added.height),
        (914_400, 914_400, 2_000_000, 500_000)
    );
    let run = &added.text_stories[0].paragraphs[0].runs[0];
    assert_eq!(run.text, "hello box");
    assert_eq!(run.style.bold, Some(true));
    assert_eq!(run.style.color.as_deref(), Some("#325EE6"));
}

/// A preset shape with a fill and default adjustments round-trips.
#[test]
fn an_added_preset_shape_saves_with_fill_and_adjustments() {
    let session = DeckSession::open(FIXTURE, 100).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    session
        .add_shape(
            &EditCtx::local("test"),
            &slide.id,
            &pptx_edit::PresetShapeDraft {
                name: "Badge".to_owned(),
                geometry: "roundRect".to_owned(),
                rect: pptx_edit::ShapeRect {
                    x: 0,
                    y: 0,
                    width: 1_000_000,
                    height: 1_000_000,
                },
                fill: Some("#FF0000".to_owned()),
            },
        )
        .unwrap();

    let saved = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&saved, 101).unwrap();
    let added = reopened.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Badge")
        .cloned()
        .expect("the new shape is in the reopened deck");
    assert_eq!(added.geometry, "roundRect");
    assert!(!added.adjust_values.is_empty(), "adjustments survive");
    assert_eq!(
        added
            .fill
            .as_ref()
            .and_then(|fill| fill.color.as_ref())
            .and_then(|color| color.rgb.as_deref()),
        Some("FF0000")
    );
}

/// Adding, removing and typing in one save all land together.
#[test]
fn an_add_a_remove_and_a_text_edit_share_one_save() {
    let session = session_with(r#"<a:p><a:r><a:t>base</a:t></a:r></a:p>"#);
    let snapshot = session.snapshot().unwrap();
    let second = &snapshot.slides[1];
    let victim = second.shapes[0].clone();
    session
        .remove_shape(&EditCtx::local("test"), &second.id, &victim.id)
        .unwrap();
    session
        .add_text_box(
            &EditCtx::local("test"),
            &snapshot.slides[0].id,
            &pptx_edit::ShapeDraft {
                name: "Note".to_owned(),
                rect: pptx_edit::ShapeRect {
                    x: 0,
                    y: 0,
                    width: 914_400,
                    height: 914_400,
                },
                text: "n".to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap();
    let story = first_story(&session);
    type_text(&session, &story.id, 4, "!");

    let saved = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&saved, 102).unwrap();
    let reopened_snapshot = reopened.snapshot().unwrap();
    assert_eq!(first_story(&reopened).paragraphs[0].runs[0].text, "base!");
    assert_eq!(
        reopened_snapshot.slides[1].shapes.len(),
        second.shapes.len() - 1
    );
    assert!(
        reopened_snapshot.slides[0]
            .shapes
            .iter()
            .any(|shape| shape.name == "Note")
    );
}

/// A shape an animation still targets cannot be deleted: the timing block is
/// invisible to the parse model, so the writer refuses rather than leaving a
/// dangling id behind.
#[test]
fn removing_an_animated_shape_is_refused_by_name() {
    let deck = {
        let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let xml = slide_xml(r#"<a:p><a:r><a:t>x</a:t></a:r></a:p>"#).replace(
            "</p:cSld></p:sld>",
            r#"</p:cSld><p:timing><p:spTgt spid="2"/></p:timing></p:sld>"#,
        );
        assert!(package.replace_part(SLIDE_PART, xml.into_bytes()));
        pptx_parse::write_pptx(&package).unwrap()
    };
    let session = DeckSession::open(&deck, 103).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    session
        .remove_shape(&EditCtx::local("test"), &slide.id, &slide.shapes[0].id)
        .unwrap();

    let error = session.save_bytes().unwrap_err();
    assert_eq!(error.save_fault(), SaveFault::Unprojectable, "{error}");
    assert!(error.to_string().contains("animation"), "{error}");
}

/// A reordered deck rewrites only the presentation's slide list; the slide
/// parts keep their bytes and the reopened deck shows the new order.
#[test]
fn a_reordered_deck_saves_and_reads_back() {
    let session = DeckSession::open(FIXTURE, 104).unwrap();
    let snapshot = session.snapshot().unwrap();
    let names: Vec<String> = snapshot
        .slides
        .iter()
        .map(|slide| slide.id.clone())
        .collect();
    session
        .move_slide(&EditCtx::local("test"), &names[0], 2)
        .unwrap();

    let saved = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&saved, 105).unwrap();
    let reopened_snapshot = reopened.snapshot().unwrap();
    assert_eq!(reopened_snapshot.slides.len(), 3);
    // Slide part paths rotate with the order.
    let original_paths: Vec<_> = snapshot
        .slides
        .iter()
        .map(|s| s.source_part_path.clone())
        .collect();
    let reopened_paths: Vec<_> = reopened_snapshot
        .slides
        .iter()
        .map(|s| s.source_part_path.clone())
        .collect();
    assert_eq!(
        reopened_paths,
        vec![
            original_paths[1].clone(),
            original_paths[2].clone(),
            original_paths[0].clone()
        ]
    );
}

/// A deleted slide drops out of the slide list; its part goes unreferenced
/// rather than corrupt, and the other slides keep their bytes.
#[test]
fn a_deleted_slide_saves_and_stays_gone() {
    let session = DeckSession::open(FIXTURE, 106).unwrap();
    let snapshot = session.snapshot().unwrap();
    let victim = snapshot.slides[1].clone();
    session
        .delete_slide(&EditCtx::local("test"), &victim.id)
        .unwrap();

    let saved = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&saved, 107).unwrap();
    let reopened_snapshot = reopened.snapshot().unwrap();
    assert_eq!(reopened_snapshot.slides.len(), 2);
    assert!(
        reopened_snapshot
            .slides
            .iter()
            .all(|slide| slide.source_part_path != victim.source_part_path),
        "the deleted slide does not come back"
    );
}

/// Deleting one slide, reordering the rest and typing into a survivor all
/// land in one save.
#[test]
fn slide_structure_and_text_edits_share_one_save() {
    let session = session_with(r#"<a:p><a:r><a:t>first</a:t></a:r></a:p>"#);
    let snapshot = session.snapshot().unwrap();
    // The story of the slide that is about to move, captured before the
    // structural edits shuffle the snapshot order.
    let story = first_story(&session);
    session
        .delete_slide(&EditCtx::local("test"), &snapshot.slides[2].id)
        .unwrap();
    session
        .move_slide(&EditCtx::local("test"), &snapshot.slides[0].id, 1)
        .unwrap();
    type_text(&session, &story.id, 5, "!");

    let saved = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&saved, 108).unwrap();
    let reopened_snapshot = reopened.snapshot().unwrap();
    assert_eq!(reopened_snapshot.slides.len(), 2);
    assert_eq!(
        reopened_snapshot.slides[1].source_part_path, snapshot.slides[0].source_part_path,
        "the first slide moved behind the second"
    );
    let moved_story = reopened_snapshot.slides[1]
        .shapes
        .iter()
        .find_map(|shape| shape.text_stories.first())
        .expect("the moved slide keeps its story");
    assert_eq!(moved_story.paragraphs[0].runs[0].text, "first!");
}

/// A slide the deck created is spelled as a fresh part: it joins the slide
/// list, references its layout, and its shapes survive the round trip.
#[test]
fn an_inserted_slide_saves_and_reads_back() {
    let session = DeckSession::open(FIXTURE, 109).unwrap();
    let snapshot = session.snapshot().unwrap();
    let layout = snapshot.slides[0].layout_part_path.clone();
    let receipt = session
        .insert_slide(&EditCtx::local("test"), 1, layout.as_deref())
        .unwrap();
    session
        .add_text_box(
            &EditCtx::local("test"),
            &receipt.slide_id,
            &pptx_edit::ShapeDraft {
                name: "Fresh".to_owned(),
                rect: pptx_edit::ShapeRect {
                    x: 914_400,
                    y: 914_400,
                    width: 3_000_000,
                    height: 600_000,
                },
                text: "brand new".to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap();

    let saved = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&saved, 110).unwrap();
    let reopened_snapshot = reopened.snapshot().unwrap();
    assert_eq!(reopened_snapshot.slides.len(), 4);
    let inserted = &reopened_snapshot.slides[1];
    assert_eq!(inserted.layout_part_path, layout);
    assert_eq!(inserted.shapes.len(), 1);
    assert_eq!(
        inserted.shapes[0].text_stories[0].paragraphs[0].runs[0].text,
        "brand new"
    );
    // The other slides keep their identities and order.
    assert_eq!(
        reopened_snapshot.slides[0].source_part_path,
        snapshot.slides[0].source_part_path
    );
    assert_eq!(
        reopened_snapshot.slides[2].source_part_path,
        snapshot.slides[1].source_part_path
    );
}

/// An inserted slide without a layout still saves, with no rels part at all.
#[test]
fn an_inserted_slide_without_a_layout_saves() {
    let session = DeckSession::open(FIXTURE, 111).unwrap();
    session
        .insert_slide(&EditCtx::local("test"), 3, None)
        .unwrap();

    let saved = session.save_bytes().unwrap();
    let reopened = DeckSession::open(&saved, 112).unwrap();
    let reopened_snapshot = reopened.snapshot().unwrap();
    assert_eq!(reopened_snapshot.slides.len(), 4);
    assert_eq!(reopened_snapshot.slides[3].layout_part_path, None);
    assert!(reopened_snapshot.slides[3].shapes.is_empty());
}

/// Deleting a whole group refuses when an animation targets one of its
/// children — the child's id leaves with the group.
#[test]
fn removing_a_group_with_an_animated_child_is_refused() {
    let deck = {
        let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let xml = concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#,
            r#" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#,
            r#" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">"#,
            r#"<p:cSld><p:spTree>"#,
            r#"<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
            r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"#,
            r#"<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/>"#,
            r#"<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>"#,
            r#"<p:sp><p:nvSpPr><p:cNvPr id="3" name="Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"#,
            r#"<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr></p:sp>"#,
            r#"</p:grpSp>"#,
            r#"</p:spTree></p:cSld><p:timing><p:spTgt spid="3"/></p:timing></p:sld>"#,
        );
        assert!(package.replace_part(SLIDE_PART, xml.as_bytes().to_vec()));
        pptx_parse::write_pptx(&package).unwrap()
    };
    let session = DeckSession::open(&deck, 113).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    session
        .remove_shape(&EditCtx::local("test"), &slide.id, &slide.shapes[0].id)
        .unwrap();

    let error = session.save_bytes().unwrap_err();
    assert_eq!(error.save_fault(), SaveFault::Unprojectable, "{error}");
    assert!(error.to_string().contains("animation"), "{error}");
}

/// Deleting a slide a custom show still presents refuses by name.
#[test]
fn deleting_a_slide_in_a_custom_show_is_refused() {
    let deck = {
        let mut package = pptx_parse::parse_pptx(FIXTURE).unwrap();
        let presentation_path = package.presentation.part_path.clone();
        let bytes = package.part_bytes(&presentation_path).unwrap().to_vec();
        let text = String::from_utf8(bytes).unwrap();
        let with_show = text.replace(
            "</p:presentation>",
            r#"<p:custShowLst><p:custShow name="short" id="0"><p:sldLst><p:sld r:id="rId3"/></p:sldLst></p:custShow></p:custShowLst></p:presentation>"#,
        );
        assert!(package.replace_part(&presentation_path, with_show.into_bytes()));
        pptx_parse::write_pptx(&package).unwrap()
    };
    let session = DeckSession::open(&deck, 114).unwrap();
    let snapshot = session.snapshot().unwrap();
    let shown = snapshot
        .slides
        .iter()
        .find(|slide| {
            slide.source_part_path.as_deref().is_some_and(|path| {
                pptx_parse::parse_pptx(&deck)
                    .unwrap()
                    .presentation
                    .slides
                    .iter()
                    .any(|entry| entry.relationship_id == "rId3" && entry.part_path == path)
            })
        })
        .cloned();
    let Some(shown) = shown else {
        panic!("the fixture has no slide bound to rId3");
    };
    session
        .delete_slide(&EditCtx::local("test"), &shown.id)
        .unwrap();

    let error = session.save_bytes().unwrap_err();
    assert_eq!(error.save_fault(), SaveFault::Unprojectable, "{error}");
    assert!(error.to_string().contains("custom show"), "{error}");
}
