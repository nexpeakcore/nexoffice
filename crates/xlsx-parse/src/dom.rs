//! A small owned XML tree for the low-volume parts that are easier to walk
//! than to stream — drawings and their charts. Worksheets stay streamed.

use std::collections::BTreeMap;

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::ParseError;

const MAX_DOM_BYTES: usize = 8 * 1024 * 1024;
const MAX_DOM_DEPTH: usize = 64;
const MAX_DOM_NODES: usize = 200_000;

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct XmlElement {
    pub name: String,
    pub attributes: BTreeMap<String, String>,
    pub children: Vec<XmlNode>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum XmlNode {
    Element(XmlElement),
    Text(String),
}

impl XmlElement {
    pub fn local_name(&self) -> &str {
        local_name(&self.name)
    }

    pub fn attribute(&self, name: &str) -> Option<&str> {
        self.attributes.get(name).map(String::as_str)
    }

    pub fn attribute_local(&self, name: &str) -> Option<&str> {
        self.attributes
            .iter()
            .find(|(key, _)| local_name(key) == name)
            .map(|(_, value)| value.as_str())
    }

    pub fn child(&self, name: &str) -> Option<&XmlElement> {
        self.child_elements()
            .find(|child| child.local_name() == name)
    }

    pub fn children_named<'a>(
        &'a self,
        name: &'a str,
    ) -> impl Iterator<Item = &'a XmlElement> + 'a {
        self.child_elements()
            .filter(move |child| child.local_name() == name)
    }

    pub fn child_elements(&self) -> impl Iterator<Item = &XmlElement> {
        self.children.iter().filter_map(|child| match child {
            XmlNode::Element(element) => Some(element),
            XmlNode::Text(_) => None,
        })
    }

    pub fn descendants_named<'a>(&'a self, name: &'a str) -> Vec<&'a XmlElement> {
        let mut output = Vec::new();
        collect_descendants(self, name, &mut output);
        output
    }

    pub fn text_content(&self) -> String {
        let mut output = String::new();
        append_text(self, &mut output);
        output
    }
}

fn collect_descendants<'a>(element: &'a XmlElement, name: &str, output: &mut Vec<&'a XmlElement>) {
    for child in element.child_elements() {
        if child.local_name() == name {
            output.push(child);
        }
        collect_descendants(child, name, output);
    }
}

fn append_text(element: &XmlElement, output: &mut String) {
    for child in &element.children {
        match child {
            XmlNode::Element(element) => append_text(element, output),
            XmlNode::Text(text) => output.push_str(text),
        }
    }
}

pub(crate) fn local_name(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}

pub(crate) fn parse_dom(xml: &[u8], part: &str) -> Result<XmlElement, ParseError> {
    if xml.len() > MAX_DOM_BYTES {
        return Err(ParseError::Malformed(format!("{part}: part too large")));
    }
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().check_end_names = true;
    let mut stack: Vec<XmlElement> = Vec::new();
    let mut root: Option<XmlElement> = None;
    let mut nodes = 0usize;

    loop {
        let event = reader
            .read_event()
            .map_err(|error| ParseError::Malformed(format!("{part}: {error}")))?;
        nodes += 1;
        if nodes > MAX_DOM_NODES {
            return Err(ParseError::Malformed(format!("{part}: too many nodes")));
        }
        match event {
            Event::Start(start) => {
                if stack.len() >= MAX_DOM_DEPTH {
                    return Err(ParseError::Malformed(format!("{part}: nested too deeply")));
                }
                stack.push(decode(&reader, &start, part)?);
            }
            Event::Empty(start) => {
                let element = decode(&reader, &start, part)?;
                attach(element, &mut stack, &mut root, part)?;
            }
            Event::End(_) => {
                let element = stack
                    .pop()
                    .ok_or_else(|| ParseError::Malformed(format!("{part}: unbalanced close")))?;
                attach(element, &mut stack, &mut root, part)?;
            }
            Event::Text(text) => {
                if let Some(parent) = stack.last_mut() {
                    let value = text
                        .decode()
                        .map_err(|error| ParseError::Malformed(format!("{part}: {error}")))?;
                    parent.children.push(XmlNode::Text(value.into_owned()));
                }
            }
            Event::CData(data) => {
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(XmlNode::Text(
                        String::from_utf8_lossy(data.as_ref()).into_owned(),
                    ));
                }
            }
            Event::GeneralRef(reference) => {
                let decoded = reference
                    .decode()
                    .map_err(|error| ParseError::Malformed(format!("{part}: {error}")))?;
                let resolved = if reference.is_char_ref() {
                    reference
                        .resolve_char_ref()
                        .map_err(|error| ParseError::Malformed(format!("{part}: {error}")))?
                        .map(|character| character.to_string())
                } else {
                    quick_xml::escape::resolve_predefined_entity(&decoded).map(str::to_owned)
                };
                let resolved = resolved.ok_or_else(|| {
                    ParseError::Malformed(format!("{part}: unresolvable entity &{decoded};"))
                })?;
                if let Some(parent) = stack.last_mut() {
                    parent.children.push(XmlNode::Text(resolved));
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    root.ok_or_else(|| ParseError::Malformed(format!("{part}: no root element")))
}

fn decode(
    reader: &Reader<&[u8]>,
    start: &quick_xml::events::BytesStart<'_>,
    part: &str,
) -> Result<XmlElement, ParseError> {
    let _ = reader;
    let name = String::from_utf8_lossy(start.name().into_inner()).into_owned();
    let mut attributes = BTreeMap::new();
    for attribute in start.attributes() {
        let attribute =
            attribute.map_err(|error| ParseError::Malformed(format!("{part}: {error}")))?;
        attributes.insert(
            String::from_utf8_lossy(attribute.key.into_inner()).into_owned(),
            String::from_utf8_lossy(&attribute.value).into_owned(),
        );
    }
    Ok(XmlElement {
        name,
        attributes,
        children: Vec::new(),
    })
}

fn attach(
    element: XmlElement,
    stack: &mut [XmlElement],
    root: &mut Option<XmlElement>,
    part: &str,
) -> Result<(), ParseError> {
    match stack.last_mut() {
        Some(parent) => {
            parent.children.push(XmlNode::Element(element));
            Ok(())
        }
        None if root.is_none() => {
            *root = Some(element);
            Ok(())
        }
        None => Err(ParseError::Malformed(format!("{part}: multiple roots"))),
    }
}
