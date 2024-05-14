use std::cell::{RefCell, RefMut};
use std::collections::{HashMap, HashSet};
use std::fs::File;
use std::io::{Read, Write};
use std::rc::Rc;

use anyhow::Result;
use zip::read::ZipArchive;
use zip::write::{SimpleFileOptions, ZipWriter};

pub struct DocZipData {
    data: Rc<RefCell<ZipArchive<File>>>,
}

impl DocZipData {
    pub fn from_file(file_name: &str) -> Result<Self> {
        let file = File::open(file_name)?;
        let archive = ZipArchive::new(file)?;
        Ok(Self {
            data: Rc::new(RefCell::new(archive)),
        })
    }

    pub fn get_zip_data(&self) -> RefMut<ZipArchive<File>> {
        self.data.borrow_mut()
    }

    fn read_text(&mut self) -> Result<String> {
        let mut text = String::new();
        self.data
            .borrow_mut()
            .by_name("word/document.xml")?
            .read_to_string(&mut text)?;
        Ok(text)
    }

    fn read_links(&mut self) -> Result<String> {
        let mut links = String::new();
        self.data
            .borrow_mut()
            .by_name("word/_rels/document.xml.rels")?
            .read_to_string(&mut links)?;
        Ok(links)
    }

    fn read_headers(&mut self) -> Result<HashMap<String, String>> {
        let mut headers = HashMap::new();
        let mut data = self.data.borrow_mut();
        for i in 0..data.len() {
            let mut file = data.by_index(i)?;
            if file.name().contains("header") {
                let mut value = String::new();
                file.read_to_string(&mut value)?;
                headers.insert(file.name().to_string(), value);
            }
        }
        Ok(headers)
    }

    fn read_footers(&mut self) -> Result<HashMap<String, String>> {
        let mut footers = HashMap::new();
        let mut data = self.data.borrow_mut();
        for i in 0..data.len() {
            let mut file = data.by_index(i)?;
            if file.name().contains("footer") {
                let mut value = String::new();
                file.read_to_string(&mut value)?;
                footers.insert(file.name().to_string(), value);
            }
        }
        Ok(footers)
    }

    fn read_images(&mut self) -> Result<HashSet<String>> {
        let mut images = HashSet::new();
        let mut data = self.data.borrow_mut();
        for i in 0..data.len() {
            let file = data.by_index(i)?;
            if file.name().starts_with("word/media/") {
                images.insert(file.name().to_string());
            }
        }
        Ok(images)
    }
}

pub struct Docx {
    data: DocZipData,
    contents: String,
    links: String,
    headers: HashMap<String, String>,
    footers: HashMap<String, String>,
    images: HashSet<String>,
}

impl Docx {
    pub fn new(file: &str) -> Result<Self> {
        let mut data = DocZipData::from_file(file)?;
        let contents = data.read_text()?;
        let links = data.read_links()?;
        let headers = data.read_headers()?;
        let footers = data.read_footers()?;
        let images = data.read_images()?;
        Ok(Self {
            data,
            contents,
            links,
            headers,
            footers,
            images,
        })
    }

    pub fn get_content(&self) -> &str {
        &self.contents
    }

    pub fn replace(&mut self, old_str: &str, new_str: &str, count: usize) -> Result<()> {
        println!("old: {}, new: {}", old_str, new_str);
        let old_str = Docx::encode(old_str);
        let new_str = Docx::encode(new_str);
        println!("old: {}, new: {}", old_str, new_str);
        self.contents = self
            .contents
            .replacen(old_str.as_str(), new_str.as_str(), count);
        Ok(())
    }

    fn encode(s: &str) -> String {
        const TAB: &str = "</w:t><w:tab/><w:t>";
        const NEWLINE: &str = "<w:br/>";

        let mut output = s.to_string();
        output = output.replace("<string>", "");
        output = output.replace("</string>", "");
        output = output.replace("\r\n", NEWLINE); // \r\n (Windows newline)
        output = output.replace('\r', NEWLINE); // \r (earlier Mac newline)
        output = output.replace('\n', NEWLINE); // \n (unix/linux/OS X newline)
        output = output.replace('\t', TAB); // \t (tab)
        output
    }

    pub fn write_to_file(&mut self, file_name: &str) -> Result<()> {
        let file = File::create(file_name)?;
        let mut zip_writer = ZipWriter::new(file);
        let options = SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Stored)
            .unix_permissions(0o755);
        let mut data = self.data.get_zip_data();
        for i in 0..data.len() {
            let mut file = data.by_index(i)?;
            zip_writer.start_file(file.name(), options)?;
            if file.name() == "word/document.xml" {
                zip_writer.write_all(self.contents.as_bytes())?;
            } else if file.name() == "word/_rels/document.xml.rels" {
                zip_writer.write_all(self.links.as_bytes())?;
            } else if file.name().contains("header") && self.headers.contains_key(file.name()) {
                zip_writer.write_all(
                    self.headers
                        .get(file.name())
                        .unwrap_or(&"".to_string())
                        .as_bytes(),
                )?;
            } else if file.name().contains("footer") && self.footers.contains_key(file.name()) {
                zip_writer.write_all(
                    self.footers
                        .get(file.name())
                        .unwrap_or(&"".to_string())
                        .as_bytes(),
                )?;
            } else {
                let mut data = Vec::new();
                file.read_to_end(&mut data)?;
                zip_writer.write_all(&data)?;
            }
        }
        zip_writer.finish()?;

        Ok(())
    }
}
