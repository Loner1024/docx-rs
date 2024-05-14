

#[cfg(test)]
mod tests {
    use crate::docx::*;
    struct TestCase<'a> {
        name: String,
        replace_with: &'a str,
        expect: String,
    }
    const TEST_FILE: &str = "./TestDocument.docx";
    const TEST_FILE_RESULT: &str = "./TestDocumentResult.docx";

    #[test]
    fn test_replace() {
        let cases = vec![
            TestCase{name:"Windows line breaks".to_string(), replace_with:"line1\r\nline2", expect:"line1<w:br/>line2".to_string()},
            TestCase{name:"Mac line breaks".to_string(), replace_with:"line1\rline2", expect: "line1<w:br/>line2".to_string()},
            TestCase{name:"Linux line breaks".to_string(), replace_with:"line1\nline2", expect:"line1<w:br/>line2".to_string()},
            TestCase{name:"Tabs".to_string(), replace_with:"line1\tline2", expect:"line1</w:t><w:tab/><w:t>line2".to_string()},
        ];
        for case in cases {
            let mut doc = Docx::new(TEST_FILE).unwrap();
            doc.replace("document.", case.replace_with, 1).unwrap();
            doc.write_to_file(TEST_FILE_RESULT).unwrap();
            let doc = Docx::new(TEST_FILE_RESULT).unwrap();
            assert!(doc.get_content().contains(&case.expect), "{} 失败: {}", case.name, doc.get_content());
        }
    }
}