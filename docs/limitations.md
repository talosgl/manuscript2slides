## Known Issues & Limitations

This document lists current limitations, design constraints, and known issues in `manuscript2slides`.

Many of these stem from the Microsoft Office Open XML formats and the Python libraries used.

---

## 1. File and Format Support

- **Supported formats:** Only `.docx` and `.pptx`. Older formats (`.doc`, `.ppt`) are not supported. If you have `.doc` files, convert them first using Word, Google Docs, or LibreOffice.

- **Text-only conversion:** We only support text content. No images, tables, etc., are copied between the formats, and we do not have plans to support these in future.

- **Sections:** "Sections" in both docx and pptx are not supported.

---
## 2. Formatting and Rendering

### Formatting Preservation: Basic & Advanced/Experimental
Basic formatting (bold, italic, underline, etc.) is always preserved. Advanced formatting (highlighting, indentation, mixed inline styles) requires "experimental formatting" to be enabled. This is turned on by default. You can disable it if you see issues during conversion.

| **Basic Formatting**                | **Advanced / Experimental Formatting**             |
| ----------------------------------- | -------------------------------------------------- |
| bold, italic, underline             | highlighting                                       |
| strikethrough                       | paragraph indentation / spacing                    |
| superscript / subscript             | font colors & mixed inline runs                    |
| font size                           | nested formatting styles                           |
| basic alignment (left/center/right) | any multi-run or partially formatted text handling |


### Field-code hyperlinks:
Some hyperlinks (like the sample_doc.docx's first "Where are Data?" link) (often in headings or pasted content) are stored as "field codes" in Word and appear as plain text after conversion. These are uncommon; normal hyperlinks should work. manuscript2slides will log a warning when such links are detected and will try to copy them as plaintext.

### Auto-fit text resizing:
PowerPoint's automatic 'shrink text on overflow' feature is not applied programmatically. PowerPoint only applies auto-fit sizing when opened in the UI. To manually fix this:

1. Open up the output presentation in PowerPoint Desktop > View > Slide Master
2. Select the text frame object, right-click > Format Shape
3. Click the Size & Properties icon {TODO, doc: ADD SCREENCAPS}
4. Click Text Box to see the options
5. Toggle "Do not Autofit" and then back to "Shrink Text on overflow"
6. Close Master View
7. Now all the slides should have their text properly resized.


## 3. Annotation and Metadata Handling
PowerPoint does not natively support true footnotes or endnotes, and the underlying library (`python-pptx`) does not yet support adding comments.

By default, no annotations are preserved during DOCX -> PPTX conversion; however, we provide the optional feature to have each type (or all) annotations copied into the speaker notes of a relevant slide.

Note that:
- All selected annotations (comments, footnotes, endnotes) are combined and copied into a slide's speaker notes.
- Inline reference numbers (e.g. `¹`, `²`, `³`) are not preserved in the slide text.
- Comment threading is not preserved. 
- When 'display comments' is selected, the GUI will preserve the comment's text body, author, timestamp. The CLI allows you to disable copying comment metadata (author, timestamp), if you want.

## 4. Reverse Pipeline (PPTX → DOCX)
The reverse pipeline (PPTX → DOCX) is significantly less robust than the forward conversion. Your original input document to the DOCX -> PPTX flow and the output document from a follow-up PPTX -> DOCX flow will not look the same, but any prose iteration work you've done in the slide body will be preserved.

If you want us to attempt to restore advanced formatting during round-trip conversion, then when you first convert a docx -> pptx, check the **Advanced Options > Preserve metadata in speaker notes** option in the UI, or in the CLI, pass `preserve_docx_metadata_in_speaker_notes = true`. With this option enabled, manuscript2slides injects compact JSON metadata into each slide's speaker notes. During a later reverse (PPTX → DOCX) conversion, that metadata helps restore highlighting, heading formatting, and comments more accurately.


- **Lost elements:**  
  Images, tables, charts, footnotes, and endnotes are not restored. (If footnotes and endnotes are preserved in the metadata, they're restored as plaintext in a comment.)

- **Advanced formatting matching:**  
  Headings, highlighting, and other advanced formatting are re-applied using approximate text matching only.

- **Comment restoration:**  
  Comments are re-inserted, but their anchor positions may differ slightly.

- **Blank paragraph at start:**  
  When creating a new `.docx`, the `python-docx` library inserts an empty paragraph as required by the Open XML spec (the document `<w:body>` must contain at least one paragraph).