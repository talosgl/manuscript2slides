# manuscript2slides

This folder was created automatically by manuscript2slides.

## What's in here?

### 📁 input/
Optional staging area for your Word documents.
- You can put `.docx` files here, or point the program to files anywhere on your computer
- Not required to use this - just a convenient place if you want it
- Includes `sample_doc.docx` and `sample_slides_output.pptx` for testing
- If you don't specify any input document, manuscript2slides will pull the sample_doc.docx from here to do a dry run

### 📁 output/
Where your converted PowerPoint files are saved.
- Default save location for `.pptx` files
- Each file gets a timestamp so nothing gets overwritten

### 📁 templates/
PowerPoint and Word template files used for conversions.
- `pptx_template.pptx` - The slide deck template (customize fonts/colors if you want)
- `docx_template.docx` - The document template for reverse conversions
- Feel free to modify these; the program will use your customized versions. However, it *must* still contain a master slide template named "manuscript2slides".

### 📁 logs/
Program logs for debugging.
- `manuscript2slides.log` - What happened during each run
- Check here if something goes wrong
- Safe to delete these files anytime

## Need Help?

Check the main project documentation or open an issue on GitHub.

---

**Note:** This entire folder is safe to delete if you want to reset everything.
The program will recreate it automatically on next run.