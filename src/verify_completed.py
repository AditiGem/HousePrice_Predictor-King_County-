import docx
import os

doc = docx.Document("completed.docx")
print("File size:", os.path.getsize("completed.docx"))
print("Total paragraphs:", len(doc.paragraphs))
print("Total tables:", len(doc.tables))

figs = [p.text for p in doc.paragraphs if p.text.startswith("Figure")]
print(f"Total figures: {len(figs)}")
for f in figs:
    print("  ", f)

# Count table fill colors to ensure NO blue backgrounds
for i, tbl in enumerate(doc.tables):
    shds = set()
    for row in tbl.rows:
        for c in row.cells:
            tcPr = c._element.tcPr
            if tcPr is not None:
                for shd in tcPr.findall("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd"):
                    shds.add(shd.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill"))
    print(f"Table {i+1} fill colors:", shds)

# Count chapters
chapters = [p.text for p in doc.paragraphs if p.text.startswith("Chapter")]
print("\nChapters:")
for ch in chapters:
    print("  ", ch)
