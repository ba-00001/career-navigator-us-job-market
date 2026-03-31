from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape
import zipfile


OUTPUT = Path("Group Project Interim Progress Report - Completed.docx")


def paragraph(text, style=None, bold=False):
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    if not text:
        return f"<w:p>{style_xml}</w:p>"
    run_props = "<w:rPr><w:b/></w:rPr>" if bold else ""
    return (
        f"<w:p>{style_xml}<w:r>{run_props}"
        f"<w:t xml:space=\"preserve\">{escape(text)}</w:t></w:r></w:p>"
    )


def table(rows):
    tbl_pr = (
        "<w:tblPr>"
        "<w:tblW w:w=\"0\" w:type=\"auto\"/>"
        "<w:tblBorders>"
        "<w:top w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/>"
        "<w:left w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/>"
        "<w:bottom w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/>"
        "<w:right w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/>"
        "<w:insideH w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/>"
        "<w:insideV w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"/>"
        "</w:tblBorders>"
        "</w:tblPr>"
    )
    tbl_grid = (
        "<w:tblGrid>"
        "<w:gridCol w:w=\"2800\"/>"
        "<w:gridCol w:w=\"3600\"/>"
        "<w:gridCol w:w=\"1700\"/>"
        "<w:gridCol w:w=\"1700\"/>"
        "</w:tblGrid>"
    )
    row_xml = []
    for index, row in enumerate(rows):
        cell_xml = []
        for cell in row:
            weight = "<w:rPr><w:b/></w:rPr>" if index == 0 else ""
            cell_xml.append(
                "<w:tc>"
                "<w:tcPr><w:tcW w:w=\"0\" w:type=\"auto\"/></w:tcPr>"
                f"<w:p><w:r>{weight}<w:t xml:space=\"preserve\">{escape(cell)}</w:t></w:r></w:p>"
                "</w:tc>"
            )
        row_xml.append("<w:tr>" + "".join(cell_xml) + "</w:tr>")
    return "<w:tbl>" + tbl_pr + tbl_grid + "".join(row_xml) + "</w:tbl>"


title = [
    paragraph("Group Project Interim Progress Report", "Title"),
    paragraph("Course: ISM 4402"),
    paragraph("Project Topic: U.S. Entry-Level Job Market by Sector: Building a Career Navigator Framework"),
]

section_1 = [
    paragraph("Section 1 - Description of Data", "Heading1"),
    paragraph("Summary of the Data", "Heading2"),
    paragraph(
        "This project uses a multi-table labor-market dataset assembled from official U.S. Bureau of Labor Statistics and O*NET sources. "
        "The goal is to analyze the U.S. entry-level job market across sectors so the team can identify where early-career workers are most likely "
        "to find strong opportunities, what occupations are growing fastest, how wages vary across sectors, and how job characteristics differ "
        "by occupation and industry."
    ),
    paragraph(
        "The final business framing connects the analysis to the concept of a LinkedIn Career Navigator AI that could help users choose sectors, "
        "occupations, and learning paths based on labor-market evidence."
    ),
    paragraph("Data Sources Being Used", "Heading2"),
    paragraph("1. BLS Employment Projections tables: occupational projections, worker characteristics, industry-occupation matrix data, openings and separations, and employment and output by industry."),
    paragraph("2. BLS Occupational Employment and Wage Statistics tables: national industry-specific and all data tables."),
    paragraph("3. BLS State and Area Employment annual average tables: employees on nonfarm payrolls in states and selected areas by major industry."),
    paragraph("4. BLS Labor Force Characteristics tables: average hours at work by industry and by occupation."),
    paragraph("5. O*NET occupation attributes: Work Activities.xlsx."),
    paragraph("Source Websites", "Heading2"),
    paragraph("BLS Employment Projections tables: https://www.bls.gov/emp/tables.htm"),
    paragraph("BLS Occupational Employment and Wage Statistics tables: https://www.bls.gov/oes/tables.htm"),
    paragraph("BLS State and Area Employment annual average tables: https://www.bls.gov/sae/tables/annual-average/"),
    paragraph("BLS Labor Force Characteristics tables: https://www.bls.gov/cps/lfcharacteristics.htm"),
    paragraph("BLS JOLTS supplemental tables: https://www.bls.gov/web/jolts.supp.toc.htm"),
    paragraph("O*NET Work Activities data: https://www.onetcenter.org/dictionary/29.2/excel/work_activities.html"),
    paragraph("Why This Data Fits the Project", "Heading2"),
    paragraph(
        "This dataset is a strong fit for the group project because it is real, multi-table, and suitable for Tableau-based business analysis. "
        "It supports data preparation, joins, calculated fields, clustering, forecasting, dashboards, and storytelling. It also allows the "
        "project to compare sectors and occupations using multiple criteria instead of only one measure such as wage."
    ),
    paragraph("Data Preparation Plan", "Heading2"),
    paragraph("1. Standardize occupation codes across BLS and O*NET so occupation-level joins work correctly."),
    paragraph("2. Standardize sector and industry labels so wages, growth, employment, and openings can be compared consistently."),
    paragraph("3. Filter for entry-level occupations using education, training, and related-work-experience fields."),
    paragraph("4. Create calculated fields such as entry-level flag, growth category, wage percentile group, accessibility score, and attractiveness score."),
    paragraph("5. Handle missing, suppressed, or non-comparable values through filtering, null handling, or documented exclusions."),
    paragraph("6. Join occupation-level, industry-level, and geography-level tables using occupation code, industry code, and state fields where appropriate."),
    paragraph("7. Reshape long-format data such as O*NET work activities when needed for Tableau visuals."),
    paragraph("8. Build focused analysis-ready subsets for occupations, sectors, state-sector employment, occupation attributes, and occupation-sector mapping."),
    paragraph("Granularity / Unit of Analysis Summary", "Heading2"),
    paragraph(
        "The table below summarizes the expected unit of analysis for each major data table. For the O*NET file, the row and column counts were "
        "confirmed from the local workbook. Several BLS files saved in the local project folder currently contain blocked HTML download responses "
        "rather than the actual spreadsheets, so those counts must be updated after the official source files are downloaded manually in a browser."
    ),
]

granularity_rows = [
    ["Table Name", "Lowest Level of Granularity", "# of Columns", "# of Rows"],
    ["BLS Occupation Projections", "One row per occupation", "TBD after manual download", "TBD after manual download"],
    ["BLS Industry Employment and Output", "One row per industry", "TBD after manual download", "TBD after manual download"],
    ["BLS Industry-Occupation Matrix", "One row per occupation-industry combination, depending on selected table version", "TBD after manual download", "TBD after manual download"],
    ["BLS OEWS Wages by Occupation and Industry", "One row per occupation-industry combination", "TBD after manual download", "TBD after manual download"],
    ["BLS State Employment by Major Industry", "One row per state-sector combination", "TBD after manual download", "TBD after manual download"],
    ["O*NET Work Activities", "One row per occupation-activity record", "15", "72,079"],
]

section_2 = [
    paragraph("Screenshot placeholder: Insert Tableau Desktop or Tableau Prep screenshots here showing how the team determined the granularity for each connected table."),
    paragraph("Section 2 - Purpose of the Analysis", "Heading1"),
    paragraph("Overall Purpose", "Heading2"),
    paragraph(
        "The purpose of the analysis is to identify where the strongest entry-level opportunities exist in the U.S. job market and how those "
        "opportunities vary by sector, occupation, wages, projected growth, education requirements, and work characteristics. The group will use "
        "the analysis to support the concept of a LinkedIn Career Navigator AI that helps early-career users choose sectors and occupations using "
        "real market data."
    ),
    paragraph("Angle of the Analysis", "Heading2"),
    paragraph(
        "The most interesting angle is not simply which jobs pay the most. Instead, the team will evaluate opportunity using a combination of "
        "accessibility for entry-level workers, projected growth and openings, sector demand, wage potential, work-activity patterns, and geography. "
        "This gives the project a stronger business story and makes the Tableau output more useful for practical career decision-making."
    ),
    paragraph("Specific Questions That Will Guide the Analysis", "Heading2"),
    paragraph("1. Which U.S. sectors currently offer the strongest volume of employment for entry-level workers?"),
    paragraph("2. Which entry-level occupations have the highest projected growth and annual openings?"),
    paragraph("3. How do wages for similar entry-level occupations differ across sectors?"),
    paragraph("4. Which sectors combine stronger wage potential with stronger projected employment growth?"),
    paragraph("5. How do education and training requirements vary across entry-level occupations and sectors?"),
    paragraph("6. Which states show the strongest concentration of employment in sectors that appear attractive for early-career workers?"),
    paragraph("7. Can occupations be clustered into meaningful groups based on work activities and job characteristics?"),
    paragraph("8. Which occupations appear most attractive when balancing accessibility, wages, and growth?"),
    paragraph("Planned Tableau Story Direction", "Heading2"),
    paragraph(
        "To keep the project manageable while still showing analytical depth, the Tableau story will focus on major U.S. sectors, entry-level occupations, "
        "selected high-interest sectors such as healthcare, professional services, retail, logistics, finance, and technology-related services, and "
        "occupation clusters built from work-activity patterns."
    ),
    paragraph(
        "A strong final narrative for the project is: \"Where should an early-career worker start? A data-driven look at entry-level opportunities across U.S. sectors.\""
    ),
    paragraph("Current Progress Note", "Heading2"),
    paragraph(
        "As of March 31, 2026, the project folder includes the O*NET workbook and project planning materials. Some BLS files in the current local folder "
        "were saved as blocked HTML responses instead of the true downloadable spreadsheets. Those official BLS workbooks still need to be downloaded "
        "manually in a browser and then connected in Tableau before the team completes the granularity screenshots and later exploratory analysis sections."
    ),
]

document_body = "".join(title + section_1) + table(granularity_rows) + "".join(section_2)

document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
 xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
 xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:w10="urn:schemas-microsoft-com:office:word"
 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
 xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
 xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
 xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
 mc:Ignorable="w14 wp14">
<w:body>
{document_body}
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
</w:sectPr>
</w:body>
</w:document>
"""

styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr><w:jc w:val="center"/></w:pPr>
    <w:rPr>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:rPr><w:b/><w:sz w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:rPr><w:b/><w:sz w:val="24"/></w:rPr>
  </w:style>
</w:styles>
"""

content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""

rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""

document_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
"""

timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat()
core_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Group Project Interim Progress Report</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:modified>
</cp:coreProperties>
"""

app_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
</Properties>
"""

with zipfile.ZipFile(OUTPUT, "w", compression=zipfile.ZIP_DEFLATED) as docx:
    docx.writestr("[Content_Types].xml", content_types_xml)
    docx.writestr("_rels/.rels", rels_xml)
    docx.writestr("docProps/core.xml", core_xml)
    docx.writestr("docProps/app.xml", app_xml)
    docx.writestr("word/document.xml", document_xml)
    docx.writestr("word/styles.xml", styles_xml)
    docx.writestr("word/_rels/document.xml.rels", document_rels_xml)

print(f"Wrote {OUTPUT}")
