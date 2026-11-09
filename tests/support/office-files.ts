import { strToU8, zipSync } from "fflate";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const REL = 'xmlns="http://schemas.openxmlformats.org/package/2006/relationships"';
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const zip = (files: Record<string, string>): Buffer => Buffer.from(zipSync(Object.fromEntries(Object.entries(files).map(([name, text]) => [name, strToU8(text)]))));

/** A paragraph with plain runs; `style` is a styleId. */
export const para = (text: string, style?: string, extra = ""): string =>
  `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}<w:r><w:t xml:space="preserve">${text}</w:t></w:r>${extra}</w:p>`;

export function buildDocx(body: string, extra: Record<string, string> = {}): Buffer {
  return zip({
    "[Content_Types].xml": `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
    "word/document.xml": `${HEAD}<w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`,
    "word/styles.xml": `${HEAD}<w:styles ${W}>`
      + '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>'
      + '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>'
      + '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>',
    ...extra,
  });
}

/** A realistic report: headings, a deleted tracked phrase, an inserted one, a list item, a text box and a table. */
export function sampleDocx(): Buffer {
  const table = '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Sales</w:t></w:r></w:p></w:tc></w:tr>'
    + '<w:tr><w:tc><w:p><w:r><w:t>EMEA</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>10 &amp; up</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
  const tracked = '<w:p><w:r><w:t xml:space="preserve">Revenue grew </w:t></w:r><w:del w:id="1" w:author="A"><w:r><w:delText>sharply</w:delText></w:r></w:del>'
    + '<w:ins w:id="2" w:author="A"><w:r><w:t>steadily</w:t></w:r></w:ins></w:p>';
  const list = '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>First point</w:t></w:r></w:p>';
  const textbox = '<w:p><w:r><w:t>Before box.</w:t></w:r><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    + '<mc:Choice><w:r><w:t>Boxed words</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>Boxed words</w:t></w:r></mc:Fallback></mc:AlternateContent></w:p>';
  const comments = `${HEAD}<w:comments ${W}><w:comment w:id="0" w:author="Lan"><w:p><w:r><w:t>Check this number</w:t></w:r></w:p></w:comment></w:comments>`;
  return buildDocx(
    para("Annual report", "Heading1") + para("Overview", "Heading2") + tracked + list + textbox + table + para("Closing remarks."),
    { "word/comments.xml": comments },
  );
}

export interface SlideSpec {
  title?: string;
  body?: string[];
  table?: string[][];
  notes?: string;
  hidden?: boolean;
  picture?: { name: string; alt: string };
  footer?: string;
}

const sp = (id: number, name: string, ph: string | null, lines: string[]): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr/><p:nvPr>${ph ? `<p:ph type="${ph}"/>` : ""}</p:nvPr></p:nvSpPr><p:spPr/>`
  + `<p:txBody><a:bodyPr/>${lines.map((line) => `<a:p><a:r><a:t>${line}</a:t></a:r></a:p>`).join("")}</p:txBody></p:sp>`;

function slideXml(spec: SlideSpec): string {
  const shapes: string[] = [];
  if (spec.title) shapes.push(sp(2, "Title 1", "title", [spec.title]));
  if (spec.body) shapes.push(sp(3, "Content 2", null, spec.body));
  if (spec.footer) shapes.push(sp(4, "Footer 3", "ftr", [spec.footer]));
  if (spec.table) {
    shapes.push(`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Table 4"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl>`
      + spec.table.map((row) => `<a:tr>${row.map((cell) => `<a:tc><a:txBody><a:p><a:r><a:t>${cell}</a:t></a:r></a:p></a:txBody></a:tc>`).join("")}</a:tr>`).join("")
      + `</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`);
  }
  if (spec.picture) shapes.push(`<p:pic><p:nvPicPr><p:cNvPr id="6" name="${spec.picture.name}" descr="${spec.picture.alt}"/></p:nvPicPr></p:pic>`);
  return `${HEAD}<p:sld ${P} ${A} ${R}${spec.hidden ? ' show="0"' : ""}><p:cSld><p:spTree>${shapes.join("")}</p:spTree></p:cSld></p:sld>`;
}

const notesXml = (text: string): string =>
  `${HEAD}<p:notes ${P} ${A}><p:cSld><p:spTree>${sp(2, "Slide Image", "sldImg", [])}${sp(3, "Notes 2", "body", [text])}${sp(4, "Number", "sldNum", ["7"])}</p:spTree></p:cSld></p:notes>`;

/** Slides are listed in presentation order, whatever their file numbers; `order` shuffles the file numbers to prove that. */
export function buildPptx(slides: SlideSpec[], order: number[] = slides.map((_, i) => i + 1)): Buffer {
  const files: Record<string, string> = {
    "[Content_Types].xml": `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
  };
  const listed: string[] = [];
  const rels: string[] = [];
  slides.forEach((spec, i) => {
    const n = order[i]!;
    files[`ppt/slides/slide${n}.xml`] = slideXml(spec);
    rels.push(`<Relationship Id="rId${i + 10}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${n}.xml"/>`);
    listed.push(`<p:sldId id="${256 + i}" r:id="rId${i + 10}"/>`);
    if (spec.notes) {
      files[`ppt/notesSlides/notesSlide${n}.xml`] = notesXml(spec.notes);
      files[`ppt/slides/_rels/slide${n}.xml.rels`] = `${HEAD}<Relationships ${REL}><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide${n}.xml"/></Relationships>`;
    }
  });
  files["ppt/presentation.xml"] = `${HEAD}<p:presentation ${P} ${R}><p:sldIdLst>${listed.join("")}</p:sldIdLst></p:presentation>`;
  files["ppt/_rels/presentation.xml.rels"] = `${HEAD}<Relationships ${REL}>${rels.join("")}</Relationships>`;
  return zip(files);
}

export function sampleDeck(): Buffer {
  return buildPptx([
    { title: "Quarterly Review", body: ["Q3 results"], notes: "Say thank you first.", footer: "Confidential" },
    { title: "Numbers", table: [["Region", "Sales"], ["EMEA", "10"]], picture: { name: "Logo", alt: "Company logo" } },
    { title: "Backup", body: ["Hidden detail", "  nested"], hidden: true },
  ], [3, 1, 2]);
}
