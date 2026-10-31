import JSZip from "jszip";
import { tagTemplate, validateListOpenTag, validateListName, validateTag } from "../dist/template-tagger.js";
import { parseTemplate } from "../dist/docx-engine.js";

async function minimalDocx(bodyInner) {
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${bodyInner}
    <w:sectPr/>
  </w:body>
</w:document>`;
  const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
  const zip = new JSZip();
  zip.file("[Content_Types].xml", ct);
  zip.folder("_rels").file(".rels", rels);
  zip.folder("word").file("document.xml", doc);
  return Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));
}

function p(text) {
  return `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
}
function cell(text) {
  return `<w:tc><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
}
function row(...cells) {
  return `<w:tr>${cells.map(cell).join("")}</w:tr>`;
}

async function main() {
  console.log("validate:", validateListName("items"), validateListOpenTag("{{#items}}"), validateTag("{{name}}"));

  // --- A) table row wrap ---
  const tableBody =
    p("Header note") +
    `<w:tbl>${row("Product", "Qty", "Price")}${row("Apple", "10", "20")}${row("Banana", "5", "12")}</w:tbl>` +
    p("Footer");
  const tableDoc = await minimalDocx(tableBody);
  const taggedTable = await tagTemplate(tableDoc, [
    {
      id: "wrap1",
      plainText: "Apple",
      occurrence: 0,
      tag: "{{#items}}",
      mode: "wrap-row",
      closeTag: "{{/}}",
    },
    { id: "f1", plainText: "Apple", occurrence: 0, tag: "{{name}}" },
    { id: "f2", plainText: "10", occurrence: 0, tag: "{{qty}}" },
    { id: "f3", plainText: "20", occurrence: 0, tag: "{{price}}" },
  ]);
  const zip1 = await JSZip.loadAsync(taggedTable.buffer);
  const xml1 = await zip1.file("word/document.xml").async("string");
  console.log("\n--- table xml snippet ---");
  const trMatch = xml1.match(/<w:tr>[\s\S]*?Apple[\s\S]*?<\/w:tr>|<w:tr>[\s\S]*?\{\{#items\}\}[\s\S]*?<\/w:tr>/);
  // Find the data row with loop
  const idx = xml1.indexOf("{{#items}}");
  console.log("has open:", xml1.includes("{{#items}}"), "has close:", xml1.includes("{{/}}"));
  console.log("has name field:", xml1.includes("{{name}}"), "qty:", xml1.includes("{{qty}}"), "price:", xml1.includes("{{price}}"));
  console.log("snippet:", xml1.slice(Math.max(0, idx - 40), idx + 120).replace(/\s+/g, " "));
  // Trailing rows stay (may be footer) — delete manually if unwanted.
  console.log("Banana row kept:", xml1.includes("Banana"));
  if (!xml1.includes("Banana")) {
    throw new Error("expected trailing row Banana to remain after wrap-row");
  }

  const parsed = await parseTemplate(taggedTable.buffer);
  console.log("parse fields:", JSON.stringify(parsed.fields, null, 2));

  // --- B) multi-paragraph wrap ---
  const multiBody = p("Intro") + p("LineOne") + p("LineTwo") + p("Outro");
  const multiDoc = await minimalDocx(multiBody);
  const taggedMulti = await tagTemplate(multiDoc, [
    {
      id: "wrap2",
      plainText: "LineOneLineTwo",
      occurrence: 0,
      tag: "{{#notes}}",
      mode: "wrap",
      closeTag: "{{/}}",
    },
    { id: "l1", plainText: "LineOne", occurrence: 0, tag: "{{title}}" },
    { id: "l2", plainText: "LineTwo", occurrence: 0, tag: "{{body}}" },
  ]);
  const zip2 = await JSZip.loadAsync(taggedMulti.buffer);
  const xml2 = await zip2.file("word/document.xml").async("string");
  console.log("\n--- multi wrap ---");
  console.log("open:", xml2.includes("{{#notes}}"), "close:", xml2.includes("{{/}}"));
  console.log("fields:", xml2.includes("{{title}}"), xml2.includes("{{body}}"));
  console.log("plain:", taggedMulti.plainText);
  const parsed2 = await parseTemplate(taggedMulti.buffer);
  console.log("parse fields:", JSON.stringify(parsed2.fields, null, 2));

  // --- C) single non-table paragraph wrap (จำเลย-style) ---
  const singleBody =
    p("จำเลยที่1เลขตินิติบุคคล 0115557012850ชื่อบริษัท เนเชอร์รัล เบลนด์ จำกัด") +
    p("จำเลยที่2เลขตินิติบุคคล 0105551234567ชื่อบริษัท อื่น จำกัด");
  const singleDoc = await minimalDocx(singleBody);
  const taggedSingle = await tagTemplate(singleDoc, [
    {
      id: "wrap3",
      plainText: "จำเลยที่1เลขตินิติบุคคล 0115557012850ชื่อบริษัท เนเชอร์รัล เบลนด์ จำกัด",
      occurrence: 0,
      tag: "{{#จำเลย}}",
      mode: "wrap",
      closeTag: "{{/}}",
    },
  ]);
  const zip3 = await JSZip.loadAsync(taggedSingle.buffer);
  const xml3 = await zip3.file("word/document.xml").async("string");
  console.log("\n--- single para wrap ---");
  console.log("open:", xml3.includes("{{#จำเลย}}"), "close:", xml3.includes("{{/}}"));
  console.log("defendant2 still raw:", xml3.includes("จำเลยที่2"));
  if (!xml3.includes("{{#จำเลย}}") || !xml3.includes("{{/}}")) {
    throw new Error("single paragraph wrap failed");
  }

  // --- D) delete paragraph ---
  const delParaDoc = await minimalDocx(p("KeepA") + p("RemoveMe") + p("KeepB"));
  const delPara = await tagTemplate(delParaDoc, [
    { id: "dp1", plainText: "RemoveMe", occurrence: 0, tag: "", mode: "delete-p" },
  ]);
  const zipDelP = await JSZip.loadAsync(delPara.buffer);
  const xmlDelP = await zipDelP.file("word/document.xml").async("string");
  console.log("\n--- delete-p ---");
  console.log("kept:", xmlDelP.includes("KeepA") && xmlDelP.includes("KeepB"), "removed:", !xmlDelP.includes("RemoveMe"));
  if (xmlDelP.includes("RemoveMe") || !xmlDelP.includes("KeepA") || !xmlDelP.includes("KeepB")) {
    throw new Error("delete-p failed");
  }

  // --- E) delete table row (+ keep replace on another para) ---
  const delRowBody =
    p("Title") +
    `<w:tbl>${row("H1", "H2")}${row("RowKeep", "A")}${row("RowDrop", "B")}</w:tbl>` +
    p("Foot");
  const delRowDoc = await minimalDocx(delRowBody);
  const delRow = await tagTemplate(delRowDoc, [
    { id: "dt1", plainText: "RowDrop", occurrence: 0, tag: "", mode: "delete-tr" },
    { id: "f", plainText: "Title", occurrence: 0, tag: "{{title}}" },
  ]);
  const zipDelTr = await JSZip.loadAsync(delRow.buffer);
  const xmlDelTr = await zipDelTr.file("word/document.xml").async("string");
  console.log("\n--- delete-tr ---");
  console.log(
    "row kept:",
    xmlDelTr.includes("RowKeep"),
    "row dropped:",
    !xmlDelTr.includes("RowDrop"),
    "title tagged:",
    xmlDelTr.includes("{{title}}"),
  );
  if (xmlDelTr.includes("RowDrop") || !xmlDelTr.includes("RowKeep") || !xmlDelTr.includes("{{title}}")) {
    throw new Error("delete-tr failed");
  }

  // --- F) two pinned deletes in table 0 must not touch table 1 ---
  const twoTblBody =
    `<w:tbl>${row("T0H")}${row("T0R1")}${row("T0DropA")}${row("T0DropB")}${row("T0R2")}</w:tbl>` +
    `<w:tbl>${row("T1H")}${row("T1KeepA")}${row("T1KeepB")}</w:tbl>`;
  const twoTblDoc = await minimalDocx(twoTblBody);
  const twoDel = await tagTemplate(twoTblDoc, [
    {
      id: "dA",
      plainText: "",
      occurrence: 0,
      tag: "",
      mode: "delete-tr",
      tableIndex: 0,
      rowIndexInTable: 2,
    },
    {
      id: "dB",
      plainText: "",
      occurrence: 0,
      tag: "",
      mode: "delete-tr",
      tableIndex: 0,
      rowIndexInTable: 3,
    },
  ]);
  const zipTwo = await JSZip.loadAsync(twoDel.buffer);
  const xmlTwo = await zipTwo.file("word/document.xml").async("string");
  console.log("\n--- double pinned delete-tr (table 0 only) ---");
  console.log(
    "T0 kept:",
    xmlTwo.includes("T0R1") && xmlTwo.includes("T0R2"),
    "T0 dropped:",
    !xmlTwo.includes("T0DropA") && !xmlTwo.includes("T0DropB"),
    "T1 intact:",
    xmlTwo.includes("T1KeepA") && xmlTwo.includes("T1KeepB") && xmlTwo.includes("T1H"),
  );
  if (
    xmlTwo.includes("T0DropA") ||
    xmlTwo.includes("T0DropB") ||
    !xmlTwo.includes("T0R1") ||
    !xmlTwo.includes("T0R2") ||
    !xmlTwo.includes("T1KeepA") ||
    !xmlTwo.includes("T1KeepB")
  ) {
    throw new Error("double pinned delete-tr touched wrong table/rows");
  }

  console.log("\nSMOKE OK");
}

main().catch((e) => {
  console.error("SMOKE FAIL", e);
  process.exit(1);
});
