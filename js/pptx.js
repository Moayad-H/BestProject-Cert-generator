/**
 * Certificate Forge - PPTX Generator Module
 * Duplicates slide1 and its relationships directly inside the PPTX zip structure,
 * replaces text placeholders, dynamically adjusts box width & font size for long names,
 * and maintains full OpenXML compliance so PowerPoint opens without any repair prompt.
 */

// Escape XML entities
export function escapeXml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

/**
 * Generate certificates presentation
 * @param {ArrayBuffer} templateBuffer - Binary buffer of the template .pptx
 * @param {Array<{course: string, studentName: string}>} students - List of students to generate
 * @param {Function} onProgress - Callback (current, total, message) => void
 * @returns {Promise<Blob>} PPTX Blob
 */
export async function generateCertificatesPresentation(
  templateBuffer,
  students,
  onProgress = () => {}
) {
  const jszipInstance = typeof window !== "undefined" ? window.JSZip : globalThis.JSZip;
  if (!jszipInstance) {
    throw new Error("JSZip library is not loaded.");
  }

  if (!students || students.length === 0) {
    throw new Error("No student records provided for generation.");
  }

  onProgress(0, students.length, "Loading PowerPoint template...");
  const zip = await jszipInstance.loadAsync(templateBuffer);

  // 1. Read key template XML files
  const presXmlFile = zip.file("ppt/presentation.xml");
  const presRelsFile = zip.file("ppt/_rels/presentation.xml.rels");
  const contentTypesFile = zip.file("[Content_Types].xml");
  const slide1XmlFile = zip.file("ppt/slides/slide1.xml");
  const slide1RelsFile = zip.file("ppt/slides/_rels/slide1.xml.rels");

  if (!presXmlFile || !presRelsFile || !contentTypesFile || !slide1XmlFile) {
    throw new Error("Invalid PowerPoint template: Missing core OpenXML presentation files.");
  }

  let presXmlStr = await presXmlFile.async("text");
  let presRelsStr = await presRelsFile.async("text");
  let contentTypesStr = await contentTypesFile.async("text");
  const slide1XmlStr = await slide1XmlFile.async("text");
  const slide1RelsBytes = slide1RelsFile ? await slide1RelsFile.async("uint8array") : null;

  // Extract slide dimensions cx, cy
  const sldSzMatch = presXmlStr.match(/<p:sldSz\s+[^>]*cx="(\d+)"[^>]*cy="(\d+)"/) ||
                     presXmlStr.match(/<p:sldSz\s+[^>]*cy="(\d+)"[^>]*cx="(\d+)"/);
  const slideWidth = sldSzMatch ? parseInt(sldSzMatch[1], 10) : 10693400;

  // 2. Determine highest existing rId in presentation.xml.rels
  const rIdMatches = [...presRelsStr.matchAll(/Id="rId(\d+)"/g)];
  let maxRId = 0;
  for (const m of rIdMatches) {
    const num = parseInt(m[1], 10);
    if (num > maxRId) maxRId = num;
  }
  let nextRId = maxRId + 1;

  // Strip existing slide relationships from presentation.xml.rels
  presRelsStr = presRelsStr.replace(
    /<Relationship\s+[^>]*Type="[^"]*relationships\/slide"[^>]*\/>\s*/gi,
    ""
  );

  // Strip existing slide overrides from [Content_Types].xml
  contentTypesStr = contentTypesStr.replace(
    /<Override\s+[^>]*PartName="\/ppt\/slides\/slide\d+\.xml"[^>]*\/>\s*/gi,
    ""
  );

  // Clean out any existing slide files from the zip
  const existingSlideFiles = Object.keys(zip.files).filter(
    (name) => name.startsWith("ppt/slides/slide") && (name.endsWith(".xml") || name.endsWith(".xml.rels"))
  );
  for (const name of existingSlideFiles) {
    zip.remove(name);
  }

  const sldIdElements = [];
  const relElements = [];
  const ctOverrides = [];

  const sldIdBase = 256;
  const total = students.length;

  // 3. Generate slides for each student
  for (let i = 0; i < total; i++) {
    const slideNum = i + 1;
    const student = students[i];
    const rId = `rId${nextRId++}`;
    const sldId = sldIdBase + i;

    // Report progress periodically
    if (i % 10 === 0 || i === total - 1) {
      onProgress(i + 1, total, `Generating slide ${slideNum} of ${total} (${student.studentName})...`);
      // Allow event loop cycle
      await new Promise((resolve) => setTimeout(resolve, 0));
    }

    // 3a. Add Relationship entry
    relElements.push(
      `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${slideNum}.xml"/>`
    );

    // 3b. Add Content Types Override
    ctOverrides.push(
      `<Override PartName="/ppt/slides/slide${slideNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
    );

    // 3c. Modify Slide XML
    let slideXml = slide1XmlStr;

    // Replace placeholders with escaped text
    const escapedStudentName = escapeXml(student.studentName);
    const escapedCourseName = escapeXml(student.course);

    slideXml = slideXml.replace(/\{Student_Name\}/g, escapedStudentName);
    slideXml = slideXml.replace(/\{Course_Name\}/g, escapedCourseName);

    // Adapt layout for long student names
    const nameLen = student.studentName.length;

    // If name is long (> 17 chars), widen text box to 90% of slide width and center it
    if (nameLen > 17) {
      const newWidth = Math.round(slideWidth * 0.90);
      const newOffsetX = Math.round((slideWidth - newWidth) / 2);

      // Regex matching TextBox 11 or student shape xfrm
      slideXml = slideXml.replace(
        /(<[a-zA-Z0-9:]*cNvPr[^>]*name="TextBox 11"[^>]*>[\s\S]*?<[a-zA-Z0-9:]*spPr>[\s\S]*?<[a-zA-Z0-9:]*xfrm[^>]*>[\s\S]*?<[a-zA-Z0-9:]*off\s+x=")\d+("[\s\S]*?<[a-zA-Z0-9:]*ext\s+cx=")\d+(")/,
        `$1${newOffsetX}$2${newWidth}$3`
      );
    }

    // If name is longer than 20 chars, reduce font size slightly
    if (nameLen > 20) {
      let newFontSize = 4700; // 47 pt
      if (nameLen > 24) {
        newFontSize = 4300; // 43 pt
      }
      if (nameLen > 30) {
        newFontSize = 3800; // 38 pt
      }

      slideXml = slideXml.replace(/sz="5489"/g, `sz="${newFontSize}"`);
    }

    // Write slide file and rels to zip
    zip.file(`ppt/slides/slide${slideNum}.xml`, slideXml);
    if (slide1RelsBytes) {
      zip.file(`ppt/slides/_rels/slide${slideNum}.xml.rels`, slide1RelsBytes);
    }

    sldIdElements.push(`<p:sldId id="${sldId}" r:id="${rId}"/>`);
  }

  // 4. Update ppt/presentation.xml
  const joinedSldIds = sldIdElements.join("");
  const newSldIdListXml = `<p:sldIdLst>${joinedSldIds}</p:sldIdLst>`;

  if (/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/.test(presXmlStr)) {
    presXmlStr = presXmlStr.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, newSldIdListXml);
  } else if (/<p:sldMasterIdLst>[\s\S]*?<\/p:sldMasterIdLst>/.test(presXmlStr)) {
    presXmlStr = presXmlStr.replace(
      /(<p:sldMasterIdLst>[\s\S]*?<\/p:sldMasterIdLst>)/,
      `$1${newSldIdListXml}`
    );
  }
  zip.file("ppt/presentation.xml", presXmlStr);

  // 5. Update ppt/_rels/presentation.xml.rels
  const joinedRels = relElements.join("");
  presRelsStr = presRelsStr.replace("</Relationships>", `${joinedRels}</Relationships>`);
  zip.file("ppt/_rels/presentation.xml.rels", presRelsStr);

  // 6. Update [Content_Types].xml
  const joinedOverrides = ctOverrides.join("");
  contentTypesStr = contentTypesStr.replace("</Types>", `${joinedOverrides}</Types>`);
  zip.file("[Content_Types].xml", contentTypesStr);

  // 7. Update docProps/app.xml with accurate slide count if present
  const appXmlFile = zip.file("docProps/app.xml");
  if (appXmlFile) {
    let appXmlStr = await appXmlFile.async("text");
    if (/<Slides>\d+<\/Slides>/.test(appXmlStr)) {
      appXmlStr = appXmlStr.replace(/<Slides>\d+<\/Slides>/, `<Slides>${total}</Slides>`);
      zip.file("docProps/app.xml", appXmlStr);
    }
  }

  // 8. Generate final binary PPTX
  onProgress(total, total, "Packing final PowerPoint file...");
  
  if (typeof Blob !== "undefined") {
    return await zip.generateAsync({
      type: "blob",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
      mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    });
  } else {
    // Node.js fallback
    return await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
  }
}
