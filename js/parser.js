/**
 * Certificate Generator - Parser Module
 * Handles Excel parsing, student name extraction, cleaning, compound name handling,
 * Arabic script detection, and official course matching.
 */

// Check if string contains Arabic characters
export function isArabicText(text) {
  return /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF]/.test(text);
}

// Convert word to Title Case while respecting hyphenated compound parts
export function toTitleCase(nameStr) {
  if (!nameStr) return "";

  const titleWord = (w) => {
    if (w.includes("-")) {
      return w.split("-").map(titleWord).join("-");
    }
    const wl = w.toLowerCase();
    if (wl === "el" || wl === "al") {
      return "El";
    }
    return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
  };

  return nameStr
    .trim()
    .split(/\s+/)
    .map(titleWord)
    .join(" ");
}

/**
 * Shorten name to first + last name only, preserving compound parts:
 * Compound first names: "Abd Allah", "Menna Allah", "Menat Allah"
 * Compound last names: "El Sayed", "Alaa Eldin", "Zein-Eldeen", "Abd Elnaser", etc.
 * Single-word names stay as is.
 */
export function shortenName(fullName) {
  if (!fullName) return "";
  const cleaned = fullName.replace(/\s+/g, " ").trim();
  const words = cleaned.split(" ");

  if (words.length <= 1) {
    return toTitleCase(cleaned);
  }

  // Check compound first name
  const w0 = words[0].toLowerCase();
  const w1 = words[1]?.toLowerCase() || "";

  let firstName = "";
  let startIdx = 1;

  if (
    words.length >= 3 &&
    ((["abd", "abdel", "abdul"].includes(w0) && ["allah", "ellah", "elah"].includes(w1)) ||
      (["menna", "menat", "menet"].includes(w0) && ["allah", "ellah", "elah"].includes(w1)))
  ) {
    firstName = `${words[0]} ${words[1]}`;
    startIdx = 2;
  } else {
    firstName = words[0];
    startIdx = 1;
  }

  const remaining = words.slice(startIdx);
  if (remaining.length === 0) {
    return toTitleCase(firstName);
  }
  if (remaining.length === 1) {
    return toTitleCase(`${firstName} ${remaining[0]}`);
  }

  // Check compound last name from the last 2 tokens
  const last2_0 = remaining[remaining.length - 2].toLowerCase();
  const last2_1 = remaining[remaining.length - 1].toLowerCase();

  let lastName = "";
  if (
    ["el", "al", "abou", "abu", "ibn", "ben"].includes(last2_0) ||
    ["eldin", "el-din", "eldeen", "el-deen", "eldein", "el-dein"].some((s) => last2_1.includes(s)) ||
    (last2_0 === "alaa" && ["eldin", "eldeen", "eldein", "el-din", "el-deen"].some((s) => last2_1.includes(s))) ||
    (["abd", "abdel"].includes(last2_0) &&
      ["allah", "elnaser", "elnasser", "salam", "motal", "meguid", "nasser"].includes(last2_1))
  ) {
    lastName = `${remaining[remaining.length - 2]} ${remaining[remaining.length - 1]}`;
  } else {
    lastName = remaining[remaining.length - 1];
  }

  return toTitleCase(`${firstName} ${lastName}`);
}

/**
 * Extract clean student names from a messy Column C cell.
 * Returns { englishNames: string[], skippedArabic: string[] }
 */
export function extractStudentsFromCell(cellStr) {
  if (!cellStr || typeof cellStr !== "string") {
    return { englishNames: [], skippedArabic: [] };
  }

  // Pre-process lines
  const rawLines = cellStr.split(/\r?\n/);
  const filteredLines = [];

  for (let line of rawLines) {
    let s = line.trim();
    if (!s) continue;

    // Remove metadata lines
    if (/^(team\s*name\b|project\s*(#|\d+|name)|automated\s*reverse|banner|notified|notes:)/i.test(s)) {
      continue;
    }

    // Strip leading "NAME:"
    if (/\bname\s*:\s*/i.test(s)) {
      s = s.replace(/^.*?name\s*:\s*/i, "");
    }

    // Skip pure email / id / phone lines
    if (/^(email|id|number|phone)\s*:\s*/i.test(s)) continue;
    if (/^[0-9+\s()-]+$/.test(s)) continue;
    if (/^[\w.-]+@[\w.-]+\.\w+$/.test(s)) continue;

    filteredLines.push(s);
  }

  let text = filteredLines.join("\n");

  // 1. Detach emails directly joined to names, e.g. "foo@gmail.comEyad Yasser"
  text = text.replace(
    /([\w.-]+@(?:student\.)?[a-zA-Z0-9.-]+\.(?:com|edu|net|org|gov|eg|io))([A-Z\u0600-\u06FF])/gi,
    "$1\n$2"
  );

  // 2. Remove emails before digit splitting so digits inside email usernames are not detached
  text = text.replace(/[\w.-]+@(?:student\.)?[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "");

  // 3. Detach digits directly prepended to names, e.g. "231014378Mahmoud"
  text = text.replace(/(\d+)([a-zA-Z\u0600-\u06FF])/g, "$1 $2");

  // 4. Detach digits appended to names if followed by space or newline
  text = text.replace(/([a-zA-Z\u0600-\u06FF])(\d+)/g, "$1 $2");

  // Check parenthesized patterns like (Reem Mousa Elshamly&ID:231006275),(Judy Hesham...)
  let chunks = [];
  const parenMatches = [...text.matchAll(/\(([^)]+)\)/g)];
  if (parenMatches.length > 0 && /(&ID:|&id:|ID:|id:)/i.test(text)) {
    chunks = parenMatches.map((m) => m[1]);
  } else {
    // Standard splitting by comma, slash, backslash, or newline
    const normalized = text.replace(/[،؛]/g, ",");
    chunks = normalized
      .split(/[\n,/\\]+/)
      .map((c) => c.trim())
      .filter(Boolean);
  }

  const englishNames = [];
  const skippedArabic = [];

  for (let rawChunk of chunks) {
    let c = rawChunk;

    // Filter metadata phrases
    if (/project\s*\d+|report\s*generator|reverse\s*engineering/i.test(c)) {
      continue;
    }

    // Remove emails
    c = c.replace(/[\w.-]+@(?:student\.)?[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "");

    // Remove phone numbers
    c = c.replace(/(\+?20\s*)?01[0125][0-9\s-]{8,}/g, "");

    // Remove labels: ID:, NAME:, EMAIL:, NUMBER:, PHONE:
    c = c.replace(/\b(id|name|email|number|phone)\s*:?\s*/gi, "");

    // Remove & followed by digits or ID
    c = c.replace(/&\s*(id\s*:?\s*)?\d+/gi, "");

    // Remove all digits
    c = c.replace(/\d+/g, "");

    // Remove leftover punctuation except hyphens
    c = c.replace(/[()\[\]{}#:_\"*&+=|]/g, " ");
    c = c.replace(/\s+/g, " ").trim();

    if (!c) continue;

    if (isArabicText(c)) {
      skippedArabic.push(c);
    } else {
      // Must contain at least 2 letters
      if (/[a-zA-Z]{2,}/.test(c)) {
        const lower = c.toLowerCase();
        if (!["team", "project", "certificates", "banner", "id", "email", "generator", "report"].includes(lower)) {
          const shortened = shortenName(c);
          if (shortened) {
            englishNames.push(shortened);
          }
        }
      }
    }
  }

  return { englishNames, skippedArabic };
}

/**
 * Course name cleaner and normalizer
 */
export function cleanCourseString(str) {
  if (!str) return "";
  let s = str.trim();
  // Strip bracketed code: e.g. [CCY4005
  s = s.replace(/\[.*$/, "");
  // Strip trailing course code e.g. CCS2305, CCIT, etc.
  s = s.replace(/\b[A-Z]{2,4}\s*\d{3,4}\b/g, "");
  s = s.replace(/\bCCIT\b/g, "");
  s = s.replace(/[\u00a0\s]+/g, " ").trim();
  return s;
}

export function normalizeCourseName(str) {
  let s = cleanCourseString(str).toLowerCase();
  s = s.replace(/\bintro\s+to\b/g, "introduction to");
  s = s.replace(/&/g, "and");
  s = s.replace(/[^\w\s]/g, "");
  s = s.replace(/\s+/g, " ").trim();
  return s;
}

/**
 * Dice coefficient string similarity (0 to 1)
 */
function diceCoefficient(s1, s2) {
  if (s1 === s2) return 1.0;
  if (s1.length < 2 || s2.length < 2) return 0.0;

  const getBigrams = (str) => {
    const bigrams = new Map();
    for (let i = 0; i < str.length - 1; i++) {
      const bg = str.slice(i, i + 2);
      bigrams.set(bg, (bigrams.get(bg) || 0) + 1);
    }
    return bigrams;
  };

  const bg1 = getBigrams(s1);
  const bg2 = getBigrams(s2);

  let intersection = 0;
  for (const [bg, count1] of bg1.entries()) {
    if (bg2.has(bg)) {
      intersection += Math.min(count1, bg2.get(bg));
    }
  }

  return (2.0 * intersection) / (s1.length - 1 + s2.length - 1);
}

/**
 * Match raw course name against the list of official courses.
 * Returns { matchedName: string, confidence: number, isFlagged: boolean }
 */
export function matchCourse(rawCourse, officialCourses) {
  const cleaned = cleanCourseString(rawCourse);
  if (!officialCourses || officialCourses.length === 0) {
    return {
      matchedName: cleaned || rawCourse,
      confidence: 1.0,
      isFlagged: false,
    };
  }

  const normInput = normalizeCourseName(rawCourse);

  // 1. Exact match on normalized
  for (const off of officialCourses) {
    if (normalizeCourseName(off.name) === normInput) {
      return {
        matchedName: off.name,
        confidence: 1.0,
        isFlagged: false,
      };
    }
  }

  // 2. Substring match
  for (const off of officialCourses) {
    const offNorm = normalizeCourseName(off.name);
    if (normInput.includes(offNorm) || offNorm.includes(normInput)) {
      return {
        matchedName: off.name,
        confidence: 0.95,
        isFlagged: false,
      };
    }
  }

  // 3. Fuzzy similarity
  let bestMatch = null;
  let bestScore = 0;

  for (const off of officialCourses) {
    const score = diceCoefficient(normInput, normalizeCourseName(off.name));
    if (score > bestScore) {
      bestScore = score;
      bestMatch = off.name;
    }
  }

  const CONFIDENCE_THRESHOLD = 0.70;

  if (bestScore >= CONFIDENCE_THRESHOLD && bestMatch) {
    return {
      matchedName: bestMatch,
      confidence: bestScore,
      isFlagged: false,
    };
  }

  return {
    matchedName: cleaned || rawCourse,
    confidence: bestScore,
    isFlagged: true,
    suggestedName: bestMatch,
  };
}

/**
 * Parse an entire Excel workbook buffer using SheetJS.
 * Returns { students: Array<{id, course, studentName, isFlagged, confidence, originalRow}>, skipped: Array<{course, name, row}> }
 */
export function parseExcelWorkbook(workbook) {
  // 1. Parse official courses sheet
  const officialCourses = [];
  const courseSheetName = workbook.SheetNames.find((s) =>
    /courses\s*code\s*&\s*names/i.test(s)
  );

  if (courseSheetName) {
    const cSheet = workbook.Sheets[courseSheetName];
    const cData = XLSX.utils.sheet_to_json(cSheet, { header: 1 });
    for (let r = 1; r < cData.length; r++) {
      const row = cData[r];
      if (row) {
        const code = String(row[0] || "").trim();
        const name = String(row[1] || "").trim();
        if (name && !officialCourses.some((c) => c.name === name)) {
          officialCourses.push({ code, name });
        }
      }
    }
  }

  // 2. Parse Projects sheet
  const projectSheetName = workbook.SheetNames.find((s) => /projects/i.test(s)) || workbook.SheetNames[0];
  const pSheet = workbook.Sheets[projectSheetName];
  const pData = XLSX.utils.sheet_to_json(pSheet, { header: 1 });

  const resultStudents = [];
  const skippedEntries = [];
  let studentCounter = 1;

  // Data starts at row index 2 (row 3 in Excel)
  for (let r = 2; r < pData.length; r++) {
    const row = pData[r];
    if (!row) continue;

    const courseCell = row[0]; // Column A
    const certCell = row[5];   // Column F
    const studentsCell = row[2]; // Column C

    if (!courseCell || !String(courseCell).trim()) {
      continue;
    }

    const certStr = String(certCell ?? "").trim().toUpperCase();
    const isCertTrue = certCell === true || certStr === "TRUE" || certStr === "1";

    if (!isCertTrue) {
      continue;
    }

    const rawCourse = String(courseCell).trim();
    const courseMatch = matchCourse(rawCourse, officialCourses);

    const { englishNames, skippedArabic } = extractStudentsFromCell(String(studentsCell || ""));

    for (const arName of skippedArabic) {
      skippedEntries.push({
        id: `skip-${skippedEntries.length + 1}`,
        excelRow: r + 1,
        course: courseMatch.matchedName,
        rawText: arName,
        reason: "Arabic script (manual handling required)",
      });
    }

    for (const name of englishNames) {
      resultStudents.push({
        id: `stu-${studentCounter++}`,
        excelRow: r + 1,
        course: courseMatch.matchedName,
        studentName: name,
        isFlagged: courseMatch.isFlagged,
        confidence: courseMatch.confidence,
        suggestedCourse: courseMatch.suggestedName,
      });
    }
  }

  return {
    students: resultStudents,
    skipped: skippedEntries,
    officialCourses,
  };
}
