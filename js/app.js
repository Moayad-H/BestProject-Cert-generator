/**
 * Certificate Forge - Main Application Controller
 */

import { parseExcelWorkbook, shortenName } from "./parser.js";
import { generateCertificatesPresentation } from "./pptx.js";

// Global App State
const state = {
  templateBuffer: null,
  templateName: "",
  templateValid: false,
  
  excelWorkbook: null,
  excelName: "",
  
  students: [],
  skipped: [],
  officialCourses: [],
  
  currentStep: 1,
  searchTerm: "",
  filterFlaggedOnly: false,
  
  isGenerating: false,
};

// DOM Element References
const elements = {
  // Navigation
  stepItems: [1, 2, 3, 4].map(i => document.getElementById(`nav-step-1`.replace('1', i))),
  panels: [1, 2, 3, 4].map(i => document.getElementById(`panel-step-1`.replace('1', i))),
  
  // Step 1: Template
  templateDropzone: document.getElementById("template-dropzone"),
  templateFileInput: document.getElementById("template-file-input"),
  templateLoadedCard: document.getElementById("template-loaded-card"),
  templateFilename: document.getElementById("template-filename"),
  templateMeta: document.getElementById("template-meta"),
  templateStatusBadge: document.getElementById("template-status-badge"),
  btnChangeTemplate: document.getElementById("btn-change-template"),
  btnLoadSampleTemplate: document.getElementById("btn-load-sample-template"),
  btnToStep2: document.getElementById("btn-to-step-2"),
  
  // Step 2: Excel
  excelDropzone: document.getElementById("excel-dropzone"),
  excelFileInput: document.getElementById("excel-file-input"),
  excelLoadedCard: document.getElementById("excel-loaded-card"),
  excelFilename: document.getElementById("excel-filename"),
  excelMeta: document.getElementById("excel-meta"),
  excelStatusBadge: document.getElementById("excel-status-badge"),
  btnChangeExcel: document.getElementById("btn-change-excel"),
  btnLoadSampleExcel: document.getElementById("btn-load-sample-excel"),
  btnBackToStep1: document.getElementById("btn-back-to-step-1"),
  btnToStep3: document.getElementById("btn-to-step-3"),
  
  // Step 3: Review Table
  statTotalStudents: document.getElementById("stat-total-students"),
  statFlaggedPill: document.getElementById("stat-flagged-pill"),
  statSkippedPill: document.getElementById("stat-skipped-pill"),
  tableSearchInput: document.getElementById("table-search-input"),
  certTableBody: document.getElementById("cert-table-body"),
  btnAddStudentModal: document.getElementById("btn-add-student-modal"),
  btnOpenSkippedDrawer: document.getElementById("btn-open-skipped-drawer"),
  btnBackToStep2: document.getElementById("btn-back-to-step-2"),
  btnToStep4: document.getElementById("btn-to-step-4"),
  
  // Step 4: Generate
  generateSummaryText: document.getElementById("generate-summary-text"),
  progressContainer: document.getElementById("progress-container"),
  progressBarFill: document.getElementById("progress-bar-fill"),
  progressStatusText: document.getElementById("progress-status-text"),
  progressPercentText: document.getElementById("progress-percent-text"),
  btnGeneratePptx: document.getElementById("btn-generate-pptx"),
  btnBackToStep3: document.getElementById("btn-back-to-step-3"),
  
  // Modals
  modalAddStudent: document.getElementById("modal-add-student"),
  formAddStudent: document.getElementById("form-add-student"),
  inputAddCourse: document.getElementById("input-add-course"),
  inputAddName: document.getElementById("input-add-name"),
  btnCloseAddModal: document.getElementById("btn-close-add-modal"),
  btnCancelAddModal: document.getElementById("btn-cancel-add-modal"),
  
  modalSkippedDrawer: document.getElementById("modal-skipped-drawer"),
  skippedListContainer: document.getElementById("skipped-list-container"),
  btnCloseSkippedModal: document.getElementById("btn-close-skipped-modal"),
  btnDismissSkippedModal: document.getElementById("btn-dismiss-skipped-modal"),
};

// Initialize Application
function init() {
  bindNavigation();
  bindTemplateUpload();
  bindExcelUpload();
  bindReviewTable();
  bindGenerateStep();
  bindModals();
}

/* ==========================================================================
   Navigation
   ========================================================================== */

function goToStep(step) {
  if (step < 1 || step > 4) return;
  
  // Validation checks
  if (step > 1 && !state.templateValid) {
    alert("Please select a valid PowerPoint template first.");
    return;
  }
  if (step > 2 && state.students.length === 0) {
    alert("Please upload an Excel nomination sheet with valid students.");
    return;
  }

  state.currentStep = step;

  // Update Stepper Nav
  elements.stepItems.forEach((item, index) => {
    const stepNum = index + 1;
    item.classList.remove("active", "completed");
    if (stepNum === step) {
      item.classList.add("active");
    } else if (stepNum < step) {
      item.classList.add("completed");
    }
  });

  // Update Step Panels
  elements.panels.forEach((panel, index) => {
    panel.classList.toggle("active", index + 1 === step);
  });

  if (step === 3) {
    renderTable();
  } else if (step === 4) {
    updateGenerateSummary();
  }
}

function bindNavigation() {
  elements.stepItems.forEach((item, index) => {
    item.addEventListener("click", () => {
      goToStep(index + 1);
    });
  });

  elements.btnToStep2.addEventListener("click", () => goToStep(2));
  elements.btnBackToStep1.addEventListener("click", () => goToStep(1));
  elements.btnToStep3.addEventListener("click", () => goToStep(3));
  elements.btnBackToStep2.addEventListener("click", () => goToStep(2));
  elements.btnToStep4.addEventListener("click", () => goToStep(4));
  elements.btnBackToStep3.addEventListener("click", () => goToStep(3));
}

/* ==========================================================================
   Step 1: Template Upload & Validation
   ========================================================================== */

async function processTemplateBuffer(buffer, fileName) {
  try {
    const zip = await window.JSZip.loadAsync(buffer);
    const slide1 = zip.file("ppt/slides/slide1.xml");
    if (!slide1) {
      alert("Invalid PowerPoint: slide1.xml was not found.");
      return;
    }

    const slide1Text = await slide1.async("text");
    const hasStudentPlaceholder = slide1Text.includes("{Student_Name}");
    const hasCoursePlaceholder = slide1Text.includes("{Course_Name}");

    if (!hasStudentPlaceholder || !hasCoursePlaceholder) {
      alert("Warning: The template is missing required placeholders {Student_Name} or {Course_Name}.");
    }

    state.templateBuffer = buffer;
    state.templateName = fileName;
    state.templateValid = true;

    // Update UI
    elements.templateDropzone.style.display = "none";
    elements.templateLoadedCard.style.display = "flex";
    elements.templateFilename.textContent = fileName;
    elements.templateMeta.textContent = `1 slide • {Student_Name} & {Course_Name} verified`;
    elements.templateStatusBadge.textContent = "Ready";
    elements.templateStatusBadge.style.color = "#34d399";
    elements.btnToStep2.disabled = false;

  } catch (err) {
    console.error("Error reading PPTX:", err);
    alert("Failed to parse the PowerPoint template: " + err.message);
  }
}

function bindTemplateUpload() {
  const dropzone = elements.templateDropzone;
  const fileInput = elements.templateFileInput;

  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });
  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
  });
  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    if (e.dataTransfer.files?.length) {
      const file = e.dataTransfer.files[0];
      if (file.name.endsWith(".pptx")) {
        const reader = new FileReader();
        reader.onload = (ev) => processTemplateBuffer(ev.target.result, file.name);
        reader.readAsArrayBuffer(file);
      } else {
        alert("Please upload a .pptx PowerPoint file.");
      }
    }
  });

  fileInput.addEventListener("change", (e) => {
    if (e.target.files?.length) {
      const file = e.target.files[0];
      const reader = new FileReader();
      reader.onload = (ev) => processTemplateBuffer(ev.target.result, file.name);
      reader.readAsArrayBuffer(file);
    }
  });

  elements.btnChangeTemplate.addEventListener("click", () => {
    elements.templateDropzone.style.display = "block";
    elements.templateLoadedCard.style.display = "none";
    elements.btnToStep2.disabled = true;
    state.templateValid = false;
    fileInput.value = "";
  });

  elements.btnLoadSampleTemplate.addEventListener("click", async () => {
    try {
      elements.btnLoadSampleTemplate.disabled = true;
      elements.btnLoadSampleTemplate.textContent = "Loading...";
      const res = await fetch("./samples/Best Project Placeholder.pptx");
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const buffer = await res.arrayBuffer();
      await processTemplateBuffer(buffer, "Best Project Placeholder.pptx");
      elements.btnLoadSampleTemplate.textContent = "Loaded!";
      setTimeout(() => {
        elements.btnLoadSampleTemplate.disabled = false;
        elements.btnLoadSampleTemplate.textContent = "Use Included Template";
      }, 1500);
    } catch (err) {
      console.error(err);
      alert("Failed to load bundled sample template: " + err.message);
      elements.btnLoadSampleTemplate.disabled = false;
      elements.btnLoadSampleTemplate.textContent = "Use Included Template";
    }
  });
}

/* ==========================================================================
   Step 2: Excel Upload & Parsing
   ========================================================================== */

function processExcelBuffer(buffer, fileName) {
  try {
    const workbook = XLSX.read(buffer, { type: "array" });
    const parsed = parseExcelWorkbook(workbook);

    state.excelWorkbook = workbook;
    state.excelName = fileName;
    state.students = parsed.students;
    state.skipped = parsed.skipped;
    state.officialCourses = parsed.officialCourses;

    // Update UI
    elements.excelDropzone.style.display = "none";
    elements.excelLoadedCard.style.display = "flex";
    elements.excelFilename.textContent = fileName;
    elements.excelMeta.textContent = `${parsed.students.length} students extracted • ${parsed.skipped.length} Arabic skipped • ${parsed.officialCourses.length} official courses`;
    elements.excelStatusBadge.textContent = "Ready";
    elements.excelStatusBadge.style.color = "#34d399";
    elements.btnToStep3.disabled = parsed.students.length === 0;

  } catch (err) {
    console.error("Error reading Excel:", err);
    alert("Failed to parse Excel file: " + err.message);
  }
}

function bindExcelUpload() {
  const dropzone = elements.excelDropzone;
  const fileInput = elements.excelFileInput;

  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });
  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
  });
  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    if (e.dataTransfer.files?.length) {
      const file = e.dataTransfer.files[0];
      const reader = new FileReader();
      reader.onload = (ev) => processExcelBuffer(ev.target.result, file.name);
      reader.readAsArrayBuffer(file);
    }
  });

  fileInput.addEventListener("change", (e) => {
    if (e.target.files?.length) {
      const file = e.target.files[0];
      const reader = new FileReader();
      reader.onload = (ev) => processExcelBuffer(ev.target.result, file.name);
      reader.readAsArrayBuffer(file);
    }
  });

  elements.btnChangeExcel.addEventListener("click", () => {
    elements.excelDropzone.style.display = "block";
    elements.excelLoadedCard.style.display = "none";
    elements.btnToStep3.disabled = true;
    fileInput.value = "";
  });

  elements.btnLoadSampleExcel.addEventListener("click", async () => {
    try {
      elements.btnLoadSampleExcel.disabled = true;
      elements.btnLoadSampleExcel.textContent = "Loading...";
      const res = await fetch("./samples/Best Project Spring 2026.xlsx");
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const buffer = await res.arrayBuffer();
      processExcelBuffer(buffer, "Best Project Spring 2026.xlsx");
      elements.btnLoadSampleExcel.textContent = "Loaded!";
      setTimeout(() => {
        elements.btnLoadSampleExcel.disabled = false;
        elements.btnLoadSampleExcel.textContent = "Use Included Excel";
      }, 1500);
    } catch (err) {
      console.error(err);
      alert("Failed to load bundled sample Excel: " + err.message);
      elements.btnLoadSampleExcel.disabled = false;
      elements.btnLoadSampleExcel.textContent = "Use Included Excel";
    }
  });
}

/* ==========================================================================
   Step 3: Review Table
   ========================================================================== */

function renderTable() {
  const tbody = elements.certTableBody;
  tbody.innerHTML = "";

  const term = state.searchTerm.toLowerCase().trim();
  const filtered = state.students.filter((st) => {
    if (state.filterFlaggedOnly && !st.isFlagged) {
      return false;
    }
    if (!term) return true;
    return (
      st.course.toLowerCase().includes(term) ||
      st.studentName.toLowerCase().includes(term)
    );
  });

  // Update Toolbar Stats
  const flaggedCount = state.students.filter((s) => s.isFlagged).length;
  elements.statTotalStudents.innerHTML = `
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
    <strong>${state.students.length}</strong> Students Ready
  `;
  elements.statFlaggedPill.innerHTML = `
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3Z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>
    <strong>${flaggedCount}</strong> Flagged
  `;
  elements.statFlaggedPill.style.display = flaggedCount > 0 ? "inline-flex" : "none";
  elements.statSkippedPill.innerHTML = `
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>
    <strong>${state.skipped.length}</strong> Skipped Arabic
  `;

  if (filtered.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td colspan="4" style="text-align: center; padding: 2.5rem; color: var(--text-muted);">
        No certificate entries match the filter criteria.
      </td>
    `;
    tbody.appendChild(tr);
    return;
  }

  const fragment = document.createDocumentFragment();

  filtered.forEach((student, index) => {
    const tr = document.createElement("tr");
    if (student.isFlagged) {
      tr.classList.add("row-flagged");
    }

    const nameLen = student.studentName.length;
    let badgeHtml = "";
    if (nameLen > 20) {
      badgeHtml = `<span class="badge-char-count long" title="Length: ${nameLen} chars. Box widened to 90% and font size scaled down.">(${nameLen} chars • auto-fit)</span>`;
    } else if (nameLen > 17) {
      badgeHtml = `<span class="badge-char-count long" title="Length: ${nameLen} chars. Box widened to 90%.">(${nameLen} chars • wide box)</span>`;
    }

    let flagBadge = "";
    if (student.isFlagged) {
      const tip = student.suggestedCourse ? `Best suggestion: ${student.suggestedCourse}` : "Please check official course name";
      flagBadge = `<span class="badge-flagged" title="${tip}">Needs Match Review</span>`;
    }

    tr.innerHTML = `
      <td style="color: var(--text-muted); font-size: 0.8125rem;">${index + 1}</td>
      <td>
        <div style="display: flex; align-items: center; gap: 0.5rem;">
          <input type="text" class="table-input input-course" data-id="${student.id}" value="${escapeAttr(student.course)}">
          ${flagBadge}
        </div>
      </td>
      <td>
        <div style="display: flex; align-items: center; justify-content: space-between;">
          <input type="text" class="table-input input-student" data-id="${student.id}" value="${escapeAttr(student.studentName)}">
          ${badgeHtml}
        </div>
      </td>
      <td style="text-align: center;">
        <button class="btn btn-outline btn-sm btn-delete-row" data-id="${student.id}" title="Delete this certificate" style="padding: 0.25rem 0.5rem; color: #f87171;">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
        </button>
      </td>
    `;

    fragment.appendChild(tr);
  });

  tbody.appendChild(fragment);

  // Bind inline edit events
  tbody.querySelectorAll(".input-course").forEach((input) => {
    input.addEventListener("change", (e) => {
      const id = e.target.dataset.id;
      const val = e.target.value.trim();
      const target = state.students.find((s) => s.id === id);
      if (target) {
        target.course = val;
        target.isFlagged = false; // User confirmed manual edit
      }
    });
  });

  tbody.querySelectorAll(".input-student").forEach((input) => {
    input.addEventListener("change", (e) => {
      const id = e.target.dataset.id;
      const val = e.target.value.trim();
      const target = state.students.find((s) => s.id === id);
      if (target) {
        target.studentName = val;
        renderTable(); // Update length badges
      }
    });
  });

  tbody.querySelectorAll(".btn-delete-row").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      const id = e.currentTarget.dataset.id;
      state.students = state.students.filter((s) => s.id !== id);
      renderTable();
    });
  });
}

function bindReviewTable() {
  elements.tableSearchInput.addEventListener("input", (e) => {
    state.searchTerm = e.target.value;
    renderTable();
  });

  elements.statFlaggedPill.addEventListener("click", () => {
    state.filterFlaggedOnly = !state.filterFlaggedOnly;
    elements.statFlaggedPill.style.border = state.filterFlaggedOnly
      ? "2px solid #f59e0b"
      : "1px solid var(--warning-border)";
    renderTable();
  });

  elements.statSkippedPill.addEventListener("click", () => {
    openSkippedDrawer();
  });

  elements.btnOpenSkippedDrawer.addEventListener("click", () => {
    openSkippedDrawer();
  });
}

/* ==========================================================================
   Modals & Drawers
   ========================================================================== */

function bindModals() {
  // Add Student Modal
  elements.btnAddStudentModal.addEventListener("click", () => {
    elements.modalAddStudent.classList.add("active");
    elements.inputAddCourse.value = "";
    elements.inputAddName.value = "";
    elements.inputAddCourse.focus();
  });

  const closeAddModal = () => elements.modalAddStudent.classList.remove("active");
  elements.btnCloseAddModal.addEventListener("click", closeAddModal);
  elements.btnCancelAddModal.addEventListener("click", closeAddModal);

  elements.formAddStudent.addEventListener("submit", (e) => {
    e.preventDefault();
    const course = elements.inputAddCourse.value.trim();
    const name = elements.inputAddName.value.trim();

    if (course && name) {
      state.students.unshift({
        id: `custom-${Date.now()}`,
        excelRow: 0,
        course,
        studentName: shortenName(name),
        isFlagged: false,
        confidence: 1.0,
      });
      closeAddModal();
      renderTable();
    }
  });

  // Skipped Drawer
  const closeSkippedModal = () => elements.modalSkippedDrawer.classList.remove("active");
  elements.btnCloseSkippedModal.addEventListener("click", closeSkippedModal);
  elements.btnDismissSkippedModal.addEventListener("click", closeSkippedModal);
}

function openSkippedDrawer() {
  const container = elements.skippedListContainer;
  container.innerHTML = "";

  if (state.skipped.length === 0) {
    container.innerHTML = `<p style="color: var(--text-muted); text-align: center; padding: 2rem;">No skipped Arabic entries.</p>`;
  } else {
    state.skipped.forEach((item, index) => {
      const div = document.createElement("div");
      div.className = "skipped-item";
      div.innerHTML = `
        <div class="skipped-info">
          <div class="skipped-name">${escapeAttr(item.rawText)}</div>
          <div class="skipped-course">Course: <strong>${escapeAttr(item.course)}</strong> (Row ${item.excelRow})</div>
        </div>
        <div style="display: flex; gap: 0.5rem;">
          <button class="btn btn-outline btn-sm btn-copy-skipped" data-text="${escapeAttr(item.rawText)}" title="Copy Arabic text">
            Copy
          </button>
          <button class="btn btn-secondary btn-sm btn-transcribe-skipped" data-index="${index}" title="Transcribe and add to certificates">
            Add as English
          </button>
        </div>
      `;
      container.appendChild(div);
    });

    // Copy action
    container.querySelectorAll(".btn-copy-skipped").forEach((btn) => {
      btn.addEventListener("click", async (e) => {
        const text = e.target.dataset.text;
        await navigator.clipboard.writeText(text);
        e.target.textContent = "Copied!";
        setTimeout(() => (e.target.textContent = "Copy"), 1500);
      });
    });

    // Transcribe action
    container.querySelectorAll(".btn-transcribe-skipped").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        const item = state.skipped[idx];
        const engName = prompt(`Enter English name for "${item.rawText}":`);
        if (engName && engName.trim()) {
          state.students.push({
            id: `transcribed-${Date.now()}`,
            excelRow: item.excelRow,
            course: item.course,
            studentName: shortenName(engName.trim()),
            isFlagged: false,
            confidence: 1.0,
          });
          state.skipped.splice(idx, 1);
          openSkippedDrawer();
          renderTable();
        }
      });
    });
  }

  elements.modalSkippedDrawer.classList.add("active");
}

/* ==========================================================================
   Step 4: Generate & Export
   ========================================================================== */

function updateGenerateSummary() {
  const count = state.students.length;
  elements.generateSummaryText.textContent = `Ready to generate ${count} certificates into "Best_Project_Certificates.pptx".`;
  elements.progressContainer.classList.remove("active");
  elements.btnGeneratePptx.disabled = count === 0;
}

function bindGenerateStep() {
  elements.btnGeneratePptx.addEventListener("click", async () => {
    if (state.isGenerating || state.students.length === 0) return;

    try {
      state.isGenerating = true;
      elements.btnGeneratePptx.disabled = true;
      elements.progressContainer.classList.add("active");
      elements.progressBarFill.style.width = "0%";

      const outputBlob = await generateCertificatesPresentation(
        state.templateBuffer,
        state.students,
        (current, total, message) => {
          const percent = total > 0 ? Math.round((current / total) * 100) : 0;
          elements.progressBarFill.style.width = `${percent}%`;
          elements.progressPercentText.textContent = `${percent}%`;
          elements.progressStatusText.textContent = message;
        }
      );

      // Trigger automatic download
      const downloadUrl = URL.createObjectURL(outputBlob);
      const a = document.createElement("a");
      a.href = downloadUrl;
      a.download = "Best_Project_Certificates.pptx";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);

      setTimeout(() => URL.revokeObjectURL(downloadUrl), 10000);

      elements.progressStatusText.textContent = "Download started! Best_Project_Certificates.pptx generated successfully.";
      elements.progressBarFill.style.width = "100%";
      elements.progressPercentText.textContent = "100%";

    } catch (err) {
      console.error("PPTX Generation error:", err);
      alert("Error generating PowerPoint certificates: " + err.message);
    } finally {
      state.isGenerating = false;
      elements.btnGeneratePptx.disabled = false;
    }
  });
}

/* ==========================================================================
   Utility Helpers
   ========================================================================== */

function escapeAttr(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Start application when DOM is ready
document.addEventListener("DOMContentLoaded", init);
