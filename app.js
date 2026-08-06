/**
 * Competency Assessment 5.0 Dashboard - Application Controller
 * Handles Data State, Dynamic Charts, IDP Management, and Backend API Integration
 */

// --------------------------------------------------------------------------
// 1. DATA STATE & BACKEND CONTRACT (Extracted from PPTX)
// --------------------------------------------------------------------------
const CompetencyData = {
  metadata: {
    year: 2024,
    company: "RX Company",
    totalEmployees: 281,
    corporateTargetGap: 30.0,
    groups: {
      office: { name: "กลุ่มงานสำนักงาน", count: 124 },
      sales: { name: "กลุ่มงานขาย", count: 102 },
      lasik: { name: "กลุ่มงานเลสิก", count: 56 }
    },
    levels: {
      staff: { level: 1, name: "ระดับพนักงาน", expectedLevel: 3, count: 233 },
      supervisor: { level: 2, name: "ระดับหัวหน้างาน", expectedLevel: 4, count: 34 },
      manager: { level: 3, name: "ระดับผู้จัดการ", expectedLevel: 5, count: 13 }
    }
  },

  // Target Goal vs History
  entrepreneurshipGoal: {
    target: 30.0,
    year2023: 22.67,
    year2024: 20.64,
    status: "PASSED"
  },

  // Core Competency Detailed Data (5 Pillars)
  coreCompetency: {
    overall: [
      { id: "relationship", name: "Relationship Excellence", total: 281, gapCount: 18, gapPct: 6.41 },
      { id: "fact", name: "Fact", total: 281, gapCount: 36, gapPct: 12.81 },
      { id: "innovative", name: "Innovative Thinking", total: 281, gapCount: 72, gapPct: 25.62 },
      { id: "learning", name: "Learning", total: 281, gapCount: 88, gapPct: 31.32 },
      { id: "entrepreneurship", name: "Entrepreneurship", total: 281, gapCount: 58, gapPct: 20.64 }
    ],
    byLevel: {
      staff: [
        { name: "Relationship", count: 233, gapCount: 3, gapPct: 1.29 },
        { name: "Fact", count: 233, gapCount: 11, gapPct: 4.72 },
        { name: "Innovative", count: 233, gapCount: 44, gapPct: 18.88 },
        { name: "Learning", count: 233, gapCount: 68, gapPct: 29.18 },
        { name: "Entrepreneurship", count: 233, gapCount: 35, gapPct: 15.02 }
      ],
      supervisor: [
        { name: "Relationship", count: 34, gapCount: 8, gapPct: 23.53 },
        { name: "Fact", count: 34, gapCount: 15, gapPct: 44.12 },
        { name: "Innovative", count: 34, gapCount: 20, gapPct: 58.82 },
        { name: "Learning", count: 34, gapCount: 13, gapPct: 38.24 },
        { name: "Entrepreneurship", count: 34, gapCount: 17, gapPct: 50.00 }
      ],
      manager: [
        { name: "Relationship", count: 14, gapCount: 7, gapPct: 50.00 },
        { name: "Fact", count: 14, gapCount: 10, gapPct: 71.43 },
        { name: "Innovative", count: 14, gapCount: 8, gapPct: 57.14 },
        { name: "Learning", count: 14, gapCount: 7, gapPct: 50.00 },
        { name: "Entrepreneurship", count: 14, gapCount: 6, gapPct: 42.86 }
      ]
    },
    byGroup: {
      office: [
        { name: "Relationship", count: 123, gapCount: 3, gapPct: 2.44 },
        { name: "Fact", count: 123, gapCount: 13, gapPct: 10.57 },
        { name: "Innovative", count: 123, gapCount: 35, gapPct: 28.46 },
        { name: "Learning", count: 123, gapCount: 49, gapPct: 39.84 },
        { name: "Entrepreneurship", count: 123, gapCount: 27, gapPct: 21.95 }
      ],
      sales: [
        { name: "Relationship", count: 102, gapCount: 8, gapPct: 7.84 },
        { name: "Fact", count: 102, gapCount: 15, gapPct: 14.71 },
        { name: "Innovative", count: 102, gapCount: 26, gapPct: 25.49 },
        { name: "Learning", count: 102, gapCount: 23, gapPct: 22.55 },
        { name: "Entrepreneurship", count: 102, gapCount: 19, gapPct: 18.63 }
      ],
      lasik: [
        { name: "Relationship", count: 56, gapCount: 7, gapPct: 12.50 },
        { name: "Fact", count: 56, gapCount: 8, gapPct: 14.29 },
        { name: "Innovative", count: 56, gapCount: 11, gapPct: 19.64 },
        { name: "Learning", count: 56, gapCount: 16, gapPct: 28.57 },
        { name: "Entrepreneurship", count: 56, gapCount: 12, gapPct: 21.43 }
      ]
    }
  },

  // Job Competencies Top Highlights
  jobCompetency: {
    overall: [
      { name: "ความรู้ด้านระบบคุณภาพ", total: 121, gapCount: 55, gapPct: 45.45, category: "Knowledge" },
      { name: "การวิเคราะห์ข้อมูล", total: 128, gapCount: 40, gapPct: 31.25, category: "Skill" },
      { name: "การสื่อสาร (Communication)", total: 162, gapCount: 33, gapPct: 20.37, category: "Skill" },
      { name: "การแก้ไขปัญหาและการตัดสินใจ", total: 238, gapCount: 41, gapPct: 17.23, category: "Skill" },
      { name: "การปฏิบัติงานตามขั้นตอน & เอกสาร", total: 158, gapCount: 14, gapPct: 8.86, category: "Skill" }
    ],
    byLevelHigh: {
      staff: [
        { name: "ความรู้ด้านระบบคุณภาพ", total: 103, gapCount: 48, gapPct: 46.60 },
        { name: "การสื่อสาร", total: 161, gapCount: 33, gapPct: 20.50 },
        { name: "การแก้ไขปัญหาและการตัดสินใจ", total: 231, gapCount: 35, gapPct: 15.15 },
        { name: "การปฏิบัติงานตามขั้นตอน & เอกสาร", total: 157, gapCount: 14, gapPct: 8.92 }
      ],
      supervisor: [
        { name: "การวิเคราะห์ข้อมูล", total: 34, gapCount: 22, gapPct: 64.71 },
        { name: "การบริหารจัดการ", total: 24, gapCount: 12, gapPct: 50.00 },
        { name: "การปฏิบัติงานเชิงกลยุทธ์", total: 30, gapCount: 14, gapPct: 46.67 }
      ],
      manager: [
        { name: "ภาวะผู้นำ (Leadership)", total: 11, gapCount: 9, gapPct: 81.82 },
        { name: "การสอนงานและพัฒนาพนักงาน (Coaching)", total: 10, gapCount: 7, gapPct: 70.00 },
        { name: "การวิเคราะห์ข้อมูล", total: 10, gapCount: 7, gapPct: 70.00 },
        { name: "การปฏิบัติงานเชิงกลยุทธ์", total: 12, gapCount: 8, gapPct: 66.67 },
        { name: "การบริหารบุคลากร", total: 11, gapCount: 7, gapPct: 63.64 }
      ]
    }
  },

  // IDP Action Plans Mock
  idpActions: [
    { id: "IDP-001", name: "สมชาย มีสุข", level: "หัวหน้างาน", group: "สำนักงาน", gapItem: "การวิเคราะห์ข้อมูล (64.71%)", status: "กำลังดำเนินการ", dueDate: "2024-12-15" },
    { id: "IDP-002", name: "วิภาดา เด่นดี", level: "ผู้จัดการ", group: "ขาย", gapItem: "ภาวะผู้นำ (81.82%)", status: "รอดำเนินการ", dueDate: "2024-12-30" },
    { id: "IDP-003", name: "อนันต์ ชัยชนะ", level: "พนักงาน", group: "เลสิก", gapItem: "Learning (28.57%)", status: "อนุมัติแล้ว", dueDate: "2024-11-30" },
    { id: "IDP-004", name: "นภาลัย สว่าง", level: "พนักงาน", group: "สำนักงาน", gapItem: "ความรู้ด้านระบบคุณภาพ (46.6%)", status: "กำลังดำเนินการ", dueDate: "2024-12-20" }
  ]
};

// --------------------------------------------------------------------------
// 2. CHART INSTANCES & SETUP
// --------------------------------------------------------------------------
let charts = {};

document.addEventListener("DOMContentLoaded", () => {
  initTabsNav();
  initCharts();
  renderSummaryTables();
  renderIDPTable();
  initModalListeners();
});

// Tab Navigation logic
function initTabsNav() {
  const navButtons = document.querySelectorAll(".nav-item button");
  const tabPanels = document.querySelectorAll(".tab-panel");

  navButtons.forEach(btn => {
    btn.addEventListener("click", () => {
      const tabId = btn.getAttribute("data-tab");

      navButtons.forEach(b => b.parentElement.classList.remove("active"));
      tabPanels.forEach(p => p.classList.remove("active"));

      btn.parentElement.classList.add("active");
      const activePanel = document.getElementById(`tab-${tabId}`);
      if (activePanel) {
        activePanel.classList.add("active");
      }
    });
  });
}

// Chart.js Initialization
function initCharts() {
  Chart.defaults.color = '#94a3b8';
  Chart.defaults.font.family = "'Plus Jakarta Sans', 'Prompt', sans-serif";

  // 1. Core Competency Radar Chart
  const ctxRadar = document.getElementById("chartCoreRadar")?.getContext("2d");
  if (ctxRadar) {
    charts.coreRadar = new Chart(ctxRadar, {
      type: "radar",
      data: {
        labels: ["Relationship", "Fact", "Innovative", "Learning", "Entrepreneurship"],
        datasets: [
          {
            label: "Overall Gap (%)",
            data: [6.41, 12.81, 25.62, 31.32, 20.64],
            backgroundColor: "rgba(99, 102, 241, 0.25)",
            borderColor: "#6366f1",
            pointBackgroundColor: "#6366f1",
            borderWidth: 2
          },
          {
            label: "Target Gap Max (%)",
            data: [30, 30, 30, 30, 30],
            backgroundColor: "transparent",
            borderColor: "rgba(244, 63, 94, 0.5)",
            borderDash: [5, 5],
            pointRadius: 0
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          r: {
            angleLines: { color: "rgba(255, 255, 255, 0.1)" },
            grid: { color: "rgba(255, 255, 255, 0.08)" },
            pointLabels: { color: "#f8fafc", font: { size: 12, weight: "bold" } },
            ticks: { backdropColor: "transparent", color: "#64748b" }
          }
        },
        plugins: {
          legend: { position: "bottom" }
        }
      }
    });
  }

  // 2. Core Gap by Job Level Bar Chart
  const ctxLevel = document.getElementById("chartCoreByLevel")?.getContext("2d");
  if (ctxLevel) {
    charts.coreByLevel = new Chart(ctxLevel, {
      type: "bar",
      data: {
        labels: ["Relationship", "Fact", "Innovative", "Learning", "Entrepreneurship"],
        datasets: [
          {
            label: "พนักงาน (233 ท่าน)",
            data: [1.29, 4.72, 18.88, 29.18, 15.02],
            backgroundColor: "#06b6d4"
          },
          {
            label: "หัวหน้างาน (34 ท่าน)",
            data: [23.53, 44.12, 58.82, 38.24, 50.00],
            backgroundColor: "#f59e0b"
          },
          {
            label: "ผู้จัดการ (14 ท่าน)",
            data: [50.00, 71.43, 57.14, 50.00, 42.86],
            backgroundColor: "#f43f5e"
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { grid: { display: false } },
          y: {
            grid: { color: "rgba(255, 255, 255, 0.08)" },
            ticks: { callback: v => v + "%" }
          }
        },
        plugins: {
          legend: { position: "bottom" }
        }
      }
    });
  }

  // 3. Top Job Competencies Gap Bar Chart
  const ctxJobTop = document.getElementById("chartJobTop")?.getContext("2d");
  if (ctxJobTop) {
    charts.jobTop = new Chart(ctxJobTop, {
      type: "bar",
      indexAxis: "y",
      data: {
        labels: [
          "ความรู้ด้านระบบคุณภาพ",
          "การวิเคราะห์ข้อมูล",
          "การสื่อสาร",
          "การแก้ไขปัญหา & ตัดสินใจ",
          "การปฏิบัติงานตามขั้นตอน & เอกสาร"
        ],
        datasets: [{
          label: "% Gap ในองค์กร",
          data: [45.45, 31.25, 20.37, 17.23, 8.86],
          backgroundColor: ["#f43f5e", "#f59e0b", "#6366f1", "#06b6d4", "#10b981"]
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { ticks: { callback: v => v + "%" }, grid: { color: "rgba(255, 255, 255, 0.08)" } },
          y: { grid: { display: false } }
        },
        plugins: { legend: { display: false } }
      }
    });
  }
}

// --------------------------------------------------------------------------
// 3. TABLE RENDERING FUNCTIONS
// --------------------------------------------------------------------------
function renderSummaryTables() {
  // Core Summary Table
  const tableBody = document.getElementById("tableCoreSummaryBody");
  if (tableBody) {
    const data = CompetencyData.coreCompetency.overall;
    tableBody.innerHTML = data.map(item => `
      <tr>
        <td><strong>${item.name}</strong></td>
        <td>${item.total}</td>
        <td><span class="gap-tag gap-${item.gapPct > 25 ? 'high' : item.gapPct > 15 ? 'medium' : 'low'}">${item.gapCount} ท่าน</span></td>
        <td>
          <div style="display:flex; align-items:center; gap:0.75rem;">
            <div class="progress-bar-bg" style="flex:1;">
              <div class="progress-bar-fill fill-${item.gapPct > 25 ? 'danger' : item.gapPct > 15 ? 'warning' : 'primary'}" style="width: ${item.gapPct}%;"></div>
            </div>
            <strong style="min-width: 45px;">${item.gapPct}%</strong>
          </div>
        </td>
      </tr>
    `).join("");
  }
}

function renderIDPTable() {
  const tbody = document.getElementById("tableIDPBody");
  if (tbody) {
    tbody.innerHTML = CompetencyData.idpActions.map(action => `
      <tr>
        <td><code>${action.id}</code></td>
        <td><strong>${action.name}</strong></td>
        <td>${action.level}</td>
        <td>${action.group}</td>
        <td><span class="gap-tag gap-high">${action.gapItem}</span></td>
        <td><span class="kpi-badge badge-${action.status === 'อนุมัติแล้ว' ? 'success' : 'warning'}">${action.status}</span></td>
        <td>${action.dueDate}</td>
        <td><button class="btn-secondary" style="padding:0.3rem 0.6rem; font-size:0.75rem;" onclick="alert('แก้ไขแผน IDP สำหรับ ${action.name}')">จัดการ</button></td>
      </tr>
    `).join("");
  }
}

// --------------------------------------------------------------------------
// 4. BACKEND INTEGRATION & DEVELOPER DRAWER
// --------------------------------------------------------------------------
function initModalListeners() {
  const modal = document.getElementById("apiModal");
  const openBtn = document.getElementById("btnOpenApiSpec");
  const closeBtn = document.getElementById("btnCloseApiSpec");

  if (openBtn && modal) {
    openBtn.addEventListener("click", () => {
      modal.classList.add("active");
    });
  }

  if (closeBtn && modal) {
    closeBtn.addEventListener("click", () => {
      modal.classList.remove("active");
    });
  }
}

// Global API Object exposes clean endpoints contract for backend developers
window.CompetencyAPIService = {
  getSummary: () => CompetencyData,
  getCoreGaps: () => CompetencyData.coreCompetency,
  getJobGaps: () => CompetencyData.jobCompetency,
  getIDPActions: () => CompetencyData.idpActions,
  createIDPPlan: (idpPayload) => {
    CompetencyData.idpActions.push(idpPayload);
    renderIDPTable();
    return { success: true, message: "IDP Plan Created", data: idpPayload };
  }
};
