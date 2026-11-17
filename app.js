// Main Application Logic for PICT Marks Portal - Rank Upfront (Prominent & Immediately Visible)
document.addEventListener("DOMContentLoaded", () => {
  if (typeof PICT_DATA === "undefined") {
    console.error("PICT_DATA not loaded!");
    return;
  }

  const { meta, divisions, students } = PICT_DATA;

  // App State
  const state = {
    currentTab: "overview",
    selectedDivision: "all",
    searchQuery: "",
    sortColumn: "roll_no",
    sortAsc: true,
    currentPage: 1,
    rowsPerPage: 25,
    compareDiv1: "FY1",
    compareDiv2: "FY8"
  };

  // DOM Elements
  const tabButtons = document.querySelectorAll(".nav-tab");
  const viewSections = document.querySelectorAll(".view-section");
  const divSelectorBar = document.getElementById("divisionSelectorBar");
  const tableSearchInput = document.getElementById("tableSearchInput");
  const globalSearchInput = document.getElementById("globalSearchInput");
  const rowsPerPageSelect = document.getElementById("rowsPerPageSelect");
  const tableHead = document.getElementById("tableHead");
  const tableBody = document.getElementById("tableBody");
  const tableCountInfo = document.getElementById("tableCountInfo");
  const paginationControls = document.getElementById("paginationControls");
  const studentModal = document.getElementById("studentModal");
  const closeModalBtn = document.getElementById("closeModalBtn");
  const themeToggleBtn = document.getElementById("themeToggleBtn");
  const exportCsvBtn = document.getElementById("exportCsvBtn");

  // Theme Init
  const savedTheme = localStorage.getItem("pict_theme") || "light";
  document.documentElement.setAttribute("data-theme", savedTheme);
  updateThemeIcon(savedTheme);

  themeToggleBtn.addEventListener("click", () => {
    const cur = document.documentElement.getAttribute("data-theme");
    const next = cur === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", next);
    localStorage.setItem("pict_theme", next);
    updateThemeIcon(next);
    renderCharts();
  });

  function updateThemeIcon(theme) {
    themeToggleBtn.innerHTML = theme === "dark" 
      ? `<svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364 6.364l-.707-.707M6.343 6.343l-.707-.707m12.728 0l-.707.707M6.343 17.657l-.707.707M16 12a4 4 0 11-8 0 4 4 0 018 0z"></path></svg>`
      : `<svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z"></path></svg>`;
  }

  // Tab Navigation
  tabButtons.forEach(btn => {
    btn.addEventListener("click", () => {
      const target = btn.dataset.tab;
      switchTab(target);
    });
  });

  function switchTab(tabId) {
    state.currentTab = tabId;
    tabButtons.forEach(b => b.classList.toggle("active", b.dataset.tab === tabId));
    viewSections.forEach(sec => sec.classList.toggle("active", sec.id === `${tabId}Section`));

    if (tabId === "overview") {
      renderCharts();
    } else if (tabId === "tables") {
      renderTable();
    } else if (tabId === "toppers") {
      renderToppersView();
    } else if (tabId === "compare") {
      renderComparisonTool();
    }
  }

  // Initialize KPI Cards
  function renderKPICards() {
    document.getElementById("kpiTotalStudents").textContent = meta.total_students;
    document.getElementById("kpiAverageScore").textContent = `${meta.overall_average} / 100`;
    document.getElementById("kpiTopScore").textContent = `${meta.highest_score} / 100`;
    if (document.getElementById("kpiDivCount")) {
      document.getElementById("kpiDivCount").textContent = "13 Divisions";
    }
    
    const top = meta.college_toppers[0];
    if (top) {
      document.getElementById("kpiTopperName").textContent = `${top.name} (${top.div})`;
    }
  }

  // Division Selector Pills
  function renderDivisionSelector() {
    divSelectorBar.innerHTML = "";
    
    const allBtn = document.createElement("button");
    allBtn.className = `div-btn ${state.selectedDivision === "all" ? "active" : ""}`;
    allBtn.innerHTML = `<span>All Divisions (${meta.total_students})</span>`;
    allBtn.addEventListener("click", () => {
      state.selectedDivision = "all";
      state.currentPage = 1;
      updateDivPills();
      renderTable();
    });
    divSelectorBar.appendChild(allBtn);

    for (let i = 1; i <= 13; i++) {
      const code = `FY${i}`;
      const divInfo = divisions[code];
      const count = divInfo ? divInfo.total_students : 0;
      const btn = document.createElement("button");
      btn.className = `div-btn ${state.selectedDivision === code ? "active" : ""}`;
      btn.innerHTML = `<span>Div ${i} (${code}) [${count}]</span>`;
      btn.addEventListener("click", () => {
        state.selectedDivision = code;
        state.currentPage = 1;
        updateDivPills();
        renderTable();
      });
      divSelectorBar.appendChild(btn);
    }
  }

  function updateDivPills() {
    const btns = divSelectorBar.querySelectorAll(".div-btn");
    btns.forEach((btn, idx) => {
      if (idx === 0) {
        btn.classList.toggle("active", state.selectedDivision === "all");
      } else {
        btn.classList.toggle("active", state.selectedDivision === `FY${idx}`);
      }
    });
  }

  // Filter & Search Event Listeners
  tableSearchInput.addEventListener("input", (e) => {
    state.searchQuery = e.target.value.trim().toLowerCase();
    state.currentPage = 1;
    renderTable();
  });

  rowsPerPageSelect.addEventListener("change", (e) => {
    state.rowsPerPage = e.target.value === "all" ? 999999 : parseInt(e.target.value);
    state.currentPage = 1;
    renderTable();
  });

  // Table Data Filtering & Sorting
  function getFilteredStudents() {
    return students.filter(s => {
      if (state.selectedDivision !== "all" && s.division !== state.selectedDivision) {
        return false;
      }
      if (state.searchQuery) {
        const q = state.searchQuery;
        const matchesRoll = String(s.roll_no).includes(q);
        const matchesMis = s.mis_id.toLowerCase().includes(q);
        const matchesName = s.name.toLowerCase().includes(q);
        const matchesDiv = s.division.toLowerCase().includes(q);
        if (!matchesRoll && !matchesMis && !matchesName && !matchesDiv) {
          return false;
        }
      }
      return true;
    }).sort((a, b) => {
      let valA = a[state.sortColumn];
      let valB = b[state.sortColumn];

      if (state.sortColumn.startsWith("subj_")) {
        const sub = state.sortColumn.replace("subj_", "");
        valA = a.marks[sub] !== null && a.marks[sub] !== undefined ? a.marks[sub] : -1;
        valB = b.marks[sub] !== null && b.marks[sub] !== undefined ? b.marks[sub] : -1;
      }

      if (valA === undefined) valA = "";
      if (valB === undefined) valB = "";

      if (typeof valA === "number" && typeof valB === "number") {
        return state.sortAsc ? valA - valB : valB - valA;
      }
      return state.sortAsc ? String(valA).localeCompare(String(valB)) : String(valB).localeCompare(String(valA));
    });
  }

  // Render Table: Class Rank and College Rank placed UPFRONT right next to Roll No and Name
  function renderTable() {
    const filtered = getFilteredStudents();
    const totalCount = filtered.length;

    let headers = [];
    let isSingleDiv = state.selectedDivision !== "all";
    let divNum = isSingleDiv ? parseInt(state.selectedDivision.replace("FY", "")) : null;

    if (!isSingleDiv) {
      // All Divisions View
      headers = [
        { key: "division", label: "Div", sortable: true },
        { key: "roll_no", label: "Roll No.", sortable: true },
        { key: "div_rank", label: "Class Rank", sortable: true, highlight: "div_rank" },
        { key: "college_rank", label: "College Rank", sortable: true, highlight: "college_rank" },
        { key: "name", label: "Student Name", sortable: true },
        { key: "mis_id", label: "MIS ID", sortable: true },
        { key: "subj_0", label: "Sub 1 (/20)", sortable: false },
        { key: "subj_1", label: "Sub 2 (/20)", sortable: false },
        { key: "subj_2", label: "Sub 3 (/20)", sortable: false },
        { key: "subj_3", label: "Sub 4 (/20)", sortable: false },
        { key: "subj_4", label: "Sub 5 (/20)", sortable: false },
        { key: "total", label: "Total (/100)", sortable: true },
        { key: "percentage", label: "Percentage (%)", sortable: true }
      ];
    } else {
      // Single Division View
      const subjs = divNum <= 7 
        ? ['lac', 'cpps', 'mfr', 'qp', 'ieee']
        : ['lac', 'cpps', 'cst', 'ese', 'eg'];
      
      const subLabels = divNum <= 7
        ? ['LAC (/20)', 'CPPS (/20)', 'MFR (/20)', 'QP (/20)', 'IEEE (/20)']
        : ['LAC (/20)', 'CPPS (/20)', 'CST (/20)', 'ESE (/20)', 'EG (/20)'];

      headers = [
        { key: "roll_no", label: "Roll No.", sortable: true },
        { key: "div_rank", label: "Class Rank", sortable: true, highlight: "div_rank" },
        { key: "college_rank", label: "College Rank", sortable: true, highlight: "college_rank" },
        { key: "name", label: "Student Name", sortable: true },
        { key: "mis_id", label: "MIS ID", sortable: true },
        { key: `subj_${subjs[0]}`, label: subLabels[0], sortable: true },
        { key: `subj_${subjs[1]}`, label: subLabels[1], sortable: true },
        { key: `subj_${subjs[2]}`, label: subLabels[2], sortable: true },
        { key: `subj_${subjs[3]}`, label: subLabels[3], sortable: true },
        { key: `subj_${subjs[4]}`, label: subLabels[4], sortable: true },
        { key: "total", label: "Total (/100)", sortable: true },
        { key: "percentage", label: "Percentage (%)", sortable: true }
      ];
    }

    const tableTitleEl = document.getElementById("currentDivisionTableTitle");
    if (tableTitleEl) {
      if (isSingleDiv) {
        tableTitleEl.textContent = `Division ${divNum} (${state.selectedDivision}) Student Marks & Rankings`;
      } else {
        tableTitleEl.textContent = `All Divisions (Master List) - Student Marks & Rankings`;
      }
    }

    // Build Table Head with distinct styling for Rank columns
    let headHtml = "<tr>";
    headers.forEach(h => {
      let sortClass = "";
      if (h.sortable) {
        sortClass = "sortable";
        if (state.sortColumn === h.key) {
          sortClass += state.sortAsc ? " sorted-asc" : " sorted-desc";
        }
      }
      let extraStyle = "";
      if (h.highlight === "div_rank") {
        extraStyle = "background: rgba(245, 158, 11, 0.15); color: #d97706; text-align: center;";
      } else if (h.highlight === "college_rank") {
        extraStyle = "background: rgba(147, 51, 234, 0.15); color: #9333ea; text-align: center;";
      }
      headHtml += `<th class="${sortClass}" data-key="${h.key}" style="${extraStyle}">${h.label}</th>`;
    });
    headHtml += "</tr>";
    tableHead.innerHTML = headHtml;

    tableHead.querySelectorAll("th.sortable").forEach(th => {
      th.addEventListener("click", () => {
        const key = th.dataset.key;
        if (state.sortColumn === key) {
          state.sortAsc = !state.sortAsc;
        } else {
          state.sortColumn = key;
          // default ascending for ranks and rolls (rank 1 first), descending for marks
          if (["div_rank", "college_rank", "roll_no"].includes(key)) {
            state.sortAsc = true;
          } else {
            state.sortAsc = false;
          }
        }
        renderTable();
      });
    });

    const totalPages = Math.ceil(totalCount / state.rowsPerPage) || 1;
    if (state.currentPage > totalPages) state.currentPage = totalPages;
    const startIndex = (state.currentPage - 1) * state.rowsPerPage;
    const paginated = filtered.slice(startIndex, startIndex + state.rowsPerPage);

    if (paginated.length === 0) {
      tableBody.innerHTML = `<tr><td colspan="${headers.length}" style="text-align:center; padding: 3rem; color: var(--text-muted);">No students found matching current filters.</td></tr>`;
    } else {
      let bodyHtml = "";
      paginated.forEach(s => {
        bodyHtml += `<tr data-roll="${s.roll_no}">`;
        
        if (!isSingleDiv) {
          bodyHtml += `<td><strong>${s.division}</strong></td>`;
        }

        bodyHtml += `
          <td><strong>${s.roll_no}</strong></td>
          <td style="text-align:center; font-weight:800; font-size:1rem; color:#d97706; background: rgba(245, 158, 11, 0.05);">${s.div_rank}</td>
          <td style="text-align:center; font-weight:800; font-size:1rem; color:#9333ea; background: rgba(147, 51, 234, 0.05);">${s.college_rank}</td>
          <td><strong>${s.name}</strong></td>
          <td><code>${s.mis_id}</code></td>
        `;

        // 5 subject marks
        s.subjects.forEach(sub => {
          const val = s.marks[sub];
          const displayVal = (val !== null && val !== undefined) ? `${val}/20` : 'AB';
          const style = (val === null || val === undefined) ? 'color: var(--danger); font-weight:700;' : '';
          bodyHtml += `<td style="text-align:center; ${style}">${displayVal}</td>`;
        });

        // Total, Percentage
        bodyHtml += `
          <td style="text-align:center; font-weight:800; color:var(--brand-blue);">${s.total} / 100</td>
          <td style="text-align:center; font-weight:700;">${s.percentage}%</td>
        </tr>`;
      });
      tableBody.innerHTML = bodyHtml;
    }

    tableBody.querySelectorAll("tr[data-roll]").forEach(tr => {
      tr.addEventListener("click", () => {
        const roll = parseInt(tr.dataset.roll);
        const stu = students.find(s => s.roll_no === roll);
        if (stu) openStudentModal(stu);
      });
    });

    const endCount = Math.min(startIndex + state.rowsPerPage, totalCount);
    tableCountInfo.textContent = `Showing ${totalCount > 0 ? startIndex + 1 : 0} to ${endCount} of ${totalCount} students`;

    renderPagination(totalPages);
  }

  function renderPagination(totalPages) {
    paginationControls.innerHTML = "";
    if (totalPages <= 1) return;

    const prevBtn = document.createElement("button");
    prevBtn.className = "page-btn";
    prevBtn.textContent = "« Prev";
    prevBtn.disabled = state.currentPage === 1;
    prevBtn.addEventListener("click", () => {
      if (state.currentPage > 1) {
        state.currentPage--;
        renderTable();
      }
    });
    paginationControls.appendChild(prevBtn);

    let pages = [];
    if (totalPages <= 7) {
      for (let p = 1; p <= totalPages; p++) pages.push(p);
    } else {
      pages.push(1);
      if (state.currentPage > 3) pages.push("...");
      const start = Math.max(2, state.currentPage - 1);
      const end = Math.min(totalPages - 1, state.currentPage + 1);
      for (let p = start; p <= end; p++) pages.push(p);
      if (state.currentPage < totalPages - 2) pages.push("...");
      pages.push(totalPages);
    }

    pages.forEach(p => {
      if (p === "...") {
        const span = document.createElement("span");
        span.textContent = "...";
        span.style.padding = "0.4rem 0.5rem";
        paginationControls.appendChild(span);
      } else {
        const pBtn = document.createElement("button");
        pBtn.className = `page-btn ${state.currentPage === p ? "active" : ""}`;
        pBtn.textContent = p;
        pBtn.addEventListener("click", () => {
          state.currentPage = p;
          renderTable();
        });
        paginationControls.appendChild(pBtn);
      }
    });

    const nextBtn = document.createElement("button");
    nextBtn.className = "page-btn";
    nextBtn.textContent = "Next »";
    nextBtn.disabled = state.currentPage === totalPages;
    nextBtn.addEventListener("click", () => {
      if (state.currentPage < totalPages) {
        state.currentPage++;
        renderTable();
      }
    });
    paginationControls.appendChild(nextBtn);
  }

  // Student Report Card Modal
  function openStudentModal(stu) {
    const initials = stu.name.split(" ").map(w => w[0]).filter(Boolean).slice(0, 2).join("").toUpperCase();
    
    document.getElementById("modalStudentAvatar").textContent = initials;
    document.getElementById("modalStudentName").textContent = stu.name;
    document.getElementById("modalStudentMeta").textContent = `Roll No: ${stu.roll_no} | MIS ID: ${stu.mis_id} | Division: ${stu.division}`;
    document.getElementById("modalClassRank").textContent = `Class Rank: #${stu.div_rank}`;
    document.getElementById("modalCollegeRank").textContent = `College Rank: #${stu.college_rank}`;
    document.getElementById("modalTotalScore").textContent = `${stu.total} / 100`;
    document.getElementById("modalPercentage").textContent = `${stu.percentage}%`;

    const grid = document.getElementById("modalSubjectsGrid");
    grid.innerHTML = "";

    stu.subjects.forEach(sub => {
      const val = stu.marks[sub];
      const subMeta = meta.subjects_meta[sub] || { short: sub.toUpperCase(), name: sub };
      const divInfo = divisions[stu.division];
      const classAvg = divInfo && divInfo.subject_stats[sub] ? divInfo.subject_stats[sub].avg : 0;

      const card = document.createElement("div");
      card.className = "subject-mark-card";
      card.innerHTML = `
        <div class="sub-code">${subMeta.short}</div>
        <div class="sub-score">${val !== null ? val : 'AB'}</div>
        <div class="sub-max">out of 20</div>
        <div style="font-size:0.75rem; color:var(--text-muted); margin-top:0.35rem;">Class Avg: ${classAvg}</div>
      `;
      grid.appendChild(card);
    });

    studentModal.classList.add("active");
  }

  closeModalBtn.addEventListener("click", () => {
    studentModal.classList.remove("active");
  });

  studentModal.addEventListener("click", (e) => {
    if (e.target === studentModal) studentModal.classList.remove("active");
  });

  // Global Student Search
  globalSearchInput.addEventListener("input", (e) => {
    const q = e.target.value.trim().toLowerCase();
    const resultsContainer = document.getElementById("globalSearchResults");
    if (!q) {
      resultsContainer.innerHTML = `<div style="text-align:center; padding: 2rem; color:var(--text-muted);">Type student name, roll number, or MIS ID above to search...</div>`;
      return;
    }

    const matches = students.filter(s => 
      String(s.roll_no).includes(q) || 
      s.mis_id.toLowerCase().includes(q) || 
      s.name.toLowerCase().includes(q)
    ).slice(0, 15);

    if (matches.length === 0) {
      resultsContainer.innerHTML = `<div style="text-align:center; padding: 2rem; color:var(--text-muted);">No students found matching "${q}".</div>`;
      return;
    }

    let html = `<div style="display:grid; grid-template-columns: repeat(auto-fill, minmax(280px, 1fr)); gap: 1rem;">`;
    matches.forEach(s => {
      html += `
        <div class="card" style="margin-bottom:0; padding:1.25rem; cursor:pointer;" onclick="window.viewStudent(${s.roll_no})">
          <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:0.5rem; font-size:0.85rem; font-weight:700;">
            <span>${s.division}</span>
            <span>Roll: ${s.roll_no}</span>
          </div>
          <h4 style="font-size:1rem; font-weight:700; margin-bottom:0.25rem;">${s.name}</h4>
          <p style="font-size:0.8rem; color:var(--text-muted); margin-bottom:0.75rem;">MIS: ${s.mis_id}</p>
          <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:0.5rem; font-size:0.85rem; font-weight:700;">
            <span style="color:#d97706;">Class Rank: #${s.div_rank}</span>
            <span style="color:#9333ea;">College Rank: #${s.college_rank}</span>
          </div>
          <div style="display:flex; align-items:center; justify-content:space-between; border-top:1px solid var(--border-color); padding-top:0.5rem;">
            <span style="font-size:0.95rem; font-weight:800; color:var(--brand-blue);">${s.total} / 100</span>
            <span style="font-size:0.95rem; font-weight:700;">${s.percentage}%</span>
          </div>
        </div>
      `;
    });
    html += `</div>`;
    resultsContainer.innerHTML = html;
  });

  window.viewStudent = function(roll) {
    const stu = students.find(s => s.roll_no === roll);
    if (stu) openStudentModal(stu);
  };

  // Charts
  let divisionChartInst = null;
  let gradeChartInst = null;

  function renderCharts() {
    if (typeof Chart === "undefined") return;

    const isDark = document.documentElement.getAttribute("data-theme") === "dark";
    const gridColor = isDark ? "rgba(255, 255, 255, 0.08)" : "rgba(0, 0, 0, 0.06)";
    const textColor = isDark ? "#cbd5e1" : "#475569";

    const divLabels = [];
    const divAvgs = [];
    for (let i = 1; i <= 13; i++) {
      const code = `FY${i}`;
      divLabels.push(code);
      divAvgs.push(divisions[code] ? divisions[code].avg_total : 0);
    }

    const ctx1 = document.getElementById("divisionChart");
    if (ctx1) {
      if (divisionChartInst) divisionChartInst.destroy();
      divisionChartInst = new Chart(ctx1, {
        type: "bar",
        data: {
          labels: divLabels,
          datasets: [{
            label: "Class Average (/100)",
            data: divAvgs,
            backgroundColor: divLabels.map((_, i) => i < 7 ? "rgba(37, 99, 235, 0.85)" : "rgba(16, 185, 129, 0.85)"),
            borderColor: divLabels.map((_, i) => i < 7 ? "#2563eb" : "#10b981"),
            borderWidth: 1.5,
            borderRadius: 6
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: (ctx) => `Average: ${ctx.raw} / 100 (${ctx.raw}%)`
              }
            }
          },
          scales: {
            x: {
              grid: { color: gridColor },
              ticks: { color: textColor, font: { weight: "600" } }
            },
            y: {
              min: 0,
              max: 100,
              grid: { color: gridColor },
              ticks: { color: textColor }
            }
          }
        }
      });
    }

    let totDistinction = 0, totFirst = 0, totSecond = 0, totPass = 0, totFail = 0;
    students.forEach(s => {
      if (s.percentage >= 75) totDistinction++;
      else if (s.percentage >= 60) totFirst++;
      else if (s.percentage >= 50) totSecond++;
      else if (s.percentage >= 40) totPass++;
      else totFail++;
    });

    const ctx2 = document.getElementById("gradeChart");
    if (ctx2) {
      if (gradeChartInst) gradeChartInst.destroy();
      gradeChartInst = new Chart(ctx2, {
        type: "doughnut",
        data: {
          labels: ["Distinction (75%+)", "First Class (60-74%)", "Second Class (50-59%)", "Pass (40-49%)", "Below 40%"],
          datasets: [{
            data: [totDistinction, totFirst, totSecond, totPass, totFail],
            backgroundColor: ["#10b981", "#3b82f6", "#f59e0b", "#8b5cf6", "#ef4444"],
            borderWidth: 2,
            borderColor: isDark ? "#131c31" : "#ffffff"
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: {
              position: "bottom",
              labels: { color: textColor, boxWidth: 14, font: { weight: "600" } }
            }
          }
        }
      });
    }
  }

  // Toppers View
  function renderToppersView() {
    const top3 = meta.college_toppers;
    if (top3 && top3.length >= 3) {
      document.getElementById("podium1Name").textContent = top3[0].name;
      document.getElementById("podium1Score").textContent = `${top3[0].total}/100 (${top3[0].pct}%)`;
      document.getElementById("podium1Meta").textContent = `${top3[0].div} | Roll ${top3[0].roll} | College Rank 1`;

      document.getElementById("podium2Name").textContent = top3[1].name;
      document.getElementById("podium2Score").textContent = `${top3[1].total}/100 (${top3[1].pct}%)`;
      document.getElementById("podium2Meta").textContent = `${top3[1].div} | Roll ${top3[1].roll} | College Rank 2`;

      document.getElementById("podium3Name").textContent = top3[2].name;
      document.getElementById("podium3Score").textContent = `${top3[2].total}/100 (${top3[2].pct}%)`;
      document.getElementById("podium3Meta").textContent = `${top3[2].div} | Roll ${top3[2].roll} | College Rank 3`;
    }

    const divToppersBody = document.getElementById("divToppersBody");
    let html = "";
    for (let i = 1; i <= 13; i++) {
      const code = `FY${i}`;
      const div = divisions[code];
      if (div) {
        html += `
          <tr>
            <td><strong>Division ${i} (${code})</strong></td>
            <td><strong>${div.topper_roll}</strong></td>
            <td><strong>${div.topper_name}</strong></td>
            <td style="text-align:center; font-weight:800; color:#d97706;">1</td>
            <td style="text-align:center; font-weight:800; color:var(--brand-blue);">${div.topper_score} / 100</td>
            <td style="text-align:center; font-weight:600;">${div.avg_total} / 100</td>
          </tr>
        `;
      }
    }
    divToppersBody.innerHTML = html;
  }

  // Division Comparison Tool
  function renderComparisonTool() {
    const div1Select = document.getElementById("compareDiv1Select");
    const div2Select = document.getElementById("compareDiv2Select");

    div1Select.value = state.compareDiv1;
    div2Select.value = state.compareDiv2;

    updateComparisonView();

    div1Select.onchange = (e) => {
      state.compareDiv1 = e.target.value;
      updateComparisonView();
    };
    div2Select.onchange = (e) => {
      state.compareDiv2 = e.target.value;
      updateComparisonView();
    };
  }

  function updateComparisonView() {
    const d1 = divisions[state.compareDiv1];
    const d2 = divisions[state.compareDiv2];
    const c1 = document.getElementById("compareCard1");
    const c2 = document.getElementById("compareCard2");

    if (!d1 || !d2) return;

    function buildCard(d) {
      return `
        <h3 style="font-size:1.25rem; font-weight:800; color:var(--brand-blue); margin-bottom:1rem;">Division ${d.div_num} (${d.division})</h3>
        <p style="color:var(--text-secondary); margin-bottom:1.25rem;">Subjects: ${d.subjects.map(s => s.toUpperCase()).join(", ")}</p>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:0.75rem; margin-bottom:1.5rem;">
          <div class="stat-card" style="padding:0.75rem;">
            <div class="stat-info">
              <h3>Total Students</h3>
              <div class="stat-value" style="font-size:1.4rem;">${d.total_students}</div>
              <div class="stat-sub">Roll ${d.min_roll} – ${d.max_roll}</div>
            </div>
          </div>
          <div class="stat-card" style="padding:0.75rem;">
            <div class="stat-info">
              <h3>Class Average</h3>
              <div class="stat-value" style="font-size:1.4rem; color:var(--success);">${d.avg_total}</div>
              <div class="stat-sub">out of 100</div>
            </div>
          </div>
          <div class="stat-card" style="padding:0.75rem;">
            <div class="stat-info">
              <h3>Highest Score</h3>
              <div class="stat-value" style="font-size:1.4rem; color:var(--brand-blue);">${d.max_total}</div>
              <div class="stat-sub">${d.topper_name}</div>
            </div>
          </div>
          <div class="stat-card" style="padding:0.75rem;">
            <div class="stat-info">
              <h3>Distinction (75%+)</h3>
              <div class="stat-value" style="font-size:1.4rem;">${d.grade_dist.distinction}</div>
              <div class="stat-sub">${Math.round((d.grade_dist.distinction / d.total_students)*100)}% of class</div>
            </div>
          </div>
        </div>
      `;
    }

    c1.innerHTML = buildCard(d1);
    c2.innerHTML = buildCard(d2);
  }

  // Export to CSV
  exportCsvBtn.addEventListener("click", () => {
    const filtered = getFilteredStudents();
    if (filtered.length === 0) {
      alert("No data to export!");
      return;
    }

    let csv = "Roll No,Class Rank,College Rank,Student Name,MIS ID,Division,Total (/100),Percentage\n";
    filtered.forEach(s => {
      csv += `"${s.roll_no}",${s.div_rank},${s.college_rank},"${s.name.replace(/"/g, '""')}","${s.mis_id}","${s.division}",${s.total},${s.percentage}\n`;
    });

    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `PICT_Marks_with_Ranks_${state.selectedDivision}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  });

  document.getElementById("printReportBtn").addEventListener("click", () => {
    window.print();
  });

  // Initialize
  renderKPICards();
  renderDivisionSelector();
  renderCharts();
  renderTable();
});
