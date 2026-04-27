const STORAGE_KEY = "requirement-platform-state";
const IMPORT_HISTORY_KEY = "requirement-import-history";
const MAX_IMPORT_HISTORY = 20;

const STATUS_OPTIONS = ["待评审", "设计中", "开发中", "已阻塞", "已上线", "已归档"];
const TYPE_OPTIONS = ["业务需求", "产品需求", "项目需求", "流程优化", "系统改造"];
const PRIORITY_OPTIONS = ["P0", "P1", "P2", "P3"];
const REVIEW_STATUS_OPTIONS = ["待评审", "评审中", "已评审", "开发中", "待验证", "已上线"];

const seedRequirements = [
  {
    id: "REQ-2026-001",
    title: "直营门店补货预测优化",
    type: "业务需求",
    priority: "P0",
    status: "开发中",
    department: "供应链中心",
    owner: "赵敏",
    submittedAt: "2026-04-04",
    plannedRelease: "2026-05-10",
    background: "门店补货主要依赖人工经验，预测偏差导致库存积压与断货并存。",
    description: "引入门店历史销售、节假日和活动因子，输出补货建议并支持审批回写。",
    reviewNotes: "已完成业务评审与方案评审，需要与 ERP 团队同步接口窗口。",
    auditTrail: [
      "2026-04-04 需求创建，提出部门：供应链中心",
      "2026-04-07 完成需求澄清，确认补货建议计算口径",
      "2026-04-12 进入开发中"
    ]
  },
  {
    id: "REQ-2026-002",
    title: "会员积分过期提醒改版",
    type: "产品需求",
    priority: "P1",
    status: "待评审",
    department: "用户增长部",
    owner: "李川",
    submittedAt: "2026-04-08",
    plannedRelease: "2026-05-03",
    background: "积分临期提醒触达率低，短信与站内信配置分散。",
    description: "统一配置站内信、短信和小程序提醒模板，支持触达效果追踪。",
    reviewNotes: "待法务确认短信文案合规要求。",
    auditTrail: [
      "2026-04-08 需求创建，提出部门：用户增长部",
      "2026-04-10 补充短信成本测算"
    ]
  },
  {
    id: "REQ-2026-003",
    title: "财务共享报销流转提速",
    type: "流程优化",
    priority: "P1",
    status: "设计中",
    department: "财务共享中心",
    owner: "王楠",
    submittedAt: "2026-04-01",
    plannedRelease: "2026-04-30",
    background: "报销单审批链长，节点超时缺少自动升级机制。",
    description: "新增审批超时提醒、自动升级和瓶颈看板，缩短平均流转时间。",
    reviewNotes: "设计方案已通过，需要补充组织架构变更同步方案。",
    auditTrail: [
      "2026-04-01 需求创建，提出部门：财务共享中心",
      "2026-04-05 完成需求评审",
      "2026-04-09 进入设计中"
    ]
  },
  {
    id: "REQ-2026-004",
    title: "海外站商品主数据打通",
    type: "项目需求",
    priority: "P0",
    status: "已阻塞",
    department: "国际业务部",
    owner: "陈昊",
    submittedAt: "2026-03-28",
    plannedRelease: "2026-05-20",
    background: "海外站商品属性口径与国内站不一致，导致商品上架流程重复维护。",
    description: "建立商品主数据映射关系，打通主站、OMS 与海外站后台字段同步。",
    reviewNotes: "阻塞于主数据标准尚未统一，需要 PMO 牵头决策。",
    auditTrail: [
      "2026-03-28 需求创建，提出部门：国际业务部",
      "2026-04-03 架构评审通过",
      "2026-04-11 标记为已阻塞"
    ]
  },
  {
    id: "REQ-2026-005",
    title: "客服工单质检抽样自动化",
    type: "系统改造",
    priority: "P2",
    status: "已上线",
    department: "客户服务部",
    owner: "周洁",
    submittedAt: "2026-03-15",
    plannedRelease: "2026-04-15",
    background: "客服质检人工抽样效率低，缺少统一质检规则沉淀。",
    description: "基于工单标签自动抽样，生成质检任务并输出问题分类看板。",
    reviewNotes: "已上线，等待第一个完整月效果复盘。",
    auditTrail: [
      "2026-03-15 需求创建，提出部门：客户服务部",
      "2026-03-22 完成评审并排期",
      "2026-04-15 正式上线"
    ]
  }
];

const state = {
  requirements: loadState(),
  activeSection: "intake",
  filters: {
    search: "",
    reviewStatus: "all"
  },
  requirementPage: 1,
  importHistorySearch: "",
  editingId: null,
  generatedSpecs: [],
  importedFiles: [],
  importHistory: loadImportHistory(),
  detailId: null
};

const elements = {
  requirementsBoard: document.getElementById("requirementsBoard"),
  requirementsBoardSummary: document.getElementById("requirementsBoardSummary"),
  requirementsEmptyState: document.getElementById("requirementsEmptyState"),
  requirementsPagination: document.getElementById("requirementsPagination"),
  intakeFileInput: document.getElementById("intakeFileInput"),
  intakeImportSummary: document.getElementById("intakeImportSummary"),
  importHistorySearchInput: document.getElementById("importHistorySearchInput"),
  importHistorySummary: document.getElementById("importHistorySummary"),
  importHistoryList: document.getElementById("importHistoryList"),
  intakeInput: document.getElementById("intakeInput"),
  intakeResult: document.getElementById("intakeResult"),
  listSearchInput: document.getElementById("listSearchInput"),
  listReviewStatusFilter: document.getElementById("listReviewStatusFilter"),
  requirementDialog: document.getElementById("requirementDialog"),
  requirementForm: document.getElementById("requirementForm"),
  dialogTitle: document.getElementById("dialogTitle"),
  detailDialog: document.getElementById("detailDialog"),
  detailTitle: document.getElementById("detailTitle"),
  detailSubtitle: document.getElementById("detailSubtitle"),
  detailContent: document.getElementById("detailContent"),
  requirementCardTemplate: document.getElementById("requirementCardTemplate")
};

initialize();

function initialize() {
  hydrateSelectOptions();
  bindEvents();
  render();
}

function loadState() {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (!stored) {
    const seeded = cloneRequirements(seedRequirements);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(seeded));
    return seeded;
  }

  try {
    return cloneRequirements(JSON.parse(stored));
  } catch {
    const seeded = cloneRequirements(seedRequirements);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(seeded));
    return seeded;
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.requirements));
}

function loadImportHistory() {
  const stored = localStorage.getItem(IMPORT_HISTORY_KEY);
  if (!stored) return [];

  try {
    const parsed = JSON.parse(stored);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveImportHistory() {
  localStorage.setItem(IMPORT_HISTORY_KEY, JSON.stringify(state.importHistory));
}

function hydrateSelectOptions() {
  if (elements.listReviewStatusFilter) {
    fillSelect(elements.listReviewStatusFilter, ["all", ...REVIEW_STATUS_OPTIONS], "全部评审状态");
    elements.listReviewStatusFilter.value = state.filters.reviewStatus;
  }

  fillSelect(elements.requirementForm.elements.type, TYPE_OPTIONS);
  fillSelect(elements.requirementForm.elements.priority, PRIORITY_OPTIONS);
  fillSelect(elements.requirementForm.elements.reviewStatus, REVIEW_STATUS_OPTIONS);
}

function fillSelect(select, options, placeholder) {
  select.innerHTML = "";
  options.forEach((value, index) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = placeholder && index === 0 ? placeholder : value;
    select.appendChild(option);
  });
}

function bindEvents() {
  document.querySelectorAll(".nav-item").forEach((button) => {
    button.addEventListener("click", () => switchSection(button.dataset.section));
  });

  const newRequirementButton = document.getElementById("newRequirementBtn");
  if (newRequirementButton) newRequirementButton.addEventListener("click", () => openForm());

  document.getElementById("closeDialogBtn")?.addEventListener("click", closeForm);
  document.getElementById("cancelDialogBtn")?.addEventListener("click", closeForm);
  document.getElementById("closeDetailDialogBtn")?.addEventListener("click", closeDetailDialog);
  document.getElementById("detailCloseBtn")?.addEventListener("click", closeDetailDialog);
  document.getElementById("detailEditBtn")?.addEventListener("click", handleDetailEdit);
  document.getElementById("seedDataBtn")?.addEventListener("click", resetSeedData);
  document.getElementById("analyzeRequirementBtn")?.addEventListener("click", analyzeIntake);
  document.getElementById("loadDemoPromptBtn")?.addEventListener("click", loadDemoPrompt);
  document.getElementById("clearIntakeBtn")?.addEventListener("click", clearIntakeInput);
  document.getElementById("clearImportHistoryBtn")?.addEventListener("click", clearImportHistory);

  if (elements.intakeFileInput) {
    elements.intakeFileInput.addEventListener("change", handleFileImport);
  }
  if (elements.importHistorySearchInput) {
    elements.importHistorySearchInput.addEventListener("input", (event) => {
      state.importHistorySearch = event.target.value.trim();
      renderImportHistory();
    });
  }

  if (elements.listSearchInput) {
    elements.listSearchInput.addEventListener("input", (event) => {
      state.filters.search = event.target.value.trim();
      state.requirementPage = 1;
      renderRequirementsBoard();
    });
  }
  elements.listReviewStatusFilter?.addEventListener("change", (event) => {
    state.filters.reviewStatus = event.target.value;
    state.requirementPage = 1;
    renderRequirementsBoard();
  });

  elements.requirementForm?.addEventListener("submit", handleFormSubmit);
}

function switchSection(sectionId) {
  state.activeSection = sectionId;
  document.querySelectorAll(".nav-item").forEach((button) => {
    button.classList.toggle("active", button.dataset.section === sectionId);
  });
  document.querySelectorAll(".panel").forEach((panel) => {
    panel.classList.toggle("active", panel.id === sectionId);
  });
}

function render() {
  renderIntakeResult();
  renderImportHistory();
  renderRequirementsBoard();
  switchSection(state.activeSection);
}

function renderHeroMetrics() {
  const total = state.requirements.length;
  const online = state.requirements.filter((item) => item.status === "已上线").length;
  const pending = state.requirements.filter((item) => item.status === "待评审").length;
  const overdue = state.requirements.filter((item) => item.status !== "已上线" && item.status !== "已归档" && item.plannedRelease < today()).length;

  elements.heroMetrics.innerHTML = [
    metric("需求总量", total, "全平台纳管需求"),
    metric("待评审", pending, "待进入决策流程"),
    metric("已上线", online, "已完成并落地"),
    metric("延期风险", overdue, "超计划未完成")
  ].join("");
}

function metric(label, value, description) {
  return `<div class="metric-card"><span>${label}</span><strong>${value}</strong><span>${description}</span></div>`;
}

function renderStatusStats() {
  elements.statusStats.innerHTML = STATUS_OPTIONS.map((status) => {
    const count = state.requirements.filter((item) => item.status === status).length;
    return `<div class="stat-card"><span>${status}</span><strong>${count}</strong></div>`;
  }).join("");
}

function renderPriorityList() {
  const items = [...state.requirements]
    .sort((a, b) => priorityScore(a.priority) - priorityScore(b.priority))
    .slice(0, 5);

  elements.priorityList.innerHTML = items.map((item) => stackItem(item, `${item.department} · ${item.owner}`)).join("");
  bindInlineActions();
}

function renderRiskList() {
  const items = state.requirements.filter((item) => {
    const overdue = item.plannedRelease < today() && !["已上线", "已归档"].includes(item.status);
    return overdue || item.status === "待评审" || item.status === "已阻塞";
  });

  elements.riskList.innerHTML = items.length
    ? items.map((item) => {
        const riskLabel = item.status === "已阻塞" ? "存在阻塞依赖" : item.status === "待评审" ? "尚未形成评审结论" : "计划时间已超期";
        return stackItem(item, riskLabel);
      }).join("")
    : `<div class="stack-item"><h4>当前无明显风险</h4><p class="item-meta">需求推进状态正常。</p></div>`;
  bindInlineActions();
}

function stackItem(item, meta) {
  return `
    <div class="stack-item">
      <h4>${item.title}</h4>
      <p class="item-meta">${item.id} · ${item.priority} · ${meta}</p>
      <button class="text-button" data-action="view" data-id="${item.id}">查看详情</button>
    </div>
  `;
}

function renderKanban() {
  elements.kanbanBoard.innerHTML = STATUS_OPTIONS.map((status) => {
    const items = state.requirements.filter((item) => item.status === status);
    return `
      <section class="kanban-column" data-status="${status}">
        <div class="column-header">
          <strong>${status}</strong>
          <span class="hint">${items.length}</span>
        </div>
        <div class="column-body" data-status="${status}">
          ${items.map((item) => renderRequirementCard(item)).join("")}
        </div>
      </section>
    `;
  }).join("");

  bindInlineActions();
  bindDragAndDrop();
}

function renderRequirementCard(item, options = {}) {
  const actionArea = options.includeDelete
    ? `
      <div class="card-topline">
        <button class="text-button" data-action="view" data-id="${item.id}">查看详情</button>
        <button class="text-button" data-action="edit" data-id="${item.id}">编辑</button>
        <button class="danger-button" data-action="delete" data-id="${item.id}">删除</button>
      </div>
    `
    : `<button class="text-button" data-action="view" data-id="${item.id}">查看详情</button>`;

  return `
    <article class="requirement-card" draggable="true" data-id="${item.id}">
      <div class="card-topline">
        <span class="badge badge-type">${item.type}</span>
        <span class="badge badge-priority" data-priority="${item.priority}">${item.priority}</span>
        <span class="status-badge" data-status="${item.status}">${item.status}</span>
      </div>
      <h4>${item.title}</h4>
      <p class="card-meta">${item.id} · ${item.department} · ${item.owner}</p>
      <p class="card-meta">计划上线 ${item.plannedRelease}</p>
      ${actionArea}
    </article>
  `;
}

function renderRequirementsBoard() {
  const rows = getFilteredRequirements();
  const pageSize = 10;
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
  state.requirementPage = Math.min(state.requirementPage, totalPages);
  const startIndex = (state.requirementPage - 1) * pageSize;
  const pageRows = rows.slice(startIndex, startIndex + pageSize);
  elements.requirementsBoardSummary.innerHTML = [
    `<span class="badge badge-type">共 ${rows.length} 条需求单</span>`,
    state.filters.reviewStatus !== "all" ? `<span class="badge badge-type">评审状态 ${state.filters.reviewStatus}</span>` : ""
  ].join("");

  elements.requirementsEmptyState.classList.toggle("hidden", rows.length !== 0);
  elements.requirementsBoard.innerHTML = pageRows
    .map((item) => renderRequirementTableRow(item))
    .join("");
  renderRequirementsPagination(rows.length, totalPages);
  bindInlineActions();
}

function renderRequirementsPagination(totalCount, totalPages) {
  if (!elements.requirementsPagination) return;
  const shouldShow = totalCount > 10;
  elements.requirementsPagination.classList.toggle("hidden", !shouldShow);
  if (!shouldShow) {
    elements.requirementsPagination.innerHTML = "";
    return;
  }

  elements.requirementsPagination.innerHTML = `
    <button class="ghost-button" type="button" data-page-action="prev" ${state.requirementPage <= 1 ? "disabled" : ""}>上一页</button>
    <span class="pagination-info">第 ${state.requirementPage} / ${totalPages} 页</span>
    <button class="ghost-button" type="button" data-page-action="next" ${state.requirementPage >= totalPages ? "disabled" : ""}>下一页</button>
  `;

  elements.requirementsPagination.querySelectorAll("[data-page-action]").forEach((button) => {
    button.onclick = () => {
      const nextPage = button.dataset.pageAction === "prev" ? state.requirementPage - 1 : state.requirementPage + 1;
      if (nextPage < 1 || nextPage > totalPages) return;
      state.requirementPage = nextPage;
      renderRequirementsBoard();
    };
  });
}

function getFilteredRequirements() {
  return state.requirements.filter((item) => {
    const hitSearch = !state.filters.search || [
      item.title,
      item.product,
      item.sourceProject,
      item.requester
    ].some((field) => String(field || "").includes(state.filters.search));
    const hitReviewStatus = state.filters.reviewStatus === "all" || item.reviewStatus === state.filters.reviewStatus;
    return hitSearch && hitReviewStatus;
  });
}

function renderRequirementTableRow(item) {
  const reviewStatusClass = ["已评审", "开发中", "待验证", "已上线"].includes(item.reviewStatus) ? "is-created" : "is-pending";
  return `
    <tr>
      <td>${item.product}</td>
      <td>
        <div class="generated-spec-title-cell">
          <strong>${item.title}</strong>
          <p title="${item.id}">${item.id}</p>
        </div>
      </td>
      <td>${item.sourceProject}</td>
      <td>${item.requester}</td>
      <td>${item.submittedAt}</td>
      <td>${item.deliveryDirector}</td>
      <td>${item.devManager}</td>
      <td>${item.productManager}</td>
      <td><span class="generated-status-chip ${reviewStatusClass}">${item.reviewStatus}</span></td>
      <td>
        <div class="generated-spec-row-actions">
          <button class="ghost-button" data-action="view" data-id="${item.id}">查看</button>
          <button class="ghost-button" data-action="edit" data-id="${item.id}">编辑</button>
          <button class="ghost-button" data-action="delete" data-id="${item.id}">删除</button>
        </div>
      </td>
    </tr>
  `;
}

function renderReviews() {
  hydrateDynamicFilters();
  const reviewItems = state.requirements.filter((item) => {
    const inQueue = ["待评审", "设计中", "已阻塞"].includes(item.status);
    const hitSearch = !state.reviewFilters.queueSearch || [item.title, item.department, item.owner].some((field) => field.includes(state.reviewFilters.queueSearch));
    const hitStatus = state.reviewFilters.queueStatus === "all" || item.status === state.reviewFilters.queueStatus;
    const hitType = state.reviewFilters.queueType === "all" || item.type === state.reviewFilters.queueType;
    return inQueue && hitSearch && hitStatus && hitType;
  });
  elements.reviewQueue.innerHTML = reviewItems.map((item) => `
    <div class="stack-item">
      <h4>${item.title}</h4>
      <p class="item-meta">${item.status} · 评审备注：${item.reviewNotes || "暂无"}</p>
      <button class="text-button" data-action="edit" data-id="${item.id}">补充评审</button>
    </div>
  `).join("");

  elements.reviewHistory.innerHTML = [...state.requirements]
    .filter((item) => {
      const hitSearch = !state.reviewFilters.historySearch || [item.title, item.reviewNotes || ""].some((field) => field.includes(state.reviewFilters.historySearch));
      const hitOwner = state.reviewFilters.historyOwner === "all" || item.owner === state.reviewFilters.historyOwner;
      const hitPriority = state.reviewFilters.historyPriority === "all" || item.priority === state.reviewFilters.historyPriority;
      return hitSearch && hitOwner && hitPriority;
    })
    .sort((a, b) => b.submittedAt.localeCompare(a.submittedAt))
    .map((item) => `
      <div class="timeline-item">
        <h4>${item.title}</h4>
        <p>${item.reviewNotes || "暂无评审结论"}</p>
        <p class="timeline-meta">${item.id} · ${item.owner} · 计划上线 ${item.plannedRelease}</p>
      </div>
    `).join("");

  bindInlineActions();
}

function renderDelivery() {
  elements.deliveryTimeline.innerHTML = [...state.requirements]
    .filter((item) => {
      const hitSearch = !state.deliveryFilters.timelineSearch || [item.title, item.department, item.owner].some((field) => field.includes(state.deliveryFilters.timelineSearch));
      const hitStatus = state.deliveryFilters.timelineStatus === "all" || item.status === state.deliveryFilters.timelineStatus;
      const hitOwner = state.deliveryFilters.timelineOwner === "all" || item.owner === state.deliveryFilters.timelineOwner;
      return hitSearch && hitStatus && hitOwner;
    })
    .sort((a, b) => a.plannedRelease.localeCompare(b.plannedRelease))
    .map((item) => `
      <div class="timeline-item">
        <h4>${item.plannedRelease} · ${item.title}</h4>
        <p>${item.status} · ${item.owner} 负责推进</p>
        <p class="timeline-meta">${item.department} · ${item.type}</p>
      </div>
    `).join("");

  const audits = [...state.requirements]
    .filter((item) => {
      const hitSearch = !state.deliveryFilters.auditSearch || [item.title, ...(item.auditTrail || [])].some((field) => field.includes(state.deliveryFilters.auditSearch));
      const hitStatus = state.deliveryFilters.auditStatus === "all" || item.status === state.deliveryFilters.auditStatus;
      const hitType = state.deliveryFilters.auditType === "all" || item.type === state.deliveryFilters.auditType;
      return hitSearch && hitStatus && hitType;
    })
    .flatMap((item) => item.auditTrail.map((log) => ({ id: item.id, title: item.title, log })))
    .sort((a, b) => b.log.localeCompare(a.log))
    .slice(0, 12);

  elements.auditLog.innerHTML = audits.map((entry) => `
    <div class="timeline-item">
      <h4>${entry.title}</h4>
      <p>${entry.log}</p>
      <p class="timeline-meta">${entry.id}</p>
    </div>
  `).join("");
}

function renderIntakeResult() {
  elements.intakeImportSummary.innerHTML = state.importedFiles.map((file) => `<span class="badge source-badge">${file}</span>`).join("");

  const specs = state.generatedSpecs;
  if (!specs.length) {
    elements.intakeResult.innerHTML = `
      <div class="empty-state">
        <h4>尚未生成需求说明</h4>
        <p class="item-meta">输入自然语言或导入文档后，这里会生成标准化标题、背景、目标、范围、验收与风险建议，并支持批量生成需求单。</p>
      </div>
    `;
    return;
  }

  const createdCount = specs.filter((spec) => spec.createdRequirementId).length;
  const pendingCount = specs.length - createdCount;
  elements.intakeResult.innerHTML = `
    <div class="spec-batch-header">
      <div>
        <h4>已识别 ${specs.length} 条候选需求</h4>
        <p class="item-meta">已生成 ${createdCount} 条，待生成 ${pendingCount} 条${state.importedFiles.length ? ` · 来源：${state.importedFiles.join("、")}` : " · 来源：手动输入"}</p>
      </div>
      <div class="spec-actions">
        <button class="primary-button" data-action="create-all-specs" ${pendingCount === 0 ? "disabled" : ""}>批量生成需求单</button>
        <button class="ghost-button" data-action="copy-all-specs">复制全部结构化说明</button>
        <button class="ghost-button" data-action="cancel-specs">清空结果</button>
      </div>
    </div>

    <div class="generated-spec-list">
      ${specs.map((spec, index) => renderGeneratedSpecCard(spec, index)).join("")}
    </div>
  `;

  bindSpecActions();
}

function renderImportHistory() {
  const rows = getFilteredImportHistory();
  elements.importHistorySummary.innerHTML = `<span class="badge source-badge">共 ${rows.length} 条</span>`;

  if (!rows.length) {
    elements.importHistoryList.innerHTML = `
      <div class="history-empty">
        <h4>暂无匹配的导入记录</h4>
        <p class="item-meta">导入成功后会自动保存到这里，可按文件名快速查找。</p>
      </div>
    `;
    bindImportHistoryActions();
    return;
  }

  elements.importHistoryList.innerHTML = rows.map((item) => `
    <article class="history-card">
      <div class="history-card-header">
        <div>
          <h5>${item.files.join("、")}</h5>
          <p class="item-meta">${item.createdAt} · ${item.specCount} 条需求 · ${item.files.length} 个文件</p>
        </div>
        <div class="history-meta">
          <span class="badge source-badge">最近使用 ${item.lastUsedAt || item.createdAt}</span>
        </div>
      </div>
      <div class="history-meta">
        ${item.files.map((file) => `<span class="badge source-badge">${file}</span>`).join("")}
      </div>
      <div class="history-card-actions">
        <button class="primary-button" type="button" data-action="restore-import-history" data-history-id="${item.id}">恢复到当前结果</button>
        <button class="ghost-button" type="button" data-action="copy-import-history" data-history-id="${item.id}">复制说明</button>
        <button class="ghost-button" type="button" data-action="delete-import-history" data-history-id="${item.id}">删除记录</button>
      </div>
    </article>
  `).join("");

  bindImportHistoryActions();
}

function getFilteredImportHistory() {
  return [...state.importHistory]
    .filter((item) => {
      if (!state.importHistorySearch) return true;
      const query = state.importHistorySearch;
      return [item.files.join(" "), item.createdAt, ...item.specs.map((spec) => spec.title || "")]
        .some((field) => String(field).includes(query));
    })
    .sort((a, b) => String(b.lastUsedAt || b.createdAt).localeCompare(String(a.lastUsedAt || a.createdAt)));
}

function renderGeneratedSpecCard(spec, index) {
  const sourceLabel = [spec.sourceName || "手动输入", spec.sourceBlockLabel].filter(Boolean).join(" / ");
  const generationLabel = spec.createdRequirementId ? `已生成 ${spec.createdRequirementId}` : "待生成";

  return `
    <article class="generated-spec-card-lite">
      <div class="generated-spec-card-top">
        <div>
          <h4>${index + 1}. ${spec.title}</h4>
          <p class="item-meta">${spec.summary}</p>
        </div>
        <div class="spec-highlight">
          <span class="badge badge-type">${spec.type}</span>
          <span class="badge badge-priority" data-priority="${spec.priority}">${spec.priority}</span>
          <span class="status-badge" data-status="${spec.status}">${spec.status}</span>
        </div>
      </div>
      <div class="generated-spec-meta">
        <span class="badge source-badge">${sourceLabel}</span>
        <span class="generated-status-chip ${spec.createdRequirementId ? "is-created" : "is-pending"}">${generationLabel}</span>
        <span class="badge badge-type">提出部门 ${spec.department}</span>
        <span class="badge badge-type">负责人 ${spec.owner}</span>
        <span class="badge badge-type">计划上线 ${spec.plannedRelease}</span>
      </div>
      <div class="generated-spec-card-note">
        <p><strong>背景：</strong>${spec.background}</p>
        <p><strong>目标：</strong>${spec.goal}</p>
      </div>
      <div class="generated-spec-row-actions">
        <button class="primary-button" data-action="create-single-spec" data-spec-id="${spec.draftId}" ${spec.createdRequirementId ? "disabled" : ""}>${spec.createdRequirementId ? "已生成" : "生成需求单"}</button>
        <button class="ghost-button" data-action="edit-single-spec" data-spec-id="${spec.draftId}">编辑后生成</button>
        <button class="ghost-button" data-action="copy-single-spec" data-spec-id="${spec.draftId}">复制说明</button>
      </div>
    </article>
  `;
}

function openForm(id, draftSpec = null) {
  state.editingId = id || null;
  const current = state.requirements.find((item) => item.id === id);
  elements.dialogTitle.textContent = current ? `编辑需求 ${current.id}` : "新建需求";
  elements.requirementForm.reset();

  if (current) {
    Object.entries(current).forEach(([key, value]) => {
      const field = elements.requirementForm.elements[key];
      if (field) {
        field.value = value;
      }
    });
  } else {
    elements.requirementForm.elements.product.value = "待补充";
    elements.requirementForm.elements.sourceProject.value = "待补充";
    elements.requirementForm.elements.requester.value = "待补充";
    elements.requirementForm.elements.deliveryDirector.value = "待分配";
    elements.requirementForm.elements.devManager.value = "待分配";
    elements.requirementForm.elements.productManager.value = "待分配";
    elements.requirementForm.elements.reviewStatus.value = "待评审";
    elements.requirementForm.elements.submittedAt.value = today();
    elements.requirementForm.elements.priority.value = "P2";
    elements.requirementForm.elements.type.value = "业务需求";

    if (draftSpec) {
      populateFormFromSpec(draftSpec);
    }
  }

  elements.requirementDialog.showModal();
}

function closeForm() {
  elements.requirementDialog.close();
}

function openDetailDialog(id) {
  const item = findRequirement(id);
  if (!item) return;

  state.detailId = id;
  elements.detailTitle.textContent = item.title;
  elements.detailSubtitle.textContent = `${item.id} · ${item.department} · ${item.owner}`;
  elements.detailContent.innerHTML = buildDetailContent(item);
  elements.detailDialog.showModal();
}

function closeDetailDialog() {
  state.detailId = null;
  elements.detailDialog.close();
}

function handleDetailEdit() {
  if (!state.detailId) return;
  closeDetailDialog();
  openForm(state.detailId);
}

function handleFormSubmit(event) {
  event.preventDefault();
  const formData = new FormData(elements.requirementForm);
  const payload = Object.fromEntries(formData.entries());
  const timestamp = getTimestamp();

  if (state.editingId) {
    state.requirements = state.requirements.map((item) => {
      if (item.id !== state.editingId) {
        return item;
      }
      const statusLabel = payload.status || item.status || "待评审";
      const auditTrail = [...item.auditTrail, `${timestamp} 更新需求信息，当前状态：${statusLabel}`];
      return enrichRequirement({ ...item, ...payload, auditTrail });
    });
  } else {
    const nextId = buildRequirementId();
    state.requirements.unshift(createRequirementRecord(payload, nextId, timestamp));
  }

  saveState();
  closeForm();
  render();
}

function buildRequirementId(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const seconds = String(date.getSeconds()).padStart(2, "0");
  return `REQ-${year}${month}${day}${hours}${minutes}${seconds}`;
}

function resetSeedData() {
  state.requirements = cloneRequirements(seedRequirements);
  state.generatedSpecs = [];
  state.importedFiles = [];
  saveState();
  render();
}

function analyzeIntake() {
  const rawText = elements.intakeInput.value.trim();
  if (!rawText) {
    alert("请先输入原始需求描述。");
    return;
  }

  state.generatedSpecs = buildSpecsFromTextSource(rawText, "手动输入");
  state.importedFiles = [];
  renderIntakeResult();
}

function loadDemoPrompt() {
  elements.intakeInput.value = [
    "需求1：仓库主管反馈现在调拨申请依赖 Excel 和微信群，经常漏处理。希望做一个内部调拨申请功能，门店可发起调拨，仓库审核后自动同步库存，异常时要提醒相关人员。最好支持移动端，要求 6 月底前上线。",
    "",
    "需求2：财务团队希望把供应商对账从邮件附件改成系统在线确认，支持差异标记、审批和回写 ERP，并沉淀月度对账统计看板。"
  ].join("\n");
}

function clearIntakeInput() {
  elements.intakeInput.value = "";
  elements.intakeFileInput.value = "";
  state.generatedSpecs = [];
  state.importedFiles = [];
  renderIntakeResult();
}

async function handleFileImport(event) {
  const files = [...event.target.files];
  if (!files.length) return;

  const importedSpecs = [];
  const importedNames = [];
  const unsupportedNames = [];

  for (const file of files) {
    const extension = getFileExtension(file.name);
    if (!["txt", "md", "markdown", "csv", "json", "docx", "xlsx", "xls"].includes(extension)) {
      unsupportedNames.push(file.name);
      continue;
    }

    const specs = await parseImportedFile(file, extension);
    if (specs.length) {
      importedSpecs.push(...specs);
      importedNames.push(file.name);
    }
  }

  elements.intakeFileInput.value = "";

  if (unsupportedNames.length) {
    alert(`以下文件暂不支持直接导入：${unsupportedNames.join("、")}。当前支持 txt、md、csv、json、docx、xlsx、xls。`);
  }

  if (!importedSpecs.length) {
    if (!unsupportedNames.length) {
      alert("未识别到可转换的需求内容，请检查文档格式或内容结构。");
    }
    return;
  }

  state.generatedSpecs = importedSpecs;
  state.importedFiles = importedNames;
  addImportHistoryRecord(importedNames, importedSpecs);
  renderIntakeResult();
  renderImportHistory();
}

function bindSpecActions() {
  document.querySelectorAll("[data-action='create-all-specs']").forEach((button) => {
    button.onclick = () => createRequirementsFromSpecs(state.generatedSpecs.filter((spec) => !spec.createdRequirementId));
  });

  document.querySelectorAll("[data-action='copy-all-specs']").forEach((button) => {
    button.onclick = () => copyTextToClipboard(formatAllSpecsText(state.generatedSpecs), "全部结构化需求说明已复制。");
  });

  document.querySelectorAll("[data-action='cancel-specs']").forEach((button) => {
    button.onclick = () => {
      state.generatedSpecs = [];
      state.importedFiles = [];
      renderIntakeResult();
    };
  });

  document.querySelectorAll("[data-action='create-single-spec']").forEach((button) => {
    button.onclick = () => createSingleRequirementFromDraft(button.dataset.specId);
  });

  document.querySelectorAll("[data-action='edit-single-spec']").forEach((button) => {
    button.onclick = () => {
      const draft = getGeneratedSpec(button.dataset.specId);
      if (!draft) return;
      openForm(null, draft);
    };
  });

  document.querySelectorAll("[data-action='copy-single-spec']").forEach((button) => {
    button.onclick = () => {
      const draft = getGeneratedSpec(button.dataset.specId);
      if (!draft) return;
      copyTextToClipboard(formatSpecText(draft), "单条结构化需求说明已复制。");
    };
  });
}

function bindImportHistoryActions() {
  document.querySelectorAll("[data-action='restore-import-history']").forEach((button) => {
    button.onclick = () => restoreImportHistory(button.dataset.historyId);
  });

  document.querySelectorAll("[data-action='copy-import-history']").forEach((button) => {
    button.onclick = () => {
      const history = state.importHistory.find((item) => item.id === button.dataset.historyId);
      if (!history) return;
      copyTextToClipboard(formatAllSpecsText(history.specs), "历史记录中的结构化说明已复制。");
    };
  });

  document.querySelectorAll("[data-action='delete-import-history']").forEach((button) => {
    button.onclick = () => deleteImportHistory(button.dataset.historyId);
  });
}

function bindInlineActions() {
  document.querySelectorAll("[data-action='view']").forEach((button) => {
    button.onclick = () => {
      const item = findRequirement(button.dataset.id);
      if (!item) return;
      openDetailDialog(item.id);
    };
  });

  document.querySelectorAll("[data-action='edit']").forEach((button) => {
    button.onclick = () => openForm(button.dataset.id);
  });

  document.querySelectorAll("[data-action='delete']").forEach((button) => {
    button.onclick = () => deleteRequirement(button.dataset.id);
  });
}

function bindDragAndDrop() {
  document.querySelectorAll(".requirement-card").forEach((card) => {
    card.addEventListener("dragstart", () => card.classList.add("dragging"));
    card.addEventListener("dragend", () => card.classList.remove("dragging"));
  });

  document.querySelectorAll(".column-body").forEach((column) => {
    column.addEventListener("dragover", (event) => event.preventDefault());
    column.addEventListener("drop", (event) => {
      event.preventDefault();
      const dragging = document.querySelector(".requirement-card.dragging");
      if (!dragging) return;
      const requirement = findRequirement(dragging.dataset.id);
      const nextStatus = column.dataset.status;
      if (!requirement || requirement.status === nextStatus) return;

      const timestamp = new Date().toLocaleString("zh-CN", { hour12: false });
      requirement.status = nextStatus;
      requirement.auditTrail.push(`${timestamp} 状态变更为 ${nextStatus}`);
      saveState();
      render();
    });
  });
}

function deleteRequirement(id) {
  const item = findRequirement(id);
  if (!item) return;

  if (!confirm(`确认删除需求“${item.title}”吗？此操作不可撤销。`)) {
    return;
  }

  state.requirements = state.requirements.filter((requirement) => requirement.id !== id);
  saveState();
  render();
}

function addImportHistoryRecord(files, specs) {
  const timestamp = getTimestamp();
  const record = {
    id: `history-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
    files: [...files],
    specs: specs.map((spec) => ({
      ...spec,
      acceptanceCriteria: [...spec.acceptanceCriteria],
      technicalNotes: [...spec.technicalNotes],
      reviewNotes: [...spec.reviewNotes],
      scope: [...spec.scope],
      process: [...spec.process]
    })),
    specCount: specs.length,
    createdAt: timestamp,
    lastUsedAt: timestamp
  };

  state.importHistory = [record, ...state.importHistory].slice(0, MAX_IMPORT_HISTORY);
  saveImportHistory();
}

function restoreImportHistory(historyId) {
  const history = state.importHistory.find((item) => item.id === historyId);
  if (!history) return;

  state.generatedSpecs = history.specs.map((spec) => ({
    ...spec,
    acceptanceCriteria: [...spec.acceptanceCriteria],
    technicalNotes: [...spec.technicalNotes],
    reviewNotes: [...spec.reviewNotes],
    scope: [...spec.scope],
    process: [...spec.process]
  }));
  state.importedFiles = [...history.files];
  history.lastUsedAt = getTimestamp();
  saveImportHistory();
  renderIntakeResult();
  renderImportHistory();
}

function deleteImportHistory(historyId) {
  state.importHistory = state.importHistory.filter((item) => item.id !== historyId);
  saveImportHistory();
  renderImportHistory();
}

function clearImportHistory() {
  if (!state.importHistory.length) return;
  if (!confirm("确认清空全部导入历史吗？此操作不可撤销。")) return;

  state.importHistory = [];
  saveImportHistory();
  renderImportHistory();
}

function buildDetailContent(item) {
  const acceptanceList = toList(item.acceptanceCriteria);
  const technicalList = toList(item.technicalNotes);
  const reviewList = toList(item.reviewNotes);
  const auditList = item.auditTrail?.length ? item.auditTrail : ["暂无审计记录"];

  return `
    <section class="detail-main">
      <div class="detail-hero">
        <div class="detail-badges">
          <span class="badge badge-type">${item.type}</span>
          <span class="badge badge-priority" data-priority="${item.priority}">${item.priority}</span>
          <span class="status-badge" data-status="${item.status}">${item.status}</span>
        </div>
        <h4>${item.title}</h4>
        <p>${item.description || "暂无需求描述"}</p>
        <div class="detail-meta-grid">
          <span class="badge badge-type">提出部门 ${item.department}</span>
          <span class="badge badge-type">负责人 ${item.owner}</span>
          <span class="badge badge-type">提出日期 ${item.submittedAt}</span>
          <span class="badge badge-type">计划上线 ${item.plannedRelease}</span>
        </div>
      </div>

      <div class="detail-section">
        <h4>业务背景</h4>
        <p>${item.background || "暂无业务背景"}</p>
      </div>

      <div class="detail-section">
        <h4>验收标准</h4>
        ${renderDetailList(acceptanceList)}
      </div>

      <div class="detail-section">
        <h4>技术说明</h4>
        ${renderDetailList(technicalList)}
      </div>
    </section>

    <aside class="detail-side">
      <div class="detail-meta-card">
        <h4>评审结论</h4>
        ${renderDetailList(reviewList)}
      </div>

      <div class="detail-meta-card">
        <h4>审计轨迹</h4>
        <div class="detail-timeline">
          ${auditList.map((log) => `<div class="detail-timeline-item"><strong>${extractAuditDate(log)}</strong><p>${extractAuditText(log)}</p></div>`).join("")}
        </div>
      </div>
    </aside>
  `;
}

function toList(value) {
  if (!value) return ["暂无"];
  if (Array.isArray(value)) return value.length ? value : ["暂无"];
  return String(value)
    .split(/\n+/)
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => item.replace(/^\d+\.\s*/, ""));
}

function renderDetailList(items) {
  return `<ol class="detail-list">${items.map((item) => `<li>${item}</li>`).join("")}</ol>`;
}

function extractAuditDate(log) {
  const matched = String(log).match(/^\d{4}-\d{2}-\d{2}(?:\s+\d{2}:\d{2}:\d{2})?/);
  return matched ? matched[0] : "记录";
}

function extractAuditText(log) {
  const date = extractAuditDate(log);
  return String(log).replace(date, "").trim() || log;
}

function findRequirement(id) {
  return state.requirements.find((item) => item.id === id);
}

function priorityScore(priority) {
  return PRIORITY_OPTIONS.indexOf(priority);
}

function today() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function cloneRequirements(items) {
  return items.map((item) => ({
    ...enrichRequirement(item),
    auditTrail: [...(item.auditTrail || [])]
  }));
}

function enrichRequirement(item) {
  const department = item.department || "待补充";
  const owner = item.owner || "待分配";
  return {
    ...item,
    product: item.product || "待补充",
    sourceProject: item.sourceProject || `${department}数字化项目`,
    requester: item.requester || owner,
    deliveryDirector: item.deliveryDirector || "待分配",
    devManager: item.devManager || "待分配",
    productManager: item.productManager || owner,
    reviewStatus: item.reviewStatus || inferReviewStatus(item.status)
  };
}

function inferReviewStatus(status) {
  if (status === "开发中") return "开发中";
  if (status === "已上线" || status === "已归档") return "已上线";
  if (status === "设计中") return "已评审";
  return "待评审";
}

function getTimestamp() {
  return new Date().toLocaleString("zh-CN", { hour12: false });
}

function createRequirementRecord(payload, id, timestamp) {
  return enrichRequirement({
    id,
    title: payload.title || "未命名需求",
    type: payload.type || "业务需求",
    priority: payload.priority || "P2",
    status: payload.status || "待评审",
    department: payload.department || "待补充",
    owner: payload.owner || "待分配",
    submittedAt: payload.submittedAt || today(),
    plannedRelease: normalizeDateHint(payload.plannedRelease) || addDays(14),
    background: payload.background || "",
    impactScope: payload.impactScope || "",
    description: payload.description || "",
    acceptanceCriteria: listToMultiline(payload.acceptanceCriteria),
    technicalNotes: listToMultiline(payload.technicalNotes),
    reviewNotes: listToMultiline(payload.reviewNotes),
    auditTrail: payload.auditTrail?.length ? [...payload.auditTrail] : [`${timestamp} 需求创建，提出部门：${payload.department || "待补充"}`]
  });
}

function createRequirementRecordFromSpec(spec, id, timestamp) {
  return createRequirementRecord(
    {
      ...spec,
      acceptanceCriteria: spec.acceptanceCriteria,
      technicalNotes: spec.technicalNotes,
      reviewNotes: spec.reviewNotes,
      impactScope: spec.impactScope
    },
    id,
    timestamp
  );
}

function createRequirementsFromSpecs(specs) {
  if (!specs.length) return;

  const timestamp = getTimestamp();
  const createdPairs = specs.map((spec, index) => {
    const id = buildRequirementId(new Date(Date.now() + index * 1000));
    return {
      draftId: spec.draftId,
      requirementId: id,
      record: createRequirementRecordFromSpec(spec, id, timestamp)
    };
  });
  const createdRequirements = createdPairs.map((item) => item.record);

  state.requirements = [...createdRequirements, ...state.requirements];
  state.generatedSpecs = state.generatedSpecs.map((spec) => {
    const created = createdPairs.find((item) => item.draftId === spec.draftId);
    if (!created) {
      return spec;
    }
    return { ...spec, createdRequirementId: created.requirementId };
  });

  saveState();
  render();
  alert(`已生成 ${createdRequirements.length} 条需求单。`);
}

function createSingleRequirementFromDraft(draftId) {
  const draft = getGeneratedSpec(draftId);
  if (!draft || draft.createdRequirementId) return;
  createRequirementsFromSpecs([draft]);
}

function getGeneratedSpec(draftId) {
  return state.generatedSpecs.find((spec) => spec.draftId === draftId);
}

async function copyTextToClipboard(text, successMessage) {
  try {
    await navigator.clipboard.writeText(text);
    alert(successMessage);
  } catch {
    alert(text);
  }
}

function formatAllSpecsText(specs) {
  return specs
    .map((spec, index) => [`需求 ${index + 1}`, `来源：${spec.sourceName || "手动输入"}`, formatSpecText(spec)].join("\n"))
    .join("\n\n------------------------------\n\n");
}

async function parseImportedFile(file, extension) {
  if (extension === "docx") {
    return parseWordFile(file);
  }

  if (["xlsx", "xls"].includes(extension)) {
    return parseExcelFile(file);
  }

  const text = await file.text();
  if (!text.trim()) return [];

  if (extension === "csv") {
    return buildSpecsFromCsvText(text, file.name);
  }

  if (extension === "json") {
    return buildSpecsFromJsonText(text, file.name);
  }

  return buildSpecsFromTextSource(text, file.name);
}

async function parseWordFile(file) {
  if (typeof mammoth === "undefined") {
    alert("Word 导入依赖未加载，请检查网络后刷新页面。");
    return [];
  }

  try {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer });
    const text = String(result.value || "").trim();
    if (!text) return [];
    return buildSpecsFromTextSource(text, file.name);
  } catch {
    alert(`文件 ${file.name} 解析失败，请确认是标准 docx 格式。`);
    return [];
  }
}

async function parseExcelFile(file) {
  if (typeof XLSX === "undefined") {
    alert("Excel 导入依赖未加载，请检查网络后刷新页面。");
    return [];
  }

  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const specs = [];

    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) return;

      const sheetEntries = buildSpecsFromWorksheet(sheet, file.name, sheetName);
      specs.push(...sheetEntries);
    });

    return specs;
  } catch {
    alert(`文件 ${file.name} 解析失败，请确认是标准 xlsx/xls 格式。`);
    return [];
  }
}

function buildSpecsFromWorksheet(sheet, sourceName, sheetName) {
  const jsonRows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  const structuredSpecs = jsonRows
    .map((entry, index) => createSpecFromStructuredEntry(entry, `${sourceName} / ${sheetName}`, index, jsonRows.length))
    .filter(Boolean);

  if (structuredSpecs.length) {
    return structuredSpecs;
  }

  const plainText = XLSX.utils.sheet_to_csv(sheet, { blankrows: false }).trim();
  if (!plainText) return [];
  return buildSpecsFromTextSource(plainText, `${sourceName} / ${sheetName}`);
}

function buildSpecsFromTextSource(rawText, sourceName) {
  const blocks = splitRequirementBlocks(rawText);
  return blocks.map((block, index) => transformNaturalLanguageToSpec(block.text, {
    draftId: createDraftId(),
    sourceName,
    sourceBlockLabel: blocks.length > 1 ? block.label || `片段 ${index + 1}` : "",
    titleHint: block.titleHint
  }));
}

function splitRequirementBlocks(rawText) {
  const normalized = String(rawText || "").replace(/\r/g, "").trim();
  if (!normalized) return [];

  const headingBlocks = normalized
    .split(/\n\s*\n(?=(?:#{1,3}\s+|需求\s*[一二三四五六七八九十\d]+[:：]|(?:\d+)[.、]\s+))/)
    .map((block) => block.trim())
    .filter(Boolean);

  if (headingBlocks.length > 1) {
    return headingBlocks.map((block, index) => buildTextBlock(block, index));
  }

  const paragraphBlocks = normalized
    .split(/\n\s*\n/)
    .map((block) => block.trim())
    .filter((block) => block.length > 30 && /(希望|需要|支持|优化|改造|系统|需求|问题|痛点)/.test(block));

  if (paragraphBlocks.length > 1) {
    return paragraphBlocks.map((block, index) => buildTextBlock(block, index));
  }

  return [buildTextBlock(normalized, 0)];
}

function buildTextBlock(text, index) {
  const firstLine = text
    .split("\n")[0]
    .replace(/^#{1,3}\s*/, "")
    .replace(/^需求\s*[一二三四五六七八九十\d]+[:：]?\s*/, "")
    .replace(/^\d+[.、]\s*/, "")
    .trim();

  return {
    text,
    label: `片段 ${index + 1}`,
    titleHint: firstLine.length && firstLine.length <= 32 ? firstLine : ""
  };
}

function buildSpecsFromCsvText(text, sourceName) {
  const rows = parseCsvRows(text).filter((row) => row.some((cell) => String(cell).trim()));
  if (rows.length <= 1) return [];

  const headers = rows[0].map((header) => normalizeKey(header));
  return rows.slice(1).map((row, index) => {
    const record = {};
    headers.forEach((header, headerIndex) => {
      record[header] = row[headerIndex] || "";
    });
    return createSpecFromStructuredEntry(record, sourceName, index, rows.length - 1);
  }).filter(Boolean);
}

function buildSpecsFromJsonText(text, sourceName) {
  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch {
    alert(`文件 ${sourceName} 不是合法的 JSON。`);
    return [];
  }

  const entries = Array.isArray(parsed)
    ? parsed
    : Array.isArray(parsed.requirements)
      ? parsed.requirements
      : Array.isArray(parsed.data)
        ? parsed.data
        : [parsed];

  return entries.map((entry, index) => {
    if (typeof entry === "string") {
      return transformNaturalLanguageToSpec(entry, {
        draftId: createDraftId(),
        sourceName,
        sourceBlockLabel: entries.length > 1 ? `条目 ${index + 1}` : ""
      });
    }
    return createSpecFromStructuredEntry(entry, sourceName, index, entries.length);
  }).filter(Boolean);
}

function createSpecFromStructuredEntry(entry, sourceName, index, total) {
  const normalizedEntry = normalizeImportedEntry(entry);
  const rawText = [
    normalizedEntry.title,
    normalizedEntry.background,
    normalizedEntry.description,
    normalizedEntry.goal,
    normalizedEntry.reviewNotes
  ].filter(Boolean).join("；");

  if (!rawText) return null;

  return transformNaturalLanguageToSpec(rawText, {
    draftId: createDraftId(),
    sourceName,
    sourceBlockLabel: total > 1 ? `条目 ${index + 1}` : "",
    titleHint: normalizedEntry.title,
    typeHint: normalizedEntry.type,
    priorityHint: normalizedEntry.priority,
    departmentHint: normalizedEntry.department,
    ownerHint: normalizedEntry.owner,
    plannedReleaseHint: normalizedEntry.plannedRelease,
    backgroundHint: normalizedEntry.background,
    descriptionHint: normalizedEntry.description,
    goalHint: normalizedEntry.goal,
    acceptanceCriteriaHint: normalizedEntry.acceptanceCriteria,
    technicalNotesHint: normalizedEntry.technicalNotes,
    reviewNotesHint: normalizedEntry.reviewNotes
  });
}

function normalizeImportedEntry(entry) {
  const source = Object.fromEntries(
    Object.entries(entry || {}).map(([key, value]) => [normalizeKey(key), value])
  );

  return {
    title: pickFirstValue(source, ["title", "标题", "需求标题", "需求名称", "name"]),
    type: pickFirstValue(source, ["type", "类型", "需求类型"]),
    priority: pickFirstValue(source, ["priority", "优先级"]),
    department: pickFirstValue(source, ["department", "部门", "提出部门"]),
    owner: pickFirstValue(source, ["owner", "负责人", "ownername"]),
    plannedRelease: pickFirstValue(source, ["plannedrelease", "计划上线", "上线时间", "duedate"]),
    background: pickFirstValue(source, ["background", "业务背景", "背景", "problem", "painpoint"]),
    goal: pickFirstValue(source, ["goal", "目标", "产品目标"]),
    description: pickFirstValue(source, ["description", "需求描述", "描述", "content", "需求内容"]),
    acceptanceCriteria: pickFirstValue(source, ["acceptancecriteria", "验收标准", "acceptance"]),
    technicalNotes: pickFirstValue(source, ["technicalnotes", "技术说明", "技术备注"]),
    reviewNotes: pickFirstValue(source, ["reviewnotes", "评审结论", "风险提示", "notes"])
  };
}

function pickFirstValue(source, keys) {
  const normalizedKeys = keys.map((key) => normalizeKey(key));
  const hitKey = normalizedKeys.find((key) => source[key]);
  return hitKey ? String(source[hitKey]).trim() : "";
}

function normalizeKey(value) {
  return String(value || "").toLowerCase().replace(/[\s_-]/g, "");
}

function parseCsvRows(text) {
  const lines = String(text || "").replace(/\r/g, "").split("\n").filter((line) => line.trim());
  return lines.map(parseCsvLine);
}

function parseCsvLine(line) {
  const cells = [];
  let current = "";
  let inQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index];
    const nextChar = line[index + 1];

    if (char === "\"") {
      if (inQuotes && nextChar === "\"") {
        current += "\"";
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      cells.push(current.trim());
      current = "";
      continue;
    }

    current += char;
  }

  cells.push(current.trim());
  return cells;
}

function createDraftId() {
  return `draft-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function getFileExtension(fileName) {
  const matched = String(fileName).toLowerCase().match(/\.([^.]+)$/);
  return matched ? matched[1] : "";
}

function normalizeListValue(value) {
  if (!value) return [];
  if (Array.isArray(value)) {
    return value.map((item) => String(item).trim()).filter(Boolean);
  }
  return String(value)
    .split(/\n+|；|;/)
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => item.replace(/^\d+[.、]\s*/, ""));
}

function listToMultiline(value) {
  if (Array.isArray(value)) return value.join("\n");
  return String(value || "").trim();
}

function normalizeDateHint(value) {
  if (!value) return "";
  const raw = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const slashMatched = raw.match(/^(\d{4})[\/.](\d{1,2})[\/.](\d{1,2})$/);
  if (slashMatched) {
    return `${slashMatched[1]}-${String(slashMatched[2]).padStart(2, "0")}-${String(slashMatched[3]).padStart(2, "0")}`;
  }
  return inferReleaseDate(raw);
}

function transformNaturalLanguageToSpec(rawText, options = {}) {
  const segments = splitTextSegments(rawText);
  const title = options.titleHint || buildTitle(rawText);
  const type = TYPE_OPTIONS.includes(options.typeHint) ? options.typeHint : inferType(rawText);
  const priority = PRIORITY_OPTIONS.includes(options.priorityHint) ? options.priorityHint : inferPriority(rawText);
  const summary = buildSummary(rawText);
  const background = options.backgroundHint || segments.problem || `当前存在${summary}相关痛点，业务希望通过系统化能力提升处理效率和结果稳定性。`;
  const goal = options.goalHint || segments.goal || "通过产品化方案固化流程、提升处理效率，并形成可跟踪、可复盘的需求交付闭环。";
  const scope = normalizeListValue(options.scopeHint).length ? normalizeListValue(options.scopeHint) : buildScope(rawText, segments);
  const process = normalizeListValue(options.processHint).length ? normalizeListValue(options.processHint) : buildProcess(rawText, segments);
  const acceptanceCriteria = normalizeListValue(options.acceptanceCriteriaHint).length ? normalizeListValue(options.acceptanceCriteriaHint) : buildAcceptance(rawText, scope, process);
  const technicalNotes = normalizeListValue(options.technicalNotesHint).length ? normalizeListValue(options.technicalNotesHint) : buildTechnicalNotes(rawText);
  const reviewNotes = normalizeListValue(options.reviewNotesHint).length ? normalizeListValue(options.reviewNotesHint) : buildReviewNotes(rawText, priority);
  const plannedRelease = normalizeDateHint(options.plannedReleaseHint) || inferReleaseDate(rawText);

  return {
    draftId: options.draftId || createDraftId(),
    sourceName: options.sourceName || "手动输入",
    sourceBlockLabel: options.sourceBlockLabel || "",
    title,
    type,
    priority,
    status: "待评审",
    department: options.departmentHint || inferDepartment(rawText),
    owner: options.ownerHint || "待分配",
    submittedAt: today(),
    plannedRelease,
    summary,
    background,
    goal,
    scope,
    process,
    description: options.descriptionHint || [goal, ...scope].join("；"),
    acceptanceCriteria,
    technicalNotes,
    reviewNotes
  };
}

function splitTextSegments(rawText) {
  const sentences = rawText
    .replace(/\r/g, "")
    .split(/[。；;\n]/)
    .map((item) => item.trim())
    .filter(Boolean);

  return {
    problem: sentences.find((item) => /(问题|痛点|低|慢|漏|错|断货|依赖|手工|人工|重复)/.test(item)),
    goal: sentences.find((item) => /(希望|需要|想要|目标|支持|实现|上线)/.test(item))
  };
}

function buildTitle(rawText) {
  const object = matchKeyword(rawText, [
    ["补货", "补货优化"],
    ["调拨", "库存调拨协同"],
    ["审批", "审批流程优化"],
    ["报销", "报销流转提效"],
    ["提醒", "提醒触达改造"],
    ["工单", "工单处理优化"],
    ["会员", "会员运营能力升级"],
    ["库存", "库存协同优化"],
    ["商品", "商品主数据打通"]
  ]);

  return object ? `${object}需求` : "自然语言识别生成需求";
}

function buildSummary(rawText) {
  const sanitized = rawText.replace(/\s+/g, " ").trim();
  return sanitized.length > 48 ? `${sanitized.slice(0, 48)}...` : sanitized;
}

function inferType(rawText) {
  if (/(系统|接口|同步|数据|ERP|OMS|主数据|自动化)/i.test(rawText)) return "系统改造";
  if (/(项目|上线|交付|改造一期|里程碑|排期)/.test(rawText)) return "项目需求";
  if (/(流程|审批|报销|协同|提效|流转)/.test(rawText)) return "流程优化";
  if (/(用户|会员|页面|提醒|配置|触达|小程序|产品)/.test(rawText)) return "产品需求";
  return "业务需求";
}

function inferPriority(rawText) {
  if (/(紧急|尽快|马上|本周|核心|关键|严重|断货|阻塞)/.test(rawText)) return "P0";
  if (/(本月|月底|重要|必须|6月|5月|Q2|Q3)/i.test(rawText)) return "P1";
  if (/(优化|改版|提升|支持|希望)/.test(rawText)) return "P2";
  return "P3";
}

function inferDepartment(rawText) {
  return matchKeyword(rawText, [
    ["门店", "零售运营部"],
    ["仓库", "仓储运营部"],
    ["库存", "供应链中心"],
    ["会员", "用户增长部"],
    ["财务", "财务共享中心"],
    ["客服", "客户服务部"],
    ["海外", "国际业务部"],
    ["商品", "商品运营部"]
  ]) || "待补充";
}

function buildScope(rawText, segments) {
  const items = [];
  if (/(申请|发起)/.test(rawText)) items.push("支持业务角色发起需求相关申请或操作录入");
  if (/(审批|审核)/.test(rawText)) items.push("支持审核人进行审批、驳回或补充说明");
  if (/(同步|回写|接口|ERP|OMS)/i.test(rawText)) items.push("支持与上下游系统进行数据同步或结果回写");
  if (/(提醒|通知|消息|短信|站内信)/.test(rawText)) items.push("支持关键节点提醒与异常通知");
  if (/(移动端|小程序|手机)/.test(rawText)) items.push("支持移动端或小程序访问场景");
  if (/(看板|报表|统计|追踪)/.test(rawText)) items.push("支持结果追踪、看板展示和效果统计");
  if (items.length === 0) {
    items.push("明确核心业务对象、触发条件和处理流程");
    items.push("沉淀产品交互、数据口径和协同节点");
  }
  if (segments.goal) items.unshift(`围绕“${segments.goal}”落地对应功能能力`);
  return items;
}

function buildProcess(rawText, segments) {
  const actors = [];
  if (/(门店|店长)/.test(rawText)) actors.push("门店/业务发起人提交申请或触发任务");
  if (/(仓库|区域经理|主管|审核|审批)/.test(rawText)) actors.push("审核角色进行校验、审批或处理");
  if (/(系统|自动|同步|回写)/.test(rawText)) actors.push("系统自动完成计算、同步、回写或通知");
  if (!actors.length) {
    actors.push("业务方提交原始诉求并补充上下文");
    actors.push("产品与研发确认范围、规则和交付方式");
    actors.push("系统按约定流程执行并沉淀结果");
  }
  if (segments.problem) actors.push(`重点解决问题：${segments.problem}`);
  return actors;
}

function buildAcceptance(rawText, scope, process) {
  const criteria = [
    "需求涉及的核心流程可在系统中完整闭环执行",
    "关键字段、状态流转和操作日志可追踪",
    "异常场景有明确提示，且不影响主流程继续处理"
  ];

  if (/(同步|回写|接口|ERP|OMS)/i.test(rawText)) {
    criteria.push("上下游系统同步成功率、失败重试和错误提示规则明确");
  }
  if (/(提醒|通知|消息)/.test(rawText)) {
    criteria.push("关键节点提醒可配置，触达结果可查看");
  }
  if (/(移动端|小程序|手机)/.test(rawText)) {
    criteria.push("移动端核心操作链路可用，展示与交互适配小屏场景");
  }

  return criteria.slice(0, 5);
}

function buildTechnicalNotes(rawText) {
  const notes = [
    "需补充数据来源、字段定义和状态机约束",
    "建议明确角色权限与操作边界，避免越权修改"
  ];

  if (/(ERP|OMS|接口|同步|回写)/i.test(rawText)) notes.push("需定义接口协议、失败补偿和幂等处理策略");
  if (/(看板|报表|统计)/.test(rawText)) notes.push("需统一统计口径与刷新频率，避免业务和技术口径不一致");
  if (/(移动端|小程序|手机)/.test(rawText)) notes.push("需确认移动端容器、登录态和兼容性要求");

  return notes.slice(0, 5);
}

function buildReviewNotes(rawText, priority) {
  const notes = [
    "评审时需确认业务价值、目标指标和上线时间是否一致",
    "需确认涉及部门、负责人和依赖系统资源是否到位"
  ];

  if (priority === "P0" || priority === "P1") {
    notes.push("优先级较高，建议尽早完成范围冻结和排期评审");
  }
  if (/(法务|合规|短信|隐私|海外)/.test(rawText)) {
    notes.push("涉及合规或外部约束，需补充法务/安全评审");
  }
  if (/(同步|接口|ERP|OMS)/i.test(rawText)) {
    notes.push("涉及系统联动，需提前安排接口联调窗口");
  }

  return notes.slice(0, 5);
}

function inferReleaseDate(rawText) {
  const monthDayMatch = rawText.match(/(\d{1,2})月(\d{1,2})[日前号]/);
  if (monthDayMatch) {
    const year = new Date().getFullYear();
    const month = String(monthDayMatch[1]).padStart(2, "0");
    const day = String(monthDayMatch[2]).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  const monthEndMatch = rawText.match(/(\d{1,2})月底/);
  if (monthEndMatch) {
    const year = new Date().getFullYear();
    const monthNumber = Number(monthEndMatch[1]);
    const lastDay = new Date(year, monthNumber, 0).getDate();
    return `${year}-${String(monthNumber).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}`;
  }

  if (/本周/.test(rawText)) return addDays(7);
  if (/本月/.test(rawText)) return addDays(15);
  return addDays(14);
}

function matchKeyword(rawText, options) {
  const matched = options.find(([keyword]) => rawText.includes(keyword));
  return matched ? matched[1] : "";
}

function populateFormFromSpec(spec) {
  const form = elements.requirementForm.elements;
  form.product.value = spec.product || "待补充";
  form.title.value = spec.title;
  form.sourceProject.value = spec.sourceProject || `${spec.department || "待补充"}数字化项目`;
  form.requester.value = spec.requester || spec.owner;
  form.deliveryDirector.value = spec.deliveryDirector || "待分配";
  form.devManager.value = spec.devManager || "待分配";
  form.productManager.value = spec.productManager || spec.owner;
  form.reviewStatus.value = spec.reviewStatus || "待评审";
  form.type.value = spec.type;
  form.priority.value = spec.priority;
  form.submittedAt.value = spec.submittedAt;
  form.background.value = spec.background;
  form.impactScope.value = spec.impactScope || "";
  form.description.value = spec.description;
  form.acceptanceCriteria.value = listToMultiline(spec.acceptanceCriteria);
  form.technicalNotes.value = listToMultiline(spec.technicalNotes);
  form.reviewNotes.value = listToMultiline(spec.reviewNotes);
}

function formatSpecText(spec) {
  return [
    `标题：${spec.title}`,
    `来源：${spec.sourceName || "手动输入"}${spec.sourceBlockLabel ? ` / ${spec.sourceBlockLabel}` : ""}`,
    `类型：${spec.type}`,
    `优先级：${spec.priority}`,
    `状态：${spec.status}`,
    `提出部门：${spec.department}`,
    `计划上线：${spec.plannedRelease}`,
    "",
    `业务背景：${spec.background}`,
    `产品目标：${spec.goal}`,
    "",
    "功能范围：",
    ...spec.scope.map((item, index) => `${index + 1}. ${item}`),
    "",
    "角色与流程：",
    ...spec.process.map((item, index) => `${index + 1}. ${item}`),
    "",
    "验收标准：",
    ...spec.acceptanceCriteria.map((item, index) => `${index + 1}. ${item}`),
    "",
    "技术说明：",
    ...spec.technicalNotes.map((item, index) => `${index + 1}. ${item}`),
    "",
    "评审与风险提示：",
    ...spec.reviewNotes.map((item, index) => `${index + 1}. ${item}`)
  ].join("\n");
}

function addDays(days) {
  const date = new Date();
  date.setDate(date.getDate() + days);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}
