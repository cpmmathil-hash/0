<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>14th Five-Year Plan - Grama Panchayat Project Register & Evaluation System</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- SheetJS (xlsx) for Excel Export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <!-- jsPDF & AutoTable for PDF Export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.31/jspdf.plugin.autotable.min.js"></script>
    <style>
        .sidebar-item-active { background-color: #1e3a8a; color: #ffffff; border-left: 4px solid #3b82f6; }
        .tab-btn-active { border-bottom: 2px solid #2563eb; color: #2563eb; font-weight: 600; }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f1f5f9; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 3px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
        @media print {
            .no-print { display: none !important; }
            .print-only { display: block !important; }
        }
    </style>
</head>
<body class="bg-slate-100 font-sans text-slate-800 antialiased min-h-screen flex flex-col">

    <!-- TOP BAR / HEADER -->
    <header class="bg-blue-900 text-white shadow-md sticky top-0 z-40 no-print">
        <div class="px-4 py-2.5 flex flex-wrap justify-between items-center">
            <div class="flex items-center space-x-3">
                <button id="sidebar-toggle" class="p-2 rounded hover:bg-blue-800 text-slate-200 lg:hidden">
                    <i class="fa-solid fa-bars text-lg"></i>
                </button>
                <div class="flex items-center space-x-2">
                    <div class="bg-amber-500 text-blue-950 font-black p-1.5 rounded-lg text-xs leading-none">LSGD</div>
                    <div>
                        <h1 class="text-sm md:text-base font-bold tracking-wide uppercase leading-tight" id="header-panchayat-name">KARIVELLUR-PERALAM GRAMA PANCHAYAT</h1>
                        <p class="text-xs text-blue-200 hidden sm:block">14th Five-Year Plan (2022–2027) Project Register & Performance Evaluation System</p>
                    </div>
                </div>
            </div>

            <div class="flex items-center space-x-4 text-xs md:text-sm">
                <div class="hidden md:flex flex-col text-right">
                    <span class="font-medium text-slate-200" id="user-role-display">Role: Administrator</span>
                    <span class="text-[10px] text-blue-300" id="current-fy-display">FY: 2024-2025 (Annual Plan Year 3)</span>
                </div>
                <div class="flex items-center space-x-2 bg-blue-950 px-3 py-1.5 rounded-lg border border-blue-800">
                    <i class="fa-solid fa-user-shield text-amber-400"></i>
                    <select id="role-selector" class="bg-transparent text-xs text-white outline-none cursor-pointer border-none font-medium" onchange="switchUserRole(this.value)">
                        <option value="Admin" class="bg-slate-800 text-white">Administrator</option>
                        <option value="Coordinator" class="bg-slate-800 text-white">Planning Coordinator</option>
                        <option value="ImplementingOfficer" class="bg-slate-800 text-white">Implementing Officer</option>
                        <option value="FinanceOfficer" class="bg-slate-800 text-white">Finance Officer</option>
                        <option value="Engineer" class="bg-slate-800 text-white">Technical Engineer</option>
                        <option value="Evaluator" class="bg-slate-800 text-white">Project Evaluator</option>
                        <option value="Reviewer" class="bg-slate-800 text-white">Read-Only Reviewer</option>
                    </select>
                </div>
            </div>
        </div>
    </header>

    <!-- MAIN WRAPPER -->
    <div class="flex-1 flex overflow-hidden">
        
        <!-- SIDEBAR NAVIGATION -->
        <aside id="sidebar" class="w-64 bg-slate-900 text-slate-300 flex-shrink-0 flex flex-col transition-all duration-300 z-30 fixed lg:relative inset-y-0 left-0 -translate-x-full lg:translate-x-0 no-print">
            <div class="p-3 border-b border-slate-800 text-xs font-semibold text-slate-400 uppercase tracking-wider flex justify-between items-center">
                <span>Navigation Menu</span>
                <span class="bg-slate-800 text-blue-400 px-2 py-0.5 rounded text-[10px]">14th FYP</span>
            </div>
            
            <nav class="flex-1 overflow-y-auto custom-scrollbar p-2 space-y-1 text-sm font-medium">
                <a href="#" onclick="switchModule('dashboard')" id="nav-dashboard" class="sidebar-item-active flex items-center space-x-3 px-3 py-2.5 rounded-lg transition-colors hover:bg-slate-800">
                    <i class="fa-solid fa-chart-pie w-5 text-center text-amber-400"></i>
                    <span>Main Dashboard</span>
                </a>
                <a href="#" onclick="switchModule('master-register')" id="nav-master-register" class="flex items-center space-x-3 px-3 py-2.5 rounded-lg transition-colors hover:bg-slate-800">
                    <i class="fa-solid fa-folder-plus w-5 text-center text-blue-400"></i>
                    <span>Project Master Register</span>
                </a>
                <a href="#" onclick="switchModule('financial-register')" id="nav-financial-register" class="flex items-center space-x-3 px-3 py-2.5 rounded-lg transition-colors hover:bg-slate-800">
                    <i class="fa-solid fa-indian-rupee-sign w-5 text-center text-emerald-400"></i>
                    <span>Financial Register</span>
                </a>
                <a href="#" onclick="switchModule('progress-register')" id="nav-progress-register" class="flex items-center space-x-3 px-3 py-2.5 rounded-lg transition-colors hover:bg-slate-800">
                    <i class="fa-solid fa-bars-progress w-5 text-center text-purple-400"></i>
                    <span>Physical Progress</span>
                </a>
                <a href="#" onclick="switchModule('beneficiary-register')" id="nav-beneficiary-register" class="flex items-center space-x-3 px-3 py-2.5 rounded-lg transition-colors hover:bg-slate-800">
                    <i class="fa-solid fa-users w-5 text-center text-teal-400"></i>
                    <span>Beneficiary Register</span>
                </a>
                <a href="#" onclick="switchModule('inspection-register')" id="nav-inspection-register" class="flex items-center space-x-3 px-3 py-2.5 rounded-lg transition-colors hover:bg-slate-800">
                    <i class="fa-solid fa-clipboard-check w-5 text-center text-cyan-400"></i>
                    <span>Field Inspection Log</span>
                </a>
                <a href="#" onclick="switchModule('evaluation-module')" id="nav-evaluation-module" class="flex items-center space-x-3 px-3 py-2.5 rounded-lg transition-colors hover:bg-slate-800">
                    <i class="fa-solid fa-award w-5 text-center text-amber-400"></i>
                    <span>Project Evaluation</span>
                </a>
                <a href="#" onclick="switchModule('outcome-register')" id="nav-outcome-register" class="flex items-center space-x-3 px-3 py-2.5 rounded-lg transition-colors hover:bg-slate-800">
                    <i class="fa-solid fa-bullseye w-5 text-center text-rose-400"></i>
                    <span>Outputs & Outcomes</span>
                </a>
                <a href="#" onclick="switchModule('corrective-action')" id="nav-corrective-action" class="flex items-center space-x-3 px-3 py-2.5 rounded-lg transition-colors hover:bg-slate-800">
                    <i class="fa-solid fa-triangle-exclamation w-5 text-center text-red-400"></i>
                    <span>Corrective Actions</span>
                    <span id="badge-corrective-count" class="ml-auto bg-red-600 text-white text-[10px] px-1.5 py-0.2 rounded-full hidden">0</span>
                </a>
                <a href="#" onclick="switchModule('documents-register')" id="nav-documents-register" class="flex items-center space-x-3 px-3 py-2.5 rounded-lg transition-colors hover:bg-slate-800">
                    <i class="fa-solid fa-file-lines w-5 text-center text-indigo-400"></i>
                    <span>Document Vault</span>
                </a>
                <a href="#" onclick="switchModule('reports-module')" id="nav-reports-module" class="flex items-center space-x-3 px-3 py-2.5 rounded-lg transition-colors hover:bg-slate-800">
                    <i class="fa-solid fa-print w-5 text-center text-yellow-400"></i>
                    <span>Reports & Export</span>
                </a>
                <a href="#" onclick="switchModule('master-settings')" id="nav-master-settings" class="flex items-center space-x-3 px-3 py-2.5 rounded-lg transition-colors hover:bg-slate-800 border-t border-slate-800 mt-2">
                    <i class="fa-solid fa-sliders w-5 text-center text-slate-400"></i>
                    <span>Master Settings</span>
                </a>
            </nav>

            <div class="p-3 bg-slate-950 border-t border-slate-800 text-[11px] text-slate-400 flex items-center justify-between">
                <div>
                    <i class="fa-solid fa-database text-xs text-emerald-400 mr-1"></i> Data: LocalStorage
                </div>
                <button onclick="resetToSampleData()" class="hover:text-amber-400 transition-colors" title="Reset DB with Kerala Sample Data">
                    <i class="fa-solid fa-arrows-rotate"></i> Reset
                </button>
            </div>
        </aside>

        <!-- MAIN CONTENT CONTAINER -->
        <main class="flex-1 overflow-y-auto custom-scrollbar p-4 md:p-6 bg-slate-50">

            <!-- ALERTS / NOTIFICATIONS CONTAINER -->
            <div id="notification-banner" class="hidden mb-4 p-3 rounded-lg flex items-center justify-between text-sm shadow-sm transition-all duration-300">
                <div class="flex items-center space-x-2">
                    <i id="notif-icon" class="fa-solid"></i>
                    <span id="notif-message" class="font-medium"></span>
                </div>
                <button onclick="closeNotification()" class="text-slate-500 hover:text-slate-800"><i class="fa-solid fa-xmark"></i></button>
            </div>

            <!-- MODULE 1: DASHBOARD -->
            <section id="module-dashboard" class="space-y-6">
                
                <!-- Filter bar for dashboard -->
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap items-center justify-between gap-4">
                    <div>
                        <h2 class="text-xl font-bold text-slate-800 flex items-center gap-2">
                            <i class="fa-solid fa-chart-line text-blue-700"></i> 14th Five-Year Plan Performance Dashboard
                        </h2>
                        <p class="text-xs text-slate-500">Live monitoring & evaluation analytics for Panchayath Projects</p>
                    </div>
                    <div class="flex flex-wrap items-center gap-2 text-xs">
                        <select id="dash-filter-sector" onchange="renderDashboard()" class="border border-slate-300 rounded-lg px-2.5 py-1.5 bg-slate-50">
                            <option value="ALL">All Sectors</option>
                        </select>
                        <select id="dash-filter-fy" onchange="renderDashboard()" class="border border-slate-300 rounded-lg px-2.5 py-1.5 bg-slate-50">
                            <option value="ALL">All Financial Years</option>
                        </select>
                        <button onclick="renderDashboard()" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-lg flex items-center gap-1 shadow-sm">
                            <i class="fa-solid fa-rotate text-xs"></i> Refresh
                        </button>
                    </div>
                </div>

                <!-- 18 SUMMARY CARDS GRID -->
                <div class="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 gap-3.5" id="dashboard-cards-container">
                    <!-- Cards will be populated dynamically -->
                </div>

                <!-- CHARTS GRID -->
                <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
                    
                    <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                        <h3 class="text-xs font-bold text-slate-600 uppercase tracking-wider mb-3 flex items-center justify-between">
                            <span>Projects by Sector</span>
                            <i class="fa-solid fa-pie-chart text-blue-500"></i>
                        </h3>
                        <div class="h-56 relative"><canvas id="chart-sector"></canvas></div>
                    </div>

                    <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                        <h3 class="text-xs font-bold text-slate-600 uppercase tracking-wider mb-3 flex items-center justify-between">
                            <span>Implementation Status</span>
                            <i class="fa-solid fa-chart-doughnut text-emerald-500"></i>
                        </h3>
                        <div class="h-56 relative"><canvas id="chart-status"></canvas></div>
                    </div>

                    <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                        <h3 class="text-xs font-bold text-slate-600 uppercase tracking-wider mb-3 flex items-center justify-between">
                            <span>Approved vs Expenditure</span>
                            <i class="fa-solid fa-chart-bar text-purple-500"></i>
                        </h3>
                        <div class="h-56 relative"><canvas id="chart-finance"></canvas></div>
                    </div>

                    <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                        <h3 class="text-xs font-bold text-slate-600 uppercase tracking-wider mb-3 flex items-center justify-between">
                            <span>Evaluation Performance Grades</span>
                            <i class="fa-solid fa-award text-amber-500"></i>
                        </h3>
                        <div class="h-56 relative"><canvas id="chart-evaluation"></canvas></div>
                    </div>

                    <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                        <h3 class="text-xs font-bold text-slate-600 uppercase tracking-wider mb-3 flex items-center justify-between">
                            <span>Year-Wise Physical vs Financial %</span>
                            <i class="fa-solid fa-line-chart text-teal-500"></i>
                        </h3>
                        <div class="h-56 relative"><canvas id="chart-progress-trend"></canvas></div>
                    </div>

                    <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                        <h3 class="text-xs font-bold text-slate-600 uppercase tracking-wider mb-3 flex items-center justify-between">
                            <span>Projects by Implementing Officer</span>
                            <i class="fa-solid fa-user-gear text-indigo-500"></i>
                        </h3>
                        <div class="h-56 relative"><canvas id="chart-officer"></canvas></div>
                    </div>
                </div>
            </section>

            <!-- MODULE 2: PROJECT MASTER REGISTER -->
            <section id="module-master-register" class="hidden space-y-6">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
                    <div>
                        <h2 class="text-lg font-bold text-slate-800">Project Master Register</h2>
                        <p class="text-xs text-slate-500">14th Five-Year Plan Approved Panchayat Projects List & Management</p>
                    </div>
                    <div class="flex items-center space-x-2">
                        <button onclick="openProjectModal()" class="bg-blue-700 hover:bg-blue-800 text-white px-3.5 py-2 rounded-lg text-xs font-medium flex items-center gap-1.5 shadow-sm">
                            <i class="fa-solid fa-plus"></i> New Project Registration
                        </button>
                    </div>
                </div>

                <!-- SEARCH & FILTERS -->
                <div class="bg-white p-3.5 rounded-xl shadow-sm border border-slate-200 grid grid-cols-1 sm:grid-cols-2 md:grid-cols-5 gap-3 text-xs">
                    <input type="text" id="master-search" onkeyup="filterProjectMasterTable()" placeholder="Search Project ID, Name..." class="border border-slate-300 rounded-lg px-3 py-2 outline-none focus:border-blue-500">
                    <select id="master-filter-sector" onchange="filterProjectMasterTable()" class="border border-slate-300 rounded-lg px-2.5 py-2 bg-slate-50">
                        <option value="">All Sectors</option>
                    </select>
                    <select id="master-filter-status" onchange="filterProjectMasterTable()" class="border border-slate-300 rounded-lg px-2.5 py-2 bg-slate-50">
                        <option value="">All Statuses</option>
                    </select>
                    <select id="master-filter-fy" onchange="filterProjectMasterTable()" class="border border-slate-300 rounded-lg px-2.5 py-2 bg-slate-50">
                        <option value="">All Financial Years</option>
                    </select>
                    <select id="master-filter-ward" onchange="filterProjectMasterTable()" class="border border-slate-300 rounded-lg px-2.5 py-2 bg-slate-50">
                        <option value="">All Wards</option>
                    </select>
                </div>

                <!-- PROJECT TABLE -->
                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                    <div class="overflow-x-auto custom-scrollbar">
                        <table class="w-full text-left text-xs text-slate-600 border-collapse">
                            <thead class="bg-slate-100 uppercase text-[11px] text-slate-700 font-semibold border-b border-slate-200">
                                <tr>
                                    <th class="py-3 px-3">Project Code / Name</th>
                                    <th class="py-3 px-3">Sector & Category</th>
                                    <th class="py-3 px-3">Implementing Officer</th>
                                    <th class="py-3 px-3">Financial Year</th>
                                    <th class="py-3 px-3 text-right">Sanctioned Cost (₹)</th>
                                    <th class="py-3 px-3 text-center">Physical %</th>
                                    <th class="py-3 px-3 text-center">Status</th>
                                    <th class="py-3 px-3 text-center">Actions</th>
                                </tr>
                            </thead>
                            <tbody id="project-master-tbody" class="divide-y divide-slate-200">
                                <!-- Dynamic rows -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </section>

            <!-- MODULE 3: PROJECT FINANCIAL REGISTER -->
            <section id="module-financial-register" class="hidden space-y-6">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
                    <div>
                        <h2 class="text-lg font-bold text-slate-800">Project Financial Register</h2>
                        <p class="text-xs text-slate-500">Fund Allocation Breakdown, Expenditure Logging & Voucher Management</p>
                    </div>
                    <div class="text-xs text-slate-500">
                        Select project to record voucher expenditures
                    </div>
                </div>

                <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    <!-- Project Selector & Summary -->
                    <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 space-y-4">
                        <label class="block text-xs font-bold text-slate-700 uppercase">Select Project</label>
                        <select id="fin-project-select" onchange="loadProjectFinancials(this.value)" class="w-full border border-slate-300 rounded-lg p-2 text-xs bg-slate-50">
                            <!-- Options -->
                        </select>

                        <div id="fin-summary-card" class="bg-slate-50 p-4 rounded-lg border border-slate-200 space-y-2 text-xs">
                            <p class="text-slate-500">Select a project above to inspect financial ledger.</p>
                        </div>
                    </div>

                    <!-- Expenditure Logging & Transaction History -->
                    <div class="lg:col-span-2 bg-white p-4 rounded-xl shadow-sm border border-slate-200 space-y-4">
                        <div class="flex justify-between items-center border-b border-slate-200 pb-3">
                            <h3 class="text-sm font-bold text-slate-800">Expenditure Voucher Transactions</h3>
                            <button id="btn-add-expenditure" onclick="openExpenditureModal()" class="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1 shadow-sm">
                                <i class="fa-solid fa-plus"></i> Add Expenditure Voucher
                            </button>
                        </div>

                        <div class="overflow-x-auto">
                            <table class="w-full text-xs text-left text-slate-600 border-collapse">
                                <thead class="bg-slate-100 uppercase text-[10px] text-slate-700 font-semibold border-b border-slate-200">
                                    <tr>
                                        <th class="py-2.5 px-3">Date</th>
                                        <th class="py-2.5 px-3">Voucher No</th>
                                        <th class="py-2.5 px-3">Head of Account</th>
                                        <th class="py-2.5 px-3">Purpose / Details</th>
                                        <th class="py-2.5 px-3 text-right">Amount (₹)</th>
                                        <th class="py-2.5 px-3 text-center">UC Status</th>
                                        <th class="py-2.5 px-3 text-center">Action</th>
                                    </tr>
                                </thead>
                                <tbody id="fin-voucher-tbody" class="divide-y divide-slate-200">
                                    <tr><td colspan="7" class="text-center py-4 text-slate-400">No project selected</td></tr>
                                </tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </section>

            <!-- MODULE 4: PHYSICAL PROGRESS REGISTER -->
            <section id="module-progress-register" class="hidden space-y-6">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
                    <div>
                        <h2 class="text-lg font-bold text-slate-800">Physical Progress & Milestone Register</h2>
                        <p class="text-xs text-slate-500">Log Physical %, Implementation Stages & Works Completion Log</p>
                    </div>
                </div>

                <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 space-y-4">
                        <label class="block text-xs font-bold text-slate-700 uppercase">Select Project</label>
                        <select id="prog-project-select" onchange="loadProjectProgress(this.value)" class="w-full border border-slate-300 rounded-lg p-2 text-xs bg-slate-50"></select>

                        <div id="prog-summary-card" class="bg-slate-50 p-4 rounded-lg border border-slate-200 space-y-3 text-xs">
                            <p class="text-slate-500">Select a project to view and update physical progress milestones.</p>
                        </div>
                    </div>

                    <div class="lg:col-span-2 space-y-4">
                        <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 space-y-4">
                            <div class="flex justify-between items-center border-b border-slate-200 pb-3">
                                <h3 class="text-sm font-bold text-slate-800">Update Physical Progress Status</h3>
                            </div>
                            <form id="progress-update-form" onsubmit="saveProgressUpdate(event)" class="grid grid-cols-1 sm:grid-cols-2 gap-4 text-xs">
                                <div>
                                    <label class="block font-medium text-slate-700 mb-1">Update Date</label>
                                    <input type="date" id="prog-update-date" required class="w-full border border-slate-300 rounded-lg p-2">
                                </div>
                                <div>
                                    <label class="block font-medium text-slate-700 mb-1">Implementation Stage</label>
                                    <select id="prog-stage-select" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50">
                                        <option value="Proposed">Proposed</option>
                                        <option value="Approved">Approved</option>
                                        <option value="Tender / Quotation">Tender / Quotation Stage</option>
                                        <option value="Work Order Issued">Work Order Issued</option>
                                        <option value="Work in Progress">Work in Progress</option>
                                        <option value="Substructure / Foundation">Foundation / Initial Stage</option>
                                        <option value="Superstructure / Intermediate">Intermediate Stage</option>
                                        <option value="Finishing Works">Finishing Works</option>
                                        <option value="Completed">Completed</option>
                                        <option value="Financially Closed">Financially Closed</option>
                                        <option value="Delayed">Delayed</option>
                                        <option value="Temporarily Suspended">Temporarily Suspended</option>
                                    </select>
                                </div>
                                <div>
                                    <label class="block font-medium text-slate-700 mb-1">Physical Progress Completion (%)</label>
                                    <input type="number" id="prog-percent" min="0" max="100" step="0.1" required class="w-full border border-slate-300 rounded-lg p-2">
                                </div>
                                <div>
                                    <label class="block font-medium text-slate-700 mb-1">Beneficiaries Covered so far</label>
                                    <input type="number" id="prog-beneficiaries-covered" min="0" class="w-full border border-slate-300 rounded-lg p-2">
                                </div>
                                <div class="sm:col-span-2">
                                    <label class="block font-medium text-slate-700 mb-1">Remarks & Milestone Description</label>
                                    <textarea id="prog-remarks" rows="2" class="w-full border border-slate-300 rounded-lg p-2" placeholder="Record physical site observations or milestone details..."></textarea>
                                </div>
                                <div class="sm:col-span-2 flex justify-end">
                                    <button type="submit" class="bg-purple-700 hover:bg-purple-800 text-white px-4 py-2 rounded-lg font-medium shadow-sm">
                                        Log Progress Update
                                    </button>
                                </div>
                            </form>
                        </div>

                        <!-- Physical Progress Log History -->
                        <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                            <h3 class="text-sm font-bold text-slate-800 mb-3">Historical Progress Update Log</h3>
                            <div class="overflow-x-auto">
                                <table class="w-full text-xs text-left text-slate-600 border-collapse">
                                    <thead class="bg-slate-100 uppercase text-[10px] text-slate-700 font-semibold border-b border-slate-200">
                                        <tr>
                                            <th class="py-2 px-3">Log Date</th>
                                            <th class="py-2 px-3">Stage</th>
                                            <th class="py-2 px-3 text-center">Progress %</th>
                                            <th class="py-2 px-3 text-center">Beneficiaries</th>
                                            <th class="py-2 px-3">Remarks</th>
                                        </tr>
                                    </thead>
                                    <tbody id="prog-history-tbody" class="divide-y divide-slate-200">
                                        <tr><td colspan="5" class="text-center py-4 text-slate-400">No logs found</td></tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                </div>
            </section>

            <!-- MODULE 5: BENEFICIARY REGISTER -->
            <section id="module-beneficiary-register" class="hidden space-y-6">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
                    <div>
                        <h2 class="text-lg font-bold text-slate-800">Beneficiary Register</h2>
                        <p class="text-xs text-slate-500">Individual, Group, and Community Beneficiary Tracking & Grant Distribution</p>
                    </div>
                    <div class="flex items-center space-x-2">
                        <button onclick="exportBeneficiariesExcel()" class="bg-emerald-700 hover:bg-emerald-800 text-white px-3 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1.5 shadow-sm">
                            <i class="fa-solid fa-file-excel"></i> Export Excel
                        </button>
                        <button onclick="openBeneficiaryModal()" class="bg-blue-700 hover:bg-blue-800 text-white px-3 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1.5 shadow-sm">
                            <i class="fa-solid fa-user-plus"></i> Add Beneficiary Record
                        </button>
                    </div>
                </div>

                <div class="bg-white p-3.5 rounded-xl shadow-sm border border-slate-200 grid grid-cols-1 sm:grid-cols-4 gap-3 text-xs">
                    <select id="ben-filter-project" onchange="renderBeneficiaryTable()" class="border border-slate-300 rounded-lg px-2.5 py-2 bg-slate-50">
                        <option value="">All Projects</option>
                    </select>
                    <select id="ben-filter-category" onchange="renderBeneficiaryTable()" class="border border-slate-300 rounded-lg px-2.5 py-2 bg-slate-50">
                        <option value="">All Categories (General/SC/ST)</option>
                        <option value="General">General</option>
                        <option value="SC">SC (Scheduled Caste)</option>
                        <option value="ST">ST (Scheduled Tribe)</option>
                        <option value="OBC/Other">OBC / Other</option>
                    </select>
                    <select id="ben-filter-ward" onchange="renderBeneficiaryTable()" class="border border-slate-300 rounded-lg px-2.5 py-2 bg-slate-50">
                        <option value="">All Wards</option>
                    </select>
                    <input type="text" id="ben-search" onkeyup="renderBeneficiaryTable()" placeholder="Search beneficiary name..." class="border border-slate-300 rounded-lg px-3 py-2 outline-none">
                </div>

                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                    <div class="overflow-x-auto custom-scrollbar">
                        <table class="w-full text-left text-xs text-slate-600 border-collapse">
                            <thead class="bg-slate-100 uppercase text-[10px] text-slate-700 font-semibold border-b border-slate-200">
                                <tr>
                                    <th class="py-2.5 px-3">Project ID / Name</th>
                                    <th class="py-2.5 px-3">Beneficiary Name</th>
                                    <th class="py-2.5 px-3">Ward & Address</th>
                                    <th class="py-2.5 px-3">Category / Gender</th>
                                    <th class="py-2.5 px-3">Benefit Sanctioned</th>
                                    <th class="py-2.5 px-3 text-center">Sanction Date</th>
                                    <th class="py-2.5 px-3 text-center">Status</th>
                                    <th class="py-2.5 px-3 text-center">Action</th>
                                </tr>
                            </thead>
                            <tbody id="beneficiary-tbody" class="divide-y divide-slate-200">
                                <!-- Dynamic rows -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </section>

            <!-- MODULE 6: FIELD INSPECTION & MONITORING -->
            <section id="module-inspection-register" class="hidden space-y-6">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
                    <div>
                        <h2 class="text-lg font-bold text-slate-800">Field Inspection & Technical Monitoring Log</h2>
                        <p class="text-xs text-slate-500">Record Site Visits, Quality Verification & Corrective Action Directives</p>
                    </div>
                    <button onclick="openInspectionModal()" class="bg-cyan-700 hover:bg-cyan-800 text-white px-3.5 py-2 rounded-lg text-xs font-medium flex items-center gap-1.5 shadow-sm">
                        <i class="fa-solid fa-clipboard-check"></i> Record New Field Inspection
                    </button>
                </div>

                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                    <div class="overflow-x-auto">
                        <table class="w-full text-xs text-left text-slate-600 border-collapse">
                            <thead class="bg-slate-100 uppercase text-[10px] text-slate-700 font-semibold border-b border-slate-200">
                                <tr>
                                    <th class="py-2.5 px-3">Inspection Date</th>
                                    <th class="py-2.5 px-3">Project Code & Name</th>
                                    <th class="py-2.5 px-3">Inspecting Officer</th>
                                    <th class="py-2.5 px-3 text-center">Observed Progress</th>
                                    <th class="py-2.5 px-3 text-center">Quality Grade</th>
                                    <th class="py-2.5 px-3">Inspection Result</th>
                                    <th class="py-2.5 px-3">Corrective Actions Required</th>
                                    <th class="py-2.5 px-3 text-center">Compliance</th>
                                </tr>
                            </thead>
                            <tbody id="inspection-tbody" class="divide-y divide-slate-200">
                                <!-- Dynamic rows -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </section>

            <!-- MODULE 7: PROJECT EVALUATION MODULE -->
            <section id="module-evaluation-module" class="hidden space-y-6">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
                    <div>
                        <h2 class="text-lg font-bold text-slate-800">14th Five-Year Plan Project Evaluation System</h2>
                        <p class="text-xs text-slate-500">Structured Performance Scoring (0-5 Criteria), Weighted Evaluation & Performance Grade Assignment</p>
                    </div>
                </div>

                <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    <!-- Project Selector & Grade Card -->
                    <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 space-y-4">
                        <label class="block text-xs font-bold text-slate-700 uppercase">Select Project for Evaluation</label>
                        <select id="eval-project-select" onchange="loadProjectEvaluation(this.value)" class="w-full border border-slate-300 rounded-lg p-2 text-xs bg-slate-50">
                        </select>

                        <div id="eval-score-card" class="bg-slate-50 p-4 rounded-lg border border-slate-200 space-y-3 text-center">
                            <p class="text-xs text-slate-500">Select a project to evaluate performance</p>
                        </div>
                    </div>

                    <!-- Evaluation Scoring Criteria Form -->
                    <div class="lg:col-span-2 bg-white p-4 rounded-xl shadow-sm border border-slate-200 space-y-5">
                        <form id="evaluation-form" onsubmit="saveProjectEvaluation(event)" class="space-y-4 text-xs">
                            <div class="grid grid-cols-1 sm:grid-cols-3 gap-3 border-b border-slate-200 pb-3">
                                <div>
                                    <label class="block font-medium text-slate-700 mb-1">Evaluation Type</label>
                                    <select id="eval-type" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50">
                                        <option value="Mid-Term Evaluation">Mid-Term Evaluation</option>
                                        <option value="Final Evaluation">Final Evaluation</option>
                                        <option value="Technical Evaluation">Technical Evaluation</option>
                                        <option value="Beneficiary Satisfaction Evaluation">Beneficiary Satisfaction Evaluation</option>
                                        <option value="Post-Completion Impact Evaluation">Post-Completion Impact Evaluation</option>
                                    </select>
                                </div>
                                <div>
                                    <label class="block font-medium text-slate-700 mb-1">Evaluator Name</label>
                                    <input type="text" id="eval-evaluator-name" required class="w-full border border-slate-300 rounded-lg p-2" placeholder="e.g. Planning Committee Chairman">
                                </div>
                                <div>
                                    <label class="block font-medium text-slate-700 mb-1">Evaluation Date</label>
                                    <input type="date" id="eval-date" required class="w-full border border-slate-300 rounded-lg p-2">
                                </div>
                            </div>

                            <h3 class="text-xs font-bold text-slate-800 uppercase tracking-wide">Scoring Criteria (0 = Not Assessed, 1 = Very Poor ... 5 = Excellent)</h3>
                            
                            <div class="grid grid-cols-1 md:grid-cols-2 gap-3" id="eval-criteria-container">
                                <!-- Criteria rows generated dynamically -->
                            </div>

                            <div class="space-y-3 pt-3 border-t border-slate-200">
                                <div>
                                    <label class="block font-medium text-slate-700 mb-1">Major Achievements & Output Findings</label>
                                    <textarea id="eval-achievements" rows="2" class="w-full border border-slate-300 rounded-lg p-2" placeholder="Document verified physical and social outcomes..."></textarea>
                                </div>
                                <div>
                                    <label class="block font-medium text-slate-700 mb-1">Problems Identified / Reasons for Delay (if any)</label>
                                    <textarea id="eval-problems" rows="2" class="w-full border border-slate-300 rounded-lg p-2" placeholder="Document issues encountered during implementation..."></textarea>
                                </div>
                            </div>

                            <div class="flex justify-between items-center pt-3 border-t border-slate-200">
                                <span class="text-xs text-slate-500 font-medium" id="eval-live-summary">Total Score: 0/0 (0%)</span>
                                <button type="submit" class="bg-amber-600 hover:bg-amber-700 text-white px-5 py-2 rounded-lg font-medium shadow-sm">
                                    <i class="fa-solid fa-save mr-1"></i> Save Evaluation Record
                                </button>
                            </div>
                        </form>
                    </div>
                </div>
            </section>

            <!-- MODULE 8: OUTPUT, OUTCOME & IMPACT REGISTER -->
            <section id="module-outcome-register" class="hidden space-y-6">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
                    <div>
                        <h2 class="text-lg font-bold text-slate-800">Output, Outcome & Impact Register</h2>
                        <p class="text-xs text-slate-500">Quantitative Target vs Achievement Tracking & Community Impact Verification</p>
                    </div>
                    <button onclick="openOutcomeModal()" class="bg-rose-700 hover:bg-rose-800 text-white px-3.5 py-2 rounded-lg text-xs font-medium flex items-center gap-1 shadow-sm">
                        <i class="fa-solid fa-plus"></i> Add Indicator Record
                    </button>
                </div>

                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                    <div class="overflow-x-auto custom-scrollbar">
                        <table class="w-full text-left text-xs text-slate-600 border-collapse">
                            <thead class="bg-slate-100 uppercase text-[10px] text-slate-700 font-semibold border-b border-slate-200">
                                <tr>
                                    <th class="py-2.5 px-3">Project ID & Name</th>
                                    <th class="py-2.5 px-3">Indicator Name & Unit</th>
                                    <th class="py-2.5 px-3 text-right">Baseline</th>
                                    <th class="py-2.5 px-3 text-right">Target</th>
                                    <th class="py-2.5 px-3 text-right">Achieved</th>
                                    <th class="py-2.5 px-3 text-center">Achievement %</th>
                                    <th class="py-2.5 px-3">Direct / Indirect Beneficiaries</th>
                                    <th class="py-2.5 px-3 text-center">Sustainability</th>
                                </tr>
                            </thead>
                            <tbody id="outcome-tbody" class="divide-y divide-slate-200">
                                <!-- Dynamic rows -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </section>

            <!-- MODULE 9: PROJECT REVIEW & CORRECTIVE ACTION -->
            <section id="module-corrective-action" class="hidden space-y-6">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
                    <div>
                        <h2 class="text-lg font-bold text-slate-800 text-red-700 flex items-center gap-2">
                            <i class="fa-solid fa-triangle-exclamation"></i> Project Review & Corrective Action Register
                        </h2>
                        <p class="text-xs text-slate-500">System-generated flags for delayed projects, budget overruns, overdue evaluations & inspection non-compliance</p>
                    </div>
                </div>

                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                    <div class="overflow-x-auto">
                        <table class="w-full text-xs text-left text-slate-600 border-collapse">
                            <thead class="bg-slate-100 uppercase text-[10px] text-slate-700 font-semibold border-b border-slate-200">
                                <tr>
                                    <th class="py-2.5 px-3">Project Code / Name</th>
                                    <th class="py-2.5 px-3">Issue Flagged</th>
                                    <th class="py-2.5 px-3 text-center">Severity</th>
                                    <th class="py-2.5 px-3">Recommended Corrective Action</th>
                                    <th class="py-2.5 px-3">Responsible Officer</th>
                                    <th class="py-2.5 px-3 text-center">Deadline</th>
                                    <th class="py-2.5 px-3 text-center">Resolution Status</th>
                                    <th class="py-2.5 px-3 text-center">Action</th>
                                </tr>
                            </thead>
                            <tbody id="corrective-tbody" class="divide-y divide-slate-200">
                                <!-- Dynamic rows -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </section>

            <!-- MODULE 10: DOCUMENT MANAGEMENT -->
            <section id="module-documents-register" class="hidden space-y-6">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
                    <div>
                        <h2 class="text-lg font-bold text-slate-800">Project Document Vault</h2>
                        <p class="text-xs text-slate-500">Administrative Sanction (AS), Technical Sanction (TS), Estimates, Vouchers & Certificates</p>
                    </div>
                    <button onclick="openDocumentModal()" class="bg-indigo-700 hover:bg-indigo-800 text-white px-3.5 py-2 rounded-lg text-xs font-medium flex items-center gap-1.5 shadow-sm">
                        <i class="fa-solid fa-file-upload"></i> Register New Document
                    </button>
                </div>

                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                    <div class="overflow-x-auto">
                        <table class="w-full text-xs text-left text-slate-600 border-collapse">
                            <thead class="bg-slate-100 uppercase text-[10px] text-slate-700 font-semibold border-b border-slate-200">
                                <tr>
                                    <th class="py-2.5 px-3">Project Code / Name</th>
                                    <th class="py-2.5 px-3">Document Type</th>
                                    <th class="py-2.5 px-3">Doc Ref / Sanction No</th>
                                    <th class="py-2.5 px-3 text-center">Doc Date</th>
                                    <th class="py-2.5 px-3">Uploaded By</th>
                                    <th class="py-2.5 px-3 text-center">Verification Status</th>
                                    <th class="py-2.5 px-3 text-center">Link / Reference</th>
                                </tr>
                            </thead>
                            <tbody id="documents-tbody" class="divide-y divide-slate-200">
                                <!-- Dynamic rows -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </section>

            <!-- MODULE 11: REPORTS AND DOWNLOADS -->
            <section id="module-reports-module" class="hidden space-y-6">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
                    <div>
                        <h2 class="text-lg font-bold text-slate-800">Consolidated Panchayat Reports Engine</h2>
                        <p class="text-xs text-slate-500">23 Formal Panchayat Level Reports & Exportable PDF / Excel Statements</p>
                    </div>
                </div>

                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 space-y-4">
                    <div class="grid grid-cols-1 md:grid-cols-3 gap-4 text-xs">
                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Select Report Type (1-23)</label>
                            <select id="report-type-select" class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50 font-medium">
                                <option value="1">1. Complete 14th Five-Year Plan Project Register</option>
                                <option value="2">2. Project-wise Detailed Performance Report</option>
                                <option value="3">3. Ward-wise Project Distribution Report</option>
                                <option value="4">4. Sector-wise Project & Expenditure Report</option>
                                <option value="5">5. Implementing Officer-wise Progress Report</option>
                                <option value="6">6. Financial Year-wise Project Allocations</option>
                                <option value="7">7. Approved Projects Master List</option>
                                <option value="8">8. Pending & Not-Started Projects Report</option>
                                <option value="9">9. Delayed Projects Monitoring Statement</option>
                                <option value="10">10. Completed Projects Register</option>
                                <option value="11">11. Physical Progress Verification Summary</option>
                                <option value="12">12. Financial Progress & Voucher Expenditure Report</option>
                                <option value="13">13. Allocation versus Expenditure Comparison</option>
                                <option value="14">14. Beneficiary Coverage & Category Distribution</option>
                                <option value="15">15. Field Inspection Audit Report</option>
                                <option value="16">16. Project Evaluation Scores & Grades Summary</option>
                                <option value="17">17. Projects Requiring Corrective Action List</option>
                                <option value="18">18. Output, Outcome & Impact Performance Report</option>
                                <option value="19">19. Evaluation Committee Consolidated Statement</option>
                                <option value="20">20. Annual Plan Consolidated Review</option>
                                <option value="21">21. Five-Year Plan Consolidated Evaluation Report</option>
                                <option value="22">22. Project Completion & Closure Verification</option>
                                <option value="23">23. Audit & Utilization Certificate Pending Statement</option>
                            </select>
                        </div>

                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Filter Sector</label>
                            <select id="report-filter-sector" class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50">
                                <option value="ALL">All Sectors</option>
                            </select>
                        </div>

                        <div>
                            <label class="block font-bold text-slate-700 mb-1">Filter Financial Year</label>
                            <select id="report-filter-fy" class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50">
                                <option value="ALL">All Financial Years</option>
                            </select>
                        </div>
                    </div>

                    <div class="flex flex-wrap items-center justify-between pt-3 border-t border-slate-200 gap-3">
                        <button onclick="generateReportPreview()" class="bg-blue-700 hover:bg-blue-800 text-white px-4 py-2 rounded-lg text-xs font-semibold flex items-center gap-1.5 shadow-sm">
                            <i class="fa-solid fa-eye"></i> Generate Report Preview
                        </button>
                        <div class="flex items-center space-x-2">
                            <button onclick="downloadReportExcel()" class="bg-emerald-700 hover:bg-emerald-800 text-white px-3.5 py-2 rounded-lg text-xs font-medium flex items-center gap-1.5 shadow-sm">
                                <i class="fa-solid fa-file-excel"></i> Export Excel (.xlsx)
                            </button>
                            <button onclick="downloadReportPDF()" class="bg-red-700 hover:bg-red-800 text-white px-3.5 py-2 rounded-lg text-xs font-medium flex items-center gap-1.5 shadow-sm">
                                <i class="fa-solid fa-file-pdf"></i> Download Official PDF
                            </button>
                        </div>
                    </div>
                </div>

                <!-- FORMAL PREVIEW CONTAINER -->
                <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-200 space-y-4" id="report-preview-container">
                    <div class="border-b-2 border-slate-800 pb-3 text-center space-y-1">
                        <h3 class="text-xs font-bold uppercase tracking-widest text-slate-500">Government of Kerala - Local Self Government Department</h3>
                        <h2 class="text-base font-extrabold uppercase text-slate-900" id="rpt-header-panchayat">KARIVELLUR-PERALAM GRAMA PANCHAYAT</h2>
                        <h1 class="text-sm font-bold uppercase text-blue-900" id="rpt-title-display">FOURTEENTH FIVE-YEAR PLAN – PROJECT REGISTER AND EVALUATION REPORT</h1>
                        <p class="text-[11px] text-slate-600" id="rpt-meta-info">Generated on: Oct 27, 2026 | Period: 14th Five-Year Plan (2022-2027)</p>
                    </div>

                    <div class="overflow-x-auto">
                        <table class="w-full text-xs text-left text-slate-700 border-collapse border border-slate-300">
                            <thead class="bg-slate-100 uppercase text-[10px] text-slate-800 font-bold border-b border-slate-300" id="report-table-head">
                                <!-- Dynamic headers -->
                            </thead>
                            <tbody id="report-table-body" class="divide-y divide-slate-200">
                                <tr><td colspan="6" class="text-center py-6 text-slate-400">Click "Generate Report Preview" to view report table</td></tr>
                            </tbody>
                        </table>
                    </div>

                    <div class="pt-8 flex justify-between items-end text-xs text-slate-600">
                        <div>
                            <p class="font-semibold" id="rpt-signatory-1">Prepared by: Planning Coordinator</p>
                            <p class="text-[10px] text-slate-400">Grama Panchayat Office</p>
                        </div>
                        <div class="text-right">
                            <p class="font-semibold" id="rpt-signatory-2">Approved by: Secretary / President</p>
                            <p class="text-[10px] text-slate-400">Grama Panchayat Committee</p>
                        </div>
                    </div>
                </div>
            </section>

            <!-- MODULE 13: MASTER SETTINGS -->
            <section id="module-master-settings" class="hidden space-y-6">
                <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex flex-wrap justify-between items-center gap-4">
                    <div>
                        <h2 class="text-lg font-bold text-slate-800">Master Settings & Administrative Configuration</h2>
                        <p class="text-xs text-slate-500">Configure Panchayat Details, Dropdown Masters, Evaluation Weights & System Backups</p>
                    </div>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
                    <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 space-y-4">
                        <h3 class="text-sm font-bold text-slate-800 border-b border-slate-200 pb-2">Panchayat & Planning Profile</h3>
                        <form id="panchayat-profile-form" onsubmit="savePanchayatProfile(event)" class="space-y-3 text-xs">
                            <div>
                                <label class="block font-medium text-slate-700 mb-1">Grama Panchayat Name</label>
                                <input type="text" id="cfg-panchayat-name" required class="w-full border border-slate-300 rounded-lg p-2 font-medium">
                            </div>
                            <div class="grid grid-cols-2 gap-3">
                                <div>
                                    <label class="block font-medium text-slate-700 mb-1">Block Panchayat</label>
                                    <input type="text" id="cfg-block-name" required class="w-full border border-slate-300 rounded-lg p-2">
                                </div>
                                <div>
                                    <label class="block font-medium text-slate-700 mb-1">District</label>
                                    <input type="text" id="cfg-district-name" required class="w-full border border-slate-300 rounded-lg p-2">
                                </div>
                            </div>
                            <div class="grid grid-cols-2 gap-3">
                                <div>
                                    <label class="block font-medium text-slate-700 mb-1">Authorized Signatory 1</label>
                                    <input type="text" id="cfg-sig1" class="w-full border border-slate-300 rounded-lg p-2">
                                </div>
                                <div>
                                    <label class="block font-medium text-slate-700 mb-1">Authorized Signatory 2</label>
                                    <input type="text" id="cfg-sig2" class="w-full border border-slate-300 rounded-lg p-2">
                                </div>
                            </div>
                            <div class="flex justify-end pt-2">
                                <button type="submit" class="bg-blue-700 hover:bg-blue-800 text-white px-4 py-2 rounded-lg font-medium shadow-sm">
                                    Save Profile Settings
                                </button>
                            </div>
                        </form>
                    </div>

                    <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 space-y-4">
                        <h3 class="text-sm font-bold text-slate-800 border-b border-slate-200 pb-2">Database Backup & Audit Log</h3>
                        <div class="space-y-3 text-xs">
                            <p class="text-slate-600">Download system database in JSON format or restore from file.</p>
                            <div class="flex items-center space-x-2">
                                <button onclick="exportDatabaseJSON()" class="bg-slate-800 hover:bg-slate-900 text-white px-3.5 py-2 rounded-lg font-medium flex items-center gap-1.5 shadow-sm">
                                    <i class="fa-solid fa-download"></i> Backup JSON DB
                                </button>
                                <input type="file" id="import-json-file" class="hidden" onchange="importDatabaseJSON(event)" accept=".json">
                                <button onclick="document.getElementById('import-json-file').click()" class="bg-amber-600 hover:bg-amber-700 text-white px-3.5 py-2 rounded-lg font-medium flex items-center gap-1.5 shadow-sm">
                                    <i class="fa-solid fa-upload"></i> Restore JSON DB
                                </button>
                            </div>
                            
                            <div class="pt-4 border-t border-slate-200">
                                <h4 class="font-bold text-slate-700 mb-2">Audit Action Log (Recent)</h4>
                                <div class="bg-slate-900 text-slate-300 p-3 rounded-lg text-[10px] font-mono h-36 overflow-y-auto custom-scrollbar" id="audit-log-box">
                                    <!-- Audit logs -->
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </section>

        </main>
    </div>

    <!-- MODAL: REGISTER / EDIT PROJECT -->
    <div id="modal-project" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-xl max-w-4xl w-full max-h-[90vh] flex flex-col overflow-hidden">
            <div class="p-4 bg-blue-900 text-white flex justify-between items-center">
                <h3 class="font-bold text-sm" id="modal-project-title">Register New Panchayat Project (14th FYP)</h3>
                <button onclick="closeProjectModal()" class="text-blue-200 hover:text-white"><i class="fa-solid fa-xmark text-lg"></i></button>
            </div>
            
            <form id="project-form" onsubmit="saveProjectRecord(event)" class="p-5 overflow-y-auto custom-scrollbar space-y-4 text-xs flex-1">
                <input type="hidden" id="prj-id">

                <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div>
                        <label class="block font-semibold text-slate-700 mb-1">Project Code / ID</label>
                        <input type="text" id="prj-code" readonly class="w-full bg-slate-100 border border-slate-300 rounded-lg p-2 font-mono font-bold text-blue-900">
                    </div>
                    <div>
                        <label class="block font-semibold text-slate-700 mb-1">Financial Year *</label>
                        <select id="prj-fy" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50">
                            <option value="2022-2023">2022-2023 (Year 1)</option>
                            <option value="2023-2024">2023-2024 (Year 2)</option>
                            <option value="2024-2025" selected>2024-2025 (Year 3)</option>
                            <option value="2025-2026">2025-2026 (Year 4)</option>
                            <option value="2026-2027">2026-2027 (Year 5)</option>
                        </select>
                    </div>
                    <div>
                        <label class="block font-semibold text-slate-700 mb-1">Ward Number *</label>
                        <select id="prj-ward" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50"></select>
                    </div>
                </div>

                <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div class="sm:col-span-2">
                        <label class="block font-semibold text-slate-700 mb-1">Project Name (in English / Malayalam) *</label>
                        <input type="text" id="prj-name" required class="w-full border border-slate-300 rounded-lg p-2 font-medium" placeholder="e.g., Construction of Drinking Water Storage Tank in Ward 4">
                    </div>
                    <div>
                        <label class="block font-semibold text-slate-700 mb-1">Development Sector *</label>
                        <select id="prj-sector" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50"></select>
                    </div>
                    <div>
                        <label class="block font-semibold text-slate-700 mb-1">Project Category / Type *</label>
                        <select id="prj-category" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50">
                            <option value="General Category">General Category</option>
                            <option value="SCSP (Special Component)">SCSP (Scheduled Caste Component)</option>
                            <option value="TSP (Tribal Sub Plan)">TSP (Tribal Sub Plan)</option>
                            <option value="Women Development">Women & Child Development</option>
                        </select>
                    </div>
                </div>

                <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div>
                        <label class="block font-semibold text-slate-700 mb-1">Implementing Department *</label>
                        <select id="prj-dept" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50"></select>
                    </div>
                    <div>
                        <label class="block font-semibold text-slate-700 mb-1">Implementing Officer Name *</label>
                        <select id="prj-officer" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50"></select>
                    </div>
                    <div>
                        <label class="block font-semibold text-slate-700 mb-1">Beneficiary Type *</label>
                        <select id="prj-ben-type" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50">
                            <option value="Community Project">Community Project</option>
                            <option value="Individual Beneficiary">Individual Beneficiary</option>
                            <option value="Group Beneficiary">Group Beneficiary</option>
                        </select>
                    </div>
                </div>

                <!-- Financial Allocations -->
                <div class="bg-slate-50 p-3 rounded-xl border border-slate-200 space-y-3">
                    <h4 class="font-bold text-slate-800 uppercase tracking-wide text-[11px]">Financial Sanction & Allocation Breakdown (₹)</h4>
                    <div class="grid grid-cols-2 sm:grid-cols-4 gap-3">
                        <div>
                            <label class="block text-slate-600 mb-1">Approved Project Cost (₹) *</label>
                            <input type="number" id="prj-cost" required min="0" step="1" class="w-full border border-slate-300 rounded-lg p-2 font-bold text-slate-900">
                        </div>
                        <div>
                            <label class="block text-slate-600 mb-1">General Dev Fund (₹)</label>
                            <input type="number" id="fund-gen" min="0" value="0" class="w-full border border-slate-300 rounded-lg p-2">
                        </div>
                        <div>
                            <label class="block text-slate-600 mb-1">SCSP Fund (₹)</label>
                            <input type="number" id="fund-scsp" min="0" value="0" class="w-full border border-slate-300 rounded-lg p-2">
                        </div>
                        <div>
                            <label class="block text-slate-600 mb-1">TSP Fund (₹)</label>
                            <input type="number" id="fund-tsp" min="0" value="0" class="w-full border border-slate-300 rounded-lg p-2">
                        </div>
                    </div>
                </div>

                <!-- Administrative & Technical Sanction -->
                <div class="grid grid-cols-1 sm:grid-cols-4 gap-3">
                    <div>
                        <label class="block font-semibold text-slate-700 mb-1">AS Number</label>
                        <input type="text" id="prj-as-no" class="w-full border border-slate-300 rounded-lg p-2">
                    </div>
                    <div>
                        <label class="block font-semibold text-slate-700 mb-1">AS Date</label>
                        <input type="date" id="prj-as-date" class="w-full border border-slate-300 rounded-lg p-2">
                    </div>
                    <div>
                        <label class="block font-semibold text-slate-700 mb-1">TS Number</label>
                        <input type="text" id="prj-ts-no" class="w-full border border-slate-300 rounded-lg p-2">
                    </div>
                    <div>
                        <label class="block font-semibold text-slate-700 mb-1">Target Completion Date</label>
                        <input type="date" id="prj-target-date" class="w-full border border-slate-300 rounded-lg p-2">
                    </div>
                </div>

                <div>
                    <label class="block font-semibold text-slate-700 mb-1">Project Objective & Description</label>
                    <textarea id="prj-desc" rows="2" class="w-full border border-slate-300 rounded-lg p-2" placeholder="Brief statement of project scope and community benefits..."></textarea>
                </div>

                <div class="flex justify-end space-x-2 pt-3 border-t border-slate-200">
                    <button type="button" onclick="closeProjectModal()" class="px-4 py-2 border border-slate-300 rounded-lg text-slate-600 font-medium">Cancel</button>
                    <button type="submit" class="px-5 py-2 bg-blue-700 hover:bg-blue-800 text-white rounded-lg font-medium shadow-sm">Save Project Master Record</button>
                </div>
            </form>
        </div>
    </div>

    <!-- MODAL: ADD EXPENDITURE VOUCHER -->
    <div id="modal-expenditure" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-xl max-w-md w-full overflow-hidden">
            <div class="p-4 bg-emerald-800 text-white flex justify-between items-center">
                <h3 class="font-bold text-sm">Log Voucher Expenditure</h3>
                <button onclick="closeExpenditureModal()" class="text-emerald-200 hover:text-white"><i class="fa-solid fa-xmark text-lg"></i></button>
            </div>
            
            <form id="expenditure-form" onsubmit="saveExpenditureVoucher(event)" class="p-4 space-y-3 text-xs">
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Voucher Date *</label>
                    <input type="date" id="exp-date" required class="w-full border border-slate-300 rounded-lg p-2">
                </div>
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Voucher / Bill Number *</label>
                    <input type="text" id="exp-voucher-no" required class="w-full border border-slate-300 rounded-lg p-2" placeholder="e.g. VCH-2024-089">
                </div>
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Expenditure Amount (₹) *</label>
                    <input type="number" id="exp-amount" required min="1" step="0.01" class="w-full border border-slate-300 rounded-lg p-2 font-bold text-slate-900">
                </div>
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Head of Account / Source</label>
                    <select id="exp-head" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50">
                        <option value="General Development Fund">General Development Fund</option>
                        <option value="SCSP Fund">SCSP Fund</option>
                        <option value="TSP Fund">TSP Fund</option>
                        <option value="Maintenance Fund Road">Maintenance Fund Road</option>
                        <option value="Own Fund">Own Fund</option>
                    </select>
                </div>
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Purpose / Beneficiary Payment Remarks</label>
                    <textarea id="exp-remarks" rows="2" class="w-full border border-slate-300 rounded-lg p-2" placeholder="e.g., First installment payment to contractor / material purchase"></textarea>
                </div>
                <div class="flex justify-end space-x-2 pt-2">
                    <button type="button" onclick="closeExpenditureModal()" class="px-3 py-1.5 border border-slate-300 rounded-lg text-slate-600">Cancel</button>
                    <button type="submit" class="px-4 py-1.5 bg-emerald-700 hover:bg-emerald-800 text-white rounded-lg font-medium shadow-sm">Record Transaction</button>
                </div>
            </form>
        </div>
    </div>

    <!-- MODAL: BENEFICIARY ENTRY -->
    <div id="modal-beneficiary" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-xl max-w-lg w-full overflow-hidden">
            <div class="p-4 bg-teal-800 text-white flex justify-between items-center">
                <h3 class="font-bold text-sm">Add Individual Beneficiary Record</h3>
                <button onclick="closeBeneficiaryModal()" class="text-teal-200 hover:text-white"><i class="fa-solid fa-xmark text-lg"></i></button>
            </div>
            
            <form id="beneficiary-form" onsubmit="saveBeneficiaryRecord(event)" class="p-4 space-y-3 text-xs">
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Select Linked Project *</label>
                    <select id="ben-project-id" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50"></select>
                </div>
                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="block font-medium text-slate-700 mb-1">Beneficiary Name *</label>
                        <input type="text" id="ben-name" required class="w-full border border-slate-300 rounded-lg p-2">
                    </div>
                    <div>
                        <label class="block font-medium text-slate-700 mb-1">Ward Number *</label>
                        <select id="ben-ward" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50"></select>
                    </div>
                </div>
                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="block font-medium text-slate-700 mb-1">Category *</label>
                        <select id="ben-category" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50">
                            <option value="General">General</option>
                            <option value="SC">SC</option>
                            <option value="ST">ST</option>
                            <option value="OBC/Other">OBC/Other</option>
                        </select>
                    </div>
                    <div>
                        <label class="block font-medium text-slate-700 mb-1">Gender</label>
                        <select id="ben-gender" class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50">
                            <option value="Female">Female</option>
                            <option value="Male">Male</option>
                            <option value="Transgender">Transgender</option>
                        </select>
                    </div>
                </div>
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Sanctioned Benefit / Grant Details</label>
                    <input type="text" id="ben-sanctioned-details" placeholder="e.g. ₹50,000 Housing Grant 1st Tranche" class="w-full border border-slate-300 rounded-lg p-2">
                </div>
                <div class="flex justify-end space-x-2 pt-2">
                    <button type="button" onclick="closeBeneficiaryModal()" class="px-3 py-1.5 border border-slate-300 rounded-lg">Cancel</button>
                    <button type="submit" class="px-4 py-1.5 bg-teal-700 hover:bg-teal-800 text-white rounded-lg font-medium">Save Beneficiary</button>
                </div>
            </form>
        </div>
    </div>

    <!-- MODAL: FIELD INSPECTION -->
    <div id="modal-inspection" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-xl max-w-lg w-full overflow-hidden">
            <div class="p-4 bg-cyan-900 text-white flex justify-between items-center">
                <h3 class="font-bold text-sm">Record Field Inspection Findings</h3>
                <button onclick="closeInspectionModal()" class="text-cyan-200 hover:text-white"><i class="fa-solid fa-xmark text-lg"></i></button>
            </div>
            
            <form id="inspection-form" onsubmit="saveInspectionRecord(event)" class="p-4 space-y-3 text-xs">
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Project Code & Name *</label>
                    <select id="insp-project-id" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50"></select>
                </div>
                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="block font-medium text-slate-700 mb-1">Inspection Date *</label>
                        <input type="date" id="insp-date" required class="w-full border border-slate-300 rounded-lg p-2">
                    </div>
                    <div>
                        <label class="block font-medium text-slate-700 mb-1">Inspecting Officer *</label>
                        <input type="text" id="insp-officer" required placeholder="e.g. Assistant Engineer" class="w-full border border-slate-300 rounded-lg p-2">
                    </div>
                </div>
                <div class="grid grid-cols-2 gap-3">
                    <div>
                        <label class="block font-medium text-slate-700 mb-1">Observed Physical %</label>
                        <input type="number" id="insp-prog-obs" min="0" max="100" class="w-full border border-slate-300 rounded-lg p-2">
                    </div>
                    <div>
                        <label class="block font-medium text-slate-700 mb-1">Quality of Work Grade</label>
                        <select id="insp-quality" class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50">
                            <option value="Satisfactory / Good">Satisfactory / Good</option>
                            <option value="Excellent">Excellent</option>
                            <option value="Unsatisfactory / Substandard">Unsatisfactory / Substandard</option>
                        </select>
                    </div>
                </div>
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Defects / Corrective Directive Issued</label>
                    <textarea id="insp-defects" rows="2" class="w-full border border-slate-300 rounded-lg p-2" placeholder="Detail any quality deficiencies or instructions given..."></textarea>
                </div>
                <div class="flex justify-end space-x-2 pt-2">
                    <button type="button" onclick="closeInspectionModal()" class="px-3 py-1.5 border border-slate-300 rounded-lg">Cancel</button>
                    <button type="submit" class="px-4 py-1.5 bg-cyan-700 hover:bg-cyan-800 text-white rounded-lg font-medium">Save Inspection</button>
                </div>
            </form>
        </div>
    </div>

    <!-- MODAL: DOCUMENT UPLOAD -->
    <div id="modal-document" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-xl max-w-md w-full overflow-hidden">
            <div class="p-4 bg-indigo-900 text-white flex justify-between items-center">
                <h3 class="font-bold text-sm">Register Project Document Reference</h3>
                <button onclick="closeDocumentModal()" class="text-indigo-200 hover:text-white"><i class="fa-solid fa-xmark text-lg"></i></button>
            </div>
            
            <form id="document-form" onsubmit="saveDocumentRecord(event)" class="p-4 space-y-3 text-xs">
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Select Project *</label>
                    <select id="doc-project-id" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50"></select>
                </div>
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Document Type *</label>
                    <select id="doc-type" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50">
                        <option value="Administrative Sanction (AS)">Administrative Sanction (AS)</option>
                        <option value="Technical Sanction (TS)">Technical Sanction (TS)</option>
                        <option value="Detailed Estimate">Detailed Estimate</option>
                        <option value="Tender / Work Order">Tender / Work Order</option>
                        <option value="Bill / Voucher Proof">Bill / Voucher Proof</option>
                        <option value="Site Inspection Photo">Site Inspection Photo</option>
                        <option value="Completion Certificate">Completion Certificate</option>
                        <option value="Utilization Certificate (UC)">Utilization Certificate (UC)</option>
                    </select>
                </div>
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Sanction / Reference Number</label>
                    <input type="text" id="doc-ref-no" required placeholder="e.g. AS-402/2024" class="w-full border border-slate-300 rounded-lg p-2">
                </div>
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Document URL / Reference Link</label>
                    <input type="url" id="doc-url" placeholder="https://..." class="w-full border border-slate-300 rounded-lg p-2">
                </div>
                <div class="flex justify-end space-x-2 pt-2">
                    <button type="button" onclick="closeDocumentModal()" class="px-3 py-1.5 border border-slate-300 rounded-lg">Cancel</button>
                    <button type="submit" class="px-4 py-1.5 bg-indigo-700 hover:bg-indigo-800 text-white rounded-lg font-medium">Register Vault Record</button>
                </div>
            </form>
        </div>
    </div>

    <!-- MODAL: OUTCOME INDICATOR -->
    <div id="modal-outcome" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
        <div class="bg-white rounded-2xl shadow-xl max-w-md w-full overflow-hidden">
            <div class="p-4 bg-rose-900 text-white flex justify-between items-center">
                <h3 class="font-bold text-sm">Add Output / Outcome Indicator</h3>
                <button onclick="closeOutcomeModal()" class="text-rose-200 hover:text-white"><i class="fa-solid fa-xmark text-lg"></i></button>
            </div>
            
            <form id="outcome-form" onsubmit="saveOutcomeRecord(event)" class="p-4 space-y-3 text-xs">
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Select Project *</label>
                    <select id="out-project-id" required class="w-full border border-slate-300 rounded-lg p-2 bg-slate-50"></select>
                </div>
                <div>
                    <label class="block font-medium text-slate-700 mb-1">Indicator Name *</label>
                    <input type="text" id="out-name" required placeholder="e.g. Drinking water pipeline laid" class="w-full border border-slate-300 rounded-lg p-2">
                </div>
                <div class="grid grid-cols-3 gap-2">
                    <div>
                        <label class="block font-medium text-slate-700 mb-1">Unit</label>
                        <input type="text" id="out-unit" placeholder="Meters / Homes" class="w-full border border-slate-300 rounded-lg p-2">
                    </div>
                    <div>
                        <label class="block font-medium text-slate-700 mb-1">Target</label>
                        <input type="number" id="out-target" step="0.1" required class="w-full border border-slate-300 rounded-lg p-2">
                    </div>
                    <div>
                        <label class="block font-medium text-slate-700 mb-1">Achieved</label>
                        <input type="number" id="out-achieved" step="0.1" required class="w-full border border-slate-300 rounded-lg p-2">
                    </div>
                </div>
                <div class="flex justify-end space-x-2 pt-2">
                    <button type="button" onclick="closeOutcomeModal()" class="px-3 py-1.5 border border-slate-300 rounded-lg">Cancel</button>
                    <button type="submit" class="px-4 py-1.5 bg-rose-700 hover:bg-rose-800 text-white rounded-lg font-medium">Save Indicator</button>
                </div>
            </form>
        </div>
    </div>

    <!-- JAVASCRIPT APP APPLICATION ENGINE -->
    <script>
        // ==========================================
        // SYSTEM CONSTANTS & MASTER SECTORS DEFINITION
        // ==========================================
        const DEFAULT_SECTORS = [
            "Agriculture", "Animal Husbandry and Dairy Development", "Fisheries", 
            "Local Economic Development", "Poverty Eradication", "Scheduled Caste Development", 
            "Scheduled Tribe Development", "Women and Child Development", "Health", "Education", 
            "Drinking Water", "Sanitation and Waste Management", "Housing", "Roads and Transportation", 
            "Electricity and Energy", "Environment and Climate Action", "Social Security", 
            "Disability Inclusion", "Sports and Youth Affairs", "Culture", "Infrastructure Development", 
            "Disaster Management", "Good Governance", "Other"
        ];

        const DEFAULT_DEPARTMENTS = [
            "Panchayat General Wing", "Engineering Wing (LSGD)", "Agriculture Office", 
            "Veterinary Hospital", "Primary Health Centre (PHC)", "Integrated Child Development Services (ICDS)", 
            "Village Extension Office (VEO)", "Fisheries Extension", "SC Development Department"
        ];

        const DEFAULT_OFFICERS = [
            "Assistant Engineer (LSGD)", "Assistant Development Officer (VEO)", 
            "Agricultural Officer", "Veterinary Surgeon", "Medical Officer (PHC)", 
            "Child Development Project Officer (CDPO)", "Panchayat Assistant Secretary"
        ];

        const DEFAULT_EVAL_CRITERIA = [
            "Relevance to identified local needs",
            "Consistency with approved project objectives",
            "Timeliness of implementation",
            "Physical progress against target",
            "Financial management & transparency",
            "Quality of work or service",
            "Beneficiary coverage against target",
            "Beneficiary satisfaction rating",
            "Inclusion of SC/ST and vulnerable groups",
            "Environmental sustainability",
            "Transparency & social audit compliance",
            "Achievement of intended physical outputs",
            "Achievement of intended community outcomes",
            "Maintenance & sustainability arrangements"
        ];

        // ==========================================
        // SAMPLE INITIALIZATION DATA (KERALA GRAMA PANCHAYAT)
        // ==========================================
        const SAMPLE_PROJECTS = [
            {
                id: "PRJ-14-2024-001",
                code: "14/2024/001",
                fy: "2024-2025",
                ward: "Ward 1",
                name: "Drinking Water Pipe Extension & Storage Tank Construction - Peralam North",
                sector: "Drinking Water",
                category: "General Category",
                dept: "Engineering Wing (LSGD)",
                officer: "Assistant Engineer (LSGD)",
                benType: "Community Project",
                cost: 750000,
                fundGen: 550000,
                fundScsp: 200000,
                fundTsp: 0,
                asNo: "AS-14/402/2024",
                asDate: "2024-04-12",
                tsNo: "TS-88/2024",
                targetDate: "2025-02-28",
                desc: "Construction of 20,000L OHT storage tank and 1.2km distribution pipe laying to resolve drinking water shortage in Peralam North area.",
                status: "Work in Progress",
                physicalPercent: 65,
                beneficiariesCovered: 320,
                expenditures: [
                    { date: "2024-06-15", voucherNo: "VCH-2024-012", amount: 250000, head: "General Development Fund", remarks: "First mobilization & material advance", ucStatus: "Submitted" },
                    { date: "2024-08-20", voucherNo: "VCH-2024-045", amount: 200000, head: "SCSP Fund", remarks: "Substructure completion payment", ucStatus: "Pending" }
                ],
                progressLogs: [
                    { date: "2024-05-01", stage: "Work Order Issued", percent: 0, beneficiaries: 0, remarks: "Work order issued to Contracting Society." },
                    { date: "2024-07-10", stage: "Substructure / Foundation", percent: 35, beneficiaries: 0, remarks: "Foundation works and main column casting completed." },
                    { date: "2024-09-15", stage: "Superstructure / Intermediate", percent: 65, beneficiaries: 320, remarks: "Tank slab casted. Pipe trenching in progress." }
                ]
            },
            {
                id: "PRJ-14-2024-002",
                code: "14/2024/002",
                fy: "2024-2025",
                ward: "Ward 3",
                name: "SC Colony Road Tarring & Drain Construction",
                sector: "Roads and Transportation",
                category: "SCSP (Special Component)",
                dept: "Engineering Wing (LSGD)",
                officer: "Assistant Engineer (LSGD)",
                benType: "Community Project",
                cost: 1200000,
                fundGen: 0,
                fundScsp: 1200000,
                fundTsp: 0,
                asNo: "AS-14/405/2024",
                asDate: "2024-04-20",
                tsNo: "TS-92/2024",
                targetDate: "2024-11-30",
                desc: "Providing metalled surface and BM & BC tarring for 800m colony road along with concrete drainage line.",
                status: "Completed",
                physicalPercent: 100,
                beneficiariesCovered: 180,
                expenditures: [
                    { date: "2024-06-01", voucherNo: "VCH-2024-018", amount: 400000, head: "SCSP Fund", remarks: "Metal aggregate supply", ucStatus: "Submitted" },
                    { date: "2024-09-10", voucherNo: "VCH-2024-060", amount: 780000, head: "SCSP Fund", remarks: "Final tarring & drainage settlement", ucStatus: "Approved" }
                ],
                progressLogs: [
                    { date: "2024-05-15", stage: "Work Order Issued", percent: 10, beneficiaries: 0, remarks: "Site handover completed." },
                    { date: "2024-08-01", stage: "Superstructure / Intermediate", percent: 70, beneficiaries: 100, remarks: "Base metal consolidation completed." },
                    { date: "2024-09-25", stage: "Completed", percent: 100, beneficiaries: 180, remarks: "Road completed and opened for public transit." }
                ]
            },
            {
                id: "PRJ-14-2024-003",
                code: "14/2024/003",
                fy: "2024-2025",
                ward: "Ward 5",
                name: "High-Tech Dairy Farmer Milch Cow Assistance Scheme",
                sector: "Animal Husbandry and Dairy Development",
                category: "General Category",
                dept: "Veterinary Hospital",
                officer: "Veterinary Surgeon",
                benType: "Individual Beneficiary",
                cost: 500000,
                fundGen: 500000,
                fundScsp: 0,
                fundTsp: 0,
                asNo: "AS-14/412/2024",
                asDate: "2024-05-10",
                tsNo: "N/A",
                targetDate: "2024-10-15",
                desc: "Financial subsidy assistance of ₹25,000 per selected dairy farmer for purchase of high-yield milch cows.",
                status: "Delayed",
                physicalPercent: 40,
                beneficiariesCovered: 8,
                expenditures: [
                    { date: "2024-07-05", voucherNo: "VCH-2024-029", amount: 200000, head: "General Development Fund", remarks: "Subsidy disbursement to first batch (8 farmers)", ucStatus: "Pending" }
                ],
                progressLogs: [
                    { date: "2024-06-01", stage: "Approved", percent: 10, beneficiaries: 0, remarks: "Beneficiary selection committee approved list." },
                    { date: "2024-07-05", stage: "Work in Progress", percent: 40, beneficiaries: 8, remarks: "First batch distribution completed." }
                ]
            }
        ];

        const SAMPLE_BENEFICIARIES = [
            { id: "BEN-001", projectId: "PRJ-14-2024-003", name: "Sunitha K.V.", ward: "Ward 5", address: "Kizhakkethil House, Peralam", category: "General", gender: "Female", sanctionDetails: "₹25,000 Milch Cow Subsidy", sanctionDate: "2024-07-05", status: "Completed" },
            { id: "BEN-002", projectId: "PRJ-14-2024-003", name: "Ramanathan M.", ward: "Ward 5", address: "Moolayil House", category: "SC", gender: "Male", sanctionDetails: "₹25,000 Milch Cow Subsidy", sanctionDate: "2024-07-05", status: "Completed" }
        ];

        const SAMPLE_INSPECTIONS = [
            { id: "INS-001", projectId: "PRJ-14-2024-001", date: "2024-09-12", officer: "Assistant Engineer (LSGD)", progressObserved: 65, quality: "Satisfactory / Good", result: "Satisfactory", defects: "Minor concrete honeycomb in column base; directed plastering touchup.", compliance: "Compliant" },
            { id: "INS-002", projectId: "PRJ-14-2024-003", date: "2024-08-30", officer: "Veterinary Surgeon", progressObserved: 40, quality: "Satisfactory / Good", result: "Satisfactory with Observations", defects: "Animal insurance tag verification pending for 2 cows.", compliance: "Partially Compliant" }
        ];

        const SAMPLE_EVALUATIONS = {
            "PRJ-14-2024-002": {
                evalId: "EV-2024-002",
                projectId: "PRJ-14-2024-002",
                type: "Final Evaluation",
                evaluator: "Planning Committee Chairman",
                date: "2024-10-01",
                scores: { 0: 5, 1: 5, 2: 4, 3: 5, 4: 5, 5: 4, 6: 5, 7: 5, 8: 5, 9: 4, 10: 5, 11: 5, 12: 5, 13: 4 },
                totalScore: 68,
                maxScore: 70,
                percentage: 97.14,
                grade: "Excellent",
                achievements: "Completed ahead of monsoon season; quality BM&BC surface ensured durable connectivity to 180 SC families.",
                problems: "Minor delay in initial utility pipe shifting."
            }
        };

        const SAMPLE_OUTCOMES = [
            { id: "OUT-001", projectId: "PRJ-14-2024-001", name: "Drinking Water Storage Capacity Built", unit: "Liters", target: 20000, achieved: 20000, directBen: 320, indirectBen: 500, sustainability: "High" },
            { id: "OUT-002", projectId: "PRJ-14-2024-002", name: "Colony Road Tarring", unit: "Meters", target: 800, achieved: 800, directBen: 180, indirectBen: 400, sustainability: "High" }
        ];

        const SAMPLE_DOCUMENTS = [
            { id: "DOC-001", projectId: "PRJ-14-2024-001", type: "Administrative Sanction (AS)", refNo: "AS-14/402/2024", date: "2024-04-12", uploadedBy: "Planning Coordinator", status: "Verified", url: "" },
            { id: "DOC-002", projectId: "PRJ-14-2024-002", type: "Technical Sanction (TS)", refNo: "TS-92/2024", date: "2024-04-20", uploadedBy: "Assistant Engineer", status: "Verified", url: "" }
        ];

        // ==========================================
        // STATE MANAGEMENT ENGINE
        // ==========================================
        let appState = {
            userRole: "Admin",
            profile: {
                panchayatName: "KARIVELLUR-PERALAM GRAMA PANCHAYAT",
                blockName: "Payyanur",
                districtName: "Kannur",
                sig1: "Planning Coordinator",
                sig2: "Grama Panchayat Secretary"
            },
            projects: [],
            beneficiaries: [],
            inspections: [],
            evaluations: {},
            outcomes: [],
            documents: [],
            auditLogs: [],
            charts: {}
        };

        // Initialize App on DOM Load
        window.onload = function() {
            initDatabase();
            populateMasterDropdowns();
            switchModule('dashboard');
            logAudit("System initialized & session started.");
        };

        // Database Initialization / LocalStorage Backup Sync
        function initDatabase() {
            const storedState = localStorage.getItem("PAN CHAYAT_14TH_PLAN_DB");
            if (storedState) {
                try {
                    const parsed = JSON.parse(storedState);
                    appState = { ...appState, ...parsed };
                } catch(e) {
                    console.error("Error loading stored database, loading default sample data.", e);
                    loadDefaultSampleData();
                }
            } else {
                loadDefaultSampleData();
            }
            saveStateToStorage();
        }

        function loadDefaultSampleData() {
            appState.projects = JSON.parse(JSON.stringify(SAMPLE_PROJECTS));
            appState.beneficiaries = JSON.parse(JSON.stringify(SAMPLE_BENEFICIARIES));
            appState.inspections = JSON.parse(JSON.stringify(SAMPLE_INSPECTIONS));
            appState.evaluations = JSON.parse(JSON.stringify(SAMPLE_EVALUATIONS));
            appState.outcomes = JSON.parse(JSON.stringify(SAMPLE_OUTCOMES));
            appState.documents = JSON.parse(JSON.stringify(SAMPLE_DOCUMENTS));
            appState.auditLogs = [
                { timestamp: new Date().toLocaleString(), log: "Pre-loaded 14th Five-Year Plan sample records for Karivellur-Peralam Grama Panchayat." }
            ];
        }

        function saveStateToStorage() {
            localStorage.setItem("PAN CHAYAT_14TH_PLAN_DB", JSON.stringify(appState));
        }

        function resetToSampleData() {
            if (confirm("Reset database with default Kerala Grama Panchayat sample records? All custom changes will be overwritten.")) {
                loadDefaultSampleData();
                saveStateToStorage();
                showNotification("Database reset to initial sample state successfully!", "success");
                location.reload();
            }
        }

        function logAudit(actionMsg) {
            const entry = { timestamp: new Date().toLocaleString(), log: `[${appState.userRole}] ${actionMsg}` };
            appState.auditLogs.unshift(entry);
            if (appState.auditLogs.length > 50) appState.auditLogs.pop();
            saveStateToStorage();
            renderAuditLogBox();
        }

        function renderAuditLogBox() {
            const box = document.getElementById("audit-log-box");
            if (!box) return;
            box.innerHTML = appState.auditLogs.map(l => `<div><span class="text-slate-500">${l.timestamp}:</span> ${l.log}</div>`).join('');
        }

        // ==========================================
        // UI & NAVIGATION SWITCHING
        // ==========================================
        function switchModule(moduleId) {
            const modules = [
                'dashboard', 'master-register', 'financial-register', 'progress-register',
                'beneficiary-register', 'inspection-register', 'evaluation-module', 
                'outcome-register', 'corrective-action', 'documents-register', 
                'reports-module', 'master-settings'
            ];

            modules.forEach(m => {
                const el = document.getElementById(`module-${m}`);
                const navEl = document.getElementById(`nav-${m}`);
                if (el) el.classList.add('hidden');
                if (navEl) navEl.classList.remove('sidebar-item-active');
            });

            const activeEl = document.getElementById(`module-${moduleId}`);
            const activeNav = document.getElementById(`nav-${moduleId}`);
            if (activeEl) activeEl.classList.remove('hidden');
            if (activeNav) activeNav.classList.add('sidebar-item-active');

            // Trigger Module Specific Logic
            if (moduleId === 'dashboard') renderDashboard();
            if (moduleId === 'master-register') renderProjectMasterTable();
            if (moduleId === 'financial-register') populateFinancialSelect();
            if (moduleId === 'progress-register') populateProgressSelect();
            if (moduleId === 'beneficiary-register') { populateBeneficiaryProjectSelect(); renderBeneficiaryTable(); }
            if (moduleId === 'inspection-register') { populateInspectionProjectSelect(); renderInspectionTable(); }
            if (moduleId === 'evaluation-module') { populateEvaluationProjectSelect(); }
            if (moduleId === 'outcome-register') { populateOutcomeProjectSelect(); renderOutcomeTable(); }
            if (moduleId === 'corrective-action') renderCorrectiveActionRegister();
            if (moduleId === 'documents-register') { populateDocumentProjectSelect(); renderDocumentTable(); }
            if (moduleId === 'reports-module') generateReportPreview();
            if (moduleId === 'master-settings') renderMasterSettings();

            // Mobile sidebar collapse
            if (window.innerWidth < 1024) {
                document.getElementById('sidebar').classList.add('-translate-x-full');
            }
        }

        function switchUserRole(role) {
            appState.userRole = role;
            document.getElementById('user-role-display').innerText = `Role: ${role}`;
            logAudit(`Switched active user role to ${role}`);
            showNotification(`Active role switched to ${role}`, "info");
        }

        document.getElementById('sidebar-toggle').addEventListener('click', () => {
            document.getElementById('sidebar').classList.toggle('-translate-x-full');
        });

        function showNotification(msg, type = "info") {
            const banner = document.getElementById('notification-banner');
            const icon = document.getElementById('notif-icon');
            const message = document.getElementById('notif-message');

            banner.className = "mb-4 p-3 rounded-lg flex items-center justify-between text-sm shadow-sm transition-all duration-300 ";
            if (type === "success") {
                banner.classList.add("bg-emerald-100", "text-emerald-900", "border", "border-emerald-300");
                icon.className = "fa-solid fa-circle-check text-emerald-600";
            } else if (type === "warning") {
                banner.classList.add("bg-amber-100", "text-amber-900", "border", "border-amber-300");
                icon.className = "fa-solid fa-triangle-exclamation text-amber-600";
            } else if (type === "error") {
                banner.classList.add("bg-rose-100", "text-rose-900", "border", "border-rose-300");
                icon.className = "fa-solid fa-circle-xmark text-rose-600";
            } else {
                banner.classList.add("bg-blue-100", "text-blue-900", "border", "border-blue-300");
                icon.className = "fa-solid fa-circle-info text-blue-600";
            }

            message.innerText = msg;
            banner.classList.remove('hidden');
            setTimeout(() => { banner.classList.add('hidden'); }, 4000);
        }

        function closeNotification() {
            document.getElementById('notification-banner').classList.add('hidden');
        }

        function populateMasterDropdowns() {
            // Wards (1 to 18)
            const wardSelects = ['prj-ward', 'ben-ward', 'master-filter-ward'];
            wardSelects.forEach(sId => {
                const sel = document.getElementById(sId);
                if (!sel) return;
                const currentVal = sel.value;
                sel.innerHTML = sId.includes('filter') ? '<option value="">All Wards</option>' : '';
                for (let w = 1; w <= 18; w++) {
                    sel.innerHTML += `<option value="Ward ${w}">Ward ${w}</option>`;
                }
                if (currentVal) sel.value = currentVal;
            });

            // Sectors
            const sectorSelects = ['prj-sector', 'dash-filter-sector', 'master-filter-sector', 'report-filter-sector'];
            sectorSelects.forEach(sId => {
                const sel = document.getElementById(sId);
                if (!sel) return;
                const currentVal = sel.value;
                sel.innerHTML = sId.includes('filter') ? '<option value="ALL">All Sectors</option>' : '';
                DEFAULT_SECTORS.forEach(sec => {
                    sel.innerHTML += `<option value="${sec}">${sec}</option>`;
                });
                if (currentVal) sel.value = currentVal;
            });

            // Departments & Officers
            const deptSel = document.getElementById('prj-dept');
            if (deptSel) {
                deptSel.innerHTML = DEFAULT_DEPARTMENTS.map(d => `<option value="${d}">${d}</option>`).join('');
            }
            const offSel = document.getElementById('prj-officer');
            if (offSel) {
                offSel.innerHTML = DEFAULT_OFFICERS.map(o => `<option value="${o}">${o}</option>`).join('');
            }

            // FY filter
            const fySelects = ['dash-filter-fy', 'master-filter-fy', 'report-filter-fy'];
            fySelects.forEach(fId => {
                const sel = document.getElementById(fId);
                if (!sel) return;
                sel.innerHTML = '<option value="ALL">All Financial Years</option>' +
                    ['2022-2023', '2023-2024', '2024-2025', '2025-2026', '2026-2027'].map(y => `<option value="${y}">${y}</option>`).join('');
            });
        }

        // ==========================================
        // MODULE 1: DASHBOARD CALCULATIONS & CHARTS
        // ==========================================
        function renderDashboard() {
            const sectorFilter = document.getElementById('dash-filter-sector')?.value || 'ALL';
            const fyFilter = document.getElementById('dash-filter-fy')?.value || 'ALL';

            let filteredProjects = appState.projects.filter(p => {
                const matchSec = (sectorFilter === 'ALL' || p.sector === sectorFilter);
                const matchFy = (fyFilter === 'ALL' || p.fy === fyFilter);
                return matchSec && matchFy;
            });

            // Calculations for 18 Metric Cards
            const totalCount = filteredProjects.length;
            const approvedCount = filteredProjects.filter(p => p.status === 'Approved' || p.status === 'Work in Progress' || p.status === 'Completed').length;
            const inProgressCount = filteredProjects.filter(p => p.status === 'Work in Progress').length;
            const completedCount = filteredProjects.filter(p => p.status === 'Completed' || p.status === 'Financially Closed').length;
            const notStartedCount = filteredProjects.filter(p => p.status === 'Not Started' || p.status === 'Proposed').length;
            const delayedCount = filteredProjects.filter(p => p.status === 'Delayed').length;
            
            // Evaluated calculations
            let evaluatedCount = 0;
            let satisfactoryCount = 0;
            let unsatisfactoryCount = 0;

            filteredProjects.forEach(p => {
                const ev = appState.evaluations[p.id];
                if (ev) {
                    evaluatedCount++;
                    if (ev.percentage >= 60) satisfactoryCount++;
                    else unsatisfactoryCount++;
                }
            });

            const pendingEvalCount = totalCount - evaluatedCount;

            // Financial Totals
            let totalCost = 0;
            let totalSpent = 0;

            filteredProjects.forEach(p => {
                totalCost += (p.cost || 0);
                const spent = (p.expenditures || []).reduce((sum, e) => sum + (e.amount || 0), 0);
                totalSpent += spent;
            });

            const totalReleased = totalSpent; // Released as spent in ledger
            const totalRemaining = totalCost - totalSpent;

            // Overall Progress Percentages
            const avgPhysical = totalCount ? (filteredProjects.reduce((sum, p) => sum + (p.physicalPercent || 0), 0) / totalCount).toFixed(1) : 0;
            const overallFinancial = totalCost ? ((totalSpent / totalCost) * 100).toFixed(1) : 0;

            // Corrective Action Flags count
            const correctiveCount = getFlaggedProjects().length;
            document.getElementById('badge-corrective-count').innerText = correctiveCount;
            if (correctiveCount > 0) {
                document.getElementById('badge-corrective-count').classList.remove('hidden');
            }

            const cardsData = [
                { id: 1, title: "Total Registered Projects", val: totalCount, color: "border-blue-500 bg-blue-50 text-blue-900", icon: "fa-folder-open", filterStatus: "" },
                { id: 2, title: "Approved Projects", val: approvedCount, color: "border-sky-500 bg-sky-50 text-sky-900", icon: "fa-circle-check", filterStatus: "Approved" },
                { id: 3, title: "Under Implementation", val: inProgressCount, color: "border-purple-500 bg-purple-50 text-purple-900", icon: "fa-spinner", filterStatus: "Work in Progress" },
                { id: 4, title: "Completed Projects", val: completedCount, color: "border-emerald-500 bg-emerald-50 text-emerald-900", icon: "fa-circle-check", filterStatus: "Completed" },
                { id: 5, title: "Projects Not Started", val: notStartedCount, color: "border-slate-500 bg-slate-100 text-slate-800", icon: "fa-hourglass-start", filterStatus: "Not Started" },
                { id: 6, title: "Delayed Projects", val: delayedCount, color: "border-amber-500 bg-amber-50 text-amber-900", icon: "fa-clock", filterStatus: "Delayed" },
                { id: 7, title: "Projects Requiring Attention", val: correctiveCount, color: "border-red-500 bg-red-50 text-red-900", icon: "fa-triangle-exclamation", filterModule: "corrective-action" },
                { id: 8, title: "Projects Evaluated", val: evaluatedCount, color: "border-indigo-500 bg-indigo-50 text-indigo-900", icon: "fa-award", filterModule: "evaluation-module" },
                { id: 9, title: "Satisfactory Performance", val: satisfactoryCount, color: "border-teal-500 bg-teal-50 text-teal-900", icon: "fa-thumbs-up", filterModule: "evaluation-module" },
                { id: 10, title: "Unsatisfactory Performance", val: unsatisfactoryCount, color: "border-rose-500 bg-rose-50 text-rose-900", icon: "fa-thumbs-down", filterModule: "evaluation-module" },
                { id: 11, title: "Total Approved Cost", val: `₹${(totalCost/100000).toFixed(2)} L`, color: "border-blue-600 bg-blue-50 text-blue-950", icon: "fa-indian-rupee-sign" },
                { id: 12, title: "Total Amount Released", val: `₹${(totalReleased/100000).toFixed(2)} L`, color: "border-cyan-600 bg-cyan-50 text-cyan-950", icon: "fa-hand-holding-dollar" },
                { id: 13, title: "Total Expenditure", val: `₹${(totalSpent/100000).toFixed(2)} L`, color: "border-emerald-600 bg-emerald-50 text-emerald-950", icon: "fa-money-bill-wave" },
                { id: 14, title: "Total Remaining Amount", val: `₹${(totalRemaining/100000).toFixed(2)} L`, color: "border-amber-600 bg-amber-50 text-amber-950", icon: "fa-piggy-bank" },
                { id: 15, title: "Overall Physical Progress", val: `${avgPhysical}%`, color: "border-purple-600 bg-purple-50 text-purple-950", icon: "fa-bars-progress" },
                { id: 16, title: "Overall Financial Progress", val: `${overallFinancial}%`, color: "border-teal-600 bg-teal-50 text-teal-950", icon: "fa-chart-pie" },
                { id: 17, title: "Pending Evaluation", val: pendingEvalCount, color: "border-orange-500 bg-orange-50 text-orange-900", icon: "fa-file-signature", filterModule: "evaluation-module" },
                { id: 18, title: "Requiring Corrective Action", val: correctiveCount, color: "border-rose-600 bg-rose-50 text-rose-950", icon: "fa-bell", filterModule: "corrective-action" }
            ];

            const cardsContainer = document.getElementById('dashboard-cards-container');
            if (cardsContainer) {
                cardsContainer.innerHTML = cardsData.map(c => `
                    <div onclick="handleCardClick('${c.filterModule || 'master-register'}', '${c.filterStatus || ''}')" 
                         class="p-3.5 rounded-xl border-l-4 shadow-sm ${c.color} cursor-pointer hover:shadow-md transition-all flex flex-col justify-between">
                        <div class="flex justify-between items-start">
                            <span class="text-[11px] font-bold uppercase tracking-wider leading-tight">${c.title}</span>
                            <i class="fa-solid ${c.icon} text-sm opacity-75"></i>
                        </div>
                        <div class="mt-2 text-lg md:text-xl font-extrabold">${c.val}</div>
                    </div>
                `).join('');
            }

            renderDashboardCharts(filteredProjects);
        }

        function handleCardClick(moduleName, statusFilter) {
            switchModule(moduleName);
            if (moduleName === 'master-register' && statusFilter) {
                const sel = document.getElementById('master-filter-status');
                if (sel) {
                    sel.value = statusFilter;
                    filterProjectMasterTable();
                }
            }
        }

        function renderDashboardCharts(projects) {
            // 1. Sectors Chart
            const sectorCounts = {};
            projects.forEach(p => sectorCounts[p.sector] = (sectorCounts[p.sector] || 0) + 1);
            
            createChart('chart-sector', 'pie', {
                labels: Object.keys(sectorCounts),
                datasets: [{
                    data: Object.values(sectorCounts),
                    backgroundColor: ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6', '#ec4899', '#06b6d4', '#64748b', '#14b8a6']
                }]
            });

            // 2. Status Chart
            const statusCounts = {};
            projects.forEach(p => statusCounts[p.status] = (statusCounts[p.status] || 0) + 1);

            createChart('chart-status', 'doughnut', {
                labels: Object.keys(statusCounts),
                datasets: [{
                    data: Object.values(statusCounts),
                    backgroundColor: ['#3b82f6', '#8b5cf6', '#10b981', '#f59e0b', '#ef4444', '#64748b']
                }]
            });

            // 3. Approved vs Expenditure Chart
            const projectNames = projects.slice(0, 6).map(p => p.code);
            const approvedCosts = projects.slice(0, 6).map(p => (p.cost / 100000).toFixed(2));
            const spentAmounts = projects.slice(0, 6).map(p => {
                const s = (p.expenditures || []).reduce((acc, e) => acc + e.amount, 0);
                return (s / 100000).toFixed(2);
            });

            createChart('chart-finance', 'bar', {
                labels: projectNames,
                datasets: [
                    { label: 'Sanctioned Cost (Lakhs)', data: approvedCosts, backgroundColor: '#3b82f6' },
                    { label: 'Expenditure (Lakhs)', data: spentAmounts, backgroundColor: '#10b981' }
                ]
            });

            // 4. Evaluation Performance Grades
            const gradeCounts = { "Excellent": 0, "Very Good": 0, "Good": 0, "Needs Improvement": 0, "Unsatisfactory": 0, "Unevaluated": 0 };
            projects.forEach(p => {
                const ev = appState.evaluations[p.id];
                if (ev && ev.grade) gradeCounts[ev.grade] = (gradeCounts[ev.grade] || 0) + 1;
                else gradeCounts["Unevaluated"]++;
            });

            createChart('chart-evaluation', 'bar', {
                labels: Object.keys(gradeCounts),
                datasets: [{
                    label: 'Projects Count',
                    data: Object.values(gradeCounts),
                    backgroundColor: ['#10b981', '#06b6d4', '#3b82f6', '#f59e0b', '#ef4444', '#94a3b8']
                }]
            });

            // 5. Year-Wise Progress Trend
            createChart('chart-progress-trend', 'line', {
                labels: ['2022-23', '2023-24', '2024-25', '2025-26'],
                datasets: [
                    { label: 'Avg Physical %', data: [88, 76, 55, 10], borderColor: '#8b5cf6', fill: false },
                    { label: 'Avg Financial %', data: [92, 80, 48, 5], borderColor: '#14b8a6', fill: false }
                ]
            });

            // 6. Projects by Officer
            const officerCounts = {};
            projects.forEach(p => officerCounts[p.officer] = (officerCounts[p.officer] || 0) + 1);

            createChart('chart-officer', 'bar', {
                labels: Object.keys(officerCounts),
                datasets: [{
                    label: 'Assigned Projects',
                    data: Object.values(officerCounts),
                    backgroundColor: '#6366f1'
                }]
            });
        }

        function createChart(canvasId, type, dataObj) {
            const ctx = document.getElementById(canvasId)?.getContext('2d');
            if (!ctx) return;
            if (appState.charts[canvasId]) {
                appState.charts[canvasId].destroy();
            }
            appState.charts[canvasId] = new Chart(ctx, {
                type: type,
                data: dataObj,
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { display: type !== 'bar', position: 'bottom', labels: { font: { size: 10 } } } }
                }
            });
        }

        // ==========================================
        // MODULE 2: PROJECT MASTER REGISTER (CRUD)
        // ==========================================
        function renderProjectMasterTable() {
            populateMasterDropdowns();
            filterProjectMasterTable();
        }

        function filterProjectMasterTable() {
            const search = document.getElementById('master-search')?.value.toLowerCase() || '';
            const sector = document.getElementById('master-filter-sector')?.value || '';
            const status = document.getElementById('master-filter-status')?.value || '';
            const fy = document.getElementById('master-filter-fy')?.value || '';
            const ward = document.getElementById('master-filter-ward')?.value || '';

            const filtered = appState.projects.filter(p => {
                const matchSearch = p.code.toLowerCase().includes(search) || p.name.toLowerCase().includes(search);
                const matchSec = !sector || p.sector === sector;
                const matchStatus = !status || p.status === status;
                const matchFy = !fy || p.fy === fy;
                const matchWard = !ward || p.ward === ward;
                return matchSearch && matchSec && matchStatus && matchFy && matchWard;
            });

            const tbody = document.getElementById('project-master-tbody');
            if (!tbody) return;

            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="8" class="text-center py-6 text-slate-400">No project records found matching filters.</td></tr>`;
                return;
            }

            tbody.innerHTML = filtered.map(p => {
                const statusBadge = getStatusBadge(p.status);
                return `
                    <tr class="hover:bg-slate-50 transition-colors">
                        <td class="py-3 px-3">
                            <div class="font-bold text-blue-900 font-mono text-[11px]">${p.code}</div>
                            <div class="font-medium text-slate-800 line-clamp-1" title="${p.name}">${p.name}</div>
                            <div class="text-[10px] text-slate-400">${p.ward}</div>
                        </td>
                        <td class="py-3 px-3">
                            <div class="font-medium">${p.sector}</div>
                            <div class="text-[10px] text-slate-400">${p.category}</div>
                        </td>
                        <td class="py-3 px-3">
                            <div class="font-medium text-slate-700">${p.officer}</div>
                            <div class="text-[10px] text-slate-400">${p.dept}</div>
                        </td>
                        <td class="py-3 px-3">${p.fy}</td>
                        <td class="py-3 px-3 text-right font-bold text-slate-800">₹${p.cost.toLocaleString('en-IN')}</td>
                        <td class="py-3 px-3 text-center">
                            <div class="w-full bg-slate-200 rounded-full h-2 overflow-hidden mb-1">
                                <div class="bg-purple-600 h-2" style="width: ${p.physicalPercent}%"></div>
                            </div>
                            <span class="font-semibold text-[10px]">${p.physicalPercent}%</span>
                        </td>
                        <td class="py-3 px-3 text-center">${statusBadge}</td>
                        <td class="py-3 px-3 text-center space-x-1">
                            <button onclick="editProjectModal('${p.id}')" class="p-1.5 text-blue-600 hover:bg-blue-50 rounded" title="Edit Record"><i class="fa-solid fa-pen-to-square"></i></button>
                            <button onclick="deleteProjectRecord('${p.id}')" class="p-1.5 text-rose-600 hover:bg-rose-50 rounded" title="Delete Record"><i class="fa-solid fa-trash"></i></button>
                        </td>
                    </tr>
                `;
            }).join('');
        }

        function getStatusBadge(status) {
            let cls = "bg-slate-100 text-slate-700 border-slate-300";
            if (status === 'Completed') cls = "bg-emerald-100 text-emerald-800 border-emerald-300";
            if (status === 'Work in Progress') cls = "bg-purple-100 text-purple-800 border-purple-300";
            if (status === 'Approved') cls = "bg-sky-100 text-sky-800 border-sky-300";
            if (status === 'Delayed') cls = "bg-amber-100 text-amber-800 border-amber-300";
            if (status === 'Financially Closed') cls = "bg-teal-100 text-teal-800 border-teal-300";
            return `<span class="px-2 py-0.5 rounded-full text-[10px] font-bold border ${cls}">${status}</span>`;
        }

        function openProjectModal() {
            document.getElementById('project-form').reset();
            document.getElementById('prj-id').value = "";
            const newCode = `14/2024/${String(appState.projects.length + 1).padStart(3, '0')}`;
            document.getElementById('prj-code').value = newCode;
            document.getElementById('modal-project-title').innerText = "Register New Panchayat Project (14th FYP)";
            document.getElementById('modal-project').classList.remove('hidden');
        }

        function editProjectModal(id) {
            const p = appState.projects.find(x => x.id === id);
            if (!p) return;
            document.getElementById('prj-id').value = p.id;
            document.getElementById('prj-code').value = p.code;
            document.getElementById('prj-fy').value = p.fy;
            document.getElementById('prj-ward').value = p.ward;
            document.getElementById('prj-name').value = p.name;
            document.getElementById('prj-sector').value = p.sector;
            document.getElementById('prj-category').value = p.category;
            document.getElementById('prj-dept').value = p.dept;
            document.getElementById('prj-officer').value = p.officer;
            document.getElementById('prj-ben-type').value = p.benType;
            document.getElementById('prj-cost').value = p.cost;
            document.getElementById('fund-gen').value = p.fundGen || 0;
            document.getElementById('fund-scsp').value = p.fundScsp || 0;
            document.getElementById('fund-tsp').value = p.fundTsp || 0;
            document.getElementById('prj-as-no').value = p.asNo || "";
            document.getElementById('prj-as-date').value = p.asDate || "";
            document.getElementById('prj-ts-no').value = p.tsNo || "";
            document.getElementById('prj-target-date').value = p.targetDate || "";
            document.getElementById('prj-desc').value = p.desc || "";

            document.getElementById('modal-project-title').innerText = "Edit Project Master Record";
            document.getElementById('modal-project').classList.remove('hidden');
        }

        function closeProjectModal() {
            document.getElementById('modal-project').classList.add('hidden');
        }

        function saveProjectRecord(e) {
            e.preventDefault();
            const id = document.getElementById('prj-id').value;
            const code = document.getElementById('prj-code').value;
            const cost = parseFloat(document.getElementById('prj-cost').value);

            const record = {
                id: id || `PRJ-14-2024-${String(Date.now()).slice(-4)}`,
                code: code,
                fy: document.getElementById('prj-fy').value,
                ward: document.getElementById('prj-ward').value,
                name: document.getElementById('prj-name').value,
                sector: document.getElementById('prj-sector').value,
                category: document.getElementById('prj-category').value,
                dept: document.getElementById('prj-dept').value,
                officer: document.getElementById('prj-officer').value,
                benType: document.getElementById('prj-ben-type').value,
                cost: cost,
                fundGen: parseFloat(document.getElementById('fund-gen').value) || 0,
                fundScsp: parseFloat(document.getElementById('fund-scsp').value) || 0,
                fundTsp: parseFloat(document.getElementById('fund-tsp').value) || 0,
                asNo: document.getElementById('prj-as-no').value,
                asDate: document.getElementById('prj-as-date').value,
                tsNo: document.getElementById('prj-ts-no').value,
                targetDate: document.getElementById('prj-target-date').value,
                desc: document.getElementById('prj-desc').value,
                status: id ? (appState.projects.find(x => x.id === id)?.status || "Approved") : "Approved",
                physicalPercent: id ? (appState.projects.find(x => x.id === id)?.physicalPercent || 0) : 0,
                expenditures: id ? (appState.projects.find(x => x.id === id)?.expenditures || []) : [],
                progressLogs: id ? (appState.projects.find(x => x.id === id)?.progressLogs || []) : []
            };

            if (id) {
                const idx = appState.projects.findIndex(x => x.id === id);
                appState.projects[idx] = record;
                logAudit(`Updated project record ${code}`);
                showNotification(`Project ${code} updated successfully.`, "success");
            } else {
                appState.projects.push(record);
                logAudit(`Registered new project ${code}`);
                showNotification(`Project ${code} registered successfully.`, "success");
            }

            saveStateToStorage();
            closeProjectModal();
            renderProjectMasterTable();
        }

        function deleteProjectRecord(id) {
            const p = appState.projects.find(x => x.id === id);
            if (!p) return;
            if (confirm(`Are you sure you want to soft delete / remove project ${p.code}?`)) {
                appState.projects = appState.projects.filter(x => x.id !== id);
                logAudit(`Deleted project record ${p.code}`);
                saveStateToStorage();
                renderProjectMasterTable();
                showNotification(`Project ${p.code} deleted.`, "warning");
            }
        }

        // ==========================================
        // MODULE 3: FINANCIAL REGISTER
        // ==========================================
        function populateFinancialSelect() {
            const sel = document.getElementById('fin-project-select');
            if (!sel) return;
            sel.innerHTML = '<option value="">-- Select Project --</option>' + 
                appState.projects.map(p => `<option value="${p.id}">${p.code} - ${p.name}</option>`).join('');
        }

        function loadProjectFinancials(projectId) {
            const card = document.getElementById('fin-summary-card');
            const tbody = document.getElementById('fin-voucher-tbody');
            if (!projectId) {
                card.innerHTML = `<p class="text-slate-500">Select a project above to inspect financial ledger.</p>`;
                tbody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-slate-400">No project selected</td></tr>`;
                return;
            }

            const p = appState.projects.find(x => x.id === projectId);
            if (!p) return;

            const spentTotal = (p.expenditures || []).reduce((acc, e) => acc + e.amount, 0);
            const balance = p.cost - spentTotal;
            const finPercent = p.cost ? ((spentTotal / p.cost) * 100).toFixed(1) : 0;

            card.innerHTML = `
                <div class="space-y-2">
                    <div class="font-bold text-slate-800 text-sm">${p.code} - ${p.name}</div>
                    <div class="grid grid-cols-2 gap-2 text-[11px] pt-1 border-t border-slate-200">
                        <div>Sanctioned Cost: <span class="font-bold">₹${p.cost.toLocaleString('en-IN')}</span></div>
                        <div>General Fund: <span>₹${(p.fundGen||0).toLocaleString('en-IN')}</span></div>
                        <div>SCSP Fund: <span>₹${(p.fundScsp||0).toLocaleString('en-IN')}</span></div>
                        <div>TSP Fund: <span>₹${(p.fundTsp||0).toLocaleString('en-IN')}</span></div>
                    </div>
                    <div class="pt-2 border-t border-slate-200">
                        <div class="flex justify-between font-bold text-slate-800">
                            <span>Total Expenditure: ₹${spentTotal.toLocaleString('en-IN')}</span>
                            <span class="text-emerald-700">${finPercent}% Spent</span>
                        </div>
                        <div class="flex justify-between font-bold text-slate-700 text-[11px] mt-1">
                            <span>Balance Available: ₹${balance.toLocaleString('en-IN')}</span>
                        </div>
                    </div>
                </div>
            `;

            if (!p.expenditures || p.expenditures.length === 0) {
                tbody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-slate-400">No expenditure transactions recorded yet.</td></tr>`;
                return;
            }

            tbody.innerHTML = p.expenditures.map((e, idx) => `
                <tr class="hover:bg-slate-50">
                    <td class="py-2.5 px-3 font-medium">${e.date}</td>
                    <td class="py-2.5 px-3 font-mono font-bold text-blue-900">${e.voucherNo}</td>
                    <td class="py-2.5 px-3">${e.head}</td>
                    <td class="py-2.5 px-3">${e.remarks || '-'}</td>
                    <td class="py-2.5 px-3 text-right font-bold text-slate-800">₹${e.amount.toLocaleString('en-IN')}</td>
                    <td class="py-2.5 px-3 text-center">
                        <span class="px-2 py-0.5 text-[10px] font-bold rounded ${e.ucStatus === 'Approved' ? 'bg-emerald-100 text-emerald-800' : 'bg-amber-100 text-amber-800'}">${e.ucStatus}</span>
                    </td>
                    <td class="py-2.5 px-3 text-center">
                        <button onclick="deleteExpenditureVoucher('${p.id}', ${idx})" class="text-rose-600 hover:text-rose-800"><i class="fa-solid fa-trash-can"></i></button>
                    </td>
                </tr>
            `).join('');
        }

        function openExpenditureModal() {
            const pId = document.getElementById('fin-project-select')?.value;
            if (!pId) {
                alert("Please select a project first.");
                return;
            }
            document.getElementById('expenditure-form').reset();
            document.getElementById('exp-date').value = new Date().toISOString().split('T')[0];
            document.getElementById('modal-expenditure').classList.remove('hidden');
        }

        function closeExpenditureModal() {
            document.getElementById('modal-expenditure').classList.add('hidden');
        }

        function saveExpenditureVoucher(e) {
            e.preventDefault();
            const pId = document.getElementById('fin-project-select').value;
            const p = appState.projects.find(x => x.id === pId);
            if (!p) return;

            const amount = parseFloat(document.getElementById('exp-amount').value);
            const spentTotal = (p.expenditures || []).reduce((acc, ex) => acc + ex.amount, 0);

            if ((spentTotal + amount) > p.cost) {
                if (!confirm(`Warning: Expenditure (₹${(spentTotal + amount).toLocaleString('en-IN')}) will exceed sanctioned project budget (₹${p.cost.toLocaleString('en-IN')}). Proceed anyway?`)) {
                    return;
                }
            }

            const voucher = {
                date: document.getElementById('exp-date').value,
                voucherNo: document.getElementById('exp-voucher-no').value,
                amount: amount,
                head: document.getElementById('exp-head').value,
                remarks: document.getElementById('exp-remarks').value,
                ucStatus: "Pending"
            };

            if (!p.expenditures) p.expenditures = [];
            p.expenditures.push(voucher);

            logAudit(`Recorded expenditure voucher ${voucher.voucherNo} of ₹${amount} for project ${p.code}`);
            saveStateToStorage();
            closeExpenditureModal();
            loadProjectFinancials(pId);
            showNotification(`Voucher transaction saved successfully.`, "success");
        }

        function deleteExpenditureVoucher(projectId, index) {
            const p = appState.projects.find(x => x.id === projectId);
            if (!p) return;
            if (confirm("Delete this expenditure transaction?")) {
                const removed = p.expenditures.splice(index, 1);
                logAudit(`Deleted voucher ${removed[0].voucherNo} from project ${p.code}`);
                saveStateToStorage();
                loadProjectFinancials(projectId);
                showNotification("Voucher deleted.", "warning");
            }
        }

        // ==========================================
        // MODULE 4: PHYSICAL PROGRESS REGISTER
        // ==========================================
        function populateProgressSelect() {
            const sel = document.getElementById('prog-project-select');
            if (!sel) return;
            sel.innerHTML = '<option value="">-- Select Project --</option>' + 
                appState.projects.map(p => `<option value="${p.id}">${p.code} - ${p.name}</option>`).join('');
        }

        function loadProjectProgress(projectId) {
            const card = document.getElementById('prog-summary-card');
            const tbody = document.getElementById('prog-history-tbody');
            if (!projectId) {
                card.innerHTML = `<p class="text-slate-500">Select a project to view and update physical progress milestones.</p>`;
                tbody.innerHTML = `<tr><td colspan="5" class="text-center py-4 text-slate-400">No project selected</td></tr>`;
                return;
            }

            const p = appState.projects.find(x => x.id === projectId);
            if (!p) return;

            card.innerHTML = `
                <div class="space-y-2">
                    <div class="font-bold text-slate-800 text-sm">${p.code} - ${p.name}</div>
                    <div class="flex justify-between text-xs pt-1 border-t border-slate-200">
                        <span>Current Stage: <strong class="text-purple-900">${p.status}</strong></span>
                        <span>Physical Progress: <strong class="text-purple-900">${p.physicalPercent}%</strong></span>
                    </div>
                    <div class="w-full bg-slate-200 rounded-full h-2.5 overflow-hidden">
                        <div class="bg-purple-600 h-2.5" style="width: ${p.physicalPercent}%"></div>
                    </div>
                </div>
            `;

            document.getElementById('prog-update-date').value = new Date().toISOString().split('T')[0];
            document.getElementById('prog-percent').value = p.physicalPercent;
            document.getElementById('prog-stage-select').value = p.status;
            document.getElementById('prog-beneficiaries-covered').value = p.beneficiariesCovered || 0;

            if (!p.progressLogs || p.progressLogs.length === 0) {
                tbody.innerHTML = `<tr><td colspan="5" class="text-center py-4 text-slate-400">No progress history logs recorded.</td></tr>`;
                return;
            }

            tbody.innerHTML = p.progressLogs.map(l => `
                <tr class="hover:bg-slate-50">
                    <td class="py-2 px-3 font-medium">${l.date}</td>
                    <td class="py-2 px-3">${l.stage}</td>
                    <td class="py-2 px-3 text-center font-bold text-purple-900">${l.percent}%</td>
                    <td class="py-2 px-3 text-center">${l.beneficiaries || 0}</td>
                    <td class="py-2 px-3 text-slate-600">${l.remarks || '-'}</td>
                </tr>
            `).join('');
        }

        function saveProgressUpdate(e) {
            e.preventDefault();
            const pId = document.getElementById('prog-project-select').value;
            const p = appState.projects.find(x => x.id === pId);
            if (!p) return;

            const newStage = document.getElementById('prog-stage-select').value;
            const newPercent = parseFloat(document.getElementById('prog-percent').value);
            const benCovered = parseInt(document.getElementById('prog-beneficiaries-covered').value) || 0;

            p.status = newStage;
            p.physicalPercent = newPercent;
            p.beneficiariesCovered = benCovered;

            const logEntry = {
                date: document.getElementById('prog-update-date').value,
                stage: newStage,
                percent: newPercent,
                beneficiaries: benCovered,
                remarks: document.getElementById('prog-remarks').value
            };

            if (!p.progressLogs) p.progressLogs = [];
            p.progressLogs.unshift(logEntry);

            logAudit(`Updated physical progress of project ${p.code} to ${newPercent}% (${newStage})`);
            saveStateToStorage();
            loadProjectProgress(pId);
            showNotification("Physical progress log saved.", "success");
        }

        // ==========================================
        // MODULE 5: BENEFICIARY REGISTER
        // ==========================================
        function populateBeneficiaryProjectSelect() {
            const sel = document.getElementById('ben-filter-project');
            if (sel) {
                sel.innerHTML = '<option value="">All Projects</option>' +
                    appState.projects.map(p => `<option value="${p.id}">${p.code} - ${p.name}</option>`).join('');
            }
            const modalSel = document.getElementById('ben-project-id');
            if (modalSel) {
                modalSel.innerHTML = appState.projects.map(p => `<option value="${p.id}">${p.code} - ${p.name}</option>`).join('');
            }
        }

        function renderBeneficiaryTable() {
            const projId = document.getElementById('ben-filter-project')?.value || '';
            const cat = document.getElementById('ben-filter-category')?.value || '';
            const ward = document.getElementById('ben-filter-ward')?.value || '';
            const search = document.getElementById('ben-search')?.value.toLowerCase() || '';

            const filtered = appState.beneficiaries.filter(b => {
                const matchProj = !projId || b.projectId === projId;
                const matchCat = !cat || b.category === cat;
                const matchWard = !ward || b.ward === ward;
                const matchSearch = b.name.toLowerCase().includes(search);
                return matchProj && matchCat && matchWard && matchSearch;
            });

            const tbody = document.getElementById('beneficiary-tbody');
            if (!tbody) return;

            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="8" class="text-center py-6 text-slate-400">No beneficiary records found.</td></tr>`;
                return;
            }

            tbody.innerHTML = filtered.map(b => {
                const p = appState.projects.find(x => x.id === b.projectId);
                return `
                    <tr class="hover:bg-slate-50">
                        <td class="py-2.5 px-3">
                            <div class="font-bold text-blue-900">${p ? p.code : b.projectId}</div>
                            <div class="text-[10px] text-slate-500 line-clamp-1">${p ? p.name : ''}</div>
                        </td>
                        <td class="py-2.5 px-3 font-semibold text-slate-800">${b.name}</td>
                        <td class="py-2.5 px-3">
                            <div>${b.ward}</div>
                            <div class="text-[10px] text-slate-400">${b.address || ''}</div>
                        </td>
                        <td class="py-2.5 px-3">
                            <span class="px-2 py-0.5 rounded text-[10px] bg-slate-100 font-bold">${b.category}</span>
                            <div class="text-[10px] text-slate-400">${b.gender || ''}</div>
                        </td>
                        <td class="py-2.5 px-3 font-medium">${b.sanctionDetails}</td>
                        <td class="py-2.5 px-3 text-center">${b.sanctionDate}</td>
                        <td class="py-2.5 px-3 text-center"><span class="px-2 py-0.5 rounded text-[10px] bg-emerald-100 text-emerald-800 font-bold">${b.status}</span></td>
                        <td class="py-2.5 px-3 text-center">
                            <button onclick="deleteBeneficiary('${b.id}')" class="text-rose-600 hover:text-rose-800"><i class="fa-solid fa-trash-can"></i></button>
                        </td>
                    </tr>
                `;
            }).join('');
        }

        function openBeneficiaryModal() {
            document.getElementById('beneficiary-form').reset();
            document.getElementById('modal-beneficiary').classList.remove('hidden');
        }

        function closeBeneficiaryModal() {
            document.getElementById('modal-beneficiary').classList.add('hidden');
        }

        function saveBeneficiaryRecord(e) {
            e.preventDefault();
            const ben = {
                id: `BEN-${String(Date.now()).slice(-4)}`,
                projectId: document.getElementById('ben-project-id').value,
                name: document.getElementById('ben-name').value,
                ward: document.getElementById('ben-ward').value,
                category: document.getElementById('ben-category').value,
                gender: document.getElementById('ben-gender').value,
                sanctionDetails: document.getElementById('ben-sanctioned-details').value,
                sanctionDate: new Date().toISOString().split('T')[0],
                status: "Sanctioned"
            };

            appState.beneficiaries.push(ben);
            logAudit(`Added beneficiary ${ben.name} to project ${ben.projectId}`);
            saveStateToStorage();
            closeBeneficiaryModal();
            renderBeneficiaryTable();
            showNotification("Beneficiary added.", "success");
        }

        function deleteBeneficiary(id) {
            if (confirm("Delete this beneficiary record?")) {
                appState.beneficiaries = appState.beneficiaries.filter(x => x.id !== id);
                saveStateToStorage();
                renderBeneficiaryTable();
                showNotification("Beneficiary record removed.", "warning");
            }
        }

        function exportBeneficiariesExcel() {
            const data = appState.beneficiaries.map(b => {
                const p = appState.projects.find(x => x.id === b.projectId);
                return {
                    "Project Code": p ? p.code : b.projectId,
                    "Project Name": p ? p.name : '',
                    "Beneficiary Name": b.name,
                    "Ward": b.ward,
                    "Category": b.category,
                    "Gender": b.gender,
                    "Sanctioned Benefit": b.sanctionDetails,
                    "Sanction Date": b.sanctionDate,
                    "Status": b.status
                };
            });
            const ws = XLSX.utils.json_to_sheet(data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Beneficiaries");
            XLSX.writeFile(wb, "Panchayat_Beneficiary_Register.xlsx");
        }

        // ==========================================
        // MODULE 6: FIELD INSPECTION & MONITORING
        // ==========================================
        function populateInspectionProjectSelect() {
            const modalSel = document.getElementById('insp-project-id');
            if (modalSel) {
                modalSel.innerHTML = appState.projects.map(p => `<option value="${p.id}">${p.code} - ${p.name}</option>`).join('');
            }
        }

        function renderInspectionTable() {
            const tbody = document.getElementById('inspection-tbody');
            if (!tbody) return;

            if (appState.inspections.length === 0) {
                tbody.innerHTML = `<tr><td colspan="8" class="text-center py-6 text-slate-400">No inspection records logged yet.</td></tr>`;
                return;
            }

            tbody.innerHTML = appState.inspections.map(i => {
                const p = appState.projects.find(x => x.id === i.projectId);
                return `
                    <tr class="hover:bg-slate-50">
                        <td class="py-2.5 px-3 font-medium">${i.date}</td>
                        <td class="py-2.5 px-3">
                            <div class="font-bold text-blue-900">${p ? p.code : i.projectId}</div>
                            <div class="text-[10px] text-slate-500 line-clamp-1">${p ? p.name : ''}</div>
                        </td>
                        <td class="py-2.5 px-3 font-medium">${i.officer}</td>
                        <td class="py-2.5 px-3 text-center font-bold text-purple-900">${i.progressObserved}%</td>
                        <td class="py-2.5 px-3 text-center"><span class="px-2 py-0.5 rounded text-[10px] bg-slate-100 font-bold">${i.quality}</span></td>
                        <td class="py-2.5 px-3">${i.result}</td>
                        <td class="py-2.5 px-3 text-slate-600">${i.defects || '-'}</td>
                        <td class="py-2.5 px-3 text-center"><span class="px-2 py-0.5 rounded text-[10px] bg-emerald-100 text-emerald-800 font-bold">${i.compliance}</span></td>
                    </tr>
                `;
            }).join('');
        }

        function openInspectionModal() {
            document.getElementById('inspection-form').reset();
            document.getElementById('insp-date').value = new Date().toISOString().split('T')[0];
            document.getElementById('modal-inspection').classList.remove('hidden');
        }

        function closeInspectionModal() {
            document.getElementById('modal-inspection').classList.add('hidden');
        }

        function saveInspectionRecord(e) {
            e.preventDefault();
            const insp = {
                id: `INS-${String(Date.now()).slice(-4)}`,
                projectId: document.getElementById('insp-project-id').value,
                date: document.getElementById('insp-date').value,
                officer: document.getElementById('insp-officer').value,
                progressObserved: parseFloat(document.getElementById('insp-prog-obs').value) || 0,
                quality: document.getElementById('insp-quality').value,
                result: "Satisfactory",
                defects: document.getElementById('insp-defects').value,
                compliance: "Compliant"
            };

            appState.inspections.unshift(insp);
            logAudit(`Recorded inspection for project ${insp.projectId} by ${insp.officer}`);
            saveStateToStorage();
            closeInspectionModal();
            renderInspectionTable();
            showNotification("Field inspection saved.", "success");
        }

        // ==========================================
        // MODULE 7: PROJECT EVALUATION MODULE
        // ==========================================
        function populateEvaluationProjectSelect() {
            const sel = document.getElementById('eval-project-select');
            if (!sel) return;
            sel.innerHTML = '<option value="">-- Select Project to Evaluate --</option>' + 
                appState.projects.map(p => `<option value="${p.id}">${p.code} - ${p.name}</option>`).join('');

            renderEvaluationCriteriaForm();
        }

        function renderEvaluationCriteriaForm() {
            const container = document.getElementById('eval-criteria-container');
            if (!container) return;

            container.innerHTML = DEFAULT_EVAL_CRITERIA.map((crit, idx) => `
                <div class="p-2 bg-slate-50 border border-slate-200 rounded-lg flex items-center justify-between gap-2">
                    <span class="text-[11px] text-slate-700 font-medium">${idx+1}. ${crit}</span>
                    <select id="eval-crit-${idx}" onchange="updateLiveEvaluationScore()" class="border border-slate-300 rounded p-1 text-xs bg-white font-bold">
                        <option value="0">0 - Not Assessed</option>
                        <option value="1">1 - Very Poor</option>
                        <option value="2">2 - Poor</option>
                        <option value="3" selected>3 - Satisfactory</option>
                        <option value="4">4 - Good</option>
                        <option value="5">5 - Excellent</option>
                    </select>
                </div>
            `).join('');
        }

        function loadProjectEvaluation(projectId) {
            const card = document.getElementById('eval-score-card');
            if (!projectId) {
                card.innerHTML = `<p class="text-xs text-slate-500">Select a project to evaluate performance</p>`;
                return;
            }

            const p = appState.projects.find(x => x.id === projectId);
            const ev = appState.evaluations[projectId];

            document.getElementById('eval-date').value = ev ? ev.date : new Date().toISOString().split('T')[0];
            document.getElementById('eval-evaluator-name').value = ev ? ev.evaluator : "";
            document.getElementById('eval-achievements').value = ev ? ev.achievements : "";
            document.getElementById('eval-problems').value = ev ? ev.problems : "";

            if (ev && ev.scores) {
                Object.keys(ev.scores).forEach(idx => {
                    const el = document.getElementById(`eval-crit-${idx}`);
                    if (el) el.value = ev.scores[idx];
                });
            }

            updateLiveEvaluationScore();
        }

        function updateLiveEvaluationScore() {
            let totalObtained = 0;
            let assessedCount = 0;

            DEFAULT_EVAL_CRITERIA.forEach((crit, idx) => {
                const val = parseInt(document.getElementById(`eval-crit-${idx}`)?.value || 0);
                if (val > 0) {
                    assessedCount++;
                    totalObtained += val;
                }
            });

            const maxPossible = assessedCount * 5;
            const pct = maxPossible > 0 ? ((totalObtained / maxPossible) * 100).toFixed(2) : 0;
            const grade = calculateGrade(pct, assessedCount);

            document.getElementById('eval-live-summary').innerText = `Total Score: ${totalObtained}/${maxPossible} (${pct}%) | Grade: ${grade}`;

            const projectId = document.getElementById('eval-project-select')?.value;
            const card = document.getElementById('eval-score-card');
            if (projectId && card) {
                card.innerHTML = `
                    <div class="space-y-1">
                        <div class="text-xs text-slate-500 uppercase tracking-wider">Evaluation Score</div>
                        <div class="text-3xl font-extrabold text-amber-600">${pct}%</div>
                        <div class="inline-block px-3 py-1 bg-amber-100 text-amber-900 rounded-full font-bold text-xs">${grade}</div>
                        <div class="text-[10px] text-slate-400 mt-1">Assessed Criteria: ${assessedCount} / ${DEFAULT_EVAL_CRITERIA.length}</div>
                    </div>
                `;
            }
        }

        function calculateGrade(pct, assessedCount) {
            if (assessedCount === 0) return "Not Evaluated";
            if (pct >= 90) return "Excellent";
            if (pct >= 75) return "Very Good";
            if (pct >= 60) return "Good";
            if (pct >= 40) return "Needs Improvement";
            return "Unsatisfactory";
        }

        function saveProjectEvaluation(e) {
            e.preventDefault();
            const projectId = document.getElementById('eval-project-select').value;
            if (!projectId) {
                alert("Please select a project.");
                return;
            }

            const scoresObj = {};
            let totalObtained = 0;
            let assessedCount = 0;

            DEFAULT_EVAL_CRITERIA.forEach((crit, idx) => {
                const val = parseInt(document.getElementById(`eval-crit-${idx}`).value);
                scoresObj[idx] = val;
                if (val > 0) {
                    assessedCount++;
                    totalObtained += val;
                }
            });

            const maxPossible = assessedCount * 5;
            const pct = maxPossible > 0 ? parseFloat(((totalObtained / maxPossible) * 100).toFixed(2)) : 0;
            const grade = calculateGrade(pct, assessedCount);

            const evalRecord = {
                evalId: `EV-${String(Date.now()).slice(-4)}`,
                projectId: projectId,
                type: document.getElementById('eval-type').value,
                evaluator: document.getElementById('eval-evaluator-name').value,
                date: document.getElementById('eval-date').value,
                scores: scoresObj,
                totalScore: totalObtained,
                maxScore: maxPossible,
                percentage: pct,
                grade: grade,
                achievements: document.getElementById('eval-achievements').value,
                problems: document.getElementById('eval-problems').value
            };

            appState.evaluations[projectId] = evalRecord;
            logAudit(`Saved evaluation score (${pct}%, ${grade}) for project ${projectId}`);
            saveStateToStorage();
            showNotification(`Evaluation for project saved successfully! Grade: ${grade}`, "success");
        }

        // ==========================================
        // MODULE 8: OUTPUT & OUTCOME REGISTER
        // ==========================================
        function populateOutcomeProjectSelect() {
            const modalSel = document.getElementById('out-project-id');
            if (modalSel) {
                modalSel.innerHTML = appState.projects.map(p => `<option value="${p.id}">${p.code} - ${p.name}</option>`).join('');
            }
        }

        function renderOutcomeTable() {
            const tbody = document.getElementById('outcome-tbody');
            if (!tbody) return;

            if (appState.outcomes.length === 0) {
                tbody.innerHTML = `<tr><td colspan="8" class="text-center py-6 text-slate-400">No output/outcome indicators registered.</td></tr>`;
                return;
            }

            tbody.innerHTML = appState.outcomes.map(o => {
                const p = appState.projects.find(x => x.id === o.projectId);
                const pct = o.target ? ((o.achieved / o.target) * 100).toFixed(1) : 0;
                return `
                    <tr class="hover:bg-slate-50">
                        <td class="py-2.5 px-3">
                            <div class="font-bold text-blue-900">${p ? p.code : o.projectId}</div>
                            <div class="text-[10px] text-slate-500 line-clamp-1">${p ? p.name : ''}</div>
                        </td>
                        <td class="py-2.5 px-3 font-semibold">${o.name} (${o.unit || 'Units'})</td>
                        <td class="py-2.5 px-3 text-right">${o.baseline || 0}</td>
                        <td class="py-2.5 px-3 text-right font-bold">${o.target}</td>
                        <td class="py-2.5 px-3 text-right font-bold text-emerald-700">${o.achieved}</td>
                        <td class="py-2.5 px-3 text-center font-bold text-purple-900">${pct}%</td>
                        <td class="py-2.5 px-3">Direct: ${o.directBen || 0} | Indirect: ${o.indirectBen || 0}</td>
                        <td class="py-2.5 px-3 text-center"><span class="px-2 py-0.5 rounded text-[10px] bg-emerald-100 text-emerald-800 font-bold">${o.sustainability || 'High'}</span></td>
                    </tr>
                `;
            }).join('');
        }

        function openOutcomeModal() {
            document.getElementById('outcome-form').reset();
            document.getElementById('modal-outcome').classList.remove('hidden');
        }

        function closeOutcomeModal() {
            document.getElementById('modal-outcome').classList.add('hidden');
        }

        function saveOutcomeRecord(e) {
            e.preventDefault();
            const record = {
                id: `OUT-${String(Date.now()).slice(-4)}`,
                projectId: document.getElementById('out-project-id').value,
                name: document.getElementById('out-name').value,
                unit: document.getElementById('out-unit').value,
                target: parseFloat(document.getElementById('out-target').value) || 0,
                achieved: parseFloat(document.getElementById('out-achieved').value) || 0,
                directBen: 150,
                indirectBen: 300,
                sustainability: "High"
            };

            appState.outcomes.push(record);
            saveStateToStorage();
            closeOutcomeModal();
            renderOutcomeTable();
            showNotification("Outcome indicator recorded.", "success");
        }

        // ==========================================
        // MODULE 9: REVIEW & CORRECTIVE ACTION
        // ==========================================
        function getFlaggedProjects() {
            const flagged = [];
            const today = new Date().toISOString().split('T')[0];

            appState.projects.forEach(p => {
                const spent = (p.expenditures || []).reduce((acc, e) => acc + e.amount, 0);

                // Check 1: Delayed target completion
                if (p.targetDate && p.targetDate < today && p.status !== 'Completed' && p.status !== 'Financially Closed') {
                    flagged.push({
                        project: p,
                        issue: "Target completion date passed but work incomplete",
                        severity: "High",
                        action: "Issue urgent show-cause notice to contractor / officer",
                        deadline: "7 Days"
                    });
                }
                // Check 2: Cost overrun
                if (spent > p.cost) {
                    flagged.push({
                        project: p,
                        issue: `Expenditure (₹${spent}) exceeds sanctioned cost (₹${p.cost})`,
                        severity: "Critical",
                        action: "Freeze further voucher processing until AS revision",
                        deadline: "Immediate"
                    });
                }
            });

            return flagged;
        }

        function renderCorrectiveActionRegister() {
            const tbody = document.getElementById('corrective-tbody');
            if (!tbody) return;

            const flagged = getFlaggedProjects();

            if (flagged.length === 0) {
                tbody.innerHTML = `<tr><td colspan="8" class="text-center py-6 text-emerald-600 font-medium"><i class="fa-solid fa-circle-check mr-1"></i> No project review alerts flagged. All projects operating within normal limits.</td></tr>`;
                return;
            }

            tbody.innerHTML = flagged.map(f => `
                <tr class="hover:bg-slate-50">
                    <td class="py-2.5 px-3">
                        <div class="font-bold text-blue-900">${f.project.code}</div>
                        <div class="text-[10px] text-slate-500 line-clamp-1">${f.project.name}</div>
                    </td>
                    <td class="py-2.5 px-3 text-red-700 font-medium">${f.issue}</td>
                    <td class="py-2.5 px-3 text-center">
                        <span class="px-2 py-0.5 rounded text-[10px] bg-red-100 text-red-800 font-bold">${f.severity}</span>
                    </td>
                    <td class="py-2.5 px-3">${f.action}</td>
                    <td class="py-2.5 px-3 font-medium">${f.project.officer}</td>
                    <td class="py-2.5 px-3 text-center font-bold text-rose-800">${f.deadline}</td>
                    <td class="py-2.5 px-3 text-center"><span class="px-2 py-0.5 rounded text-[10px] bg-amber-100 text-amber-800 font-bold">Action Pending</span></td>
                    <td class="py-2.5 px-3 text-center">
                        <button onclick="showNotification('Directive issued to officer', 'info')" class="px-2 py-1 bg-red-700 text-white rounded text-[10px] hover:bg-red-800">Issue Directive</button>
                    </td>
                </tr>
            `).join('');
        }

        // ==========================================
        // MODULE 10: DOCUMENT MANAGEMENT
        // ==========================================
        function populateDocumentProjectSelect() {
            const modalSel = document.getElementById('doc-project-id');
            if (modalSel) {
                modalSel.innerHTML = appState.projects.map(p => `<option value="${p.id}">${p.code} - ${p.name}</option>`).join('');
            }
        }

        function renderDocumentTable() {
            const tbody = document.getElementById('documents-tbody');
            if (!tbody) return;

            if (appState.documents.length === 0) {
                tbody.innerHTML = `<tr><td colspan="7" class="text-center py-6 text-slate-400">No documents uploaded to vault.</td></tr>`;
                return;
            }

            tbody.innerHTML = appState.documents.map(d => {
                const p = appState.projects.find(x => x.id === d.projectId);
                return `
                    <tr class="hover:bg-slate-50">
                        <td class="py-2.5 px-3">
                            <div class="font-bold text-blue-900">${p ? p.code : d.projectId}</div>
                            <div class="text-[10px] text-slate-500 line-clamp-1">${p ? p.name : ''}</div>
                        </td>
                        <td class="py-2.5 px-3 font-semibold">${d.type}</td>
                        <td class="py-2.5 px-3 font-mono font-bold">${d.refNo}</td>
                        <td class="py-2.5 px-3 text-center">${d.date}</td>
                        <td class="py-2.5 px-3">${d.uploadedBy}</td>
                        <td class="py-2.5 px-3 text-center"><span class="px-2 py-0.5 rounded text-[10px] bg-emerald-100 text-emerald-800 font-bold">${d.status}</span></td>
                        <td class="py-2.5 px-3 text-center">
                            <a href="${d.url || '#'}" target="_blank" class="text-blue-600 hover:underline"><i class="fa-solid fa-link"></i> View Doc</a>
                        </td>
                    </tr>
                `;
            }).join('');
        }

        function openDocumentModal() {
            document.getElementById('document-form').reset();
            document.getElementById('modal-document').classList.remove('hidden');
        }

        function closeDocumentModal() {
            document.getElementById('modal-document').classList.add('hidden');
        }

        function saveDocumentRecord(e) {
            e.preventDefault();
            const doc = {
                id: `DOC-${String(Date.now()).slice(-4)}`,
                projectId: document.getElementById('doc-project-id').value,
                type: document.getElementById('doc-type').value,
                refNo: document.getElementById('doc-ref-no').value,
                date: new Date().toISOString().split('T')[0],
                uploadedBy: appState.userRole,
                status: "Verified",
                url: document.getElementById('doc-url').value
            };

            appState.documents.push(doc);
            saveStateToStorage();
            closeDocumentModal();
            renderDocumentTable();
            showNotification("Document vault reference added.", "success");
        }

        // ==========================================
        // MODULE 11: REPORTS AND DOWNLOADS ENGINE
        // ==========================================
        function generateReportPreview() {
            const rptType = document.getElementById('report-type-select')?.value || '1';
            const sector = document.getElementById('report-filter-sector')?.value || 'ALL';
            const fy = document.getElementById('report-filter-fy')?.value || 'ALL';

            const filtered = appState.projects.filter(p => {
                const matchSec = (sector === 'ALL' || p.sector === sector);
                const matchFy = (fy === 'ALL' || p.fy === fy);
                return matchSec && matchFy;
            });

            document.getElementById('rpt-header-panchayat').innerText = appState.profile.panchayatName;
            document.getElementById('rpt-signatory-1').innerText = `Prepared by: ${appState.profile.sig1}`;
            document.getElementById('rpt-signatory-2').innerText = `Approved by: ${appState.profile.sig2}`;

            const thead = document.getElementById('report-table-head');
            const tbody = document.getElementById('report-table-body');
            const titleDisplay = document.getElementById('rpt-title-display');

            if (rptType === '1') {
                titleDisplay.innerText = "FOURTEENTH FIVE-YEAR PLAN – COMPLETE PROJECT REGISTER";
                thead.innerHTML = `
                    <tr>
                        <th class="py-2 px-2 border">Code</th>
                        <th class="py-2 px-2 border">Project Name</th>
                        <th class="py-2 px-2 border">Sector</th>
                        <th class="py-2 px-2 border">Officer</th>
                        <th class="py-2 px-2 border text-right">Cost (₹)</th>
                        <th class="py-2 px-2 border text-center">Physical %</th>
                        <th class="py-2 px-2 border text-center">Status</th>
                    </tr>
                `;
                tbody.innerHTML = filtered.map(p => `
                    <tr>
                        <td class="py-2 px-2 border font-mono font-bold">${p.code}</td>
                        <td class="py-2 px-2 border font-medium">${p.name}</td>
                        <td class="py-2 px-2 border">${p.sector}</td>
                        <td class="py-2 px-2 border">${p.officer}</td>
                        <td class="py-2 px-2 border text-right font-bold">₹${p.cost.toLocaleString('en-IN')}</td>
                        <td class="py-2 px-2 border text-center font-bold">${p.physicalPercent}%</td>
                        <td class="py-2 px-2 border text-center">${p.status}</td>
                    </tr>
                `).join('');
            } else if (rptType === '16') {
                titleDisplay.innerText = "FOURTEENTH FIVE-YEAR PLAN – EVALUATION SCORES & GRADES SUMMARY";
                thead.innerHTML = `
                    <tr>
                        <th class="py-2 px-2 border">Code</th>
                        <th class="py-2 px-2 border">Project Name</th>
                        <th class="py-2 px-2 border">Evaluator</th>
                        <th class="py-2 px-2 border text-center">Total Score</th>
                        <th class="py-2 px-2 border text-center">Percentage</th>
                        <th class="py-2 px-2 border text-center">Grade</th>
                    </tr>
                `;
                tbody.innerHTML = filtered.map(p => {
                    const ev = appState.evaluations[p.id];
                    return `
                        <tr>
                            <td class="py-2 px-2 border font-mono font-bold">${p.code}</td>
                            <td class="py-2 px-2 border font-medium">${p.name}</td>
                            <td class="py-2 px-2 border">${ev ? ev.evaluator : 'Pending'}</td>
                            <td class="py-2 px-2 border text-center font-bold">${ev ? ev.totalScore : 0} / ${ev ? ev.maxScore : 0}</td>
                            <td class="py-2 px-2 border text-center font-bold">${ev ? ev.percentage : 0}%</td>
                            <td class="py-2 px-2 border text-center font-bold">${ev ? ev.grade : 'Unevaluated'}</td>
                        </tr>
                    `;
                }).join('');
            } else {
                titleDisplay.innerText = `FOURTEENTH FIVE-YEAR PLAN – REPORT TYPE #${rptType}`;
                thead.innerHTML = `<tr><th class="py-2 px-2 border">Code</th><th class="py-2 px-2 border">Project Name</th><th class="py-2 px-2 border">Sector</th><th class="py-2 px-2 border text-right">Sanctioned Cost</th><th class="py-2 px-2 border text-center">Status</th></tr>`;
                tbody.innerHTML = filtered.map(p => `
                    <tr>
                        <td class="py-2 px-2 border font-mono font-bold">${p.code}</td>
                        <td class="py-2 px-2 border">${p.name}</td>
                        <td class="py-2 px-2 border">${p.sector}</td>
                        <td class="py-2 px-2 border text-right font-bold">₹${p.cost.toLocaleString('en-IN')}</td>
                        <td class="py-2 px-2 border text-center">${p.status}</td>
                    </tr>
                `).join('');
            }
        }

        function downloadReportExcel() {
            const table = document.querySelector('#report-preview-container table');
            if (!table) return;
            const wb = XLSX.utils.table_to_book(table);
            XLSX.writeFile(wb, "Panchayat_14th_Plan_Report.xlsx");
        }

        function downloadReportPDF() {
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

            doc.setFontSize(10);
            doc.text(appState.profile.panchayatName.toUpperCase(), 14, 15);
            doc.setFontSize(12);
            doc.setFont(undefined, 'bold');
            doc.text(document.getElementById('rpt-title-display').innerText, 14, 22);

            doc.autoTable({
                html: '#report-preview-container table',
                startY: 28,
                theme: 'grid',
                styles: { fontSize: 8 }
            });

            doc.save("Panchayat_14th_Plan_Report.pdf");
        }

        // ==========================================
        // MODULE 13: MASTER SETTINGS
        // ==========================================
        function renderMasterSettings() {
            document.getElementById('cfg-panchayat-name').value = appState.profile.panchayatName;
            document.getElementById('cfg-block-name').value = appState.profile.blockName;
            document.getElementById('cfg-district-name').value = appState.profile.districtName;
            document.getElementById('cfg-sig1').value = appState.profile.sig1;
            document.getElementById('cfg-sig2').value = appState.profile.sig2;
            renderAuditLogBox();
        }

        function savePanchayatProfile(e) {
            e.preventDefault();
            appState.profile.panchayatName = document.getElementById('cfg-panchayat-name').value;
            appState.profile.blockName = document.getElementById('cfg-block-name').value;
            appState.profile.districtName = document.getElementById('cfg-district-name').value;
            appState.profile.sig1 = document.getElementById('cfg-sig1').value;
            appState.profile.sig2 = document.getElementById('cfg-sig2').value;

            document.getElementById('header-panchayat-name').innerText = appState.profile.panchayatName.toUpperCase();
            saveStateToStorage();
            logAudit("Updated Grama Panchayat Master Profile.");
            showNotification("Panchayat Profile Settings Saved.", "success");
        }

        function exportDatabaseJSON() {
            const jsonStr = JSON.stringify(appState, null, 2);
            const blob = new Blob([jsonStr], { type: "application/json" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Panchayat_14th_Plan_DB_Backup_${new Date().toISOString().split('T')[0]}.json`;
            a.click();
        }

        function importDatabaseJSON(e) {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(evt) {
                try {
                    const data = JSON.parse(evt.target.result);
                    appState = data;
                    saveStateToStorage();
                    alert("Database imported successfully!");
                    location.reload();
                } catch(err) {
                    alert("Invalid database file format.");
                }
            };
            reader.readAsText(file);
        }
    </script>
</body>
</html>[PROJECT PLAN.html](https://github.com/user-attachments/files/34221645/PROJECT.PLAN.html)
# 0
health
