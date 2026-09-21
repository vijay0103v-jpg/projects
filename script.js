/**
 * ============================================================================
 * Project Tracking & Management System - Main Script
 * Full Client-Side SPA with LocalStorage & SheetJS Excel Integration
 * ============================================================================
 */

(function () {
    'use strict';

    // Current Reference Date for manufacturing calculations (Real-time Today)
    function getTodayDateStr() {
        const d = new Date();
        const yyyy = d.getFullYear();
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${yyyy}-${mm}-${dd}`;
    }

    let SYSTEM_DATE_STR = getTodayDateStr();
    let SYSTEM_DATE = new Date(SYSTEM_DATE_STR + 'T00:00:00');

    function refreshSystemDate() {
        SYSTEM_DATE_STR = getTodayDateStr();
        SYSTEM_DATE = new Date(SYSTEM_DATE_STR + 'T00:00:00');
    }

    // LocalStorage Keys
    const STORAGE_KEY_PROJECTS = 'apex_mfg_projects_db_v2';
    const STORAGE_KEY_MATERIALS = 'apex_mfg_materials_db_v2';
    const STORAGE_KEY_TIMELINE = 'apex_mfg_timeline_db_v2';
    const STORAGE_KEY_BLOCKERS = 'apex_mfg_blockers_db_v2';
    const STORAGE_KEY_CURRENT_USER = 'apex_mfg_current_user_v2';

    // Application Global State
    let projects = [];
    let materials = [];
    let timelineEntries = [];
    let blockers = [];
    let currentUser = null;

    // Table State Parameters
    let tableSearchQuery = '';
    let filterStatus = '';
    let filterPriority = '';
    let filterLead = '';
    let filterCustomer = '';
    let activeQuickFilter = 'all';
    let sortColumn = 'id';
    let sortDirection = 'asc'; // 'asc' or 'desc'
    let currentPage = 1;
    let pageSize = 25;

    // Column Configuration & Visibility State
    const columnsConfig = [
        { id: 'sno', label: 'S.No', visible: true },
        { id: 'name', label: 'Project Name', visible: true, sortable: true },
        { id: 'type', label: 'New / Repeated', visible: true, sortable: true },
        { id: 'customer', label: 'Customer', visible: true, sortable: true },
        { id: 'poNumber', label: 'PO Number', visible: true, sortable: true },
        { id: 'priority', label: 'Priority', visible: true, sortable: true },
        { id: 'lead', label: 'Project Lead', visible: true, sortable: true },
        { id: 'poDate', label: 'PO Date', visible: false, sortable: true },
        { id: 'committedDate', label: 'Committed Date', visible: true, sortable: true },
        { id: 'revisedDate', label: 'Revised Date', visible: true, sortable: true },
        { id: 'deliveredDate', label: 'Actual Delivered Date', visible: false, sortable: true },
        { id: 'totalQty', label: 'Total Qty', visible: true, sortable: true },
        { id: 'deliveredQty', label: 'Delivered Qty', visible: true, sortable: true },
        { id: 'pendingQty', label: 'Pending Qty', visible: true, sortable: true },
        { id: 'matAvailability', label: 'Mat. Avail %', visible: true, sortable: true },
        { id: 'status', label: 'Current Status', visible: true, sortable: true },
        { id: 'completion', label: 'Overall %', visible: true, sortable: true },
        { id: 'invoiceNumber', label: 'Invoice No.', visible: false, sortable: true },
        { id: 'actions', label: 'Actions', visible: true, sortable: false }
    ];

    // Chart.js Instances
    let chartStatusInstance = null;
    let chartLeadInstance = null;
    let chartPriorityInstance = null;
    let chartCustomerInstance = null;
    let chartMonthlyInstance = null;

    // Modals References
    let projectModalBs = null;
    let detailModalBs = null;
    let importModalBs = null;
    let addMaterialModalBs = null;
    let notificationsModalBs = null;
    let blockerModalBs = null;
    let addBatchModalBs = null;

    // Default Stages List for Manufacturing Progress
    const PIPELINE_STAGES = [
        "Design", "Procurement", "CNC", "Assembly", 
        "Wiring", "Testing", "Delivery", "Completed"
    ];

    const STORAGE_KEY_INITIALIZED = 'apex_mfg_initialized_v2';

    /* ==========================================================================
       1. DATA STORAGE & INITIALIZATION
       ========================================================================== */

    function initData() {
        const isInitialized = localStorage.getItem(STORAGE_KEY_INITIALIZED);

        if (!isInitialized) {
            seedSampleData();
            localStorage.setItem(STORAGE_KEY_INITIALIZED, 'true');
            saveAllData();
            return;
        }

        const storedProj = localStorage.getItem(STORAGE_KEY_PROJECTS);
        const storedMat = localStorage.getItem(STORAGE_KEY_MATERIALS);
        const storedTL = localStorage.getItem(STORAGE_KEY_TIMELINE);
        const storedBlk = localStorage.getItem(STORAGE_KEY_BLOCKERS);

        try {
            projects = storedProj !== null ? JSON.parse(storedProj) : [];
            materials = storedMat !== null ? JSON.parse(storedMat) : [];
            timelineEntries = storedTL !== null ? JSON.parse(storedTL) : [];
            blockers = storedBlk !== null ? JSON.parse(storedBlk) : [];
        } catch (e) {
            projects = [];
            materials = [];
            timelineEntries = [];
            blockers = [];
        }

        if (!Array.isArray(projects)) projects = [];
        if (!Array.isArray(materials)) materials = [];
        if (!Array.isArray(timelineEntries)) timelineEntries = [];
        if (!Array.isArray(blockers)) blockers = [];

        recalculateAllProjectsState();
        saveAllData();
    }

    function saveAllData() {
        localStorage.setItem(STORAGE_KEY_PROJECTS, JSON.stringify(projects));
        localStorage.setItem(STORAGE_KEY_MATERIALS, JSON.stringify(materials));
        localStorage.setItem(STORAGE_KEY_TIMELINE, JSON.stringify(timelineEntries));
        localStorage.setItem(STORAGE_KEY_BLOCKERS, JSON.stringify(blockers));
        localStorage.setItem(STORAGE_KEY_INITIALIZED, 'true');
        updateStorageStatusText();
        populateDropdownOptions();
    }

    function updateStorageStatusText() {
        const textEl = document.getElementById('storage-status-text');
        const badgeEl = document.getElementById('nav-total-projects-badge');
        const blkBadge = document.getElementById('nav-active-blockers-badge');
        const delivBadge = document.getElementById('nav-weekly-deliveries-badge');

        if (textEl) textEl.innerText = `${projects.length} Projects stored locally`;
        if (badgeEl) badgeEl.innerText = projects.length;
        if (blkBadge) blkBadge.innerText = blockers.filter(b => b.status !== "Resolved").length;
        if (delivBadge) delivBadge.innerText = getThisWeekDeliveriesCount();
    }

    /* ==========================================================================
       1B. AUTHENTICATION & PIN SECURITY CONTROLLER
       ========================================================================== */

    const STORAGE_KEY_APP_PIN = 'apex_mfg_app_pin_v2';
    const STORAGE_KEY_PIN_UNLOCKED = 'apex_mfg_pin_unlocked_v2';
    let enteredPinBuffer = '';

    function getStoredPin() {
        return localStorage.getItem(STORAGE_KEY_APP_PIN) || '1234';
    }

    function updatePinHintDisplay() {
        const hintEl = document.getElementById('pin-hint-text');
        if (hintEl) {
            const stored = getStoredPin();
            if (stored === '1234') {
                hintEl.innerHTML = `<i class="fa-solid fa-key me-1 text-primary"></i> Default Access PIN: <strong class="text-dark">1234</strong>`;
            } else {
                hintEl.innerHTML = `<i class="fa-solid fa-shield-halved me-1 text-success"></i> Custom Security PIN Active`;
            }
        }
    }

    function initAuth() {
        const storedUser = localStorage.getItem(STORAGE_KEY_CURRENT_USER);
        if (storedUser) {
            try {
                currentUser = JSON.parse(storedUser);
            } catch (e) {
                currentUser = null;
            }
        }

        if (!currentUser) {
            currentUser = {
                name: 'Project Manager',
                role: 'Project Manager',
                initials: 'PM'
            };
            localStorage.setItem(STORAGE_KEY_CURRENT_USER, JSON.stringify(currentUser));
        }

        updateUserUI();
        updatePinHintDisplay();

        const isUnlocked = sessionStorage.getItem(STORAGE_KEY_PIN_UNLOCKED) === 'true';
        const loginScr = document.getElementById('login-screen');

        if (isUnlocked) {
            if (loginScr) loginScr.classList.add('d-none');
        } else {
            if (loginScr) loginScr.classList.remove('d-none');
            resetPinBuffer();
        }
    }

    function resetPinBuffer() {
        enteredPinBuffer = '';
        updatePinDotsDisplay();
        const errText = document.getElementById('pin-error-text');
        if (errText) errText.classList.add('d-none');
    }

    function updatePinDotsDisplay() {
        const dots = document.querySelectorAll('#pin-dots-display .pin-dot');
        dots.forEach((dot, idx) => {
            if (idx < enteredPinBuffer.length) {
                dot.classList.add('filled');
            } else {
                dot.classList.remove('filled');
            }
        });
    }

    function handlePinKeyInput(key) {
        const errText = document.getElementById('pin-error-text');
        if (errText) errText.classList.add('d-none');

        if (key >= '0' && key <= '9') {
            if (enteredPinBuffer.length < 4) {
                enteredPinBuffer += key;
                updatePinDotsDisplay();
                if (enteredPinBuffer.length === 4) {
                    setTimeout(verifyEnteredPin, 150);
                }
            }
        } else if (key === 'backspace') {
            enteredPinBuffer = enteredPinBuffer.slice(0, -1);
            updatePinDotsDisplay();
        } else if (key === 'clear') {
            resetPinBuffer();
        }
    }

    function verifyEnteredPin() {
        const targetPin = getStoredPin();
        if (enteredPinBuffer === targetPin) {
            sessionStorage.setItem(STORAGE_KEY_PIN_UNLOCKED, 'true');
            const loginScr = document.getElementById('login-screen');
            if (loginScr) loginScr.classList.add('d-none');
            resetPinBuffer();
            renderActiveView(getCurrentViewName());
        } else {
            const errText = document.getElementById('pin-error-text');
            if (errText) errText.classList.remove('d-none');

            const card = document.querySelector('.pin-card-container');
            if (card) {
                card.classList.add('pin-shake');
                setTimeout(() => card.classList.remove('pin-shake'), 400);
            }
            enteredPinBuffer = '';
            setTimeout(updatePinDotsDisplay, 250);
        }
    }

    function lockApp() {
        sessionStorage.setItem(STORAGE_KEY_PIN_UNLOCKED, 'false');
        resetPinBuffer();
        const loginScr = document.getElementById('login-screen');
        if (loginScr) loginScr.classList.remove('d-none');
    }

    function promptChangePin() {
        const currentStored = getStoredPin();
        const inputOld = prompt("Enter Current PIN (Default is 1234):", "");
        if (inputOld === null) return;

        if (inputOld.trim() !== currentStored) {
            alert("Incorrect current PIN. PIN change failed.");
            return;
        }

        const newPin = prompt("Enter New 4-Digit Security PIN (numbers only):", "");
        if (!newPin || !/^\d{4}$/.exec(newPin.trim())) {
            alert("Invalid PIN! PIN must be exactly 4 numeric digits (e.g. 5678).");
            return;
        }

        localStorage.setItem(STORAGE_KEY_APP_PIN, newPin.trim());
        updatePinHintDisplay();
        alert(`Security PIN updated successfully! Your custom PIN (${newPin.trim()}) is now saved and active.`);
    }

    function formatDisplayName(input) {
        if (!input) return { name: 'Project Manager', initials: 'PM' };
        let raw = input.includes('@') ? input.split('@')[0] : input;
        
        let textOnly = raw.replace(/\d+/g, ' ').trim();
        if (!textOnly) textOnly = raw;

        const parts = textOnly.split(/[._\-\s]+/).filter(Boolean);
        const cleanName = parts.map(p => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()).join(' ');
        
        let initials = 'PM';
        if (parts.length >= 2) {
            initials = (parts[0][0] + parts[1][0]).toUpperCase();
        } else if (parts.length === 1 && parts[0].length >= 2) {
            initials = parts[0].substring(0, 2).toUpperCase();
        } else if (parts.length === 1) {
            initials = parts[0].toUpperCase();
        }

        return { name: cleanName, initials: initials };
    }

    function getTimeBasedGreeting() {
        const hours = new Date().getHours();
        if (hours >= 5 && hours < 12) {
            return "Good morning";
        } else if (hours >= 12 && hours < 17) {
            return "Good afternoon";
        } else {
            return "Good evening";
        }
    }

    function updateUserUI() {
        refreshSystemDate();
        if (currentUser) {
            const sidebarAvatar = document.getElementById('sidebar-user-avatar');
            const sidebarName = document.getElementById('sidebar-user-name');
            const sidebarRole = document.getElementById('sidebar-user-role');

            const headerAvatar = document.getElementById('header-user-avatar');
            const headerName = document.getElementById('header-user-name');
            const dropdownName = document.getElementById('dropdown-user-name');
            const dropdownRole = document.getElementById('dropdown-user-role');

            if (sidebarAvatar) sidebarAvatar.innerText = currentUser.initials || 'PM';
            if (sidebarName) sidebarName.innerText = currentUser.name;
            if (sidebarRole) sidebarRole.innerText = currentUser.role;

            if (headerAvatar) headerAvatar.innerText = currentUser.initials || 'PM';
            if (headerName) headerName.innerText = currentUser.name;
            if (dropdownName) dropdownName.innerText = currentUser.name;
            if (dropdownRole) dropdownRole.innerText = currentUser.role;

            const welcomeTitle = document.querySelector('.welcome-title');
            if (welcomeTitle) {
                const firstName = currentUser.name ? currentUser.name.split(' ')[0] : 'User';
                const greeting = getTimeBasedGreeting();
                welcomeTitle.innerText = `${greeting}, ${firstName}`;
            }
        }

        const currentDateDisplay = document.getElementById('current-date-display');
        if (currentDateDisplay) {
            currentDateDisplay.innerText = formatDate(SYSTEM_DATE_STR);
        }
    }

    function editUserProfile() {
        const curName = currentUser ? currentUser.name : "Project Manager";
        const curRole = currentUser ? currentUser.role : "Project Manager";

        const newName = prompt("Edit Display Name:", curName);
        if (!newName || !newName.trim()) return;

        const newRole = prompt("Edit Role / Title (Project Manager / Project Coordinator):", curRole);
        const { name, initials } = formatDisplayName(newName.trim());

        currentUser = {
            name: name,
            role: newRole ? newRole.trim() : "Project Manager",
            initials: initials
        };
        localStorage.setItem(STORAGE_KEY_CURRENT_USER, JSON.stringify(currentUser));
        updateUserUI();
    }

    function seedSampleData() {
        projects = [
            {
                id: 101,
                name: "Cell Modem Board",
                type: "New",
                customer: "Enphase",
                poNumber: "ENIND205415",
                description: "Cellular communication modem PCBA with external antenna connector.",
                priority: "4",
                lead: "Raja Rajan",
                poDate: "2026-08-01",
                committedDate: "2026-09-15",
                revisedDate: "2026-09-15",
                totalQty: 500,
                deliveredQty: 150,
                pendingQty: 350,
                matAvailability: 95,
                status: "In Progress",
                completion: 62,
                invoiceNumber: "INV-2026-801",
                stages: {
                    "Design": "Completed", "Procurement": "Completed", "CNC": "Completed",
                    "Assembly": "Completed", "Wiring": "In Progress", "Testing": "Not Started",
                    "Delivery": "Not Started", "Completed": "Not Started"
                }
            },
            {
                id: 102,
                name: "IQSS CB+PCB Assembly",
                type: "Repeated",
                customer: "Enphase",
                poNumber: "ENIND205435",
                description: "Combiner box controller PCB assembly with main switch gear interface.",
                priority: "3",
                lead: "Raja Rajan",
                poDate: "2026-08-05",
                committedDate: "2026-09-10",
                revisedDate: "2026-09-12",
                totalQty: 250,
                deliveredQty: 50,
                pendingQty: 200,
                matAvailability: 88,
                status: "In Progress",
                completion: 50,
                invoiceNumber: "INV-2026-802",
                stages: {
                    "Design": "Completed", "Procurement": "Completed", "CNC": "Completed",
                    "Assembly": "In Progress", "Wiring": "Not Started", "Testing": "Not Started",
                    "Delivery": "Not Started", "Completed": "Not Started"
                }
            },
            {
                id: 103,
                name: "IQ Meter CAN Board",
                type: "New",
                customer: "Enphase",
                poNumber: "ENIND205454",
                description: "CAN bus communication interface card for commercial power meters.",
                priority: "2",
                lead: "Raja Rajan",
                poDate: "2026-07-20",
                committedDate: "2026-09-02", // DELAYED!
                revisedDate: "2026-09-08",
                totalQty: 1000,
                deliveredQty: 300,
                pendingQty: 700,
                matAvailability: 70,
                status: "Delayed",
                completion: 45,
                invoiceNumber: "INV-2026-780",
                stages: {
                    "Design": "Completed", "Procurement": "Completed", "CNC": "In Progress",
                    "Assembly": "Not Started", "Wiring": "Not Started", "Testing": "Not Started",
                    "Delivery": "Not Started", "Completed": "Not Started"
                }
            },
            {
                id: 104,
                name: "IQ Bat R5 Control Board",
                type: "New",
                customer: "Enphase",
                poNumber: "ENUSA228384",
                description: "Battery management system controller board Revision 5.",
                priority: "5",
                lead: "Raja Rajan",
                poDate: "2026-08-10",
                committedDate: "2026-10-01",
                revisedDate: "2026-10-01",
                totalQty: 120,
                deliveredQty: 0,
                pendingQty: 120,
                matAvailability: 100,
                status: "Under Design",
                completion: 25,
                invoiceNumber: "",
                stages: {
                    "Design": "In Progress", "Procurement": "Not Started", "CNC": "Not Started",
                    "Assembly": "Not Started", "Wiring": "Not Started", "Testing": "Not Started",
                    "Delivery": "Not Started", "Completed": "Not Started"
                }
            },
            {
                id: 105,
                name: "Harness Cable Assembly",
                type: "Repeated",
                customer: "Enphase",
                poNumber: "ENIND205490",
                description: "Heavy duty copper wiring harness with molex connectors.",
                priority: "1A",
                lead: "Tamil",
                poDate: "2026-07-15",
                committedDate: "2026-08-28", // DELAYED!
                revisedDate: "2026-09-05",
                totalQty: 800,
                deliveredQty: 200,
                pendingQty: 600,
                matAvailability: 60,
                status: "Delayed",
                completion: 70,
                invoiceNumber: "INV-2026-710",
                stages: {
                    "Design": "Completed", "Procurement": "Completed", "CNC": "Completed",
                    "Assembly": "Completed", "Wiring": "In Progress", "Testing": "Not Started",
                    "Delivery": "Not Started", "Completed": "Not Started"
                }
            },
            {
                id: 106,
                name: "IQ GW 3 Controller",
                type: "New",
                customer: "Enphase",
                poNumber: "ENUSA228410",
                description: "Next-gen Gateway 3 central control enclosure & PCB.",
                priority: "1",
                lead: "Prabhu",
                poDate: "2026-08-12",
                committedDate: "2026-09-25",
                revisedDate: "2026-09-25",
                totalQty: 300,
                deliveredQty: 0,
                pendingQty: 300,
                matAvailability: 100,
                status: "Under Design",
                completion: 20,
                invoiceNumber: "",
                stages: {
                    "Design": "In Progress", "Procurement": "Not Started", "CNC": "Not Started",
                    "Assembly": "Not Started", "Wiring": "Not Started", "Testing": "Not Started",
                    "Delivery": "Not Started", "Completed": "Not Started"
                }
            },
            {
                id: 107,
                name: "IQ Meter Gen2 Box",
                type: "Repeated",
                customer: "Enphase",
                poNumber: "ENIND205501",
                description: "Gen2 Meter outdoor rated aluminum enclosure box.",
                priority: "2",
                lead: "Tamil",
                poDate: "2026-08-18",
                committedDate: "2026-09-18",
                revisedDate: "2026-09-18",
                totalQty: 400,
                deliveredQty: 100,
                pendingQty: 300,
                matAvailability: 90,
                status: "Under Production",
                completion: 55,
                invoiceNumber: "INV-2026-890",
                stages: {
                    "Design": "Completed", "Procurement": "Completed", "CNC": "In Progress",
                    "Assembly": "Not Started", "Wiring": "Not Started", "Testing": "Not Started",
                    "Delivery": "Not Started", "Completed": "Not Started"
                }
            },
            {
                id: 108,
                name: "IQSS GID Module",
                type: "New",
                customer: "Enphase",
                poNumber: "ENUSA228450",
                description: "Grid interconnect device control module.",
                priority: "3",
                lead: "Prabhu",
                poDate: "2026-08-20",
                committedDate: "2026-09-30",
                revisedDate: "2026-09-30",
                totalQty: 150,
                deliveredQty: 0,
                pendingQty: 150,
                matAvailability: 80,
                status: "Under Wiring",
                completion: 60,
                invoiceNumber: "",
                stages: {
                    "Design": "Completed", "Procurement": "Completed", "CNC": "Completed",
                    "Assembly": "Completed", "Wiring": "In Progress", "Testing": "Not Started",
                    "Delivery": "Not Started", "Completed": "Not Started"
                }
            },
            {
                id: 109,
                name: "BI-DI Inverter Board",
                type: "New",
                customer: "Enphase",
                poNumber: "ENIND205520",
                description: "Bidirectional power conversion board assembly.",
                priority: "1A",
                lead: "Raja Rajan",
                poDate: "2026-07-01",
                committedDate: "2026-08-30",
                revisedDate: "2026-08-30",
                totalQty: 200,
                deliveredQty: 200,
                pendingQty: 0,
                matAvailability: 100,
                status: "Completed",
                completion: 100,
                invoiceNumber: "INV-2026-699",
                stages: {
                    "Design": "Completed", "Procurement": "Completed", "CNC": "Completed",
                    "Assembly": "Completed", "Wiring": "Completed", "Testing": "Completed",
                    "Delivery": "Completed", "Completed": "Completed"
                }
            },
            {
                id: 110,
                name: "CAL-A Interface Box",
                type: "Repeated",
                customer: "Enphase",
                poNumber: "ENUSA228499",
                description: "Calibration & testing interface box assembly.",
                priority: "4",
                lead: "Tamil",
                poDate: "2026-08-02",
                committedDate: "2026-09-22",
                revisedDate: "2026-09-22",
                totalQty: 60,
                deliveredQty: 0,
                pendingQty: 60,
                matAvailability: 40,
                status: "On Hold",
                completion: 30,
                invoiceNumber: "",
                stages: {
                    "Design": "Completed", "Procurement": "In Progress", "CNC": "Hold",
                    "Assembly": "Not Started", "Wiring": "Not Started", "Testing": "Not Started",
                    "Delivery": "Not Started", "Completed": "Not Started"
                }
            }
        ];

        materials = [
            { id: 1, projectId: 101, name: "Modem Chipset SMT", partNo: "MDM-9021", reqQty: 500, availQty: 500, pendingQty: 0, supplier: "Qualcomm", expDate: "2026-08-10", status: "Available" },
            { id: 2, projectId: 101, name: "FR4 4-Layer PCB", partNo: "PCB-CM-04", reqQty: 500, availQty: 475, pendingQty: 25, supplier: "JLCPCB", expDate: "2026-08-15", status: "Partially Available" },
            { id: 3, projectId: 103, name: "CAN Transceiver IC", partNo: "TCAN-1042", reqQty: 1000, availQty: 700, pendingQty: 300, supplier: "Texas Inst.", expDate: "2026-09-08", status: "Delayed" },
            { id: 4, projectId: 105, name: "Molex 12-Pin Connectors", partNo: "MOL-43025", reqQty: 800, availQty: 480, pendingQty: 320, supplier: "Molex Inc.", expDate: "2026-09-06", status: "Pending" }
        ];

        timelineEntries = [
            { id: 1, projectId: 101, date: "2026-08-26", update: "Under CNC Milling for Aluminum housing.", person: "Raja Rajan", status: "Completed" },
            { id: 2, projectId: 101, date: "2026-08-27", update: "SMT Assembly Started on Line 2.", person: "Raja Rajan", status: "Completed" },
            { id: 3, projectId: 101, date: "2026-08-29", update: "Harness Wiring & Crimping Started.", person: "Raja Rajan", status: "In Progress" },
            { id: 4, projectId: 102, date: "2026-08-28", update: "Component placement verified. PCB assembly in progress on Line 1.", person: "Raja Rajan", status: "In Progress" },
            { id: 5, projectId: 102, date: "2026-09-01", update: "Combiner box housing alignment check passed.", person: "Suresh", status: "In Progress" },
            { id: 6, projectId: 103, date: "2026-09-02", update: "Continuity testing pending CAN chip arrival.", person: "Raja Rajan", status: "Hold" },
            { id: 7, projectId: 103, date: "2026-09-03", update: "CAN transceiver IC shipment received from supplier.", person: "Tamil", status: "In Progress" },
            { id: 8, projectId: 104, date: "2026-08-20", update: "Schematic design and Gerber files approved by engineering team.", person: "Raja Rajan", status: "Completed" },
            { id: 9, projectId: 104, date: "2026-08-25", update: "Initial prototype BOM requisition submitted to procurement.", person: "Prabhu", status: "In Progress" },
            { id: 10, projectId: 105, date: "2026-08-28", update: "Wiring harness inspection complete. Ready for final testing.", person: "Tamil", status: "In Progress" },
            { id: 11, projectId: 105, date: "2026-09-01", update: "Crimping tensile test completed successfully.", person: "Sathish", status: "Completed" },
            { id: 12, projectId: 106, date: "2026-08-22", update: "Mechanical enclosure 3D prototype printed and validated.", person: "Prabhu", status: "Completed" },
            { id: 13, projectId: 107, date: "2026-08-30", update: "Laser cutting and sheet metal bending completed for chassis.", person: "Tamil", status: "Completed" },
            { id: 14, projectId: 107, date: "2026-09-02", update: "Powder coating applied to enclosures.", person: "Sathish", status: "In Progress" },
            { id: 15, projectId: 108, date: "2026-09-01", update: "Wiring loom assembly completed. Final pinout continuity test passed.", person: "Prabhu", status: "Completed" },
            { id: 16, projectId: 109, date: "2026-08-28", update: "High-voltage insulation test completed with 100% pass rate.", person: "Raja Rajan", status: "Completed" },
            { id: 18, projectId: 110, date: "2026-08-25", update: "On hold pending customer specification clarification for DB9 pinouts.", person: "Tamil", status: "Hold" }
        ];

        blockers = [
            {
                id: 1,
                projectId: 103,
                title: "CAN Transceiver IC Supply Shortage",
                category: "Component Shortage",
                severity: "High",
                owner: "Raja Rajan",
                date: "2026-09-02",
                status: "Open",
                notes: "Supplier Texas Instruments delayed shipment due to wafer backlog. Expedited shipment expected Sept 8."
            },
            {
                id: 2,
                projectId: 105,
                title: "12-Pin Molex Connector Crimping Tool Clearance",
                category: "Design Clearance",
                severity: "Critical",
                owner: "Tamil",
                date: "2026-08-28",
                status: "In Progress",
                notes: "Awaiting tooling engineering signoff for custom crimp height specifications."
            },
            {
                id: 3,
                projectId: 110,
                title: "DB9 Pinout Specification Customer Clarification Pending",
                category: "Customer Specs Pending",
                severity: "Medium",
                owner: "Tamil",
                date: "2026-08-25",
                status: "Open",
                notes: "Customer Enphase lead contacted to verify RS232 pin mapping."
            },
            {
                id: 4,
                projectId: 101,
                title: "Antenna Signal Attenuation on Enclosure",
                category: "Quality Failure",
                severity: "Low",
                owner: "Raja Rajan",
                date: "2026-08-24",
                status: "Resolved",
                notes: "Replaced standard EMI gasket with 3M conductive foil. Signal attenuation cleared."
            }
        ];
    }

    /* ==========================================================================
       2. AUTOMATED COMPUTATION & DELAY ENGINE
       ========================================================================== */

    
    function ensureProjectDeliveryBatches(proj) {
        if (!proj) return;
        if (!proj.deliveryBatches || !Array.isArray(proj.deliveryBatches) || proj.deliveryBatches.length === 0) {
            const total = parseInt(proj.totalQty) || 100;
            const b1Qty = Math.round(total * 0.3);
            const b2Qty = Math.round(total * 0.35);
            const b3Qty = Math.max(0, total - b1Qty - b2Qty);

            const poDateStr = proj.poDate || SYSTEM_DATE_STR;
            const commDateStr = proj.committedDate || SYSTEM_DATE_STR;
            const revDateStr = proj.revisedDate || commDateStr;

            proj.deliveryBatches = [
                {
                    id: 1,
                    batchNo: 'Batch 1 (Initial)',
                    date: poDateStr,
                    qty: b1Qty,
                    status: (proj.deliveredQty || 0) >= b1Qty ? 'Delivered' : 'Scheduled',
                    note: 'First partial shipment batch'
                },
                {
                    id: 2,
                    batchNo: 'Batch 2 (Mid Run)',
                    date: commDateStr,
                    qty: b2Qty,
                    status: (proj.deliveredQty || 0) >= (b1Qty + b2Qty) ? 'Delivered' : 'In Transit',
                    note: 'Second partial shipment batch'
                },
                {
                    id: 3,
                    batchNo: 'Batch 3 (Final)',
                    date: revDateStr,
                    qty: b3Qty,
                    status: (proj.deliveredQty || 0) >= total ? 'Delivered' : 'Scheduled',
                    note: 'Final shipment batch'
                }
            ];
        }

        const deliveredFromBatches = proj.deliveryBatches
            .filter(b => b.status === 'Delivered')
            .reduce((sum, b) => sum + (parseInt(b.qty) || 0), 0);

        proj.deliveredQty = deliveredFromBatches;
        proj.pendingQty = Math.max(0, (parseInt(proj.totalQty) || 0) - proj.deliveredQty);
    }

    function recalculateAllProjectsState() {
        refreshSystemDate();
        projects.forEach(proj => {
            ensureProjectDeliveryBatches(proj);
            // Auto pending qty = Total - Delivered
            proj.pendingQty = Math.max(0, (parseInt(proj.totalQty) || 0) - (parseInt(proj.deliveredQty) || 0));

            // Auto calculate materials availability % based on BOM items
            const pMats = materials.filter(m => m.projectId === proj.id);
            if (pMats.length > 0) {
                const totalReq = pMats.reduce((acc, m) => acc + (parseInt(m.reqQty) || 0), 0);
                const totalAvail = pMats.reduce((acc, m) => acc + (parseInt(m.availQty) || 0), 0);
                proj.matAvailability = totalReq > 0 ? Math.min(100, Math.round((totalAvail / totalReq) * 100)) : 100;
            }

            // Auto completion percentage check based on stages if available
            if (proj.stages) {
                let stageCount = 0;
                let completedCount = 0;
                let inProgressCount = 0;

                PIPELINE_STAGES.forEach(stg => {
                    stageCount++;
                    const st = proj.stages[stg] || "Not Started";
                    if (st === "Completed") completedCount += 1;
                    else if (st === "In Progress") inProgressCount += 0.5;
                });

                if (stageCount > 0) {
                    proj.completion = Math.min(100, Math.round(((completedCount + inProgressCount) / stageCount) * 100));
                }
            }

            // Auto status override if 100% completed
            if (proj.completion >= 100 || proj.pendingQty === 0) {
                proj.status = "Completed";
                proj.completion = 100;
            }

            // Auto delay detection logic:
            // Check if committed or revised date passed and status is not Completed
            const targetDateStr = proj.revisedDate || proj.committedDate;
            if (targetDateStr && proj.status !== "Completed") {
                const targetDate = new Date(targetDateStr);
                // Strip time for exact date comparison
                targetDate.setHours(0,0,0,0);
                const currentCheckDate = new Date(SYSTEM_DATE);
                currentCheckDate.setHours(0,0,0,0);

                if (currentCheckDate > targetDate && proj.status !== "On Hold") {
                    proj.status = "Delayed";
                } else if (currentCheckDate <= targetDate && proj.status === "Delayed") {
                    proj.status = "In Progress";
                }
            }
        });
    }

    
    function calculateDelayDays(proj) {
        const targetDateStr = proj.revisedDate || proj.committedDate;
        if (!targetDateStr) return 0;
        const targetDate = new Date(targetDateStr);
        targetDate.setHours(0,0,0,0);
        const currentCheckDate = new Date(SYSTEM_DATE);
        currentCheckDate.setHours(0,0,0,0);

        const diffTime = currentCheckDate - targetDate;
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        return diffDays > 0 ? diffDays : 0;
    }

    /* ==========================================================================
       3. DOM READY & INITIALIZATION
       ========================================================================== */

    document.addEventListener('DOMContentLoaded', () => {
        initData();
        initAuth();
        setupBootstrapModals();
        setupEventListeners();
        renderActiveView('dashboard');
    });

    function setupBootstrapModals() {
        const addProjEl = document.getElementById('addProjectModal');
        const detailEl = document.getElementById('projectDetailModal');
        const importEl = document.getElementById('importModal');
        const addMatEl = document.getElementById('addMaterialModal');
        const notifEl = document.getElementById('notificationsModal');
        const blkEl = document.getElementById('blockerModal');

        if (addProjEl) projectModalBs = new bootstrap.Modal(addProjEl);
        if (detailEl) detailModalBs = new bootstrap.Modal(detailEl);
        if (importEl) importModalBs = new bootstrap.Modal(importEl);
        if (addMatEl) addMaterialModalBs = new bootstrap.Modal(addMatEl);
        if (notifEl) notificationsModalBs = new bootstrap.Modal(notifEl);
        if (blkEl) blockerModalBs = new bootstrap.Modal(blkEl);
        const addBatchEl = document.getElementById('addBatchModal');
        if (addBatchEl) addBatchModalBs = new bootstrap.Modal(addBatchEl);
    }

    /* ==========================================================================
       4. NAVIGATION & VIEW SWITCHING
       ========================================================================== */

    function setupEventListeners() {
        
        // Dashboard Sprint Filter Button
        const btnFilterSprint = document.getElementById('btn-filter-sprint');
        if (btnFilterSprint) {
            btnFilterSprint.addEventListener('click', () => {
                const choice = prompt("Select Sprint Cycle:\n1. Sprint 14 (Current • 4 days left)\n2. Sprint 13 (Previous • Completed)\n3. Sprint 15 (Upcoming)\n4. All Sprints\n\nEnter option (1-4):", "1");
                if (!choice) return;

                const badge = document.getElementById('dashboard-sprint-badge');
                if (choice === '1') {
                    if (badge) badge.innerText = 'Sprint 14 • 4 days left';
                    btnFilterSprint.innerHTML = `<i class="fa-regular fa-calendar-check me-1"></i> Sprint 14`;
                } else if (choice === '2') {
                    if (badge) badge.innerText = 'Sprint 13 • Completed';
                    btnFilterSprint.innerHTML = `<i class="fa-regular fa-calendar-check me-1"></i> Sprint 13`;
                } else if (choice === '3') {
                    if (badge) badge.innerText = 'Sprint 15 • Starts in 4d';
                    btnFilterSprint.innerHTML = `<i class="fa-regular fa-calendar-check me-1"></i> Sprint 15`;
                } else {
                    if (badge) badge.innerText = 'All Sprints Overview';
                    btnFilterSprint.innerHTML = `<i class="fa-regular fa-calendar-check me-1"></i> All Sprints`;
                }
                renderDashboardView();
            });
        }

        // Dashboard Filter Options Button
        const btnFilterOptions = document.getElementById('btn-filter-options');
        if (btnFilterOptions) {
            btnFilterOptions.addEventListener('click', () => {
                const status = prompt("Filter Dashboard by Status (In Progress, Delayed, Under Design, Completed, or Leave Empty for All):", "");
                if (status !== null) {
                    filterStatus = status.trim();
                    switchNavTo('projects');
                }
            });
        }

        // Dashboard Export Report Button
        const btnExportDash = document.getElementById('btn-export-dashboard-report');
        if (btnExportDash) {
            btnExportDash.addEventListener('click', () => {
                exportDeliveriesToExcel();
            });
        }

        // Weekly Deliveries View Mode Toggle (Table vs Cards)
        document.querySelectorAll('#deliveries-view-mode-toggle [data-view-mode]').forEach(btn => {
            btn.addEventListener('click', () => {
                const mode = btn.getAttribute('data-view-mode');
                document.querySelectorAll('#deliveries-view-mode-toggle [data-view-mode]').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');

                const tableCont = document.getElementById('deliveries-summary-table-container');
                const cardsCont = document.getElementById('deliveries-summary-cards-container');

                if (mode === 'cards') {
                    if (tableCont) tableCont.classList.add('d-none');
                    if (cardsCont) cardsCont.classList.remove('d-none');
                } else {
                    if (cardsCont) cardsCont.classList.add('d-none');
                    if (tableCont) tableCont.classList.remove('d-none');
                }
            });
        });

        // Weekly Deliveries Week Filter Chips
        document.querySelectorAll('#deliveries-week-chips [data-week-chip]').forEach(chip => {
            chip.addEventListener('click', () => {
                const val = chip.getAttribute('data-week-chip');
                document.querySelectorAll('#deliveries-week-chips [data-week-chip]').forEach(c => c.classList.remove('active'));
                chip.classList.add('active');

                const weekSelect = document.getElementById('deliveries-week-filter');
                if (weekSelect) {
                    weekSelect.value = val === 'later' ? 'week_4_plus' : val;
                }
                renderDeliveriesView();
            });
        });

        // PIN Keypad Button Click Handlers
        document.querySelectorAll('.btn-pin-key').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                const key = btn.getAttribute('data-key');
                if (key) handlePinKeyInput(key);
            });
        });

        // Physical Keyboard Support for PIN Lock Screen
        document.addEventListener('keydown', (e) => {
            const loginScr = document.getElementById('login-screen');
            if (loginScr && !loginScr.classList.contains('d-none')) {
                if (e.key >= '0' && e.key <= '9') {
                    handlePinKeyInput(e.key);
                } else if (e.key === 'Backspace') {
                    handlePinKeyInput('backspace');
                } else if (e.key === 'Escape' || e.key === 'Delete') {
                    handlePinKeyInput('clear');
                }
            }
        });

        // Sidebar link clicks
        document.querySelectorAll('.nav-link-item').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const viewName = btn.getAttribute('data-view');
                if (viewName) {
                    e.preventDefault();
                    switchNavActive(btn);
                    renderActiveView(viewName);
                }
            });
        });

        // Mobile Sidebar Toggle
        const btnToggleSidebar = document.getElementById('btn-toggle-sidebar');
        if (btnToggleSidebar) {
            btnToggleSidebar.addEventListener('click', () => {
                document.getElementById('sidebar').classList.toggle('show');
            });
        }

        // Global Search Bar
        const globalSearch = document.getElementById('global-search-input');
        if (globalSearch) {
            globalSearch.addEventListener('input', (e) => {
                tableSearchQuery = e.target.value.trim().toLowerCase();
                const currentView = getCurrentViewName();
                if (currentView !== 'projects') {
                    // Navigate to projects database view on search
                    switchNavTo('projects');
                }
                renderProjectsTable();
            });
        }

        // Header Quick Add Button
        const headerAddBtn = document.getElementById('btn-header-add-project');
        if (headerAddBtn) {
            headerAddBtn.addEventListener('click', () => openAddProjectModal());
        }

        const mainAddBtn = document.getElementById('btn-add-new-project-main');
        if (mainAddBtn) {
            mainAddBtn.addEventListener('click', () => openAddProjectModal());
        }

        // Notifications Button
        const btnNotif = document.getElementById('btn-open-notifications');
        if (btnNotif) {
            btnNotif.addEventListener('click', () => {
                renderNotificationsModal();
                if (notificationsModalBs) notificationsModalBs.show();
            });
        }

        // Quick Filters Buttons on Dashboard
        document.querySelectorAll('.btn-filter-chip').forEach(chip => {
            chip.addEventListener('click', () => {
                document.querySelectorAll('.btn-filter-chip').forEach(c => c.classList.remove('active'));
                chip.classList.add('active');
                activeQuickFilter = chip.getAttribute('data-filter');
                switchNavTo('projects');
            });
        });

        // Project Form Auto Calculations
        const totalQtyInput = document.getElementById('proj-input-total-qty');
        const delivQtyInput = document.getElementById('proj-input-delivered-qty');
        if (totalQtyInput && delivQtyInput) {
            const updatePending = () => {
                const tot = parseInt(totalQtyInput.value) || 0;
                const del = parseInt(delivQtyInput.value) || 0;
                document.getElementById('proj-input-pending-qty').value = Math.max(0, tot - del);
            };
            totalQtyInput.addEventListener('input', updatePending);
            delivQtyInput.addEventListener('input', updatePending);
        }

        // Project Form Submit
        
        const btnAddBatchDetail = document.getElementById('btn-add-batch-for-detail');
        if (btnAddBatchDetail) {
            btnAddBatchDetail.addEventListener('click', () => {
                const noteProjIdEl = document.getElementById('detail-note-proj-id');
                const projId = noteProjIdEl ? parseInt(noteProjIdEl.value) : null;
                if (projId) openAddBatchModal(projId);
            });
        }

        const batchForm = document.getElementById('form-add-batch');
        if (batchForm) {
            batchForm.addEventListener('submit', handleSaveBatch);
        }

        const projForm = document.getElementById('form-project');
        if (projForm) {
            projForm.addEventListener('submit', handleSaveProject);
        }

        // Table Search & Filter Controls
        const tableSearch = document.getElementById('table-search-input');
        if (tableSearch) {
            tableSearch.addEventListener('input', (e) => {
                tableSearchQuery = e.target.value.trim().toLowerCase();
                currentPage = 1;
                renderProjectsTable();
            });
        }

        const statusSel = document.getElementById('filter-status-select');
        if (statusSel) {
            statusSel.addEventListener('change', (e) => {
                filterStatus = e.target.value;
                currentPage = 1;
                renderProjectsTable();
            });
        }

        const prioSel = document.getElementById('filter-priority-select');
        if (prioSel) {
            prioSel.addEventListener('change', (e) => {
                filterPriority = e.target.value;
                currentPage = 1;
                renderProjectsTable();
            });
        }

        const leadSel = document.getElementById('filter-lead-select');
        if (leadSel) {
            leadSel.addEventListener('change', (e) => {
                filterLead = e.target.value;
                currentPage = 1;
                renderProjectsTable();
            });
        }

        const custSel = document.getElementById('filter-customer-select');
        if (custSel) {
            custSel.addEventListener('change', (e) => {
                filterCustomer = e.target.value;
                currentPage = 1;
                renderProjectsTable();
            });
        }

        const pageSizeSel = document.getElementById('table-page-size');
        if (pageSizeSel) {
            pageSizeSel.addEventListener('change', (e) => {
                pageSize = parseInt(e.target.value) || 25;
                currentPage = 1;
                renderProjectsTable();
            });
        }

        // Exports & Print Buttons
        const exportExcelBtn = document.getElementById('btn-export-excel');
        if (exportExcelBtn) exportExcelBtn.addEventListener('click', exportToExcel);

        const exportCsvBtn = document.getElementById('btn-export-csv');
        if (exportCsvBtn) exportCsvBtn.addEventListener('click', exportToCSV);

        const printBtn = document.getElementById('btn-print-table');
        if (printBtn) printBtn.addEventListener('click', () => window.print());

        // Import Modal & Template Buttons
        const openImportBtn = document.getElementById('btn-open-import-modal');
        if (openImportBtn) openImportBtn.addEventListener('click', () => {
            if (importModalBs) importModalBs.show();
        });

        const downloadTemplateBtn = document.getElementById('btn-download-template');
        if (downloadTemplateBtn) downloadTemplateBtn.addEventListener('click', downloadExcelTemplate);

        const importFileInput = document.getElementById('import-file-input');
        if (importFileInput) importFileInput.addEventListener('change', handleImportFileSelect);

        const execImportBtn = document.getElementById('btn-execute-import');
        if (execImportBtn) execImportBtn.addEventListener('click', executeImportData);

        // Materials Modal & Forms
        const openAddMatBtn = document.getElementById('btn-open-add-material-modal');
        if (openAddMatBtn) openAddMatBtn.addEventListener('click', () => openAddMaterialModal());

        const formAddMat = document.getElementById('form-add-material');
        if (formAddMat) formAddMat.addEventListener('submit', handleSaveMaterial);

        const matSearch = document.getElementById('mat-search-input');
        if (matSearch) matSearch.addEventListener('input', () => renderMaterialsView());

        const matFilterProj = document.getElementById('mat-filter-project');
        if (matFilterProj) matFilterProj.addEventListener('change', () => renderMaterialsView());

        const matFilterStat = document.getElementById('mat-filter-status');
        if (matFilterStat) matFilterStat.addEventListener('change', () => renderMaterialsView());

        // Project Blockers Modal & View Filters
        const openAddBlkBtn = document.getElementById('btn-open-add-blocker-modal');
        if (openAddBlkBtn) openAddBlkBtn.addEventListener('click', () => openAddBlockerModal());

        const formBlk = document.getElementById('form-blocker');
        if (formBlk) formBlk.addEventListener('submit', handleSaveBlocker);

        const blkSearch = document.getElementById('blocker-search-input');
        if (blkSearch) blkSearch.addEventListener('input', () => renderBlockersView());

        const blkFilterProj = document.getElementById('blocker-filter-project');
        if (blkFilterProj) blkFilterProj.addEventListener('change', () => renderBlockersView());

        const blkFilterStat = document.getElementById('blocker-filter-status');
        if (blkFilterStat) blkFilterStat.addEventListener('change', () => renderBlockersView());

        const blkFilterSev = document.getElementById('blocker-filter-severity');
        if (blkFilterSev) blkFilterSev.addEventListener('change', () => renderBlockersView());

        const btnAddBlkDetail = document.getElementById('btn-add-blocker-for-detail');
        if (btnAddBlkDetail) btnAddBlkDetail.addEventListener('click', () => {
            const curProjId = parseInt(document.getElementById('detail-note-proj-id')?.value);
            openAddBlockerModal(curProjId);
        });

        const btnAddMatDetail = document.getElementById('btn-add-material-for-detail');
        if (btnAddMatDetail) btnAddMatDetail.addEventListener('click', () => {
            const curProjId = parseInt(document.getElementById('detail-note-proj-id')?.value);
            openAddMaterialModal(curProjId);
        });

        // Timeline Entry Form
        const btnAddGlobalTL = document.getElementById('btn-add-global-timeline');
        if (btnAddGlobalTL) {
            btnAddGlobalTL.addEventListener('click', () => {
                const formTL = document.getElementById('form-add-timeline-entry');
                if (formTL) {
                    formTL.scrollIntoView({ behavior: 'smooth' });
                    const updateInput = document.getElementById('tl-input-update');
                    if (updateInput) updateInput.focus();
                }
            });
        }

        const formAddTL = document.getElementById('form-add-timeline-entry');
        if (formAddTL) formAddTL.addEventListener('submit', handleSaveTimelineEntry);

        const formDetailNote = document.getElementById('form-detail-add-note');
        if (formDetailNote) formDetailNote.addEventListener('submit', handleSaveDetailNote);

        // Timeline Filter Dropdown Listener
        const tlFilterProj = document.getElementById('timeline-filter-project');
        if (tlFilterProj) {
            tlFilterProj.addEventListener('change', () => {
                renderTimelineView();
            });
        }

        // Reports View Controls
        const btnGenReport = document.getElementById('btn-generate-report');
        if (btnGenReport) btnGenReport.addEventListener('click', generateReport);

        const btnExportReportExcel = document.getElementById('btn-export-report-excel');
        if (btnExportReportExcel) btnExportReportExcel.addEventListener('click', exportReportToExcel);

        const btnPrintReport = document.getElementById('btn-print-report');
        if (btnPrintReport) btnPrintReport.addEventListener('click', () => window.print());

        // Weekly Deliveries View Controls
        const delivSearch = document.getElementById('deliveries-search-input');
        if (delivSearch) delivSearch.addEventListener('input', () => renderDeliveriesView());

        const delivWeekFilter = document.getElementById('deliveries-week-filter');
        if (delivWeekFilter) delivWeekFilter.addEventListener('change', () => {
            const val = delivWeekFilter.value;
            document.querySelectorAll('#deliveries-week-chips button').forEach(b => {
                b.classList.toggle('active', b.getAttribute('data-week-chip') === val);
            });
            renderDeliveriesView();
        });

        const delivStatusFilter = document.getElementById('deliveries-status-filter');
        if (delivStatusFilter) delivStatusFilter.addEventListener('change', () => renderDeliveriesView());

        const btnExportDeliv = document.getElementById('btn-export-deliveries-excel');
        if (btnExportDeliv) btnExportDeliv.addEventListener('click', exportDeliveriesToExcel);

        const btnPrintDeliv = document.getElementById('btn-print-deliveries-summary');
        if (btnPrintDeliv) btnPrintDeliv.addEventListener('click', () => window.print());

        document.querySelectorAll('#deliveries-view-mode-toggle button').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const mode = e.currentTarget.getAttribute('data-view-mode');
                document.querySelectorAll('#deliveries-view-mode-toggle button').forEach(b => b.classList.remove('active'));
                e.currentTarget.classList.add('active');

                const tableCont = document.getElementById('deliveries-summary-table-container');
                const cardsCont = document.getElementById('deliveries-summary-cards-container');

                if (mode === 'cards') {
                    if (tableCont) tableCont.classList.add('d-none');
                    if (cardsCont) cardsCont.classList.remove('d-none');
                } else {
                    if (cardsCont) cardsCont.classList.add('d-none');
                    if (tableCont) tableCont.classList.remove('d-none');
                }
            });
        });

        document.querySelectorAll('#deliveries-week-chips button').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const chipVal = e.currentTarget.getAttribute('data-week-chip');
                document.querySelectorAll('#deliveries-week-chips button').forEach(b => b.classList.remove('active'));
                e.currentTarget.classList.add('active');
                if (delivWeekFilter) delivWeekFilter.value = chipVal;
                renderDeliveriesView();
            });
        });

        // Settings Buttons
        const btnReseed = document.getElementById('btn-reseed-data-setting');
        if (btnReseed) btnReseed.addEventListener('click', () => {
            if (confirm("Re-seed factory sample projects into LocalStorage? Existing changes will be replaced.")) {
                seedSampleData();
                recalculateAllProjectsState();
                saveAllData();
                renderActiveView(getCurrentViewName());
                alert("Factory sample data restored successfully!");
            }
        });

        const btnQuickReset = document.getElementById('btn-quick-reset-data');
        if (btnQuickReset) btnQuickReset.addEventListener('click', (e) => {
            e.preventDefault();
            if (confirm("Reset application to sample data?")) {
                seedSampleData();
                recalculateAllProjectsState();
                saveAllData();
                renderActiveView('dashboard');
            }
        });

        const btnClearAll = document.getElementById('btn-clear-all-data');
        if (btnClearAll) btnClearAll.addEventListener('click', () => {
            if (confirm("ARE YOU SURE? This will clear all projects, materials, timeline entries, and blockers permanently.")) {
                projects = [];
                materials = [];
                timelineEntries = [];
                blockers = [];
                saveAllData();
                renderActiveView(getCurrentViewName());
                alert("All application data cleared successfully! Your empty workspace state is preserved across reloads.");
            }
        });

        // Backup JSON
        const btnDownloadJson = document.getElementById('btn-download-json-backup');
        if (btnDownloadJson) btnDownloadJson.addEventListener('click', exportJsonBackup);

        const inputJsonRestore = document.getElementById('input-json-restore');
        if (inputJsonRestore) inputJsonRestore.addEventListener('change', importJsonBackup);
    }

    function getCurrentViewName() {
        const activeSec = document.querySelector('.view-section.active');
        return activeSec ? activeSec.id.replace('-view', '') : 'dashboard';
    }

    function switchNavActive(activeEl) {
        document.querySelectorAll('.nav-link-item').forEach(el => el.classList.remove('active'));
        activeEl.classList.add('active');
        // Hide mobile sidebar if open
        document.getElementById('sidebar').classList.remove('show');
    }

    function switchNavTo(viewName) {
        const link = document.querySelector(`.nav-link-item[data-view="${viewName}"]`);
        if (link) {
            switchNavActive(link);
            renderActiveView(viewName);
        }
    }

    function renderActiveView(viewName) {
        document.querySelectorAll('.view-section').forEach(sec => sec.classList.remove('active'));
        const targetSec = document.getElementById(`${viewName}-view`);
        if (targetSec) targetSec.classList.add('active');

        updateUserUI();
        populateDropdownOptions();

        switch (viewName) {
            case 'dashboard':
                renderDashboardView();
                break;
            case 'projects':
                renderProjectsTable();
                break;
            case 'timeline':
                renderTimelineView();
                break;
            case 'materials':
                renderMaterialsView();
                break;
            case 'blockers':
                renderBlockersView();
                break;
            case 'customers':
                renderCustomersView();
                break;
            case 'leads':
                renderLeadsView();
                break;
            case 'reports':
                // Ready for user generation
                break;
            case 'deliveries':
                renderDeliveriesView();
                break;
            case 'settings':
                // Settings loaded
                break;
        }
    }

    function populateDropdownOptions() {
        const activeProjects = projects.filter(p => p.status !== 'Completed' && p.completion < 100);

        // Project Leads set
        const leadsSet = new Set(projects.map(p => p.lead).filter(Boolean));
        const custsSet = new Set(projects.map(p => p.customer).filter(Boolean));

        const leadSel = document.getElementById('filter-lead-select');
        if (leadSel) {
            const curVal = leadSel.value;
            leadSel.innerHTML = '<option value="">All Leads</option>' + 
                Array.from(leadsSet).map(l => `<option value="${l}">${l}</option>`).join('');
            leadSel.value = curVal;
        }

        const custSel = document.getElementById('filter-customer-select');
        if (custSel) {
            const curVal = custSel.value;
            custSel.innerHTML = '<option value="">All Customers</option>' + 
                Array.from(custsSet).map(c => `<option value="${c}">${c}</option>`).join('');
            custSel.value = curVal;
        }

        // Timeline project selectors
        const tlProjSel = document.getElementById('timeline-filter-project');
        if (tlProjSel) {
            const curVal = tlProjSel.value;
            tlProjSel.innerHTML = '<option value="">All Projects</option>' + 
                projects.map(p => `<option value="${p.id}">${p.name} (${p.poNumber})</option>`).join('');
            tlProjSel.value = curVal;
        }

        // Exclude completed projects from New Activity Log Entry modal dropdown
        const tlInputProj = document.getElementById('tl-input-project-id');
        if (tlInputProj) {
            if (activeProjects.length > 0) {
                tlInputProj.innerHTML = activeProjects.map(p => `<option value="${p.id}">${p.name} (${p.poNumber})</option>`).join('');
            } else {
                tlInputProj.innerHTML = '<option value="">No Active Projects Available</option>';
            }
        }

        // Materials project selectors
        const matProjSel = document.getElementById('mat-filter-project');
        if (matProjSel) {
            const curVal = matProjSel.value;
            matProjSel.innerHTML = '<option value="">All Projects</option>' + 
                projects.map(p => `<option value="${p.id}">${p.name} (${p.poNumber})</option>`).join('');
            matProjSel.value = curVal;
        }

        const matModalProj = document.getElementById('mat-modal-proj-id');
        if (matModalProj) {
            matModalProj.innerHTML = activeProjects.map(p => `<option value="${p.id}">${p.name} (${p.poNumber})</option>`).join('');
        }

        // Blockers project selectors
        const blkProjSel = document.getElementById('blocker-filter-project');
        if (blkProjSel) {
            const curVal = blkProjSel.value;
            blkProjSel.innerHTML = '<option value="">All Projects</option>' + 
                projects.map(p => `<option value="${p.id}">${p.name} (${p.poNumber})</option>`).join('');
            blkProjSel.value = curVal;
        }

        const blkModalProj = document.getElementById('blocker-input-project-id') || document.getElementById('blocker-modal-proj-id');
        if (blkModalProj) {
            blkModalProj.innerHTML = activeProjects.map(p => `<option value="${p.id}">${p.name} (${p.poNumber})</option>`).join('');
        }
    }

    let chartBurndownInstance = null;

    let countdownIntervalId = null;

    function renderTargetCountdownWidget() {
        const cardContainer = document.querySelector('.dark-countdown-card');
        if (!cardContainer) return;

        if (countdownIntervalId) {
            clearInterval(countdownIntervalId);
            countdownIntervalId = null;
        }

        const activeProjects = projects.filter(p => p.status !== 'Completed' && (p.revisedDate || p.committedDate));

        if (activeProjects.length === 0) {
            cardContainer.innerHTML = `
                <div class="d-flex justify-content-between align-items-center mb-2">
                    <span class="badge bg-white text-dark fw-bold" style="font-size: 0.68rem;">MILESTONE DEADLINE</span>
                    <i class="fa-solid fa-flag-checkered text-white-50"></i>
                </div>
                <h4 class="fw-bold text-white mb-1" style="font-size: 1.15rem;">No Active Deadlines</h4>
                <div class="small text-white-50 mb-3">All project milestones completed or no active target date set.</div>
                <button class="btn-dark-card-action" onclick="ApexApp.openAddProjectModal()">
                    <i class="fa-solid fa-plus me-1"></i> Add Target Project
                </button>
            `;
            return;
        }

        activeProjects.sort((a, b) => {
            const dateA = new Date(a.revisedDate || a.committedDate);
            const dateB = new Date(b.revisedDate || b.committedDate);
            return dateA - dateB;
        });

        const targetProj = activeProjects[0];
        const targetDateStr = targetProj.revisedDate || targetProj.committedDate;
        const targetDate = new Date(`${targetDateStr}T23:59:59`);

        const isDelayed = targetProj.status === 'Delayed' || new Date(SYSTEM_DATE) > targetDate;
        const badgeLabel = isDelayed ? '🔴 URGENT OVERDUE' : '🚀 NEXT SHIPMENT TARGET';
        const badgeClass = isDelayed ? 'bg-danger text-white' : 'bg-white text-dark';

        cardContainer.innerHTML = `
            <div class="d-flex justify-content-between align-items-center mb-1">
                <span class="badge ${badgeClass} fw-bold" style="font-size: 0.68rem;">${badgeLabel}</span>
                <i class="fa-solid fa-box-archive text-white-50"></i>
            </div>
            <h4 class="fw-bold mt-2 mb-0 text-white text-truncate" style="font-size: 1.1rem;" title="${escapeHtml(targetProj.name)}">
                ${escapeHtml(targetProj.name)}
            </h4>
            <div class="small text-white-50 text-truncate mb-2" style="font-size: 0.78rem;">
                PO: ${escapeHtml(targetProj.poNumber)} | ${escapeHtml(targetProj.customer)}
            </div>

            <div class="countdown-timer-grid">
                <div class="timer-box">
                    <div class="timer-val" id="cd-days">00</div>
                    <div class="timer-unit">DAYS</div>
                </div>
                <div class="timer-box">
                    <div class="timer-val" id="cd-hours">00</div>
                    <div class="timer-unit">HOURS</div>
                </div>
                <div class="timer-box">
                    <div class="timer-val" id="cd-mins">00</div>
                    <div class="timer-unit">MINS</div>
                </div>
                <div class="timer-box">
                    <div class="timer-val" id="cd-secs">00</div>
                    <div class="timer-unit">SECS</div>
                </div>
            </div>

            <button class="btn-dark-card-action mt-2" onclick="ApexApp.viewProjectDetails(${targetProj.id})">
                <span>Inspect Project #${targetProj.id}</span>
                <i class="fa-solid fa-arrow-right"></i>
            </button>
        `;

        function updateTimer() {
            const now = new Date();
            const diff = targetDate - now;

            const daysEl = document.getElementById('cd-days');
            const hoursEl = document.getElementById('cd-hours');
            const minsEl = document.getElementById('cd-mins');
            const secsEl = document.getElementById('cd-secs');

            if (!daysEl || !hoursEl || !minsEl || !secsEl) return;

            if (diff <= 0) {
                daysEl.innerText = "00";
                hoursEl.innerText = "00";
                minsEl.innerText = "00";
                secsEl.innerText = "00";
                return;
            }

            const days = Math.floor(diff / (1000 * 60 * 60 * 24));
            const hours = Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
            const mins = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
            const secs = Math.floor((diff % (1000 * 60)) / 1000);

            daysEl.innerText = String(days).padStart(2, '0');
            hoursEl.innerText = String(hours).padStart(2, '0');
            minsEl.innerText = String(mins).padStart(2, '0');
            secsEl.innerText = String(secs).padStart(2, '0');
        }

        updateTimer();
        countdownIntervalId = setInterval(updateTimer, 1000);
    }

    function renderDashboardView() {
        recalculateAllProjectsState();
        calculateDashboardKPIs();
        renderDashboardActiveCards();
        renderDashboardPriorityTable();
        renderDashboardActivityFeed();
        renderTargetCountdownWidget();
        renderBurndownChart();
        renderDashboardCharts();
        renderSidebarActiveProjects();
    }

    function renderSidebarActiveProjects() {
        const container = document.getElementById('sidebar-active-projects-list');
        if (!container) return;

        const activeList = projects.filter(p => p.status !== "Completed" && p.status !== "Yet to Start").slice(0, 5);

        if (activeList.length === 0) {
            container.innerHTML = `<div class="text-muted px-3 py-1" style="font-size: 0.72rem;">No active projects</div>`;
            return;
        }

        const colorMap = {
            'In Progress': 'text-primary',
            'Under Design': 'text-info',
            'Under Production': 'text-warning',
            'Under Wiring': 'text-purple',
            'Delayed': 'text-danger',
            'On Hold': 'text-secondary'
        };

        container.innerHTML = activeList.map(p => {
            const dotClass = colorMap[p.status] || 'text-primary';
            return `
                <a href="#" class="nav-link-item text-truncate" onclick="ApexApp.viewProjectDetails(${p.id}); return false;" title="${escapeHtml(p.name)} (${p.status})">
                    <i class="fa-solid fa-circle ${dotClass}" style="font-size: 0.5rem;"></i>
                    <span class="text-truncate">${escapeHtml(p.name)}</span>
                </a>
            `;
        }).join('');
    }

    function calculateDashboardKPIs() {
        const total = projects.length;
        const newProj = projects.filter(p => p.type === "New").length;
        const repProj = projects.filter(p => p.type === "Repeated").length;
        const inProg = projects.filter(p => p.status === "In Progress" || p.status.includes("Under")).length;
        const comp = projects.filter(p => p.status === "Completed").length;
        const onHold = projects.filter(p => p.status === "On Hold").length;
        const delayed = projects.filter(p => p.status === "Delayed").length;
        const highPrio = projects.filter(p => p.priority === "1A" || p.priority === "1").length;
        
        const totalPendingQty = projects.reduce((acc, p) => acc + (p.pendingQty || 0), 0);
        const avgComp = total > 0 ? Math.round(projects.reduce((acc, p) => acc + (p.completion || 0), 0) / total) : 0;

        // Update DOM KPI values
        setElemText('kpi-total-projects', total);
        setElemText('kpi-new-projects', newProj);
        setElemText('kpi-repeated-projects', repProj);
        setElemText('kpi-in-progress', inProg);
        setElemText('kpi-completed', comp);
        setElemText('kpi-on-hold', onHold);
        setElemText('kpi-delayed', delayed);
        setElemText('kpi-high-priority', highPrio);
        setElemText('kpi-pending-qty', totalPendingQty.toLocaleString());
        setElemText('kpi-avg-completion', `${avgComp}%`);

        // Update Quick Filter Counts
        setElemText('cnt-all', total);
        setElemText('cnt-critical', projects.filter(p => p.priority === "1A").length);
        setElemText('cnt-delayed', delayed);
        setElemText('cnt-due-week', getDueThisWeekProjects().length);
        setElemText('cnt-in-progress', inProg);
        setElemText('cnt-completed', comp);
        setElemText('cnt-on-hold', onHold);
        setElemText('cnt-mat-pending', projects.filter(p => (p.matAvailability || 100) < 80).length);
    }

    function setElemText(id, val) {
        const el = document.getElementById(id);
        if (el) el.innerText = val;
    }

    function renderDashboardActiveCards() {
        const container = document.getElementById('active-projects-cards-container');
        if (!container) return;

        const activeProjects = projects.filter(p => p.status !== "Completed").slice(0, 2);

        if (activeProjects.length === 0) {
            container.innerHTML = `<div class="col-12 text-center py-4 text-muted border rounded bg-white">No active projects running.</div>`;
            return;
        }

        container.innerHTML = activeProjects.map(p => {
            const stagesList = p.stages ? Object.keys(p.stages).slice(0, 3) : PIPELINE_STAGES.slice(0, 3);
            
            return `
                <div class="col-md-6">
                    <div class="tf-card h-100 mb-0">
                        <div class="d-flex justify-content-between align-items-start mb-2">
                            <div class="d-flex align-items-center gap-2">
                                <div class="brand-icon" style="width: 28px; height: 28px; font-size: 0.75rem;">
                                    <i class="fa-solid fa-cube"></i>
                                </div>
                                <div>
                                    <h4 class="fw-bold text-dark mb-0" style="font-size: 0.95rem;">
                                        <a href="#" onclick="ApexApp.viewProjectDetails(${p.id}); return false;" class="text-dark text-decoration-none hover-primary">${escapeHtml(p.name)}</a>
                                    </h4>
                                    <div class="text-muted" style="font-size: 0.72rem;">${escapeHtml(p.customer)} • PO: ${escapeHtml(p.poNumber)}</div>
                                </div>
                            </div>
                            ${getPriorityBadgeHtml(p.priority)}
                        </div>

                        <div class="my-3">
                            <div class="d-flex justify-content-between align-items-center small mb-1">
                                <span class="text-muted">Overall Progress</span>
                                <span class="fw-bold text-primary">${p.completion}%</span>
                            </div>
                            <div class="progress" style="height: 6px;">
                                <div class="progress-bar bg-primary" role="progressbar" style="width: ${p.completion}%"></div>
                            </div>
                        </div>

                        <div class="border-top pt-2 mt-2">
                            <div class="text-muted uppercase fw-bold mb-2" style="font-size: 0.68rem; letter-spacing: 0.05em;">KEY DELIVERABLES / STAGES</div>
                            ${stagesList.map(stg => {
                                const status = (p.stages && p.stages[stg]) || 'Not Started';
                                const badgeClass = status === 'Completed' ? 'bg-success-subtle text-success' : (status === 'In Progress' ? 'bg-warning-subtle text-warning' : 'bg-light text-muted');
                                return `
                                    <div class="d-flex justify-content-between align-items-center py-1 border-bottom" style="font-size: 0.78rem;">
                                        <span><i class="fa-solid fa-circle-check text-muted me-1" style="font-size: 0.7rem;"></i>${stg}</span>
                                        <span class="badge ${badgeClass}" style="font-size: 0.65rem;">${status}</span>
                                    </div>
                                `;
                            }).join('')}
                        </div>

                        <div class="d-flex justify-content-between align-items-center mt-3 pt-2 border-top">
                            <div class="d-flex align-items-center">
                                <div class="user-avatar me-1" style="width: 24px; height: 24px; font-size: 0.65rem;">
                                    ${p.lead ? escapeHtml(p.lead.substring(0, 2).toUpperCase()) : 'RR'}
                                </div>
                                <span class="small text-muted">${escapeHtml(p.lead)}</span>
                            </div>
                            <div class="small text-muted"><i class="fa-regular fa-clock me-1"></i>Target: ${formatDate(p.committedDate)}</div>
                        </div>
                    </div>
                </div>
            `;
        }).join('');
    }

    function renderDashboardPriorityTable() {
        const tbody = document.getElementById('dashboard-priority-tbody');
        if (!tbody) return;

        const urgentList = [...projects].sort((a, b) => {
            if (a.priority === "1A") return -1;
            if (b.priority === "1A") return 1;
            return b.id - a.id;
        }).slice(0, 4);

        tbody.innerHTML = urgentList.map((p, idx) => `
            <tr>
                <td><input class="form-check-input" type="checkbox" ${p.status === 'Completed' ? 'checked' : ''} onchange="ApexApp.viewProjectDetails(${p.id})"></td>
                <td>
                    <a href="#" onclick="ApexApp.viewProjectDetails(${p.id}); return false;" class="fw-bold text-dark text-decoration-none hover-primary">${escapeHtml(p.name)}</a>
                    <div class="text-muted" style="font-size: 0.72rem;">PO: ${escapeHtml(p.poNumber)} • ${escapeHtml(p.customer)}</div>
                </td>
                <td><span class="badge bg-light text-dark border">${p.type}</span></td>
                <td>${getPriorityBadgeHtml(p.priority)}</td>
                <td>
                    <div class="d-flex align-items-center gap-1">
                        <div class="user-avatar" style="width: 22px; height: 22px; font-size: 0.6rem;">${p.lead ? escapeHtml(p.lead.substring(0, 2).toUpperCase()) : 'PM'}</div>
                        <span>${escapeHtml(p.lead)}</span>
                    </div>
                </td>
                <td>${formatDate(p.revisedDate || p.committedDate)}</td>
                <td>${getStatusBadgeHtml(p.status)}</td>
            </tr>
        `).join('');
    }

    function renderDashboardActivityFeed() {
        const feedContainer = document.getElementById('dashboard-activity-feed');
        if (!feedContainer) return;

        const recentLogs = timelineEntries.slice(0, 4);
        if (recentLogs.length === 0) {
            feedContainer.innerHTML = `<div class="text-center text-muted py-3 small">No recent activity.</div>`;
            return;
        }

        feedContainer.innerHTML = recentLogs.map(t => {
            const p = projects.find(item => item.id === t.projectId);
            const initial = t.person ? t.person.substring(0, 2).toUpperCase() : 'AP';

            return `
                <div class="activity-feed-item">
                    <div class="activity-dot">${escapeHtml(initial)}</div>
                    <div style="flex-grow: 1; line-height: 1.2;">
                        <div class="fw-bold text-dark" style="font-size: 0.8rem;">
                            ${escapeHtml(t.person)}
                            <span class="fw-normal text-muted">posted update on</span>
                            <span class="text-primary">${p ? escapeHtml(p.name) : 'Project'}</span>
                        </div>
                        <div class="text-secondary mt-1" style="font-size: 0.75rem;">"${escapeHtml(t.update)}"</div>
                        <div class="text-muted mt-1" style="font-size: 0.68rem;"><i class="fa-regular fa-clock me-1"></i>${formatDate(t.date)}</div>
                    </div>
                </div>
            `;
        }).join('');
    }

    function renderBurndownChart() {
        const ctx = document.getElementById('chart-burndown');
        if (!ctx) return;

        if (chartBurndownInstance) chartBurndownInstance.destroy();
        chartBurndownInstance = new Chart(ctx, {
            type: 'line',
            data: {
                labels: ['Day 1', 'Day 3', 'Today (Day 5)', 'Day 10'],
                datasets: [
                    {
                        label: 'Remaining Target',
                        data: [45, 32, 18, 0],
                        borderColor: '#4f46e5',
                        backgroundColor: 'rgba(79, 70, 229, 0.1)',
                        fill: true,
                        tension: 0.3
                    },
                    {
                        label: 'Ideal Velocity',
                        data: [45, 30, 15, 0],
                        borderColor: '#cbd5e1',
                        borderDash: [5, 5],
                        fill: false,
                        tension: 0
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: { legend: { display: false } },
                scales: { y: { beginAtZero: true, ticks: { font: { size: 10 } } }, x: { ticks: { font: { size: 10 } } } }
            }
        });
    }

    function getDueThisWeekProjects() {
        const now = new Date(SYSTEM_DATE);
        const nextWeek = new Date(now);
        nextWeek.setDate(now.getDate() + 7);

        return projects.filter(p => {
            if (p.status === "Completed") return false;
            const targetDateStr = p.revisedDate || p.committedDate;
            if (!targetDateStr) return false;
            const tDate = new Date(targetDateStr);
            return tDate >= now && tDate <= nextWeek;
        });
    }

    function renderDashboardAlerts() {
        const alertsGrid = document.getElementById('dashboard-alerts-grid');
        if (!alertsGrid) return;

        const overdueList = projects.filter(p => p.status === "Delayed");
        const dueWeekList = getDueThisWeekProjects();
        const matPendingList = projects.filter(p => (p.matAvailability || 100) < 80);
        const onHoldList = projects.filter(p => p.status === "On Hold");
        const custInputList = projects.filter(p => p.status === "Under Design" || p.status === "Yet to Start");

        let cardsHtml = '';
        let totalAlertsCount = overdueList.length + dueWeekList.length + matPendingList.length + onHoldList.length;

        // 1. Overdue Alert Card
        if (overdueList.length > 0) {
            cardsHtml += `
                <div class="alert-item-card alert-danger" onclick="ApexApp.filterAndNavProjects('status', 'Delayed')">
                    <div class="alert-icon-box"><i class="fa-solid fa-triangle-exclamation"></i></div>
                    <div class="alert-content">
                        <div class="alert-title">
                            <span>Overdue Projects</span>
                            <span class="badge bg-danger">${overdueList.length}</span>
                        </div>
                        <div class="alert-desc">${overdueList.map(p => p.name).slice(0, 2).join(', ')}${overdueList.length > 2 ? '...' : ''} breached committed date.</div>
                    </div>
                </div>
            `;
        }

        // 2. Due within 7 days
        if (dueWeekList.length > 0) {
            cardsHtml += `
                <div class="alert-item-card alert-warning" onclick="ApexApp.filterAndNavProjects('quick', 'due-week')">
                    <div class="alert-icon-box"><i class="fa-solid fa-clock"></i></div>
                    <div class="alert-content">
                        <div class="alert-title">
                            <span>Due Within 7 Days</span>
                            <span class="badge bg-warning text-dark">${dueWeekList.length}</span>
                        </div>
                        <div class="alert-desc">${dueWeekList.map(p => p.name).slice(0, 2).join(', ')} delivery date approaching.</div>
                    </div>
                </div>
            `;
        }

        // 3. Material Pending
        if (matPendingList.length > 0) {
            cardsHtml += `
                <div class="alert-item-card alert-info" onclick="ApexApp.filterAndNavProjects('quick', 'mat-pending')">
                    <div class="alert-icon-box"><i class="fa-solid fa-boxes-stacked"></i></div>
                    <div class="alert-content">
                        <div class="alert-title">
                            <span>Material Shortage Risk</span>
                            <span class="badge bg-primary">${matPendingList.length}</span>
                        </div>
                        <div class="alert-desc">Projects with less than 80% material availability.</div>
                    </div>
                </div>
            `;
        }

        // 4. On Hold Alert
        if (onHoldList.length > 0) {
            cardsHtml += `
                <div class="alert-item-card alert-dark" onclick="ApexApp.filterAndNavProjects('status', 'On Hold')">
                    <div class="alert-icon-box"><i class="fa-solid fa-pause"></i></div>
                    <div class="alert-content">
                        <div class="alert-title">
                            <span>Projects On Hold</span>
                            <span class="badge bg-secondary">${onHoldList.length}</span>
                        </div>
                        <div class="alert-desc">Awaiting client specs or component approval.</div>
                    </div>
                </div>
            `;
        }

        if (cardsHtml === '') {
            cardsHtml = `<div class="p-3 text-center text-muted w-100"><i class="fa-solid fa-circle-check text-success me-2"></i> All systems clear. No critical alerts.</div>`;
        }

        alertsGrid.innerHTML = cardsHtml;
        setElemText('alert-summary-badge', `${totalAlertsCount} Alerts Active`);
        setElemText('header-alert-count', totalAlertsCount);
    }

    /* Chart.js Rendering Engine */
    function renderDashboardCharts() {
        renderStatusChart();
        renderLeadChart();
        renderPriorityChart();
        renderCustomerChart();
        renderMonthlyChart();
    }

    function renderStatusChart() {
        const ctx = document.getElementById('chart-status');
        if (!ctx) return;

        const statusCounts = {};
        projects.forEach(p => {
            statusCounts[p.status] = (statusCounts[p.status] || 0) + 1;
        });

        const labels = Object.keys(statusCounts);
        const data = Object.values(statusCounts);
        const colors = labels.map(l => getStatusColorHex(l));

        if (chartStatusInstance) chartStatusInstance.destroy();
        chartStatusInstance = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: labels,
                datasets: [{
                    data: data,
                    backgroundColor: colors,
                    borderWidth: 2,
                    borderColor: '#ffffff'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { position: 'right', labels: { boxWidth: 12, font: { size: 11 } } }
                }
            }
        });
    }

    function renderLeadChart() {
        const ctx = document.getElementById('chart-lead');
        if (!ctx) return;

        const leadCounts = {};
        projects.forEach(p => {
            leadCounts[p.lead] = (leadCounts[p.lead] || 0) + 1;
        });

        if (chartLeadInstance) chartLeadInstance.destroy();
        chartLeadInstance = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: Object.keys(leadCounts),
                datasets: [{
                    label: 'Assigned Projects',
                    data: Object.values(leadCounts),
                    backgroundColor: '#2563eb',
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: { beginAtZero: true, ticks: { stepSize: 1 } }
                },
                plugins: { legend: { display: false } }
            }
        });
    }

    function renderPriorityChart() {
        const ctx = document.getElementById('chart-priority');
        if (!ctx) return;

        const prioMap = { '1A': 0, '1': 0, '2': 0, '3': 0, '4': 0, '5': 0 };
        projects.forEach(p => {
            if (prioMap[p.priority] !== undefined) prioMap[p.priority]++;
        });

        if (chartPriorityInstance) chartPriorityInstance.destroy();
        chartPriorityInstance = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: ['1A Critical', '1 High', '2 Medium', '3 Low', '4 Normal', '5 Lowest'],
                datasets: [{
                    label: 'Projects',
                    data: [prioMap['1A'], prioMap['1'], prioMap['2'], prioMap['3'], prioMap['4'], prioMap['5']],
                    backgroundColor: ['#dc2626', '#ea580c', '#d97706', '#0d9488', '#0284c7', '#64748b'],
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } },
                plugins: { legend: { display: false } }
            }
        });
    }

    function renderCustomerChart() {
        const ctx = document.getElementById('chart-customer');
        if (!ctx) return;

        const custMap = {};
        projects.forEach(p => {
            custMap[p.customer] = (custMap[p.customer] || 0) + 1;
        });

        if (chartCustomerInstance) chartCustomerInstance.destroy();
        chartCustomerInstance = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: Object.keys(custMap),
                datasets: [{
                    label: 'Project Volume',
                    data: Object.values(custMap),
                    backgroundColor: '#06b6d4',
                    borderRadius: 6
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                indexAxis: 'y',
                scales: { x: { beginAtZero: true, ticks: { stepSize: 1 } } },
                plugins: { legend: { display: false } }
            }
        });
    }

    function renderMonthlyChart() {
        const ctx = document.getElementById('chart-monthly');
        if (!ctx) return;

        // Group projects by delivery month (Committed/Revised Date)
        const monthMap = {};
        projects.forEach(p => {
            const dateStr = p.revisedDate || p.committedDate;
            if (dateStr) {
                const mKey = dateStr.substring(0, 7); // YYYY-MM
                if (!monthMap[mKey]) monthMap[mKey] = { completed: 0, active: 0, delayed: 0 };
                if (p.status === "Completed") monthMap[mKey].completed++;
                else if (p.status === "Delayed") monthMap[mKey].delayed++;
                else monthMap[mKey].active++;
            }
        });

        const months = Object.keys(monthMap).sort();

        if (chartMonthlyInstance) chartMonthlyInstance.destroy();
        chartMonthlyInstance = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: months,
                datasets: [
                    { label: 'Completed', data: months.map(m => monthMap[m].completed), backgroundColor: '#10b981' },
                    { label: 'In Progress', data: months.map(m => monthMap[m].active), backgroundColor: '#f59e0b' },
                    { label: 'Delayed', data: months.map(m => monthMap[m].delayed), backgroundColor: '#ef4444' }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    x: { stacked: true },
                    y: { stacked: true, beginAtZero: true, ticks: { stepSize: 1 } }
                }
            }
        });
    }

    function renderDashboardRecentTable() {
        const tbody = document.getElementById('dashboard-recent-tbody');
        if (!tbody) return;

        // Top 5 urgent or recent projects
        const sortedList = [...projects].sort((a, b) => {
            if (a.priority === "1A") return -1;
            if (b.priority === "1A") return 1;
            return b.id - a.id;
        }).slice(0, 5);

        tbody.innerHTML = sortedList.map(p => `
            <tr>
                <td class="fw-bold text-dark">${escapeHtml(p.name)}</td>
                <td>${escapeHtml(p.customer)}</td>
                <td><code class="text-primary">${escapeHtml(p.poNumber)}</code></td>
                <td>${escapeHtml(p.lead)}</td>
                <td>${getPriorityBadgeHtml(p.priority)}</td>
                <td>${getStatusBadgeHtml(p.status)}</td>
                <td>${formatDate(p.revisedDate || p.committedDate)}</td>
                <td>
                    <div class="table-progress-bar">
                        <div class="table-progress-fill" style="width: ${p.completion}%"></div>
                    </div>
                    <span class="fw-bold text-dark" style="font-size: 0.78rem;">${p.completion}%</span>
                </td>
                <td>
                    <button class="btn btn-xs btn-outline-primary py-0 px-2" onclick="ApexApp.viewProjectDetails(${p.id})">
                        <i class="fa-solid fa-eye"></i> Details
                    </button>
                </td>
            </tr>
        `).join('');
    }

    /* ==========================================================================
       6. PROJECT DATABASE TABLE CONTROLLER (SEARCH, SORT, FILTER, PAGINATION)
       ========================================================================== */

    function renderProjectsTable() {
        recalculateAllProjectsState();
        renderTableHeader();

        let filtered = projects.filter(p => {
            // Text Search
            if (tableSearchQuery) {
                const q = tableSearchQuery;
                const matchText = (p.name + ' ' + p.customer + ' ' + p.poNumber + ' ' + p.lead + ' ' + p.description + ' ' + p.invoiceNumber).toLowerCase();
                if (!matchText.includes(q)) return false;
            }

            // Dropdown Filters
            if (filterStatus && p.status !== filterStatus) return false;
            if (filterPriority && p.priority !== filterPriority) return false;
            if (filterLead && p.lead !== filterLead) return false;
            if (filterCustomer && p.customer !== filterCustomer) return false;

            // Quick Filter Toolbar
            if (activeQuickFilter !== 'all') {
                if (activeQuickFilter === 'critical' && p.priority !== '1A') return false;
                if (activeQuickFilter === 'delayed' && p.status !== 'Delayed') return false;
                if (activeQuickFilter === 'due-week' && !getDueThisWeekProjects().some(dp => dp.id === p.id)) return false;
                if (activeQuickFilter === 'in-progress' && !(p.status === 'In Progress' || p.status.includes('Under'))) return false;
                if (activeQuickFilter === 'completed' && p.status !== 'Completed') return false;
                if (activeQuickFilter === 'on-hold' && p.status !== 'On Hold') return false;
                if (activeQuickFilter === 'mat-pending' && (p.matAvailability || 100) >= 80) return false;
            }

            return true;
        });

        // Sorting
        filtered.sort((a, b) => {
            let valA = a[sortColumn];
            let valB = b[sortColumn];

            if (typeof valA === 'string') valA = valA.toLowerCase();
            if (typeof valB === 'string') valB = valB.toLowerCase();

            if (valA < valB) return sortDirection === 'asc' ? -1 : 1;
            if (valA > valB) return sortDirection === 'asc' ? 1 : -1;
            return 0;
        });

        // Pagination
        const totalItems = filtered.length;
        const totalPages = Math.ceil(totalItems / pageSize) || 1;
        if (currentPage > totalPages) currentPage = totalPages;

        const startIndex = (currentPage - 1) * pageSize;
        const paginated = filtered.slice(startIndex, startIndex + pageSize);

        renderTableBody(paginated, startIndex);
        renderTablePagination(totalItems, startIndex, paginated.length, totalPages);
        renderColumnVisibilityDropdown();
    }

    function renderTableHeader() {
        const trHeader = document.getElementById('projects-table-header');
        if (!trHeader) return;

        trHeader.innerHTML = columnsConfig.filter(c => c.visible).map(col => {
            if (!col.sortable) {
                return `<th>${col.label}</th>`;
            }
            const isSorted = sortColumn === col.id;
            const sortClass = isSorted ? (sortDirection === 'asc' ? 'sorted-asc' : 'sorted-desc') : '';
            const icon = isSorted ? (sortDirection === 'asc' ? 'fa-sort-up' : 'fa-sort-down') : 'fa-sort';

            return `
                <th class="sortable ${sortClass}" onclick="ApexApp.sortTable('${col.id}')">
                    ${col.label} <i class="fa-solid ${icon} sort-icon"></i>
                </th>
            `;
        }).join('');
    }

    function renderTableBody(items, startIndex) {
        const tbody = document.getElementById('projects-table-tbody');
        if (!tbody) return;

        if (items.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="20" class="text-center py-5 text-muted">
                        <i class="fa-solid fa-folder-open fa-2x mb-2 text-secondary"></i>
                        <div>No projects found matching the selected criteria.</div>
                    </td>
                </tr>
            `;
            return;
        }

        const visibleCols = columnsConfig.filter(c => c.visible).map(c => c.id);

        tbody.innerHTML = items.map((p, idx) => {
            const sno = startIndex + idx + 1;
            let cellsHtml = '';

            visibleCols.forEach(colId => {
                switch (colId) {
                    case 'sno':
                        cellsHtml += `<td class="text-muted fw-bold">${sno}</td>`;
                        break;
                    case 'name':
                        cellsHtml += `
                            <td>
                                <a href="#" class="fw-bold text-dark text-decoration-none hover-primary" onclick="ApexApp.viewProjectDetails(${p.id}); return false;">
                                    ${escapeHtml(p.name)}
                                </a>
                            </td>
                        `;
                        break;
                    case 'type':
                        cellsHtml += `<td><span class="badge bg-light text-dark border">${escapeHtml(p.type)}</span></td>`;
                        break;
                    case 'customer':
                        cellsHtml += `<td class="fw-semibold">${escapeHtml(p.customer)}</td>`;
                        break;
                    case 'poNumber':
                        cellsHtml += `<td><code class="text-primary fw-bold">${escapeHtml(p.poNumber)}</code></td>`;
                        break;
                    case 'priority':
                        cellsHtml += `<td>${getPriorityBadgeHtml(p.priority)}</td>`;
                        break;
                    case 'lead':
                        cellsHtml += `<td><i class="fa-solid fa-user-gear me-1 text-muted"></i>${escapeHtml(p.lead)}</td>`;
                        break;
                    case 'poDate':
                        cellsHtml += `<td>${formatDate(p.poDate)}</td>`;
                        break;
                    case 'committedDate':
                        cellsHtml += `<td>${formatDate(p.committedDate)}</td>`;
                        break;
                    case 'revisedDate':
                        cellsHtml += `<td>${formatDate(p.revisedDate)}</td>`;
                        break;
                    case 'deliveredDate':
                        cellsHtml += `<td>${p.deliveredDate ? formatDate(p.deliveredDate) : '<span class="text-muted small">-</span>'}</td>`;
                        break;
                    case 'totalQty':
                        cellsHtml += `<td class="fw-semibold">${p.totalQty}</td>`;
                        break;
                    case 'deliveredQty':
                        cellsHtml += `<td class="text-success fw-bold">${p.deliveredQty}</td>`;
                        break;
                    case 'pendingQty':
                        cellsHtml += `<td class="${p.pendingQty > 0 ? 'text-danger fw-bold' : 'text-muted'}">${p.pendingQty}</td>`;
                        break;
                    case 'matAvailability':
                        const matVal = p.matAvailability || 100;
                        const matClass = matVal >= 90 ? 'text-success' : (matVal >= 70 ? 'text-warning' : 'text-danger');
                        cellsHtml += `<td class="fw-bold ${matClass}">${matVal}%</td>`;
                        break;
                    case 'status':
                        cellsHtml += `<td>${getStatusBadgeHtml(p.status)}</td>`;
                        break;
                    case 'completion':
                        cellsHtml += `
                            <td>
                                <div class="table-progress-bar">
                                    <div class="table-progress-fill" style="width: ${p.completion}%"></div>
                                </div>
                                <span class="fw-bold text-dark" style="font-size: 0.78rem;">${p.completion}%</span>
                            </td>
                        `;
                        break;
                    case 'invoiceNumber':
                        cellsHtml += `<td>${p.invoiceNumber ? `<code>${p.invoiceNumber}</code>` : '-'}</td>`;
                        break;
                    case 'actions':
                        cellsHtml += `
                            <td>
                                <div class="btn-group btn-group-sm">
                                    <button class="btn btn-outline-primary" onclick="ApexApp.viewProjectDetails(${p.id})" title="View Details">
                                        <i class="fa-solid fa-eye"></i>
                                    </button>
                                    <button class="btn btn-outline-secondary" onclick="ApexApp.editProject(${p.id})" title="Edit Project">
                                        <i class="fa-solid fa-pen-to-square"></i>
                                    </button>
                                    <button class="btn btn-outline-danger" onclick="ApexApp.deleteProject(${p.id})" title="Delete Project">
                                        <i class="fa-solid fa-trash"></i>
                                    </button>
                                </div>
                            </td>
                        `;
                        break;
                }
            });

            return `<tr>${cellsHtml}</tr>`;
        }).join('');
    }

    function renderTablePagination(totalItems, startIndex, count, totalPages) {
        const infoEl = document.getElementById('table-pagination-info');
        if (infoEl) {
            const endIdx = startIndex + count;
            infoEl.innerText = totalItems > 0 ? `Showing ${startIndex + 1} to ${endIdx} of ${totalItems} projects` : 'Showing 0 projects';
        }

        const btnsContainer = document.getElementById('table-pagination-btns');
        if (!btnsContainer) return;

        let btnsHtml = '';

        btnsHtml += `
            <button class="btn btn-sm btn-outline-secondary ${currentPage === 1 ? 'disabled' : ''}" onclick="ApexApp.changePage(${currentPage - 1})">
                <i class="fa-solid fa-chevron-left"></i>
            </button>
        `;

        for (let i = 1; i <= totalPages; i++) {
            if (i === 1 || i === totalPages || (i >= currentPage - 1 && i <= currentPage + 1)) {
                btnsHtml += `
                    <button class="btn btn-sm ${i === currentPage ? 'btn-primary' : 'btn-outline-secondary'}" onclick="ApexApp.changePage(${i})">
                        ${i}
                    </button>
                `;
            } else if (i === currentPage - 2 || i === currentPage + 2) {
                btnsHtml += `<span class="px-1 text-muted">...</span>`;
            }
        }

        btnsHtml += `
            <button class="btn btn-sm btn-outline-secondary ${currentPage === totalPages ? 'disabled' : ''}" onclick="ApexApp.changePage(${currentPage + 1})">
                <i class="fa-solid fa-chevron-right"></i>
            </button>
        `;

        btnsContainer.innerHTML = btnsHtml;
    }

    function renderColumnVisibilityDropdown() {
        const drop = document.getElementById('column-visibility-dropdown');
        if (!drop) return;

        drop.innerHTML = columnsConfig.map(col => `
            <li class="form-check mb-1">
                <input class="form-check-input" type="checkbox" id="col-toggle-${col.id}" ${col.visible ? 'checked' : ''} onchange="ApexApp.toggleColumnVisibility('${col.id}')">
                <label class="form-check-label small" for="col-toggle-${col.id}">${col.label}</label>
            </li>
        `).join('');
    }

    /* ==========================================================================
       7. ADD / EDIT / DELETE PROJECT ACTIONS
       ========================================================================== */

    function openAddProjectModal() {
        document.getElementById('addProjectModalTitle').innerText = 'Add New Manufacturing Project';
        document.getElementById('form-project').reset();
        document.getElementById('proj-input-id').value = '';
        document.getElementById('proj-input-committed-date').value = SYSTEM_DATE_STR;
        if (document.getElementById('proj-input-delivered-date')) document.getElementById('proj-input-delivered-date').value = '';
        document.getElementById('proj-input-pending-qty').value = 100;
        if (projectModalBs) projectModalBs.show();
    }

    function editProject(id) {
        const p = projects.find(item => item.id === id);
        if (!p) return;

        document.getElementById('addProjectModalTitle').innerText = `Edit Project: ${p.name}`;
        document.getElementById('proj-input-id').value = p.id;
        document.getElementById('proj-input-name').value = p.name;
        document.getElementById('proj-input-type').value = p.type || 'New';
        document.getElementById('proj-input-customer').value = p.customer;
        document.getElementById('proj-input-po').value = p.poNumber;
        document.getElementById('proj-input-lead').value = p.lead;
        document.getElementById('proj-input-priority').value = p.priority;
        document.getElementById('proj-input-po-date').value = p.poDate || '';
        document.getElementById('proj-input-committed-date').value = p.committedDate || '';
        document.getElementById('proj-input-revised-date').value = p.revisedDate || '';
        if (document.getElementById('proj-input-delivered-date')) document.getElementById('proj-input-delivered-date').value = p.deliveredDate || '';
        document.getElementById('proj-input-total-qty').value = p.totalQty;
        document.getElementById('proj-input-delivered-qty').value = p.deliveredQty;
        document.getElementById('proj-input-pending-qty').value = p.pendingQty;
        document.getElementById('proj-input-mat-avail').value = p.matAvailability || 100;
        document.getElementById('proj-input-status').value = p.status;
        document.getElementById('proj-input-invoice').value = p.invoiceNumber || '';
        document.getElementById('proj-input-team').value = p.teamMembers || '';
        document.getElementById('proj-input-desc').value = p.description || '';

        if (projectModalBs) projectModalBs.show();
    }

    function handleSaveProject(e) {
        e.preventDefault();

        const idVal = document.getElementById('proj-input-id').value;
        const totalQty = parseInt(document.getElementById('proj-input-total-qty').value) || 0;
        const deliveredQty = parseInt(document.getElementById('proj-input-delivered-qty').value) || 0;
        const pendingQty = Math.max(0, totalQty - deliveredQty);
        const statusVal = document.getElementById('proj-input-status').value;
        const delivDateVal = document.getElementById('proj-input-delivered-date')?.value || (statusVal === 'Completed' ? SYSTEM_DATE_STR : '');

        const projectData = {
            name: document.getElementById('proj-input-name').value.trim(),
            type: document.getElementById('proj-input-type').value,
            customer: document.getElementById('proj-input-customer').value.trim(),
            poNumber: document.getElementById('proj-input-po').value.trim(),
            lead: document.getElementById('proj-input-lead').value.trim(),
            priority: document.getElementById('proj-input-priority').value,
            poDate: document.getElementById('proj-input-po-date').value,
            committedDate: document.getElementById('proj-input-committed-date').value,
            revisedDate: document.getElementById('proj-input-revised-date').value,
            deliveredDate: delivDateVal,
            totalQty: totalQty,
            deliveredQty: deliveredQty,
            pendingQty: pendingQty,
            matAvailability: parseInt(document.getElementById('proj-input-mat-avail').value) || 100,
            status: statusVal,
            invoiceNumber: document.getElementById('proj-input-invoice').value.trim(),
            teamMembers: document.getElementById('proj-input-team').value.trim(),
            description: document.getElementById('proj-input-desc').value.trim()
        };

        if (idVal) {
            // Edit existing
            const index = projects.findIndex(p => p.id === parseInt(idVal));
            if (index !== -1) {
                projects[index] = { ...projects[index], ...projectData };
            }
        } else {
            // Create new
            const newId = projects.length > 0 ? Math.max(...projects.map(p => p.id)) + 1 : 101;
            const newStages = {};
            PIPELINE_STAGES.forEach(s => newStages[s] = s === "Design" ? "In Progress" : "Not Started");

            projects.unshift({
                id: newId,
                ...projectData,
                completion: 10,
                stages: newStages
            });
        }

        recalculateAllProjectsState();
        saveAllData();

        if (projectModalBs) projectModalBs.hide();
        renderActiveView(getCurrentViewName());
    }

    function deleteProject(id) {
        const p = projects.find(item => item.id === id);
        if (!p) return;

        if (confirm(`Are you sure you want to delete project "${p.name}" (${p.poNumber})?`)) {
            projects = projects.filter(item => item.id !== id);
            materials = materials.filter(m => m.projectId !== id);
            timelineEntries = timelineEntries.filter(t => t.projectId !== id);

            recalculateAllProjectsState();
            saveAllData();
            renderActiveView(getCurrentViewName());
        }
    }

    /* ==========================================================================
       8. PROJECT DETAILS INSPECTOR & INTERACTIVE PIPELINE
       ========================================================================== */

    function viewProjectDetails(id) {
        const p = projects.find(item => item.id === id);
        if (!p) return;

        document.getElementById('detail-modal-title').innerText = p.name;
        document.getElementById('detail-modal-status-badge').innerHTML = getStatusBadgeHtml(p.status);
        document.getElementById('detail-modal-priority-badge').innerHTML = getPriorityBadgeHtml(p.priority);
        document.getElementById('detail-overall-completion-badge').innerText = `${p.completion}% Complete`;

        // Delay Alert Check
        const delayBanner = document.getElementById('detail-delay-banner');
        if (p.status === "Delayed") {
            const delayDays = calculateDelayDays(p);
            delayBanner.classList.remove('d-none');
            delayBanner.innerHTML = `
                <div class="delay-banner">
                    <div>
                        <div class="fw-bold fs-6"><i class="fa-solid fa-circle-exclamation me-2"></i>🔴 PROJECT IS DELAYED</div>
                        <div class="small">Original Committed Date: <strong>${formatDate(p.committedDate)}</strong> | Revised Date: <strong>${formatDate(p.revisedDate)}</strong></div>
                    </div>
                    <div class="text-end">
                        <div class="fw-extrabold fs-4">${delayDays} Days Overdue</div>
                        <div class="small">System Current Date: ${SYSTEM_DATE_STR}</div>
                    </div>
                </div>
            `;
        } else {
            delayBanner.classList.add('d-none');
        }

        // Render Interactive Pipeline Stages
        renderPipelineTracker(p);

        // Fill Specs
        setElemText('detail-info-name', p.name);
        setElemText('detail-info-type', p.type);
        setElemText('detail-info-customer', p.customer);
        setElemText('detail-info-po', p.poNumber);
        setElemText('detail-info-lead', p.lead);
        setElemText('detail-info-invoice', p.invoiceNumber || 'N/A');
        setElemText('detail-info-desc', p.description || 'N/A');
        setElemText('detail-info-po-date', formatDate(p.poDate));
        setElemText('detail-info-committed-date', formatDate(p.committedDate));
        setElemText('detail-info-revised-date', formatDate(p.revisedDate));
        setElemText('detail-info-delivered-date', p.deliveredDate ? formatDate(p.deliveredDate) : (p.status === 'Completed' ? formatDate(SYSTEM_DATE_STR) : 'Not Delivered Yet'));
        setElemText('detail-info-total-qty', p.totalQty);
        setElemText('detail-info-delivered-qty', p.deliveredQty);
        setElemText('detail-info-pending-qty', p.pendingQty);
        // Set Note Form Context
        const noteProjIdEl = document.getElementById('detail-note-proj-id');
        if (noteProjIdEl) noteProjIdEl.value = p.id;
        const notePersonEl = document.getElementById('detail-note-person');
        if (notePersonEl) notePersonEl.value = p.lead || 'Admin User';

        // Render Materials Sub-tab
        renderDetailMaterialsTab(p.id);

        // Render Timeline Sub-tab
        renderDetailTimelineTab(p.id);

        // Render Blockers & Risks Sub-tab
        renderDetailBlockersTab(p.id);

        // Render Delivery Batches Sub-tab
        renderDetailBatchesTab(p.id);

        // Render Team & Contributors
        renderProjectContributors(p);

        // Edit button binding
        const editBtn = document.getElementById('btn-detail-edit-proj');
        if (editBtn) {
            editBtn.onclick = () => {
                if (detailModalBs) detailModalBs.hide();
                editProject(p.id);
            };
        }

        if (detailModalBs) detailModalBs.show();
    }

    function renderProjectContributors(p) {
        const container = document.getElementById('detail-contributors-container');
        if (!container) return;

        const contributorsMap = new Map();

        // 1. Project Lead
        if (p.lead) {
            contributorsMap.set(p.lead.trim().toLowerCase(), {
                name: p.lead.trim(),
                role: 'Project Lead',
                badgeClass: 'bg-primary'
            });
        }

        // 2. Assigned Team Members from p.teamMembers
        if (p.teamMembers) {
            const members = p.teamMembers.split(',');
            members.forEach(m => {
                const name = m.trim();
                if (name && !contributorsMap.has(name.toLowerCase())) {
                    contributorsMap.set(name.toLowerCase(), {
                        name: name,
                        role: 'Team Member',
                        badgeClass: 'bg-info text-dark'
                    });
                }
            });
        }

        // 3. Activity Timeline Log Authors
        const pLogs = timelineEntries.filter(t => t.projectId === p.id);
        const authorCounts = {};
        pLogs.forEach(t => {
            if (t.person) {
                const name = t.person.trim();
                authorCounts[name] = (authorCounts[name] || 0) + 1;
            }
        });

        Object.keys(authorCounts).forEach(name => {
            const key = name.toLowerCase();
            if (contributorsMap.has(key)) {
                const existing = contributorsMap.get(key);
                existing.noteCount = authorCounts[name];
            } else {
                contributorsMap.set(key, {
                    name: name,
                    role: 'Contributor',
                    badgeClass: 'bg-secondary',
                    noteCount: authorCounts[name]
                });
            }
        });

        const list = Array.from(contributorsMap.values());

        if (list.length === 0) {
            container.innerHTML = `<span class="text-muted small">No team members assigned yet.</span>`;
            return;
        }

        container.innerHTML = list.map(c => `
            <div class="d-inline-flex align-items-center gap-2 p-2 border rounded bg-light">
                <div class="user-avatar" style="width: 28px; height: 28px; font-size: 0.75rem;">
                    ${escapeHtml(c.name.substring(0, 2).toUpperCase())}
                </div>
                <div>
                    <div class="fw-bold" style="font-size: 0.8rem; line-height: 1.1;">${escapeHtml(c.name)}</div>
                    <div class="d-flex align-items-center gap-1">
                        <span class="badge ${c.badgeClass}" style="font-size: 0.65rem;">${c.role}</span>
                        ${c.noteCount ? `<span class="badge bg-light text-dark border" style="font-size: 0.65rem;">${c.noteCount} Activity Logs</span>` : ''}
                    </div>
                </div>
            </div>
        `).join('');
    }

    function renderPipelineTracker(p) {
        const container = document.getElementById('pipeline-tracker-container');
        if (!container) return;

        if (!p.stages) {
            p.stages = {};
            PIPELINE_STAGES.forEach(s => p.stages[s] = "Not Started");
        }

        container.innerHTML = PIPELINE_STAGES.map((stg, idx) => {
            const stStatus = p.stages[stg] || "Not Started";
            let nodeClass = "stage-not-started";
            let icon = `<span style="font-size: 0.85rem;">${idx + 1}</span>`;

            if (stStatus === "Completed") {
                nodeClass = "stage-completed";
                icon = `<i class="fa-solid fa-check"></i>`;
            } else if (stStatus === "In Progress") {
                nodeClass = "stage-in-progress";
                icon = `<i class="fa-solid fa-spinner fa-spin"></i>`;
            } else if (stStatus === "Hold") {
                nodeClass = "stage-hold";
                icon = `<i class="fa-solid fa-pause"></i>`;
            }

            return `
                <div class="pipeline-stage-item" onclick="ApexApp.toggleStageStatus(${p.id}, '${stg}')" title="Click to cycle stage status">
                    <div class="stage-node ${nodeClass}">${icon}</div>
                    <div class="stage-label">${stg}</div>
                    <div class="stage-status-text">${stStatus}</div>
                </div>
            `;
        }).join('');
    }

    function toggleStageStatus(projId, stageName) {
        const p = projects.find(item => item.id === projId);
        if (!p) return;

        if (!p.stages) p.stages = {};
        const current = p.stages[stageName] || "Not Started";
        let next = "In Progress";
        if (current === "Not Started") next = "In Progress";
        else if (current === "In Progress") next = "Completed";
        else if (current === "Completed") next = "Hold";
        else if (current === "Hold") next = "Not Started";

        p.stages[stageName] = next;

        // Auto timeline entry recording
        timelineEntries.unshift({
            id: Date.now(),
            projectId: p.id,
            date: SYSTEM_DATE_STR,
            update: `Stage [${stageName}] status updated to: ${next}`,
            person: p.lead || "System Lead",
            status: next
        });

        recalculateAllProjectsState();
        saveAllData();

        // Refresh modal view
        viewProjectDetails(p.id);
        renderActiveView(getCurrentViewName());
    }

    
    /* ==========================================================================
       8B. DELIVERY BATCHES CONTROLLER (3+ PARTIAL SHIPMENT DATES)
       ========================================================================== */

    function renderDetailBatchesTab(projId) {
        const tbody = document.getElementById('detail-batches-tbody');
        if (!tbody) return;

        const proj = projects.find(p => p.id === projId);
        if (!proj) return;

        ensureProjectDeliveryBatches(proj);

        const totalBatchQty = proj.deliveryBatches.reduce((acc, b) => acc + (parseInt(b.qty) || 0), 0);
        const deliveredBatchQty = proj.deliveryBatches.filter(b => b.status === 'Delivered').reduce((acc, b) => acc + (parseInt(b.qty) || 0), 0);

        let html = '';
        if (proj.deliveryBatches.length === 0) {
            html = `<tr><td colspan="6" class="text-center text-muted py-3">No delivery batches configured for this project. Click "Add Delivery Batch" to create one.</td></tr>`;
        } else {
            html = proj.deliveryBatches.map((b, idx) => {
                let statusBadge = `<span class="badge bg-secondary">Scheduled</span>`;
                if (b.status === 'Delivered') statusBadge = `<span class="badge bg-success">🟢 Delivered</span>`;
                else if (b.status === 'In Transit') statusBadge = `<span class="badge bg-warning text-dark">🚚 In Transit</span>`;

                return `
                    <tr>
                        <td class="fw-bold">#${idx + 1} (${escapeHtml(b.batchNo || 'Batch ' + (idx + 1))})</td>
                        <td><i class="fa-regular fa-calendar me-1 text-primary"></i> ${formatDate(b.date)}</td>
                        <td class="fw-bold text-dark">${(parseInt(b.qty) || 0).toLocaleString()} units</td>
                        <td>${statusBadge}</td>
                        <td class="small text-muted">${escapeHtml(b.note || '-')}</td>
                        <td class="text-end">
                            <div class="btn-group btn-group-sm">
                                ${b.status !== 'Delivered' ? `
                                    <button class="btn btn-outline-success btn-xs" onclick="ApexApp.markBatchDelivered(${proj.id}, ${b.id})" title="Mark Delivered">
                                        <i class="fa-solid fa-check me-1"></i> Delivered
                                    </button>
                                ` : ''}
                                <button class="btn btn-outline-primary btn-xs" onclick="ApexApp.openAddBatchModal(${proj.id}, ${b.id})" title="Edit Batch">
                                    <i class="fa-solid fa-pen"></i>
                                </button>
                                <button class="btn btn-outline-danger btn-xs" onclick="ApexApp.deleteBatch(${proj.id}, ${b.id})" title="Delete Batch">
                                    <i class="fa-solid fa-trash"></i>
                                </button>
                            </div>
                        </td>
                    </tr>
                `;
            }).join('');
        }

        html += `
            <tr class="table-light fw-bold">
                <td colspan="2">Total Batches: ${proj.deliveryBatches.length}</td>
                <td>Delivered: ${deliveredBatchQty.toLocaleString()} / Scheduled: ${totalBatchQty.toLocaleString()} units</td>
                <td colspan="3" class="text-end text-muted small">Project Total PO Qty: ${(parseInt(proj.totalQty) || 0).toLocaleString()} units</td>
            </tr>
        `;

        tbody.innerHTML = html;
    }

    function openAddBatchModal(projId, batchId = null) {
        const proj = projects.find(p => p.id === projId);
        if (!proj) return;

        ensureProjectDeliveryBatches(proj);

        document.getElementById('batch-modal-proj-id').value = projId;
        document.getElementById('batch-modal-id').value = batchId || '';

        const modalTitle = document.getElementById('addBatchModalTitle');

        if (batchId) {
            const b = proj.deliveryBatches.find(item => item.id === batchId);
            if (b) {
                if (modalTitle) modalTitle.innerHTML = `<i class="fa-solid fa-pen text-success me-2"></i> Edit Delivery Batch`;
                document.getElementById('batch-modal-date').value = b.date || SYSTEM_DATE_STR;
                document.getElementById('batch-modal-qty').value = b.qty || 1;
                document.getElementById('batch-modal-status').value = b.status || 'Scheduled';
                document.getElementById('batch-modal-note').value = b.note || '';
            }
        } else {
            if (modalTitle) modalTitle.innerHTML = `<i class="fa-solid fa-truck-ramp-box text-success me-2"></i> Add Delivery Batch`;
            document.getElementById('batch-modal-date').value = SYSTEM_DATE_STR;
            const pendingRem = Math.max(1, proj.pendingQty || 100);
            document.getElementById('batch-modal-qty').value = pendingRem;
            document.getElementById('batch-modal-status').value = 'Scheduled';
            document.getElementById('batch-modal-note').value = 'Batch ' + (proj.deliveryBatches.length + 1);
        }

        if (addBatchModalBs) addBatchModalBs.show();
    }

    function handleSaveBatch(e) {
        e.preventDefault();
        const projId = parseInt(document.getElementById('batch-modal-proj-id').value);
        const batchIdStr = document.getElementById('batch-modal-id').value;
        const date = document.getElementById('batch-modal-date').value;
        const qty = parseInt(document.getElementById('batch-modal-qty').value) || 0;
        const status = document.getElementById('batch-modal-status').value || 'Scheduled';
        const note = document.getElementById('batch-modal-note').value.trim();

        const proj = projects.find(p => p.id === projId);
        if (!proj) return;

        ensureProjectDeliveryBatches(proj);

        if (batchIdStr) {
            const batchId = parseInt(batchIdStr);
            const b = proj.deliveryBatches.find(item => item.id === batchId);
            if (b) {
                b.date = date;
                b.qty = qty;
                b.status = status;
                b.note = note;
            }
        } else {
            const newId = proj.deliveryBatches.length > 0 ? Math.max(...proj.deliveryBatches.map(b => b.id)) + 1 : 1;
            proj.deliveryBatches.push({
                id: newId,
                batchNo: 'Batch ' + newId,
                date: date,
                qty: qty,
                status: status,
                note: note
            });
        }

        recalculateAllProjectsState();
        saveAllData();

        if (addBatchModalBs) addBatchModalBs.hide();

        renderDetailBatchesTab(projId);
        viewProjectDetails(projId);
        renderActiveView(getCurrentViewName());
    }

    function markBatchDelivered(projId, batchId) {
        const proj = projects.find(p => p.id === projId);
        if (!proj) return;

        ensureProjectDeliveryBatches(proj);
        const b = proj.deliveryBatches.find(item => item.id === batchId);
        if (b) {
            b.status = 'Delivered';
            recalculateAllProjectsState();
            saveAllData();
            renderDetailBatchesTab(projId);
            viewProjectDetails(projId);
            renderActiveView(getCurrentViewName());
        }
    }

    function deleteBatch(projId, batchId) {
        const proj = projects.find(p => p.id === projId);
        if (!proj) return;

        ensureProjectDeliveryBatches(proj);
        if (proj.deliveryBatches.length <= 1) {
            alert("A project must have at least 1 delivery batch schedule.");
            return;
        }

        if (confirm("Are you sure you want to remove this delivery batch?")) {
            proj.deliveryBatches = proj.deliveryBatches.filter(b => b.id !== batchId);
            recalculateAllProjectsState();
            saveAllData();
            renderDetailBatchesTab(projId);
            viewProjectDetails(projId);
            renderActiveView(getCurrentViewName());
        }
    }

    function renderDetailMaterialsTab(projId) {
        const tbody = document.getElementById('detail-materials-tbody');
        if (!tbody) return;

        const pMats = materials.filter(m => m.projectId === projId);

        if (pMats.length === 0) {
            tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted py-3">No BOM materials added for this project yet.</td></tr>`;
            return;
        }

        tbody.innerHTML = pMats.map(m => `
            <tr>
                <td class="fw-bold">${escapeHtml(m.name)}</td>
                <td><code>${escapeHtml(m.partNo)}</code></td>
                <td>${m.reqQty}</td>
                <td class="text-success">${m.availQty}</td>
                <td class="text-danger">${m.pendingQty}</td>
                <td>${escapeHtml(m.supplier)}</td>
                <td>${formatDate(m.expDate)}</td>
                <td>${getMaterialStatusBadge(m.status)}</td>
                <td>
                    <div class="btn-group btn-group-sm">
                        <button class="btn btn-xs btn-outline-secondary py-0 px-2" onclick="ApexApp.editMaterial(${m.id})" title="Edit Material">
                            <i class="fa-solid fa-pen-to-square"></i>
                        </button>
                        <button class="btn btn-xs btn-outline-danger py-0 px-2" onclick="ApexApp.deleteMaterial(${m.id})" title="Delete Material">
                            <i class="fa-solid fa-trash"></i>
                        </button>
                    </div>
                </td>
            </tr>
        `).join('');
    }

    function renderDetailTimelineTab(projId) {
        const container = document.getElementById('detail-timeline-container');
        if (!container) return;

        const pLogs = timelineEntries.filter(t => t.projectId === projId);

        if (pLogs.length === 0) {
            container.innerHTML = `<div class="text-center text-muted py-3">No activity logs recorded yet.</div>`;
            return;
        }

        container.innerHTML = pLogs.map(t => `
            <div class="timeline-item">
                <div class="timeline-dot"></div>
                <div class="timeline-card">
                    <div class="timeline-meta">
                        <span><i class="fa-regular fa-calendar me-1"></i>${formatDate(t.date)}</span>
                        <div class="d-flex align-items-center gap-2">
                            <span><i class="fa-solid fa-user me-1"></i>${escapeHtml(t.person)}</span>
                            <button class="btn btn-xs text-danger p-0 border-0 ms-1" onclick="ApexApp.deleteTimelineEntry(${t.id})" title="Delete Note">
                                <i class="fa-solid fa-xmark"></i>
                            </button>
                        </div>
                    </div>
                    <div class="timeline-text">${escapeHtml(t.update)}</div>
                </div>
            </div>
        `).join('');
    }

    function handleSaveDetailNote(e) {
        e.preventDefault();
        const projId = parseInt(document.getElementById('detail-note-proj-id').value);
        const person = document.getElementById('detail-note-person').value.trim() || 'Admin User';
        const text = document.getElementById('detail-note-text').value.trim();

        if (!text || !projId) return;

        timelineEntries.unshift({
            id: Date.now(),
            projectId: projId,
            date: SYSTEM_DATE_STR,
            update: text,
            person: person,
            status: 'Note'
        });

        saveAllData();
        document.getElementById('detail-note-text').value = '';

        renderDetailTimelineTab(projId);
        if (getCurrentViewName() === 'timeline') renderTimelineView();
    }

    function deleteTimelineEntry(id) {
        const entry = timelineEntries.find(t => t.id === id);
        if (!entry) return;

        if (confirm("Delete this activity log note?")) {
            const projId = entry.projectId;
            timelineEntries = timelineEntries.filter(t => t.id !== id);
            saveAllData();

            renderDetailTimelineTab(projId);
            if (getCurrentViewName() === 'timeline') renderTimelineView();
        }
    }

    /* ==========================================================================
       9. TIMELINE & LOGS GLOBAL VIEW
       ========================================================================== */

    function renderTimelineView() {
        const container = document.getElementById('timeline-feed-container');
        if (!container) return;

        // Auto pre-fill date input with SYSTEM_DATE_STR if empty
        const dateInput = document.getElementById('tl-input-date');
        if (dateInput && !dateInput.value) {
            dateInput.value = SYSTEM_DATE_STR;
        }

        const selProjId = document.getElementById('timeline-filter-project')?.value;
        let logs = [...timelineEntries];

        if (selProjId) {
            logs = logs.filter(t => t.projectId === parseInt(selProjId));
        }

        if (logs.length === 0) {
            container.innerHTML = `<div class="text-center text-muted py-4"><i class="fa-solid fa-timeline text-secondary fa-2x mb-2"></i><div>No activity updates logged for the selected filter.</div></div>`;
            return;
        }

        container.innerHTML = logs.map(t => {
            const p = projects.find(item => item.id === t.projectId);
            const projName = p ? p.name : 'Unknown Project';

            return `
                <div class="timeline-item">
                    <div class="timeline-dot"></div>
                    <div class="timeline-card">
                        <div class="timeline-meta">
                            <span class="fw-bold text-primary">${escapeHtml(projName)}</span>
                            <div class="d-flex align-items-center gap-2">
                                <span><i class="fa-regular fa-calendar me-1"></i>${formatDate(t.date)} — ${escapeHtml(t.person)}</span>
                                <button class="btn btn-xs text-danger p-0 border-0 ms-1" onclick="ApexApp.deleteTimelineEntry(${t.id})" title="Delete Entry">
                                    <i class="fa-solid fa-xmark"></i>
                                </button>
                            </div>
                        </div>
                        <div class="timeline-text">${escapeHtml(t.update)}</div>
                    </div>
                </div>
            `;
        }).join('');
    }

    function handleSaveTimelineEntry(e) {
        e.preventDefault();

        const projIdInput = document.getElementById('tl-input-project-id').value;
        const projId = parseInt(projIdInput);
        if (!projId || isNaN(projId)) {
            alert("Please select a valid project from the dropdown.");
            return;
        }

        let entryDate = document.getElementById('tl-input-date').value;
        if (!entryDate) {
            entryDate = SYSTEM_DATE_STR;
        }

        const updateText = document.getElementById('tl-input-update').value.trim();
        if (!updateText) {
            alert("Please type an update or description.");
            return;
        }

        const personName = document.getElementById('tl-input-person').value.trim() || 'Admin User';
        const stageStatus = document.getElementById('tl-input-status').value;

        timelineEntries.unshift({
            id: Date.now(),
            projectId: projId,
            date: entryDate,
            update: updateText,
            person: personName,
            status: stageStatus
        });

        saveAllData();

        // Reset form & restore default date
        const form = document.getElementById('form-add-timeline-entry');
        if (form) form.reset();
        const dateInput = document.getElementById('tl-input-date');
        if (dateInput) dateInput.value = SYSTEM_DATE_STR;

        renderTimelineView();
        alert("Timeline activity update posted successfully!");
    }

    /* ==========================================================================
       10. MATERIALS TRACKING VIEW
       ========================================================================== */

    function renderMaterialsView() {
        const tbody = document.getElementById('materials-table-tbody');
        if (!tbody) return;

        const searchQuery = (document.getElementById('mat-search-input')?.value || '').toLowerCase();
        const filterProjId = document.getElementById('mat-filter-project')?.value;
        const filterStat = document.getElementById('mat-filter-status')?.value;

        let filteredMats = materials.filter(m => {
            if (searchQuery) {
                const matchStr = (m.name + ' ' + m.partNo + ' ' + m.supplier).toLowerCase();
                if (!matchStr.includes(searchQuery)) return false;
            }
            if (filterProjId && m.projectId !== parseInt(filterProjId)) return false;
            if (filterStat && m.status !== filterStat) return false;
            return true;
        });

        if (filteredMats.length === 0) {
            tbody.innerHTML = `<tr><td colspan="11" class="text-center text-muted py-4">No materials found.</td></tr>`;
            return;
        }

        tbody.innerHTML = filteredMats.map(m => {
            const p = projects.find(item => item.id === m.projectId);
            const projName = p ? p.name : 'Unknown';
            const availPct = m.reqQty > 0 ? Math.round((m.availQty / m.reqQty) * 100) : 100;

            return `
                <tr>
                    <td class="fw-bold text-dark">${escapeHtml(projName)}</td>
                    <td>${escapeHtml(m.name)}</td>
                    <td><code>${escapeHtml(m.partNo)}</code></td>
                    <td>${m.reqQty}</td>
                    <td class="text-success fw-bold">${m.availQty}</td>
                    <td class="text-danger fw-bold">${m.pendingQty}</td>
                    <td>
                        <div class="mat-progress-bar">
                            <div class="mat-progress-fill" style="width: ${availPct}%"></div>
                        </div>
                        <span class="small fw-bold">${availPct}%</span>
                    </td>
                    <td>${escapeHtml(m.supplier)}</td>
                    <td>${formatDate(m.expDate)}</td>
                    <td>${getMaterialStatusBadge(m.status)}</td>
                    <td>
                        <div class="btn-group btn-group-sm">
                            <button class="btn btn-xs btn-outline-secondary py-0 px-2" onclick="ApexApp.editMaterial(${m.id})" title="Edit Material">
                                <i class="fa-solid fa-pen-to-square"></i>
                            </button>
                            <button class="btn btn-xs btn-outline-danger py-0 px-2" onclick="ApexApp.deleteMaterial(${m.id})" title="Delete Material">
                                <i class="fa-solid fa-trash"></i>
                            </button>
                        </div>
                    </td>
                </tr>
            `;
        }).join('');
    }

    function openAddMaterialModal(projId) {
        document.getElementById('addMaterialModalTitle').innerHTML = '<i class="fa-solid fa-boxes-packing text-primary me-2"></i> Add Material Record';
        document.getElementById('form-add-material').reset();
        document.getElementById('mat-modal-id').value = '';
        if (projId) {
            document.getElementById('mat-modal-proj-id').value = projId;
        }
        if (addMaterialModalBs) addMaterialModalBs.show();
    }

    function editMaterial(id) {
        const m = materials.find(item => item.id === id);
        if (!m) return;

        document.getElementById('addMaterialModalTitle').innerHTML = `<i class="fa-solid fa-pen-to-square text-primary me-2"></i> Edit Material: ${escapeHtml(m.name)}`;
        document.getElementById('mat-modal-id').value = m.id;
        document.getElementById('mat-modal-proj-id').value = m.projectId;
        document.getElementById('mat-modal-name').value = m.name;
        document.getElementById('mat-modal-partno').value = m.partNo || '';
        document.getElementById('mat-modal-supplier').value = m.supplier || '';
        document.getElementById('mat-modal-reqqty').value = m.reqQty;
        document.getElementById('mat-modal-availqty').value = m.availQty;
        document.getElementById('mat-modal-expdate').value = m.expDate || '';
        document.getElementById('mat-modal-status').value = m.status || 'Available';

        if (addMaterialModalBs) addMaterialModalBs.show();
    }

    function handleSaveMaterial(e) {
        e.preventDefault();

        const matId = document.getElementById('mat-modal-id').value;
        const req = parseInt(document.getElementById('mat-modal-reqqty').value) || 0;
        const avail = parseInt(document.getElementById('mat-modal-availqty').value) || 0;
        const pend = Math.max(0, req - avail);

        const matData = {
            projectId: parseInt(document.getElementById('mat-modal-proj-id').value),
            name: document.getElementById('mat-modal-name').value.trim(),
            partNo: document.getElementById('mat-modal-partno').value.trim(),
            supplier: document.getElementById('mat-modal-supplier').value.trim(),
            reqQty: req,
            availQty: avail,
            pendingQty: pend,
            expDate: document.getElementById('mat-modal-expdate').value,
            status: document.getElementById('mat-modal-status').value
        };

        if (matId) {
            const idx = materials.findIndex(m => m.id === parseInt(matId));
            if (idx !== -1) {
                materials[idx] = { ...materials[idx], ...matData };
            }
        } else {
            materials.push({
                id: Date.now(),
                ...matData
            });
        }

        recalculateAllProjectsState();
        saveAllData();

        if (addMaterialModalBs) addMaterialModalBs.hide();

        renderMaterialsView();
        const curView = getCurrentViewName();
        if (curView === 'projects') renderProjectsTable();
        if (curView === 'dashboard') renderDashboardView();

        const curDetailProjId = matData.projectId;
        const detailModalEl = document.getElementById('projectDetailModal');
        if (detailModalEl && detailModalEl.classList.contains('show')) {
            renderDetailMaterialsTab(curDetailProjId);
        }
    }

    function deleteMaterial(id) {
        const m = materials.find(item => item.id === id);
        if (!m) return;

        if (confirm(`Delete material record "${m.name}" (${m.partNo})?`)) {
            const projId = m.projectId;
            materials = materials.filter(item => item.id !== id);

            recalculateAllProjectsState();
            saveAllData();

            renderMaterialsView();
            if (getCurrentViewName() === 'projects') renderProjectsTable();
            if (getCurrentViewName() === 'dashboard') renderDashboardView();

            const detailModalEl = document.getElementById('projectDetailModal');
            if (detailModalEl && detailModalEl.classList.contains('show')) {
                renderDetailMaterialsTab(projId);
            }
        }
    }

    /* ==========================================================================
       10B. PROJECT BLOCKERS & RISKS CONTROLLER
       ========================================================================== */

    function renderBlockersView() {
        const tbody = document.getElementById('blockers-table-tbody');
        if (!tbody) return;

        // Stat cards
        const totalCount = blockers.length;
        const openCount = blockers.filter(b => b.status === 'Open').length;
        const inProgCount = blockers.filter(b => b.status === 'In Progress').length;
        const resolvedCount = blockers.filter(b => b.status === 'Resolved').length;
        const criticalCount = blockers.filter(b => b.severity === 'Critical' && b.status !== 'Resolved').length;

        const statOpen = document.getElementById('stat-open-blockers-count');
        const statCrit = document.getElementById('stat-critical-blockers-count');
        const statRes = document.getElementById('stat-resolved-blockers-count');
        const statTot = document.getElementById('stat-total-blockers-count');

        if (statOpen) statOpen.innerText = openCount + inProgCount;
        if (statCrit) statCrit.innerText = criticalCount;
        if (statRes) statRes.innerText = resolvedCount;
        if (statTot) statTot.innerText = totalCount;

        const searchQuery = (document.getElementById('blocker-search-input')?.value || '').toLowerCase();
        const filterProjId = document.getElementById('blocker-filter-project')?.value;
        const filterStat = document.getElementById('blocker-filter-status')?.value;
        const filterSev = document.getElementById('blocker-filter-severity')?.value;

        let filtered = blockers.filter(b => {
            if (searchQuery) {
                const matchStr = (b.title + ' ' + b.category + ' ' + b.owner + ' ' + (b.notes || '')).toLowerCase();
                if (!matchStr.includes(searchQuery)) return false;
            }
            if (filterProjId && b.projectId !== parseInt(filterProjId)) return false;
            if (filterStat && b.status !== filterStat) return false;
            if (filterSev && b.severity !== filterSev) return false;
            return true;
        });

        if (filtered.length === 0) {
            tbody.innerHTML = `<tr><td colspan="9" class="text-center text-muted py-4">No project blockers or risks found for current filters.</td></tr>`;
            return;
        }

        tbody.innerHTML = filtered.map(b => {
            const p = projects.find(item => item.id === b.projectId);
            const projName = p ? p.name : 'Unknown Project';

            let sevBadgeClass = 'bg-secondary';
            if (b.severity === 'Critical') sevBadgeClass = 'bg-danger';
            else if (b.severity === 'High') sevBadgeClass = 'bg-warning text-dark';
            else if (b.severity === 'Medium') sevBadgeClass = 'bg-info text-dark';
            else if (b.severity === 'Low') sevBadgeClass = 'bg-secondary';

            let statBadgeClass = 'bg-danger';
            if (b.status === 'Resolved') statBadgeClass = 'bg-success';
            else if (b.status === 'In Progress') statBadgeClass = 'bg-primary';

            return `
                <tr>
                    <td class="fw-bold text-dark"><a href="#" onclick="ApexApp.viewProjectDetails(${b.projectId}); return false;">${escapeHtml(projName)}</a></td>
                    <td class="fw-bold text-primary">${escapeHtml(b.title)}</td>
                    <td><span class="badge bg-light text-dark border">${escapeHtml(b.category)}</span></td>
                    <td><span class="badge ${sevBadgeClass}">${escapeHtml(b.severity)}</span></td>
                    <td><i class="fa-solid fa-user-tie me-1 text-secondary small"></i>${escapeHtml(b.owner)}</td>
                    <td>${formatDate(b.date)}</td>
                    <td><span class="badge ${statBadgeClass}">${escapeHtml(b.status)}</span></td>
                    <td class="small text-muted text-truncate" style="max-width: 250px;" title="${escapeHtml(b.notes || '')}">${escapeHtml(b.notes || '-')}</td>
                    <td>
                        <div class="btn-group btn-group-sm">
                            ${b.status !== 'Resolved' ? `
                                <button class="btn btn-xs btn-outline-success py-0 px-2" onclick="ApexApp.resolveBlocker(${b.id})" title="Mark Resolved">
                                    <i class="fa-solid fa-check"></i>
                                </button>
                            ` : ''}
                            <button class="btn btn-xs btn-outline-secondary py-0 px-2" onclick="ApexApp.editBlocker(${b.id})" title="Edit Blocker">
                                <i class="fa-solid fa-pen-to-square"></i>
                            </button>
                            <button class="btn btn-xs btn-outline-danger py-0 px-2" onclick="ApexApp.deleteBlocker(${b.id})" title="Delete Blocker">
                                <i class="fa-solid fa-trash"></i>
                            </button>
                        </div>
                    </td>
                </tr>
            `;
        }).join('');
    }

    function openAddBlockerModal(projId) {
        document.getElementById('blockerModalTitle').innerHTML = '<i class="fa-solid fa-triangle-exclamation text-danger me-2"></i> Log New Project Blocker / Risk';
        const form = document.getElementById('form-blocker');
        if (form) form.reset();

        const idInput = document.getElementById('blocker-input-id');
        if (idInput) idInput.value = '';

        const dateInput = document.getElementById('blocker-input-date');
        if (dateInput) dateInput.value = SYSTEM_DATE_STR;

        if (projId) {
            const projInput = document.getElementById('blocker-input-project-id');
            if (projInput) projInput.value = projId;
        }

        if (blockerModalBs) blockerModalBs.show();
    }

    function editBlocker(id) {
        const b = blockers.find(item => item.id === id);
        if (!b) return;

        document.getElementById('blockerModalTitle').innerHTML = `<i class="fa-solid fa-pen-to-square text-primary me-2"></i> Edit Blocker: ${escapeHtml(b.title)}`;
        
        const idInput = document.getElementById('blocker-input-id');
        if (idInput) idInput.value = b.id;

        const projInput = document.getElementById('blocker-input-project-id');
        if (projInput) projInput.value = b.projectId;

        const titleInput = document.getElementById('blocker-input-title');
        if (titleInput) titleInput.value = b.title;

        const catInput = document.getElementById('blocker-input-category');
        if (catInput) catInput.value = b.category || 'Component Shortage';

        const sevInput = document.getElementById('blocker-input-severity');
        if (sevInput) sevInput.value = b.severity || 'Medium';

        const ownerInput = document.getElementById('blocker-input-owner');
        if (ownerInput) ownerInput.value = b.owner || '';

        const dateInput = document.getElementById('blocker-input-date');
        if (dateInput) dateInput.value = b.date || SYSTEM_DATE_STR;

        const statInput = document.getElementById('blocker-input-status');
        if (statInput) statInput.value = b.status || 'Open';

        const notesInput = document.getElementById('blocker-input-notes');
        if (notesInput) notesInput.value = b.notes || '';

        if (blockerModalBs) blockerModalBs.show();
    }

    function handleSaveBlocker(e) {
        e.preventDefault();

        const blkId = document.getElementById('blocker-input-id')?.value;
        const projId = parseInt(document.getElementById('blocker-input-project-id')?.value);

        if (!projId || isNaN(projId)) {
            alert("Please select a project.");
            return;
        }

        const blockerData = {
            projectId: projId,
            title: document.getElementById('blocker-input-title').value.trim(),
            category: document.getElementById('blocker-input-category').value,
            severity: document.getElementById('blocker-input-severity').value,
            owner: document.getElementById('blocker-input-owner').value.trim() || 'Unassigned',
            date: document.getElementById('blocker-input-date').value || SYSTEM_DATE_STR,
            status: document.getElementById('blocker-input-status').value,
            notes: document.getElementById('blocker-input-notes').value.trim()
        };

        if (blkId) {
            const idx = blockers.findIndex(b => b.id === parseInt(blkId));
            if (idx !== -1) {
                blockers[idx] = { ...blockers[idx], ...blockerData };
            }
        } else {
            blockers.unshift({
                id: Date.now(),
                ...blockerData
            });

            // Also record in project activity timeline
            timelineEntries.unshift({
                id: Date.now() + 1,
                projectId: projId,
                date: SYSTEM_DATE_STR,
                update: `🚨 Blocked / Risk Logged: [${blockerData.severity}] ${blockerData.title}`,
                person: blockerData.owner,
                status: 'Hold'
            });
        }

        saveAllData();
        updateStorageStatusText();

        if (blockerModalBs) blockerModalBs.hide();

        renderBlockersView();
        const curView = getCurrentViewName();
        if (curView === 'projects') renderProjectsTable();
        if (curView === 'dashboard') renderDashboardView();

        const detailModalEl = document.getElementById('projectDetailModal');
        if (detailModalEl && detailModalEl.classList.contains('show')) {
            renderDetailBlockersTab(projId);
        }
        alert(blkId ? "Blocker log updated successfully!" : "New blocker logged successfully!");
    }

    function resolveBlocker(id) {
        const b = blockers.find(item => item.id === id);
        if (!b) return;

        if (confirm(`Mark blocker "${b.title}" as Resolved?`)) {
            b.status = 'Resolved';
            timelineEntries.unshift({
                id: Date.now(),
                projectId: b.projectId,
                date: SYSTEM_DATE_STR,
                update: `✅ Blocker Resolved: ${b.title}`,
                person: b.owner || 'Admin User',
                status: 'Completed'
            });

            saveAllData();
            updateStorageStatusText();

            renderBlockersView();
            const curView = getCurrentViewName();
            if (curView === 'projects') renderProjectsTable();
            if (curView === 'dashboard') renderDashboardView();

            const detailModalEl = document.getElementById('projectDetailModal');
            if (detailModalEl && detailModalEl.classList.contains('show')) {
                renderDetailBlockersTab(b.projectId);
            }
            alert("Blocker status updated to Resolved!");
        }
    }

    function deleteBlocker(id) {
        const b = blockers.find(item => item.id === id);
        if (!b) return;

        if (confirm(`Delete blocker log "${b.title}"?`)) {
            const projId = b.projectId;
            blockers = blockers.filter(item => item.id !== id);

            saveAllData();
            updateStorageStatusText();

            renderBlockersView();
            if (getCurrentViewName() === 'projects') renderProjectsTable();
            if (getCurrentViewName() === 'dashboard') renderDashboardView();

            const detailModalEl = document.getElementById('projectDetailModal');
            if (detailModalEl && detailModalEl.classList.contains('show')) {
                renderDetailBlockersTab(projId);
            }
        }
    }

    function renderDetailBlockersTab(projId) {
        const tbody = document.getElementById('detail-blockers-tbody');
        if (!tbody) return;

        const pBlks = blockers.filter(b => b.projectId === projId);

        if (pBlks.length === 0) {
            tbody.innerHTML = `<tr><td colspan="7" class="text-center text-muted py-3">No active blockers or risks logged for this project.</td></tr>`;
            return;
        }

        tbody.innerHTML = pBlks.map(b => {
            let sevBadgeClass = 'bg-secondary';
            if (b.severity === 'Critical') sevBadgeClass = 'bg-danger';
            else if (b.severity === 'High') sevBadgeClass = 'bg-warning text-dark';
            else if (b.severity === 'Medium') sevBadgeClass = 'bg-info text-dark';

            let statBadgeClass = 'bg-danger';
            if (b.status === 'Resolved') statBadgeClass = 'bg-success';
            else if (b.status === 'In Progress') statBadgeClass = 'bg-primary';

            return `
                <tr>
                    <td class="fw-bold text-dark">${escapeHtml(b.title)}</td>
                    <td><span class="badge bg-light text-dark border">${escapeHtml(b.category)}</span></td>
                    <td><span class="badge ${sevBadgeClass}">${escapeHtml(b.severity)}</span></td>
                    <td>${escapeHtml(b.owner)}</td>
                    <td>${formatDate(b.date)}</td>
                    <td><span class="badge ${statBadgeClass}">${escapeHtml(b.status)}</span></td>
                    <td>
                        <div class="btn-group btn-group-sm">
                            ${b.status !== 'Resolved' ? `
                                <button class="btn btn-xs btn-outline-success py-0 px-2" onclick="ApexApp.resolveBlocker(${b.id})" title="Mark Resolved">
                                    <i class="fa-solid fa-check"></i>
                                </button>
                            ` : ''}
                            <button class="btn btn-xs btn-outline-secondary py-0 px-2" onclick="ApexApp.editBlocker(${b.id})" title="Edit Blocker">
                                <i class="fa-solid fa-pen-to-square"></i>
                            </button>
                            <button class="btn btn-xs btn-outline-danger py-0 px-2" onclick="ApexApp.deleteBlocker(${b.id})" title="Delete Blocker">
                                <i class="fa-solid fa-trash"></i>
                            </button>
                        </div>
                    </td>
                </tr>
            `;
        }).join('');
    }

    /* ==========================================================================
       11. CUSTOMER & PROJECT LEADS VIEWS
       ========================================================================== */

    
    function renderCustomersView() {
        const listContainer = document.getElementById('customer-selector-list');
        const detailContainer = document.getElementById('customer-details-container');
        if (!listContainer || !detailContainer) return;

        const custMap = {};
        projects.forEach(p => {
            if (!custMap[p.customer]) custMap[p.customer] = [];
            custMap[p.customer].push(p);
        });

        const customers = Object.keys(custMap);
        if (customers.length === 0) return;

        listContainer.innerHTML = customers.map((c, idx) => `
            <button class="list-group-item list-group-item-action d-flex justify-content-between align-items-center ${idx === 0 ? 'active' : ''}" onclick="ApexApp.selectCustomer('${escapeHtml(c)}', this)">
                <span class="fw-bold">${escapeHtml(c)}</span>
                <span class="badge bg-primary rounded-pill">${custMap[c].length} Projects</span>
            </button>
        `).join('');

        // Render first customer details
        selectCustomer(customers[0]);
    }

    function selectCustomer(custName, clickedEl) {
        if (clickedEl) {
            document.querySelectorAll('#customer-selector-list .list-group-item').forEach(el => el.classList.remove('active'));
            clickedEl.classList.add('active');
        }

        const detailContainer = document.getElementById('customer-details-container');
        const custProjects = projects.filter(p => p.customer === custName);

        const total = custProjects.length;
        const comp = custProjects.filter(p => p.status === 'Completed').length;
        const active = custProjects.filter(p => p.status === 'In Progress' || p.status.includes('Under')).length;
        const delayed = custProjects.filter(p => p.status === 'Delayed').length;
        const totalPendingQty = custProjects.reduce((acc, p) => acc + (p.pendingQty || 0), 0);

        detailContainer.innerHTML = `
            <div class="table-card p-4 mb-4">
                <div class="d-flex justify-content-between align-items-center mb-3">
                    <h3 class="fw-bold text-dark mb-0"><i class="fa-solid fa-building me-2 text-primary"></i>${escapeHtml(custName)}</h3>
                    <span class="badge bg-primary fs-6">${total} Total Projects</span>
                </div>

                <div class="row g-3 text-center mb-4">
                    <div class="col-md-3">
                        <div class="p-3 border rounded bg-light">
                            <div class="text-muted small uppercase">Completed</div>
                            <div class="fs-4 fw-bold text-success">${comp}</div>
                        </div>
                    </div>
                    <div class="col-md-3">
                        <div class="p-3 border rounded bg-light">
                            <div class="text-muted small uppercase">Active</div>
                            <div class="fs-4 fw-bold text-primary">${active}</div>
                        </div>
                    </div>
                    <div class="col-md-3">
                        <div class="p-3 border rounded bg-light">
                            <div class="text-muted small uppercase">Delayed</div>
                            <div class="fs-4 fw-bold text-danger">${delayed}</div>
                        </div>
                    </div>
                    <div class="col-md-3">
                        <div class="p-3 border rounded bg-light">
                            <div class="text-muted small uppercase">Pending Qty</div>
                            <div class="fs-4 fw-bold text-dark">${totalPendingQty}</div>
                        </div>
                    </div>
                </div>

                <h5 class="fw-bold mb-3 border-bottom pb-2">Assigned Projects</h5>
                <div class="table-responsive">
                    <table class="table table-hover border align-middle">
                        <thead class="table-light">
                            <tr>
                                <th>Project Name</th>
                                <th>PO Number</th>
                                <th>Lead</th>
                                <th>Committed Date</th>
                                <th>Status</th>
                                <th>Completion</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${custProjects.map(p => `
                                <tr>
                                    <td class="fw-bold"><a href="#" onclick="ApexApp.viewProjectDetails(${p.id}); return false;">${escapeHtml(p.name)}</a></td>
                                    <td><code>${escapeHtml(p.poNumber)}</code></td>
                                    <td>${escapeHtml(p.lead)}</td>
                                    <td>${formatDate(p.committedDate)}</td>
                                    <td>${getStatusBadgeHtml(p.status)}</td>
                                    <td>${p.completion}%</td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        `;
    }

    function renderLeadsView() {
        const listContainer = document.getElementById('lead-selector-list');
        const detailContainer = document.getElementById('lead-details-container');
        if (!listContainer || !detailContainer) return;

        const leadMap = {};
        projects.forEach(p => {
            if (!leadMap[p.lead]) leadMap[p.lead] = [];
            leadMap[p.lead].push(p);
        });

        const leads = Object.keys(leadMap);
        if (leads.length === 0) return;

        listContainer.innerHTML = leads.map((l, idx) => `
            <button class="list-group-item list-group-item-action d-flex justify-content-between align-items-center ${idx === 0 ? 'active' : ''}" onclick="ApexApp.selectLead('${escapeHtml(l)}', this)">
                <span class="fw-bold"><i class="fa-solid fa-user-tie me-2 text-primary"></i>${escapeHtml(l)}</span>
                <span class="badge bg-primary rounded-pill">${leadMap[l].length} Projects</span>
            </button>
        `).join('');

        selectLead(leads[0]);
    }

    function selectLead(leadName, clickedEl) {
        if (clickedEl) {
            document.querySelectorAll('#lead-selector-list .list-group-item').forEach(el => el.classList.remove('active'));
            clickedEl.classList.add('active');
        }

        const detailContainer = document.getElementById('lead-details-container');
        const leadProjects = projects.filter(p => p.lead === leadName);

        const total = leadProjects.length;
        const comp = leadProjects.filter(p => p.status === 'Completed').length;
        const inProg = leadProjects.filter(p => p.status === 'In Progress' || p.status.includes('Under')).length;
        const delayed = leadProjects.filter(p => p.status === 'Delayed').length;
        const onHold = leadProjects.filter(p => p.status === 'On Hold').length;
        const avgComp = total > 0 ? Math.round(leadProjects.reduce((acc, p) => acc + (p.completion || 0), 0) / total) : 0;

        detailContainer.innerHTML = `
            <div class="table-card p-4 mb-4">
                <div class="d-flex justify-content-between align-items-center mb-3">
                    <h3 class="fw-bold text-dark mb-0"><i class="fa-solid fa-user-gear me-2 text-primary"></i>${escapeHtml(leadName)}</h3>
                    <span class="badge bg-primary fs-6">${total} Assigned Projects</span>
                </div>

                <div class="row row-cols-1 row-cols-md-5 g-3 text-center mb-4">
                    <div class="col">
                        <div class="p-3 border rounded bg-light">
                            <div class="text-muted small uppercase">Completed</div>
                            <div class="fs-4 fw-bold text-success">${comp}</div>
                        </div>
                    </div>
                    <div class="col">
                        <div class="p-3 border rounded bg-light">
                            <div class="text-muted small uppercase">In Progress</div>
                            <div class="fs-4 fw-bold text-warning">${inProg}</div>
                        </div>
                    </div>
                    <div class="col">
                        <div class="p-3 border rounded bg-light">
                            <div class="text-muted small uppercase">Delayed</div>
                            <div class="fs-4 fw-bold text-danger">${delayed}</div>
                        </div>
                    </div>
                    <div class="col">
                        <div class="p-3 border rounded bg-light">
                            <div class="text-muted small uppercase">On Hold</div>
                            <div class="fs-4 fw-bold text-secondary">${onHold}</div>
                        </div>
                    </div>
                    <div class="col">
                        <div class="p-3 border rounded bg-light">
                            <div class="text-muted small uppercase">Avg Completion</div>
                            <div class="fs-4 fw-bold text-primary">${avgComp}%</div>
                        </div>
                    </div>
                </div>

                <h5 class="fw-bold mb-3 border-bottom pb-2">Portfolio Overview</h5>
                <div class="table-responsive">
                    <table class="table table-hover border align-middle">
                        <thead class="table-light">
                            <tr>
                                <th>Project Name</th>
                                <th>Customer</th>
                                <th>PO Number</th>
                                <th>Priority</th>
                                <th>Committed Date</th>
                                <th>Status</th>
                                <th>Completion</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${leadProjects.map(p => `
                                <tr>
                                    <td class="fw-bold"><a href="#" onclick="ApexApp.viewProjectDetails(${p.id}); return false;">${escapeHtml(p.name)}</a></td>
                                    <td>${escapeHtml(p.customer)}</td>
                                    <td><code>${escapeHtml(p.poNumber)}</code></td>
                                    <td>${getPriorityBadgeHtml(p.priority)}</td>
                                    <td>${formatDate(p.committedDate)}</td>
                                    <td>${getStatusBadgeHtml(p.status)}</td>
                                    <td>${p.completion}%</td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>
            </div>
        `;
    }

    /* ==========================================================================
       11B. WEEKLY DELIVERIES & EXECUTIVE PROJECT SUMMARY ENGINE
       ========================================================================== */

    function getWeekBounds(refDateStr) {
        const d = refDateStr ? new Date(refDateStr) : new Date(SYSTEM_DATE);
        d.setHours(0, 0, 0, 0);

        const day = d.getDay(); // 0 = Sunday, 1 = Monday...
        const diffToMonday = (day === 0 ? -6 : 1 - day);
        const monday = new Date(d);
        monday.setDate(d.getDate() + diffToMonday);

        const sunday = new Date(monday);
        sunday.setDate(monday.getDate() + 6);
        sunday.setHours(23, 59, 59, 999);

        return { monday, sunday };
    }

    function getThisWeekDeliveriesCount() {
        const { monday, sunday } = getWeekBounds();
        return projects.filter(p => {
            if (p.status === 'Completed') return false;
            const targetDateStr = p.revisedDate || p.committedDate;
            if (!targetDateStr) return false;
            const tDate = new Date(targetDateStr);
            return tDate >= monday && tDate <= sunday;
        }).length;
    }

    function renderDeliveriesView() {
        const summaryTbody = document.getElementById('deliveries-summary-tbody');
        const breakdownContainer = document.getElementById('deliveries-weekly-breakdown-container');
        if (!summaryTbody || !breakdownContainer) return;

        const searchQuery = (document.getElementById('deliveries-search-input')?.value || '').toLowerCase();
        const weekFilter = document.getElementById('deliveries-week-filter')?.value || 'all';
        const statusFilter = document.getElementById('deliveries-status-filter')?.value || '';

        const { monday: curMonday, sunday: curSunday } = getWeekBounds();

        const nextMonday = new Date(curMonday);
        nextMonday.setDate(curMonday.getDate() + 7);
        const nextSunday = new Date(curSunday);
        nextSunday.setDate(curSunday.getDate() + 7);

        const week3Monday = new Date(curMonday);
        week3Monday.setDate(curMonday.getDate() + 14);
        const week3Sunday = new Date(curSunday);
        week3Sunday.setDate(curSunday.getDate() + 14);

        function getProjectWeekCategory(p) {
            if (p.status === 'Completed') return 'completed';
            const targetDateStr = p.revisedDate || p.committedDate;
            if (!targetDateStr) return 'week_4_plus';

            const tDate = new Date(targetDateStr);
            tDate.setHours(0, 0, 0, 0);

            if (tDate < curMonday) return 'overdue';
            if (tDate >= curMonday && tDate <= curSunday) return 'this_week';
            if (tDate >= nextMonday && tDate <= nextSunday) return 'next_week';
            if (tDate >= week3Monday && tDate <= week3Sunday) return 'week_3';
            return 'week_4_plus';
        }

        function getDeliveryHealthBadge(p, category) {
            if (p.status === 'Completed') return { badge: '<span class="badge bg-success">🟢 Completed</span>', icon: '🟢' };
            if (p.status === 'Delayed' || category === 'overdue') return { badge: '<span class="badge bg-danger">🔴 Overdue</span>', icon: '🔴' };
            if ((p.matAvailability || 100) < 80 || p.status === 'On Hold') return { badge: '<span class="badge bg-warning text-dark">🟡 At Risk</span>', icon: '🟡' };
            return { badge: '<span class="badge bg-success">🟢 On Track</span>', icon: '🟢' };
        }

        // Calculate KPI values
        const totalProjectsCount = projects.length;
        const thisWeekProjects = projects.filter(p => getProjectWeekCategory(p) === 'this_week');
        const thisWeekQty = thisWeekProjects.reduce((acc, p) => acc + (parseInt(p.pendingQty) || 0), 0);

        const nextWeekProjects = projects.filter(p => getProjectWeekCategory(p) === 'next_week');
        const nextWeekQty = nextWeekProjects.reduce((acc, p) => acc + (parseInt(p.pendingQty) || 0), 0);

        const overdueProjects = projects.filter(p => getProjectWeekCategory(p) === 'overdue');

        setElemText('kpi-deliv-total-projects', totalProjectsCount);
        setElemText('kpi-deliv-this-week-count', thisWeekProjects.length);
        setElemText('kpi-deliv-this-week-sub', `${thisWeekQty.toLocaleString()} units scheduled`);
        setElemText('kpi-deliv-next-week-count', nextWeekProjects.length);
        setElemText('kpi-deliv-next-week-sub', `${nextWeekQty.toLocaleString()} units next week`);
        setElemText('kpi-deliv-overdue-count', overdueProjects.length);

        const navDelivBadge = document.getElementById('nav-weekly-deliveries-badge');
        if (navDelivBadge) navDelivBadge.innerText = thisWeekProjects.length;

        // Render Customer Portfolio Snapshot Bar
        const custBar = document.getElementById('deliveries-customer-summary-bar');
        if (custBar) {
            const custMap = {};
            projects.forEach(p => {
                const cName = p.customer || 'Unassigned';
                if (!custMap[cName]) custMap[cName] = { count: 0, pending: 0, overdue: 0 };
                custMap[cName].count++;
                custMap[cName].pending += (parseInt(p.pendingQty) || 0);
                if (getProjectWeekCategory(p) === 'overdue') custMap[cName].overdue++;
            });

            const custKeys = Object.keys(custMap);
            if (custKeys.length === 0) {
                custBar.innerHTML = `<div class="col-12 text-muted small">No customer portfolio data.</div>`;
            } else {
                custBar.innerHTML = custKeys.map(cName => {
                    const info = custMap[cName];
                    return `
                        <div class="col-md-3 col-6">
                            <div class="p-2 border rounded bg-light d-flex justify-content-between align-items-center">
                                <div style="min-width: 0;">
                                    <div class="fw-bold text-dark small text-truncate" title="${escapeHtml(cName)}">${escapeHtml(cName)}</div>
                                    <div class="text-muted" style="font-size: 0.7rem;">${info.count} Projects • ${info.pending.toLocaleString()} units</div>
                                </div>
                                ${info.overdue > 0 ? `<span class="badge bg-danger ms-1">${info.overdue} Overdue</span>` : `<span class="badge bg-success-subtle text-success ms-1">🟢 Clear</span>`}
                            </div>
                        </div>
                    `;
                }).join('');
            }
        }

        // Filter projects list
        let filteredProjects = projects.filter(p => {
            if (searchQuery) {
                const matchStr = (p.name + ' ' + p.poNumber + ' ' + p.customer + ' ' + p.lead).toLowerCase();
                if (!matchStr.includes(searchQuery)) return false;
            }
            if (statusFilter && p.status !== statusFilter) return false;

            const category = getProjectWeekCategory(p);
            if (weekFilter !== 'all') {
                if (weekFilter === 'overdue' && category !== 'overdue') return false;
                if (weekFilter === 'this_week' && category !== 'this_week') return false;
                if (weekFilter === 'next_week' && category !== 'next_week') return false;
                if (weekFilter === 'week_3' && category !== 'week_3') return false;
                if (weekFilter === 'week_4_plus' && category !== 'week_4_plus') return false;
            }
            return true;
        });

        setElemText('deliveries-filtered-count-badge', `${filteredProjects.length} Projects`);

        filteredProjects.sort((a, b) => {
            const dateA = new Date(a.revisedDate || a.committedDate || '9999-12-31');
            const dateB = new Date(b.revisedDate || b.committedDate || '9999-12-31');
            return dateA - dateB;
        });

        // SECTION A1: Executive Short Summary Table
        if (filteredProjects.length === 0) {
            summaryTbody.innerHTML = `<tr><td colspan="11" class="text-center text-muted py-4">No projects match the active delivery filters.</td></tr>`;
        } else {
            summaryTbody.innerHTML = filteredProjects.map(p => {
                const targetDateStr = p.revisedDate || p.committedDate;
                const category = getProjectWeekCategory(p);
                const health = getDeliveryHealthBadge(p, category);

                let badgeHtml = '';
                if (p.status === 'Completed') {
                    badgeHtml = `<span class="badge bg-success">Delivered</span>`;
                } else if (category === 'overdue') {
                    const days = calculateDelayDays(p);
                    badgeHtml = `<span class="badge bg-danger">Overdue ${days}d</span>`;
                } else if (category === 'this_week') {
                    badgeHtml = `<span class="badge bg-success">Due This Week</span>`;
                } else if (category === 'next_week') {
                    badgeHtml = `<span class="badge bg-primary">Next Week</span>`;
                } else if (category === 'week_3') {
                    badgeHtml = `<span class="badge bg-info">Week 3</span>`;
                } else {
                    badgeHtml = `<span class="badge bg-secondary">Later</span>`;
                }

                let statusBadgeClass = 'bg-secondary';
                if (p.status === 'In Progress') statusBadgeClass = 'bg-warning text-dark';
                else if (p.status === 'Completed') statusBadgeClass = 'bg-success';
                else if (p.status === 'Delayed') statusBadgeClass = 'bg-danger';
                else if (p.status === 'Under Design') statusBadgeClass = 'bg-info text-dark';

                return `
                    <tr>
                        <td class="text-center fs-6">${health.icon}</td>
                        <td>
                            <div class="fw-bold text-dark">${escapeHtml(p.name)}</div>
                            <div class="small text-muted">${escapeHtml(p.type || 'Standard')} • Priority ${p.priority}</div>
                        </td>
                        <td class="fw-semibold">${escapeHtml(p.customer)}</td>
                        <td><code class="text-dark">${escapeHtml(p.poNumber)}</code></td>
                        <td>${escapeHtml(p.lead)}</td>
                        <td>
                            <div class="fw-semibold">${formatDate(targetDateStr)}</div>
                            <div class="mt-1">${badgeHtml}</div>
                        </td>
                        <td class="text-center">
                            <div class="fw-bold text-dark">${(parseInt(p.deliveredQty) || 0).toLocaleString()} / ${(parseInt(p.totalQty) || 0).toLocaleString()}</div>
                            <div class="small text-primary">Pending: ${(parseInt(p.pendingQty) || 0).toLocaleString()}</div>
                        </td>
                        <td>
                            <div class="d-flex align-items-center gap-2">
                                <div class="progress flex-grow-1" style="height: 6px;">
                                    <div class="progress-bar ${p.matAvailability >= 90 ? 'bg-success' : p.matAvailability >= 70 ? 'bg-warning' : 'bg-danger'}" style="width: ${p.matAvailability || 100}%;"></div>
                                </div>
                                <span class="small fw-bold">${p.matAvailability || 100}%</span>
                            </div>
                        </td>
                        <td>
                            <div class="d-flex align-items-center gap-2">
                                <div class="progress flex-grow-1" style="height: 6px;">
                                    <div class="progress-bar bg-primary" style="width: ${p.completion || 0}%;"></div>
                                </div>
                                <span class="small fw-bold">${p.completion || 0}%</span>
                            </div>
                        </td>
                        <td><span class="badge ${statusBadgeClass}">${p.status}</span></td>
                        <td class="text-end">
                            <button class="btn btn-sm btn-outline-primary" onclick="ApexApp.viewProjectDetails(${p.id})">
                                <i class="fa-solid fa-eye me-1"></i> Inspect
                            </button>
                        </td>
                    </tr>
                `;
            }).join('');
        }

        // SECTION A2: Compact Grid Cards Container
        const cardsGrid = document.getElementById('deliveries-summary-cards-grid');
        if (cardsGrid) {
            if (filteredProjects.length === 0) {
                cardsGrid.innerHTML = `<div class="col-12 text-center text-muted py-4">No projects found for active delivery filters.</div>`;
            } else {
                cardsGrid.innerHTML = filteredProjects.map(p => {
                    const targetDateStr = p.revisedDate || p.committedDate;
                    const category = getProjectWeekCategory(p);
                    const health = getDeliveryHealthBadge(p, category);

                    return `
                        <div class="col-md-6 col-lg-4">
                            <div class="card h-100 border shadow-sm p-3 position-relative">
                                <div class="d-flex justify-content-between align-items-start mb-2">
                                    <div style="min-width: 0; flex: 1;">
                                        <h6 class="fw-bold text-dark mb-0 text-truncate" title="${escapeHtml(p.name)}">${escapeHtml(p.name)}</h6>
                                        <div class="small text-muted">${escapeHtml(p.customer)} | <code>${escapeHtml(p.poNumber)}</code></div>
                                    </div>
                                    <div class="ms-2">${health.badge}</div>
                                </div>
                                <div class="d-flex justify-content-between small text-muted my-2 border-top border-bottom py-1">
                                    <span>Lead: <strong class="text-dark">${escapeHtml(p.lead)}</strong></span>
                                    <span>Target: <strong class="text-dark">${formatDate(targetDateStr)}</strong></span>
                                </div>
                                <div class="row g-2 text-center my-1" style="font-size: 0.78rem;">
                                    <div class="col-6">
                                        <div class="bg-light p-1.5 rounded">
                                            <div class="text-muted">Pending / Total</div>
                                            <div class="fw-bold text-dark">${(parseInt(p.pendingQty) || 0).toLocaleString()} / ${(parseInt(p.totalQty) || 0).toLocaleString()}</div>
                                        </div>
                                    </div>
                                    <div class="col-6">
                                        <div class="bg-light p-1.5 rounded">
                                            <div class="text-muted">Mat. Readiness</div>
                                            <div class="fw-bold ${p.matAvailability >= 80 ? 'text-success' : 'text-danger'}">${p.matAvailability || 100}%</div>
                                        </div>
                                    </div>
                                </div>
                                <div class="mt-2">
                                    <div class="d-flex justify-content-between small mb-1">
                                        <span class="text-muted">Completion</span>
                                        <span class="fw-bold">${p.completion || 0}%</span>
                                    </div>
                                    <div class="progress" style="height: 6px;">
                                        <div class="progress-bar bg-primary" style="width: ${p.completion || 0}%;"></div>
                                    </div>
                                </div>
                                <div class="mt-3 text-end">
                                    <button class="btn btn-xs btn-outline-primary px-2" onclick="ApexApp.viewProjectDetails(${p.id})">
                                        <i class="fa-solid fa-eye me-1"></i> Inspect Project
                                    </button>
                                </div>
                            </div>
                        </div>
                    `;
                }).join('');
            }
        }

        // SECTION B: Weekly Deliveries Schedule Breakdown
        const groups = [
            { key: 'overdue', label: '🔴 Overdue Deliveries', sub: 'Action required - passed target shipment date', class: 'border-danger bg-danger-subtle' },
            { key: 'this_week', label: `🚀 This Week (${formatDate(curMonday.toISOString().slice(0,10))} – ${formatDate(curSunday.toISOString().slice(0,10))})`, sub: 'Current shipping schedule target', class: 'border-success bg-success-subtle' },
            { key: 'next_week', label: `📦 Next Week (${formatDate(nextMonday.toISOString().slice(0,10))} – ${formatDate(nextSunday.toISOString().slice(0,10))})`, sub: 'Upcoming 7-day shipment schedule', class: 'border-primary bg-primary-subtle' },
            { key: 'week_3', label: `🗓️ Week 3 (${formatDate(week3Monday.toISOString().slice(0,10))} – ${formatDate(week3Sunday.toISOString().slice(0,10))})`, sub: 'Medium-term target schedule', class: 'border-info bg-info-subtle' },
            { key: 'week_4_plus', label: '⏳ Week 4+ & Future Deliveries', sub: 'Longer term manufacturing pipeline', class: 'border-secondary bg-light' }
        ];

        let breakdownHtml = '';
        groups.forEach(g => {
            const groupProjects = filteredProjects.filter(p => getProjectWeekCategory(p) === g.key);
            if (weekFilter !== 'all' && weekFilter !== g.key) return;
            if (groupProjects.length === 0 && weekFilter === 'all') return;

            const groupPendingQty = groupProjects.reduce((acc, p) => acc + (parseInt(p.pendingQty) || 0), 0);

            breakdownHtml += `
                <div class="card mb-3 border shadow-sm">
                    <div class="card-header ${g.class} py-2 px-3 d-flex justify-content-between align-items-center">
                        <div>
                            <span class="fw-bold text-dark">${g.label}</span>
                            <span class="small text-muted ms-2 d-none d-md-inline">${g.sub}</span>
                        </div>
                        <div>
                            <span class="badge bg-white text-dark border me-1">${groupProjects.length} Projects</span>
                            <span class="badge bg-dark text-white">${groupPendingQty.toLocaleString()} Pending Units</span>
                        </div>
                    </div>
                    <div class="card-body p-2">
                        ${groupProjects.length === 0 ? '<div class="text-muted small p-2 text-center">No projects scheduled for this week slot.</div>' : `
                            <div class="table-responsive">
                                <table class="table table-sm table-hover align-middle mb-0" style="font-size: 0.85rem;">
                                    <thead class="table-light">
                                        <tr>
                                            <th>Project Name</th>
                                            <th>Customer & PO</th>
                                            <th>Lead</th>
                                            <th>Target Date</th>
                                            <th>Pending / Total Qty</th>
                                            <th>Mat. Readiness</th>
                                            <th>Progress</th>
                                            <th>Delivery Batches (3+)</th>
                                            <th>Status</th>
                                            <th class="text-end">Inspect</th>
                                        </tr>
                                    </thead>
                                    <tbody>
                                        ${groupProjects.map(p => `
                                            <tr>
                                                <td class="fw-bold text-primary">${escapeHtml(p.name)}</td>
                                                <td><strong>${escapeHtml(p.customer)}</strong> <span class="text-muted">(${escapeHtml(p.poNumber)})</span></td>
                                                <td>${escapeHtml(p.lead)}</td>
                                                <td><span class="fw-semibold">${formatDate(p.revisedDate || p.committedDate)}</span></td>
                                                <td><strong class="text-dark">${(parseInt(p.pendingQty) || 0).toLocaleString()}</strong> / ${(parseInt(p.totalQty) || 0).toLocaleString()}</td>
                                                <td><span class="badge ${p.matAvailability >= 90 ? 'bg-success' : 'bg-warning text-dark'}">${p.matAvailability || 100}% Mat.</span></td>
                                                <td>
                                                    <div class="progress" style="height: 5px; width: 80px;">
                                                        <div class="progress-bar bg-primary" style="width: ${p.completion || 0}%;"></div>
                                                    </div>
                                                </td>
                                                <td>
                                                    <div class="d-flex flex-wrap gap-1" style="max-width: 260px;">
                                                        ${(p.deliveryBatches || []).map(b => `
                                                            <span class="badge ${b.status === 'Delivered' ? 'bg-success' : b.status === 'In Transit' ? 'bg-warning text-dark' : 'bg-light text-dark border'}" style="font-size: 0.7rem;" title="${escapeHtml(b.note || '')}">
                                                                ${escapeHtml(b.batchNo || 'Batch')}: ${b.qty}u (${formatDate(b.date)})
                                                            </span>
                                                        `).join('')}
                                                    </div>
                                                </td>
                                                <td><span class="badge bg-secondary">${p.status}</span></td>
                                                <td class="text-end">
                                                    <button class="btn btn-xs btn-outline-primary px-2 py-0" onclick="ApexApp.viewProjectDetails(${p.id})">
                                                        <i class="fa-solid fa-arrow-right"></i>
                                                    </button>
                                                </td>
                                            </tr>
                                        `).join('')}
                                    </tbody>
                                </table>
                            </div>
                        `}
                    </div>
                </div>
            `;
        });

        if (!breakdownHtml) {
            breakdownHtml = `<div class="text-center text-muted py-5"><i class="fa-solid fa-calendar-xmark fa-2x mb-2 text-secondary"></i><div>No weekly deliveries found matching selected week filter.</div></div>`;
        }

        breakdownContainer.innerHTML = breakdownHtml;
    }

    function exportDeliveriesToExcel() {
        if (typeof XLSX === 'undefined') {
            alert("SheetJS library not loaded. Unable to export.");
            return;
        }

        const data = projects.map(p => {
            const targetDateStr = p.revisedDate || p.committedDate;
            const { monday, sunday } = getWeekBounds();
            const tDate = new Date(targetDateStr);
            let weekBucket = "Week 4+";
            if (p.status === 'Completed') weekBucket = "Completed";
            else if (tDate < monday) weekBucket = "Overdue";
            else if (tDate >= monday && tDate <= sunday) weekBucket = "This Week";

            return {
                "Project ID": p.id,
                "Project Name": p.name,
                "Type": p.type || "New",
                "Customer": p.customer,
                "PO Number": p.poNumber,
                "Lead Manager": p.lead,
                "Committed Date": p.committedDate || "",
                "Revised Target Date": p.revisedDate || "",
                "Delivery Window": weekBucket,
                "Total Qty": p.totalQty || 0,
                "Delivered Qty": p.deliveredQty || 0,
                "Pending Qty": p.pendingQty || 0,
                "Material Availability %": (p.matAvailability || 100) + "%",
                "Completion %": (p.completion || 0) + "%",
                "Status": p.status
            };
        });

        const worksheet = XLSX.utils.json_to_sheet(data);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Weekly Deliveries");
        XLSX.writeFile(workbook, `Weekly_Deliveries_Summary_${SYSTEM_DATE_STR}.xlsx`);
    }

    /* ==========================================================================
       12. REPORTS GENERATOR ENGINE
       ========================================================================== */

    function generateReport() {
        const rType = document.getElementById('report-type-select').value;
        const fromDate = document.getElementById('report-from-date').value;
        const toDate = document.getElementById('report-to-date').value;

        const placeholder = document.getElementById('report-empty-placeholder');
        const wrapper = document.getElementById('report-table-wrapper');
        const thead = document.getElementById('report-table-thead');
        const tbody = document.getElementById('report-table-tbody');

        if (rType === 'blockers') {
            let blkData = [...blockers];
            if (fromDate) {
                blkData = blkData.filter(b => (b.date || '') >= fromDate);
            }
            if (toDate) {
                blkData = blkData.filter(b => (b.date || '') <= toDate);
            }

            placeholder.classList.add('d-none');
            wrapper.classList.remove('d-none');

            thead.innerHTML = `
                <tr>
                    <th>S.No</th>
                    <th>Project Name</th>
                    <th>Customer</th>
                    <th>Blocker Title</th>
                    <th>Category</th>
                    <th>Severity</th>
                    <th>Owner</th>
                    <th>Date Logged</th>
                    <th>Status</th>
                    <th>Action Notes</th>
                </tr>
            `;

            tbody.innerHTML = blkData.map((b, i) => {
                const p = projects.find(item => item.id === b.projectId);
                return `
                    <tr>
                        <td>${i + 1}</td>
                        <td class="fw-bold">${escapeHtml(p ? p.name : 'Unknown')}</td>
                        <td>${escapeHtml(p ? p.customer : '-')}</td>
                        <td class="fw-bold text-danger">${escapeHtml(b.title)}</td>
                        <td>${escapeHtml(b.category)}</td>
                        <td><span class="badge ${b.severity === 'Critical' ? 'bg-danger' : b.severity === 'High' ? 'bg-warning text-dark' : 'bg-info text-dark'}">${escapeHtml(b.severity)}</span></td>
                        <td>${escapeHtml(b.owner)}</td>
                        <td>${formatDate(b.date)}</td>
                        <td><span class="badge ${b.status === 'Resolved' ? 'bg-success' : 'bg-danger'}">${escapeHtml(b.status)}</span></td>
                        <td class="small">${escapeHtml(b.notes || '-')}</td>
                    </tr>
                `;
            }).join('');
            return;
        }

        let reportData = [...projects];

        if (fromDate) {
            reportData = reportData.filter(p => (p.committedDate || '') >= fromDate);
        }
        if (toDate) {
            reportData = reportData.filter(p => (p.committedDate || '') <= toDate);
        }

        if (rType === 'delayed') {
            reportData = reportData.filter(p => p.status === 'Delayed');
        }

        placeholder.classList.add('d-none');
        wrapper.classList.remove('d-none');

        // Render Headings
        thead.innerHTML = `
            <tr>
                <th>S.No</th>
                <th>Project Name</th>
                <th>Customer</th>
                <th>PO Number</th>
                <th>Project Lead</th>
                <th>Priority</th>
                <th>Committed Date</th>
                <th>Status</th>
                <th>Delivered Qty</th>
                <th>Pending Qty</th>
                <th>Completion %</th>
            </tr>
        `;

        tbody.innerHTML = reportData.map((p, i) => `
            <tr>
                <td>${i + 1}</td>
                <td class="fw-bold">${escapeHtml(p.name)}</td>
                <td>${escapeHtml(p.customer)}</td>
                <td><code>${escapeHtml(p.poNumber)}</code></td>
                <td>${escapeHtml(p.lead)}</td>
                <td>${getPriorityBadgeHtml(p.priority)}</td>
                <td>${formatDate(p.committedDate)}</td>
                <td>${getStatusBadgeHtml(p.status)}</td>
                <td>${p.deliveredQty}</td>
                <td>${p.pendingQty}</td>
                <td>${p.completion}%</td>
            </tr>
        `).join('');
    }

    /* ==========================================================================
       13. SHEETJS EXCEL & CSV IMPORT / EXPORT
       ========================================================================== */

    function exportToExcel() {
        if (typeof XLSX === 'undefined') {
            alert("Excel export library is loading. Please try again.");
            return;
        }

        const exportData = projects.map((p, i) => ({
            "S.No": i + 1,
            "Project Name": p.name,
            "New / Repeated": p.type,
            "Customer": p.customer,
            "PO Number": p.poNumber,
            "Description": p.description || '',
            "Priority": p.priority,
            "Project Lead": p.lead,
            "PO Date": p.poDate || '',
            "Customer Committed Date": p.committedDate || '',
            "Revised Delivery Date": p.revisedDate || '',
            "Total Quantity": p.totalQty,
            "Delivered Quantity": p.deliveredQty,
            "Pending Quantity": p.pendingQty,
            "Materials Availability %": p.matAvailability || 100,
            "Current Status": p.status,
            "Overall Completion %": p.completion,
            "Invoice Number": p.invoiceNumber || ''
        }));

        const worksheet = XLSX.utils.json_to_sheet(exportData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Projects Tracker");
        XLSX.writeFile(workbook, `Manufacturing_Project_Tracker_${SYSTEM_DATE_STR}.xlsx`);
    }

    function exportToCSV() {
        if (typeof XLSX === 'undefined') return;
        const exportData = projects.map((p, i) => ({
            "S.No": i + 1,
            "Project Name": p.name,
            "Customer": p.customer,
            "PO Number": p.poNumber,
            "Lead": p.lead,
            "Status": p.status,
            "Committed Date": p.committedDate,
            "Total Qty": p.totalQty,
            "Delivered Qty": p.deliveredQty,
            "Pending Qty": p.pendingQty,
            "Completion %": p.completion
        }));

        const worksheet = XLSX.utils.json_to_sheet(exportData);
        const csvOutput = XLSX.utils.sheet_to_csv(worksheet);

        const blob = new Blob([csvOutput], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.setAttribute("download", `Project_Export_${SYSTEM_DATE_STR}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    function downloadExcelTemplate() {
        const templateData = [
            {
                "Project Name": "Cell Modem Assembly",
                "New / Repeated": "New",
                "Customer": "Enphase",
                "PO Number": "PO-9001",
                "Description": "Sample project template",
                "Priority": "4",
                "Project Lead": "Raja Rajan",
                "PO Date": "2026-08-01",
                "Customer Committed Date": "2026-09-30",
                "Revised Delivery Date": "2026-09-30",
                "Total Quantity": 100,
                "Delivered Quantity": 0,
                "Materials Availability %": 100,
                "Current Status": "In Progress",
                "Invoice Number": "INV-100"
            }
        ];

        const worksheet = XLSX.utils.json_to_sheet(templateData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Template");
        XLSX.writeFile(workbook, "Project_Tracker_Import_Template.xlsx");
    }

    let parsedImportRows = [];

    function handleImportFileSelect(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function (evt) {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            parsedImportRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            if (parsedImportRows.length < 2) {
                alert("The selected file contains no data rows.");
                return;
            }

            renderColumnMappingSection(parsedImportRows[0]);
        };
        reader.readAsArrayBuffer(file);
    }

    function renderColumnMappingSection(headers) {
        const section = document.getElementById('import-mapping-section');
        const grid = document.getElementById('import-mapping-grid');
        const execBtn = document.getElementById('btn-execute-import');

        if (!section || !grid) return;

        section.classList.remove('d-none');
        if (execBtn) execBtn.disabled = false;

        const systemFields = [
            { key: 'name', label: 'Project Name' },
            { key: 'customer', label: 'Customer' },
            { key: 'poNumber', label: 'PO Number' },
            { key: 'lead', label: 'Project Lead' },
            { key: 'priority', label: 'Priority' },
            { key: 'committedDate', label: 'Committed Date' },
            { key: 'totalQty', label: 'Total Quantity' },
            { key: 'status', label: 'Current Status' }
        ];

        grid.innerHTML = systemFields.map(field => {
            // Find auto match header index
            const matchedIdx = headers.findIndex(h => h && h.toString().toLowerCase().includes(field.label.toLowerCase()));
            
            return `
                <div class="col-md-6 mb-2">
                    <label class="form-label small fw-bold text-muted">${field.label}</label>
                    <select class="form-select form-select-sm import-map-select" data-field="${field.key}">
                        <option value="">-- Ignore --</option>
                        ${headers.map((h, i) => `
                            <option value="${i}" ${i === matchedIdx ? 'selected' : ''}>${h}</option>
                        `).join('')}
                    </select>
                </div>
            `;
        }).join('');
    }

    function executeImportData() {
        if (parsedImportRows.length < 2) return;

        const mapSelects = document.querySelectorAll('.import-map-select');
        const map = {};
        mapSelects.forEach(sel => {
            const field = sel.getAttribute('data-field');
            const colIdx = sel.value;
            if (colIdx !== '') map[field] = parseInt(colIdx);
        });

        const rowsData = parsedImportRows.slice(1);
        let importedCount = 0;

        rowsData.forEach(row => {
            if (!row || row.length === 0) return;

            const name = map.name !== undefined ? row[map.name] : 'Imported Project';
            if (!name) return;

            const newId = projects.length > 0 ? Math.max(...projects.map(p => p.id)) + 1 : 101;
            const totalQty = map.totalQty !== undefined ? parseInt(row[map.totalQty]) || 100 : 100;

            const newProj = {
                id: newId,
                name: String(name),
                type: 'New',
                customer: map.customer !== undefined ? String(row[map.customer] || 'Default') : 'Default',
                poNumber: map.poNumber !== undefined ? String(row[map.poNumber] || `PO-${newId}`) : `PO-${newId}`,
                lead: map.lead !== undefined ? String(row[map.lead] || 'Raja Rajan') : 'Raja Rajan',
                priority: map.priority !== undefined ? String(row[map.priority] || '4') : '4',
                committedDate: map.committedDate !== undefined ? String(row[map.committedDate] || SYSTEM_DATE_STR) : SYSTEM_DATE_STR,
                totalQty: totalQty,
                deliveredQty: 0,
                pendingQty: totalQty,
                matAvailability: 100,
                status: map.status !== undefined ? String(row[map.status] || 'In Progress') : 'In Progress',
                completion: 10
            };

            projects.unshift(newProj);
            importedCount++;
        });

        recalculateAllProjectsState();
        saveAllData();

        if (importModalBs) importModalBs.hide();
        renderActiveView(getCurrentViewName());
        alert(`Successfully imported ${importedCount} projects!`);
    }

    /* JSON Backup Restore */
    function exportJsonBackup() {
        const backupObj = { projects, materials, timelineEntries, exportDate: new Date().toISOString() };
        const blob = new Blob([JSON.stringify(backupObj, null, 2)], { type: 'application/json' });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.setAttribute("download", `APEX_MFG_Backup_${SYSTEM_DATE_STR}.json`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    function importJsonBackup(e) {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function (evt) {
            try {
                const parsed = JSON.parse(evt.target.result);
                if (parsed.projects && Array.isArray(parsed.projects)) {
                    projects = parsed.projects;
                    if (parsed.materials) materials = parsed.materials;
                    if (parsed.timelineEntries) timelineEntries = parsed.timelineEntries;

                    recalculateAllProjectsState();
                    saveAllData();
                    renderActiveView(getCurrentViewName());
                    alert("JSON Backup restored successfully!");
                } else {
                    alert("Invalid backup file format.");
                }
            } catch (err) {
                alert("Failed to parse JSON file: " + err.message);
            }
        };
        reader.readAsText(file);
    }

    /* Notifications Modal */
    function renderNotificationsModal() {
        const body = document.getElementById('notifications-modal-body');
        if (!body) return;

        const overdue = projects.filter(p => p.status === "Delayed");
        const dueWeek = getDueThisWeekProjects();

        if (overdue.length === 0 && dueWeek.length === 0) {
            body.innerHTML = `<div class="p-4 text-center text-muted"><i class="fa-solid fa-bell-slash fa-2x mb-2 text-secondary"></i><div>No active alert notifications.</div></div>`;
            return;
        }

        let html = '<div class="list-group list-group-flush">';

        overdue.forEach(p => {
            html += `
                <div class="list-group-item p-3 border-bottom">
                    <div class="d-flex justify-content-between align-items-center mb-1">
                        <span class="fw-bold text-danger"><i class="fa-solid fa-triangle-exclamation me-1"></i>🔴 OVERDUE: ${escapeHtml(p.name)}</span>
                        <span class="badge bg-danger">${calculateDelayDays(p)} Days Overdue</span>
                    </div>
                    <div class="small text-muted">Customer: ${escapeHtml(p.customer)} | PO: ${escapeHtml(p.poNumber)} | Lead: ${escapeHtml(p.lead)}</div>
                    <div class="mt-2">
                        <button class="btn btn-xs btn-outline-primary" onclick="ApexApp.viewProjectDetails(${p.id})">Open Inspector</button>
                    </div>
                </div>
            `;
        });

        dueWeek.forEach(p => {
            html += `
                <div class="list-group-item p-3 border-bottom">
                    <div class="d-flex justify-content-between align-items-center mb-1">
                        <span class="fw-bold text-warning"><i class="fa-solid fa-clock me-1"></i>⏰ DUE THIS WEEK: ${escapeHtml(p.name)}</span>
                        <span class="badge bg-warning text-dark">Due ${formatDate(p.revisedDate || p.committedDate)}</span>
                    </div>
                    <div class="small text-muted">Customer: ${escapeHtml(p.customer)} | PO: ${escapeHtml(p.poNumber)}</div>
                </div>
            `;
        });

        html += '</div>';
        body.innerHTML = html;
    }

    /* Export Report to Excel */
    function exportReportToExcel() {
        const table = document.getElementById('report-output-table');
        if (!table || document.getElementById('report-table-wrapper').classList.contains('d-none')) {
            alert("Please generate a report first.");
            return;
        }

        const wb = XLSX.utils.table_to_book(table, { sheet: "Report" });
        XLSX.writeFile(wb, `Manufacturing_Report_${SYSTEM_DATE_STR}.xlsx`);
    }

    /* ==========================================================================
       14. HELPER UTILITIES & BADGES
       ========================================================================== */

    function getStatusBadgeHtml(status) {
        switch (status) {
            case 'Completed':
                return `<span class="badge-status badge-status-completed"><i class="fa-solid fa-circle-check"></i> Completed</span>`;
            case 'In Progress':
                return `<span class="badge-status badge-status-in-progress"><i class="fa-solid fa-circle-notch fa-spin"></i> In Progress</span>`;
            case 'Under Design':
                return `<span class="badge-status badge-status-under-design"><i class="fa-solid fa-compass-drafting"></i> Under Design</span>`;
            case 'Under Production':
                return `<span class="badge-status badge-status-under-production"><i class="fa-solid fa-gears"></i> Under Production</span>`;
            case 'Under Wiring':
                return `<span class="badge-status badge-status-under-wiring"><i class="fa-solid fa-bolt"></i> Under Wiring</span>`;
            case 'Delayed':
                return `<span class="badge-status badge-status-delayed"><i class="fa-solid fa-triangle-exclamation"></i> Delayed</span>`;
            case 'On Hold':
                return `<span class="badge-status badge-status-on-hold"><i class="fa-solid fa-pause"></i> On Hold</span>`;
            case 'Yet to Start':
            default:
                return `<span class="badge-status badge-status-yet-to-start"><i class="fa-solid fa-circle"></i> Yet to Start</span>`;
        }
    }

    function getPriorityBadgeHtml(prio) {
        switch (prio) {
            case '1A':
                return `<span class="badge-priority badge-p1a"><i class="fa-solid fa-fire"></i> 1A - Critical</span>`;
            case '1':
                return `<span class="badge-priority badge-p1">1 - High</span>`;
            case '2':
                return `<span class="badge-priority badge-p2">2 - Medium</span>`;
            case '3':
                return `<span class="badge-priority badge-p3">3 - Low</span>`;
            case '4':
                return `<span class="badge-priority badge-p4">4 - Normal</span>`;
            case '5':
            default:
                return `<span class="badge-priority badge-p5">5 - Lowest</span>`;
        }
    }

    function getMaterialStatusBadge(st) {
        switch (st) {
            case 'Available':
                return `<span class="badge bg-success">Available</span>`;
            case 'Partially Available':
                return `<span class="badge bg-warning text-dark">Partially Avail</span>`;
            case 'Pending':
                return `<span class="badge bg-danger">Pending</span>`;
            case 'Ordered':
                return `<span class="badge bg-info text-dark">Ordered</span>`;
            case 'Delayed':
                return `<span class="badge bg-danger">Delayed</span>`;
            default:
                return `<span class="badge bg-secondary">${st}</span>`;
        }
    }

    function getStatusColorHex(st) {
        switch (st) {
            case 'Completed': return '#10b981';
            case 'In Progress': return '#f59e0b';
            case 'Under Design': return '#2563eb';
            case 'Under Production': return '#f97316';
            case 'Under Wiring': return '#8b5cf6';
            case 'Delayed': return '#ef4444';
            case 'On Hold': return '#475569';
            case 'Yet to Start': return '#94a3b8';
            default: return '#64748b';
        }
    }

    function formatDate(dateStr) {
        if (!dateStr) return '-';
        const parts = dateStr.split('-');
        if (parts.length === 3) {
            return `${parts[2]}-${parts[1]}-${parts[0]}`;
        }
        return dateStr;
    }

    function escapeHtml(str) {
        if (!str) return '';
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    function copyToClipboard(text, btnElement) {
        if (!text) return;
        const textStr = String(text);
        if (navigator.clipboard && window.isSecureContext) {
            navigator.clipboard.writeText(textStr).then(() => {
                showCopyFeedback(btnElement, textStr);
            }).catch(() => fallbackCopyText(textStr, btnElement));
        } else {
            fallbackCopyText(textStr, btnElement);
        }
    }

    function fallbackCopyText(textStr, btnElement) {
        const textArea = document.createElement("textarea");
        textArea.value = textStr;
        textArea.style.position = "fixed";
        textArea.style.left = "-999999px";
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        try {
            document.execCommand('copy');
            showCopyFeedback(btnElement, textStr);
        } catch (err) {
            alert("Copying text: " + textStr);
        }
        document.body.removeChild(textArea);
    }

    function showCopyFeedback(btnElement, textStr) {
        if (btnElement) {
            const originalHtml = btnElement.innerHTML;
            btnElement.innerHTML = `<i class="fa-solid fa-check text-success"></i> Copied!`;
            setTimeout(() => {
                btnElement.innerHTML = originalHtml;
            }, 1400);
        } else {
            alert("Copied to clipboard: " + textStr);
        }
    }

    /* ==========================================================================
       15. PUBLIC APEX APPLICATION INTERFACE FOR INLINE HTML EVENT HANDLERS
       ========================================================================== */

    window.ApexApp = {
        openAddBatchModal: openAddBatchModal,
        markBatchDelivered: markBatchDelivered,
        deleteBatch: deleteBatch,
        renderDetailBatchesTab: renderDetailBatchesTab,
        copyToClipboard: copyToClipboard,
        sortTable: function (colId) {
            if (sortColumn === colId) {
                sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
            } else {
                sortColumn = colId;
                sortDirection = 'asc';
            }
            renderProjectsTable();
        },
        changePage: function (page) {
            currentPage = page;
            renderProjectsTable();
        },
        toggleColumnVisibility: function (colId) {
            const col = columnsConfig.find(c => c.id === colId);
            if (col) {
                col.visible = !col.visible;
                renderProjectsTable();
            }
        },
        openAddProjectModal: openAddProjectModal,
        viewProjectDetails: viewProjectDetails,
        editProject: editProject,
        deleteProject: deleteProject,
        toggleStageStatus: toggleStageStatus,
        openAddMaterialModal: openAddMaterialModal,
        editMaterial: editMaterial,
        deleteMaterial: deleteMaterial,
        openAddBlockerModal: openAddBlockerModal,
        editBlocker: editBlocker,
        resolveBlocker: resolveBlocker,
        deleteBlocker: deleteBlocker,
        deleteTimelineEntry: deleteTimelineEntry,
        exportDeliveriesToExcel: exportDeliveriesToExcel,
        selectCustomer: selectCustomer,
        selectLead: selectLead,
        lockApp: lockApp,
        logoutUser: lockApp,
        promptChangePin: promptChangePin,
        editUserProfile: editUserProfile,
        filterAndNavProjects: function (type, val) {
            if (notificationsModalBs) notificationsModalBs.hide();

            if (type === 'status') {
                filterStatus = val;
                const statusSel = document.getElementById('filter-status-select');
                if (statusSel) statusSel.value = val;
            } else if (type === 'quick') {
                activeQuickFilter = val;
                document.querySelectorAll('.btn-filter-chip').forEach(c => {
                    c.classList.toggle('active', c.getAttribute('data-filter') === val);
                });
            }

            switchNavTo('projects');
        }
    };

})();
