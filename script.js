// Global variables
let employees = [];
let skills = [];
let schedule = [];
let currentView = 'schedule'; // Default view
let translations = {}; // Store loaded translations
let currentLanguage = 'en'; // Default language

// i18n functions
async function loadLanguage(lang) {
    try {
        const response = await fetch(`lang/${lang}.json`);
        if (!response.ok) {
            throw new Error(`Failed to load language file for ${lang}`);
        }
        
        translations = await response.json();
        currentLanguage = lang;
        localStorage.setItem('preferredLanguage', lang);
        
        // Update all UI text with new language
        updateUILanguage();
        
        return true;
    } catch (error) {
        console.error('Error loading language file:', error);
        return false;
    }
}

function switchLanguage(lang) {
    if (lang === currentLanguage) return;
    
    loadLanguage(lang).then(success => {
        if (success) {
            console.log(`Switched to ${lang} language`);
        } else {
            console.error(`Failed to switch to ${lang} language`);
        }
    });
}

function t(key) {
    // Split the key path by dots to navigate through nested objects
    const keys = key.split('.');
    let result = translations;
    
    // Traverse the translations object
    for (const k of keys) {
        if (result && result[k] !== undefined) {
            result = result[k];
        } else {
            // If translation not found, return the key itself
            console.warn(`Translation missing for key: ${key}`);
            return key;
        }
    }
    
    return result;
}

function updateUILanguage() {
    // Update document title
    document.title = t('app.title');
    
    // Update navbar
    document.querySelector('.navbar-brand').textContent = t('app.title');
    document.getElementById('nav-schedule').textContent = t('nav.schedule');
    document.getElementById('nav-employees').textContent = t('nav.employees');
    document.getElementById('nav-teams').textContent = t('nav.teams');
    document.getElementById('nav-skills').textContent = t('nav.skills');
    
    // Update language dropdown
    document.getElementById('language-dropdown').textContent = t('language.select');
    
    // Update data management dropdown
    document.getElementById('data-dropdown').textContent = t('nav.dataManagement');
    document.getElementById('export-data-button').textContent = t('nav.exportData');
    document.getElementById('import-data-button').textContent = t('nav.importData');
    document.getElementById('clear-data-button').textContent = t('nav.clearData');
    document.getElementById('load-demo-button').textContent = t('nav.loadDemo');
    
    // Update loading text
    document.querySelector('.spinner-border .visually-hidden').textContent = t('app.loading');
    
    // Update error container
    document.querySelector('#error-container h4').textContent = t('app.error.title');
    
    // Update tooltips
    document.getElementById('export-data-button').title = t('tooltips.exportData');
    document.getElementById('import-data-button').title = t('tooltips.importData');
    document.getElementById('clear-data-button').title = t('tooltips.clearData');
    document.getElementById('load-demo-button').title = t('tooltips.loadDemo');
    
    // Update the current view based on which view is active
    if (currentView === 'schedule') updateScheduleView();
    else if (currentView === 'employees') updateEmployeesView();
    else if (currentView === 'teams') updateTeamsView();
    else if (currentView === 'skills') updateSkillsView();
    
    // Update the date range button text
    updateDateRangeButtonText();
}

function updateScheduleView() {
    // Update schedule view labels
    document.querySelector('label[for="week-selector"]').textContent = t('schedule.selectWeek');
    document.getElementById('export-button').textContent = t('schedule.exportToExcel');
    document.getElementById('plan-button').textContent = t('schedule.generatePlan');
    
    // Add skills toggle if it doesn't exist yet
    if (!document.getElementById('skills-toggle-container')) {
        const exportButton = document.getElementById('export-button');
        
        // Create toggle container
        const toggleContainer = document.createElement('div');
        toggleContainer.id = 'skills-toggle-container';
        toggleContainer.classList.add('form-check', 'form-switch', 'ms-2', 'd-inline-block');
        
        // Create toggle input
        const toggleInput = document.createElement('input');
        toggleInput.type = 'checkbox';
        toggleInput.id = 'skills-toggle';
        toggleInput.classList.add('form-check-input');
        toggleInput.checked = true; // Set it to checked by default
        toggleInput.addEventListener('change', function() {
            const skillsRows = document.querySelectorAll('.skills-row');
            skillsRows.forEach(row => {
                if (this.checked) {
                    row.style.display = '';
                } else {
                    row.style.display = 'none';
                }
            });
        });
        
        // Create toggle label
        const toggleLabel = document.createElement('label');
        toggleLabel.classList.add('form-check-label', 'ms-1');
        toggleLabel.setAttribute('for', 'skills-toggle');
        toggleLabel.textContent = 'Show Skills';
        
        // Add toggle elements to container
        toggleContainer.appendChild(toggleInput);
        toggleContainer.appendChild(toggleLabel);
        
        // Insert toggle before the export button
        exportButton.parentNode.insertBefore(toggleContainer, exportButton);
    }
    
    // Update table headers
    const scheduleHeaders = document.querySelectorAll('#schedule-table thead th');
    if (scheduleHeaders.length >= 4) {
        scheduleHeaders[0].textContent = t('schedule.day');
        scheduleHeaders[1].textContent = t('schedule.amShift');
        scheduleHeaders[2].textContent = t('schedule.pmShift');
        scheduleHeaders[3].textContent = t('schedule.nightShift');
    }
    
    // Re-render schedule to update all cells with new language
    renderSchedule();
}

function updateEmployeesView() {
    // Update employees view labels
    document.querySelector('#employees-view .card-header h5').textContent = t('employees.title');
    document.getElementById('add-employee-button').textContent = t('employees.addEmployee');
    
    // Update table headers
    const employeeHeaders = document.querySelectorAll('#employees-table thead th');
    if (employeeHeaders.length >= 5) {
        employeeHeaders[0].textContent = t('employees.name');
        employeeHeaders[1].textContent = t('employees.team');
        employeeHeaders[2].textContent = t('employees.skills');
        employeeHeaders[3].textContent = t('employees.absentDates');
        employeeHeaders[4].textContent = t('employees.actions');
    }
    
    // Re-render employees to update all with new language
    renderEmployees();
}

function updateTeamsView() {
    // Update team headers
    const teamHeaders = document.querySelectorAll('.team-header');
    if (teamHeaders.length >= 3) {
        teamHeaders[0].textContent = t('teams.teamA');
        teamHeaders[1].textContent = t('teams.teamB');
        teamHeaders[2].textContent = t('teams.teamC');
    }
    
    // Re-render teams with new language
    renderTeams();
}

function updateSkillsView() {
    // Update skills view labels
    document.querySelector('#skills-view .card-header h5').textContent = t('skills.title');
    document.getElementById('add-skill-button').textContent = t('skills.addSkill');
    
    // Update table headers
    const skillHeaders = document.querySelectorAll('#skills-table thead th');
    if (skillHeaders.length >= 4) {
        skillHeaders[0].textContent = t('skills.skillName');
        skillHeaders[1].textContent = t('skills.description');
        skillHeaders[2].textContent = t('skills.employeesWithSkill');
        skillHeaders[3].textContent = t('skills.actions');
    }
    
    // Re-render skills with new language
    renderSkills();
}

// Initialize the application data
let appData = {
    employees: [],
    skills: [],    // Array of available skills
    schedule: {},
    weekVersions: {}, // Stores multiple versions of schedules for each week
    lockedWeeks: {},  // Tracks which weeks are locked
    currentWeek: 0,
    weekDates: []
};

// DOM elements
let elements = {};

// Constants
const DAYS_OF_WEEK = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
const DAYS_OF_WEEK_DE = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag'];
const SHIFT_TYPES = ['AM', 'PM', 'Night'];

// Get localized day name
function getLocalizedDay(day) {
    const dayIndex = DAYS_OF_WEEK.indexOf(day);
    if (dayIndex === -1) return day;
    
    return currentLanguage === 'de' ? DAYS_OF_WEEK_DE[dayIndex] : DAYS_OF_WEEK[dayIndex];
}

// Safely initialize Bootstrap modals
function initModals() {
    try {
        // Hide loading indicator first
        const loadingIndicator = document.getElementById('loading-indicator');
        if (loadingIndicator) {
            loadingIndicator.classList.add('d-none');
        }
        
        // Employee modal elements
        elements.employeeModal = new bootstrap.Modal(document.getElementById('employee-modal'));
        elements.employeeForm = document.getElementById('employee-form');
        elements.employeeId = document.getElementById('employee-id');
        elements.employeeName = document.getElementById('employee-name');
        elements.employeeTeam = document.getElementById('employee-team');
        elements.employeeSkills = document.getElementById('employee-skills');
        elements.absentDatesContainer = document.getElementById('absent-dates-container');
        elements.addDateBtn = document.getElementById('add-date-btn');
        elements.saveEmployeeBtn = document.getElementById('save-employee-btn');
        
        // Add date range button
        const addDateRangeBtn = document.createElement('button');
        addDateRangeBtn.type = 'button';
        addDateRangeBtn.id = 'add-date-range-btn';
        addDateRangeBtn.classList.add('btn', 'btn-outline-primary', 'mt-2', 'ms-2');
        addDateRangeBtn.textContent = currentLanguage === 'de' ? 'Zeitraum hinzufügen' : 'Add Date Range';
        addDateRangeBtn.addEventListener('click', addAbsentDateRange);
        
        // Insert the new button after the existing "Add Absent Date" button
        elements.addDateBtn.parentNode.appendChild(addDateRangeBtn);

        const shiftEditModalEl = document.getElementById('shift-edit-modal');
        const skillModalEl = document.getElementById('skill-modal');
        
        if (!document.getElementById('employee-modal') || !shiftEditModalEl || !skillModalEl) {
            throw new Error('Modal elements not found in the DOM');
        }
        
        elements.shiftEditModal = new bootstrap.Modal(shiftEditModalEl);
        elements.skillModal = new bootstrap.Modal(skillModalEl);
        
        // Initialize other modal elements
        elements.shiftEditForm = document.getElementById('shift-edit-form');
        elements.shiftDay = document.getElementById('shift-day');
        elements.shiftType = document.getElementById('shift-type');
        elements.shiftEmployees = document.getElementById('shift-employees');
        elements.shiftInfoLabel = document.getElementById('shift-info-label');
        elements.saveShiftBtn = document.getElementById('save-shift-btn');
        
        elements.skillForm = document.getElementById('skill-form');
        elements.skillId = document.getElementById('skill-id');
        elements.skillName = document.getElementById('skill-name');
        elements.skillDescription = document.getElementById('skill-description');
        elements.saveSkillBtn = document.getElementById('save-skill-btn');
        
        return true;
    } catch (error) {
        console.error('Error initializing Bootstrap modals:', error);
        showError('Error initializing Bootstrap modals: ' + error.message);
        return false;
    }
}

// Show error message
function showError(message) {
    const errorContainer = document.getElementById('error-container');
    const errorMessage = document.getElementById('error-message');
    
    if (errorContainer && errorMessage) {
        errorMessage.textContent = message;
        errorContainer.classList.remove('d-none');
        
        // Hide loading indicator if it's still showing
        const loadingIndicator = document.getElementById('loading-indicator');
        if (loadingIndicator) {
            loadingIndicator.classList.add('d-none');
        }
    } else {
        // Fallback to alert if the error container isn't found
        alert('Application Error: ' + message);
    }
}

// Initialize the application
function init() {
    // Wait for DOM to be fully loaded
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initAfterDOM);
    } else {
        initAfterDOM();
    }
}

// Initialize app after DOM is loaded
function initAfterDOM() {
    try {
        // Initialize DOM elements
        elements = {
            // Navigation
            navSchedule: document.getElementById('nav-schedule'),
            navEmployees: document.getElementById('nav-employees'),
            navTeams: document.getElementById('nav-teams'),
            navSkills: document.getElementById('nav-skills'),
            
            // Views
            scheduleView: document.getElementById('schedule-view'),
            employeesView: document.getElementById('employees-view'),
            teamsView: document.getElementById('teams-view'),
            skillsView: document.getElementById('skills-view'),
            
            // Schedule elements
            weekSelector: document.getElementById('week-selector'),
            scheduleTable: document.getElementById('schedule-table'),
            planButton: document.getElementById('plan-button'),
            exportButton: document.getElementById('export-button'),
            
            // Employee elements
            employeesTable: document.getElementById('employees-table'),
            addEmployeeButton: document.getElementById('add-employee-button'),
            
            // Team elements
            teamAList: document.getElementById('team-a-list'),
            teamBList: document.getElementById('team-b-list'),
            teamCList: document.getElementById('team-c-list'),
            
            // Skills elements
            skillsTable: document.getElementById('skills-table'),
            addSkillButton: document.getElementById('add-skill-button'),
            
            // Modal elements
            employeeForm: document.getElementById('employee-form'),
            employeeId: document.getElementById('employee-id'),
            employeeName: document.getElementById('employee-name'),
            employeeTeam: document.getElementById('employee-team'),
            employeeSkills: document.getElementById('employee-skills'),
            absentDatesContainer: document.getElementById('absent-dates-container'),
            addDateBtn: document.getElementById('add-date-btn'),
            saveEmployeeBtn: document.getElementById('save-employee-btn'),
            
            shiftEditForm: document.getElementById('shift-edit-form'),
            shiftDay: document.getElementById('shift-day'),
            shiftType: document.getElementById('shift-type'),
            shiftInfoLabel: document.getElementById('shift-info-label'),
            shiftEmployees: document.getElementById('shift-employees'),
            saveShiftBtn: document.getElementById('save-shift-btn'),
            
            skillForm: document.getElementById('skill-form'),
            skillId: document.getElementById('skill-id'),
            skillName: document.getElementById('skill-name'),
            skillDescription: document.getElementById('skill-description'),
            saveSkillBtn: document.getElementById('save-skill-btn'),
            
            exportDataButton: document.getElementById('export-data-button'),
            importDataButton: document.getElementById('import-data-button'),
            clearDataButton: document.getElementById('clear-data-button'),
            loadDemoButton: document.getElementById('load-demo-button')
        };

        // Verify critical elements exist
        const criticalElements = [
            'scheduleView', 'employeesView', 'teamsView', 
            'weekSelector', 'scheduleTable', 
            'employeeForm', 'addEmployeeButton'
        ];
        
        for (const element of criticalElements) {
            if (!elements[element]) {
                throw new Error(`Critical element ${element} not found in the DOM`);
            }
        }

        // Initialize bootstrap modals
        if (!initModals()) {
            return; // Exit if modal initialization fails
        }
        
        // Initialize application data
        loadData();
        
        // Load preferred language or default to English
        const preferredLanguage = localStorage.getItem('preferredLanguage') || 'en';
        loadLanguage(preferredLanguage).then(() => {
            // Initialize the rest of the UI after language is loaded
            setupEventListeners();
            setupNavigation();
            generateWeekOptions();
            renderSchedule();
            renderEmployees();
            renderTeams();
            renderSkills();
            
            // Add information about new features
            showFeatureInfo();
            
            // Hide loading indicator (if it's still showing)
            const loadingIndicator = document.getElementById('loading-indicator');
            if (loadingIndicator) {
                loadingIndicator.classList.add('d-none');
            }
            
            console.log('Application initialized successfully with language:', currentLanguage);
        }).catch(err => {
            console.error('Error loading language:', err);
            
            // If language loading fails, continue with default text
            setupEventListeners();
            setupNavigation();
            generateWeekOptions();
            renderSchedule();
            renderEmployees();
            renderTeams();
            renderSkills();
            
            // Add information about new features
            showFeatureInfo();
            
            console.log('Application initialized with default language due to error');
        });
    } catch (error) {
        console.error('Error initializing application:', error);
        showError('Error initializing application: ' + error.message);
    }
}

// Load data from localStorage or initialize with default values
function loadData() {
    const savedData = localStorage.getItem('shiftPlannerData');
    
    if (savedData) {
        appData = JSON.parse(savedData);
        
        // Ensure backward compatibility with older data format
        if (!appData.weekVersions) {
            appData.weekVersions = {};
        }
        if (!appData.lockedWeeks) {
            appData.lockedWeeks = {};
        }
        if (!appData.skills) {
            appData.skills = [];
        }
        
        // For each employee, ensure they have a skills array
        appData.employees.forEach(employee => {
            if (!employee.skills) {
                employee.skills = [];
            }
        });
        
        // For each week in the schedule, ensure there's a version entry
        Object.keys(appData.schedule).forEach(weekIndex => {
            if (!appData.weekVersions[weekIndex]) {
                // Create default version if not exists
                appData.weekVersions[weekIndex] = {
                    "v1": {
                        name: "Initial Version",
                        date: new Date().toISOString(),
                        isActive: true,
                        schedule: appData.schedule[weekIndex]
                    }
                };
            }
        });
    } else {
        // Initialize with some sample skills
        const defaultSkills = [
            { id: 1, name: 'First Aid', description: 'Certified in basic first aid' },
            { id: 2, name: 'Forklift', description: 'Licensed to operate forklifts' },
            { id: 3, name: 'Team Lead', description: 'Can lead a team' },
            { id: 4, name: 'Technical Support', description: 'Can provide technical assistance' }
        ];
        
        // Initialize with some sample data
        appData = {
            employees: [
                { id: 1, name: 'John Doe', team: 'A', absentDates: [], skills: [1, 3] },
                { id: 2, name: 'Jane Smith', team: 'A', absentDates: [], skills: [1, 4] },
                { id: 3, name: 'Bob Johnson', team: 'B', absentDates: [], skills: [2] },
                { id: 4, name: 'Alice Brown', team: 'B', absentDates: [], skills: [2, 3] },
                { id: 5, name: 'Charlie Davis', team: 'C', absentDates: [], skills: [4] }
            ],
            skills: defaultSkills,
            schedule: {},
            weekVersions: {},
            lockedWeeks: {},
            currentWeek: 0,
            weekDates: generateNextFourWeeks()
        };
        
        // Generate initial schedule
        generateSchedule();
        
        // Initialize version tracking for each week
        Object.keys(appData.schedule).forEach(weekIndex => {
            appData.weekVersions[weekIndex] = {
                "v1": {
                    name: "Initial Version",
                    date: new Date().toISOString(),
                    isActive: true,
                    schedule: JSON.parse(JSON.stringify(appData.schedule[weekIndex]))
                }
            };
            appData.lockedWeeks[weekIndex] = false;
        });
        
        saveData();
    }
}

// Save data to localStorage
function saveData() {
    localStorage.setItem('shiftPlannerData', JSON.stringify(appData));
}

// Show a temporary message to indicate operation in progress
function showProgressIndicator(message, duration = 3000) {
    // Check if there's already a progress indicator
    let indicator = document.getElementById('progress-indicator');
    
    if (!indicator) {
        // Create a new progress indicator
        indicator = document.createElement('div');
        indicator.id = 'progress-indicator';
        indicator.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background-color: rgba(0, 0, 0, 0.7);
            color: white;
            padding: 10px 15px;
            border-radius: 5px;
            z-index: 9999;
            transition: opacity 0.3s;
        `;
        document.body.appendChild(indicator);
    }
    
    // Update the message
    indicator.textContent = message;
    indicator.style.opacity = '1';
    
    // Remove after the specified duration
    setTimeout(() => {
        indicator.style.opacity = '0';
        setTimeout(() => {
            if (indicator.parentNode) {
                indicator.parentNode.removeChild(indicator);
            }
        }, 300);
    }, duration);
}

// Export data to a JSON file
function exportData() {
    // Show progress indicator
    showProgressIndicator('Preparing data for export...');
    
    // Create a JSON blob with the current data
    const dataStr = JSON.stringify(appData, null, 2);
    const dataBlob = new Blob([dataStr], { type: 'application/json' });
    
    // Create a download link and trigger the download
    const today = new Date();
    const fileName = `ShiftPlanner_Backup_${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}.json`;
    
    const downloadLink = document.createElement('a');
    downloadLink.href = URL.createObjectURL(dataBlob);
    downloadLink.download = fileName;
    
    // Append to the body, click, and remove
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
    
    // Update progress indicator
    showProgressIndicator('Data exported successfully!');
    console.log('Data exported successfully');
}

// Import data from a JSON file
function importData(file) {
    // Show progress indicator
    showProgressIndicator('Reading file...');
    
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(event) {
            try {
                // Update progress
                showProgressIndicator('Validating data...');
                
                // Parse the JSON data
                const importedData = JSON.parse(event.target.result);
                
                // Validate the imported data
                if (!validateImportedData(importedData)) {
                    reject(new Error('Invalid data format. The file does not contain valid Shift Planner data.'));
                    return;
                }
                
                // Confirm before overwriting current data
                if (confirm('This will replace your current data. Are you sure you want to continue?')) {
                    // Show progress for applying data
                    showProgressIndicator('Applying imported data...');
                    
                    // Update the application data
                    appData = importedData;
                    
                    // Save to localStorage
                    saveData();
                    
                    // Reload the application views
                    generateWeekOptions();
                    renderSchedule();
                    renderEmployees();
                    renderTeams();
                    renderSkills();
                    
                    // Success message
                    showProgressIndicator('Data imported successfully!');
                    resolve('Data imported successfully');
                } else {
                    showProgressIndicator('Import cancelled');
                    reject(new Error('Import cancelled by user'));
                }
            } catch (error) {
                reject(new Error('Error parsing the file: ' + error.message));
            }
        };
        
        reader.onerror = function() {
            reject(new Error('Error reading the file'));
        };
        
        reader.readAsText(file);
    });
}

// Validate imported data to ensure it has the expected structure
function validateImportedData(data) {
    // Check for required top-level properties
    const requiredProps = ['employees', 'skills', 'schedule', 'weekVersions', 'lockedWeeks', 'weekDates'];
    
    for (const prop of requiredProps) {
        if (!data.hasOwnProperty(prop)) {
            return false;
        }
    }
    
    // Validate employees array
    if (!Array.isArray(data.employees)) {
        return false;
    }
    
    // Validate skills array
    if (!Array.isArray(data.skills)) {
        return false;
    }
    
    // Basic validation passed
    return true;
}

// Handle file input for importing data
function handleFileImport(event) {
    const file = event.target.files[0];
    if (!file) {
        return;
    }
    
    // Check if it's a JSON file
    if (file.type !== 'application/json' && !file.name.endsWith('.json')) {
        alert('Please select a JSON file');
        return;
    }
    
    importData(file)
        .then(message => {
            alert(message);
        })
        .catch(error => {
            alert('Import error: ' + error.message);
        });
    
    // Reset the file input
    event.target.value = '';
}

// Open the file import dialog
function openImportDialog() {
    // Create a file input element
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.json,application/json';
    fileInput.style.display = 'none';
    
    // Add event listener for when a file is selected
    fileInput.addEventListener('change', handleFileImport);
    
    // Append to body, trigger click, and then remove
    document.body.appendChild(fileInput);
    fileInput.click();
    
    // Clean up the input after selection
    setTimeout(() => {
        document.body.removeChild(fileInput);
    }, 5000);
}

// Set up event listeners
function setupEventListeners() {
    // Navigation
    elements.navSchedule.addEventListener('click', (e) => {
        e.preventDefault();
        switchView('schedule');
    });
    
    elements.navEmployees.addEventListener('click', (e) => {
        e.preventDefault();
        switchView('employees');
    });
    
    elements.navTeams.addEventListener('click', (e) => {
        e.preventDefault();
        switchView('teams');
    });
    
    elements.navSkills.addEventListener('click', (e) => {
        e.preventDefault();
        switchView('skills');
    });
    
    // Language switcher
    document.getElementById('lang-en').addEventListener('click', (e) => {
        e.preventDefault();
        switchLanguage('en');
    });
    
    document.getElementById('lang-de').addEventListener('click', (e) => {
        e.preventDefault();
        switchLanguage('de');
    });
    
    // Schedule
    elements.weekSelector.addEventListener('change', handleWeekChange);
    elements.planButton.addEventListener('click', handlePlanGeneration);
    elements.exportButton.addEventListener('click', exportToExcel);
    
    // Data Management
    if (elements.exportDataButton) {
        elements.exportDataButton.addEventListener('click', exportData);
    }
    
    if (elements.importDataButton) {
        elements.importDataButton.addEventListener('click', openImportDialog);
    }
    
    if (elements.clearDataButton) {
        elements.clearDataButton.addEventListener('click', clearAllData);
    }
    
    if (elements.loadDemoButton) {
        elements.loadDemoButton.addEventListener('click', loadDemoData);
    }
    
    // Employees
    elements.addEmployeeButton.addEventListener('click', () => openEmployeeModal());
    elements.addDateBtn.addEventListener('click', addAbsentDateInput);
    elements.saveEmployeeBtn.addEventListener('click', saveEmployee);
    
    // Skills
    if (elements.addSkillButton) {
        elements.addSkillButton.addEventListener('click', () => openSkillModal());
    }
    if (elements.saveSkillBtn) {
        elements.saveSkillBtn.addEventListener('click', saveSkill);
    }
    
    // Shift edit
    elements.saveShiftBtn.addEventListener('click', saveShiftAssignment);
    
    // Table click handlers will be set up in the render functions
}

// Set up navigation
function setupNavigation() {
    const navLinks = [
        elements.navSchedule, 
        elements.navEmployees, 
        elements.navTeams,
        elements.navSkills
    ];
    
    const views = [
        elements.scheduleView, 
        elements.employeesView, 
        elements.teamsView,
        elements.skillsView
    ];
    
    // Initialize with schedule view active
    switchView('schedule');
}

// Switch between views
function switchView(viewName) {
    // Reset active class on nav links
    elements.navSchedule.classList.remove('active');
    elements.navEmployees.classList.remove('active');
    elements.navTeams.classList.remove('active');
    if (elements.navSkills) elements.navSkills.classList.remove('active');
    
    // Hide all views
    elements.scheduleView.classList.add('d-none');
    elements.employeesView.classList.add('d-none');
    elements.teamsView.classList.add('d-none');
    if (elements.skillsView) elements.skillsView.classList.add('d-none');
    
    // Show selected view and set active class
    switch(viewName) {
        case 'schedule':
            elements.scheduleView.classList.remove('d-none');
            elements.navSchedule.classList.add('active');
            renderSchedule();
            break;
        case 'employees':
            elements.employeesView.classList.remove('d-none');
            elements.navEmployees.classList.add('active');
            renderEmployees();
            break;
        case 'teams':
            elements.teamsView.classList.remove('d-none');
            elements.navTeams.classList.add('active');
            renderTeams();
            break;
        case 'skills':
            if (elements.skillsView) {
                elements.skillsView.classList.remove('d-none');
                elements.navSkills.classList.add('active');
                renderSkills();
            }
            break;
    }
}

// Generate options for the week selector
function generateWeekOptions() {
    elements.weekSelector.innerHTML = '';
    
    appData.weekDates.forEach((week, index) => {
        const option = document.createElement('option');
        option.value = index;
        
        const startDate = new Date(week[0]);
        const endDate = new Date(week[week.length - 1]);
        
        const startFormatted = formatDate(startDate);
        const endFormatted = formatDate(endDate);
        
        // Use translation for "Week" label
        option.textContent = `${t('schedule.weekLabel')} ${index + 1}: ${startFormatted} - ${endFormatted}`;
        
        // Add lock icon if week is locked
        if (isWeekLocked(index)) {
            option.textContent += ' 🔒';
        }
        
        // Add active version name
        const activeVersion = getActiveVersion(index);
        if (activeVersion) {
            option.textContent += ` (${activeVersion.name})`;
        }
        
        if (index === appData.currentWeek) {
            option.selected = true;
        }
        
        elements.weekSelector.appendChild(option);
    });
    
    // Update the week label after generating options
    updateWeekLabel(appData.currentWeek);
}

// Format date based on current language (DD/MM/YYYY for English, DD.MM.YYYY for German)
function formatDate(date) {
    const options = { year: 'numeric', month: '2-digit', day: '2-digit' };
    return date.toLocaleDateString(currentLanguage === 'de' ? 'de-DE' : 'en-US', options);
}

// Format date as YYYY-MM-DD (ISO format, used for input[type=date] values)
function formatDateISO(date) {
    return date.toISOString().split('T')[0];
}

// Handle week change
function handleWeekChange() {
    appData.currentWeek = parseInt(elements.weekSelector.value);
    renderSchedule();
    saveData();
}

// Generate the next four weeks starting from the current date
function generateNextFourWeeks() {
    const weeks = [];
    const today = new Date();
    
    // Get the current day of the week (0 = Sunday, 1 = Monday, ...)
    const currentDayOfWeek = today.getDay();
    
    // Calculate the date of the most recent Monday
    const mostRecentMonday = new Date(today);
    mostRecentMonday.setDate(today.getDate() - (currentDayOfWeek === 0 ? 6 : currentDayOfWeek - 1));
    
    // Generate 4 weeks
    for (let week = 0; week < 4; week++) {
        const weekDates = [];
        
        // For each day of the week
        for (let day = 0; day < 7; day++) {
            const date = new Date(mostRecentMonday);
            date.setDate(mostRecentMonday.getDate() + (week * 7) + day);
            weekDates.push(formatDateISO(date));
        }
        
        weeks.push(weekDates);
    }
    
    return weeks;
}

// Generate the schedule for all weeks
function generateSchedule() {
    appData.schedule = {};
    
    appData.weekDates.forEach((week, weekIndex) => {
        appData.schedule[weekIndex] = {};
        
        // Determine which team is on AM and PM shift for this week
        const teamOnAM = weekIndex % 2 === 0 ? 'A' : 'B';
        const teamOnPM = teamOnAM === 'A' ? 'B' : 'A';
        
        week.forEach((dateStr, dayIndex) => {
            const daySchedule = {
                date: dateStr,
                AM: {
                    team: teamOnAM,
                    employees: getEmployeesForShift(teamOnAM, dateStr)
                },
                PM: {
                    team: teamOnPM,
                    employees: getEmployeesForShift(teamOnPM, dateStr)
                },
                Night: {
                    team: 'C',
                    employees: getEmployeesForShift('C', dateStr)
                }
            };
            
            appData.schedule[weekIndex][DAYS_OF_WEEK[dayIndex]] = daySchedule;
        });
        
        // Initialize version tracking for this week
        if (!appData.weekVersions[weekIndex]) {
            appData.weekVersions[weekIndex] = {
                "v1": {
                    name: "Initial Version",
                    date: new Date().toISOString(),
                    isActive: true,
                    schedule: JSON.parse(JSON.stringify(appData.schedule[weekIndex]))
                }
            };
            appData.lockedWeeks[weekIndex] = false;
        }
    });
}

// Get all employees for a shift, including those who are absent
function getEmployeesForShift(team, dateStr) {
    return appData.employees
        .filter(employee => employee.team === team)
        .map(employee => employee.id);
}

// Get available employees for a shift, excluding those who are absent
function getAvailableEmployeesForShift(team, dateStr) {
    return appData.employees
        .filter(employee => employee.team === team && !isEmployeeAbsent(employee, dateStr))
        .map(employee => employee.id);
}

// Check if an employee is absent on a specific date
function isEmployeeAbsent(employee, dateStr) {
    return employee.absentDates.includes(dateStr);
}

// Handle plan generation
function handlePlanGeneration() {
    // Open a modal to select date range
    openPlanGenerationModal();
}

// Open modal for plan generation with date range selection
function openPlanGenerationModal() {
    // Remove any existing modal to avoid duplication
    let existingModal = document.getElementById('plan-modal');
    if (existingModal) {
        existingModal.remove();
    }
    
    // Create a new modal
    let planModal = document.createElement('div');
    planModal.id = 'plan-modal';
    planModal.className = 'modal fade';
    planModal.tabIndex = '-1';
    
    // Get today's date and format it for the input fields
    const today = new Date();
    const oneMonthLater = new Date(today);
    oneMonthLater.setMonth(today.getMonth() + 1);
    
    const startDateValue = formatDateISO(today);
    const endDateValue = formatDateISO(oneMonthLater);
    
    planModal.innerHTML = `
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title">${t('modals.planGeneration.title')}</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <form id="plan-form">
                        <div class="mb-3">
                            <label for="start-date" class="form-label">${t('modals.planGeneration.startDate')}</label>
                            <input type="date" class="form-control" id="start-date" value="${startDateValue}" required>
                        </div>
                        <div class="mb-3">
                            <label for="end-date" class="form-label">${t('modals.planGeneration.endDate')}</label>
                            <input type="date" class="form-control" id="end-date" value="${endDateValue}" required>
                        </div>
                        <div class="mb-3 form-check">
                            <input type="checkbox" class="form-check-input" id="create-versions" checked>
                            <label class="form-check-label" for="create-versions">${t('modals.planGeneration.createVersions')}</label>
                        </div>
                        <div class="mb-3">
                            <p class="text-info">
                                <i class="bi bi-info-circle"></i> 
                                ${currentLanguage === 'de' ? 'Gesperrte Wochen werden nicht verändert.' : 'Locked weeks will not be modified.'}
                            </p>
                        </div>
                    </form>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">${t('modals.buttons.cancel')}</button>
                    <button type="button" class="btn btn-primary" id="generate-plan-btn">${t('modals.buttons.generate')}</button>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(planModal);
    
    // Create and show the modal
    const modal = new bootstrap.Modal(planModal);
    modal.show();
    
    // Add event listener to the generate button
    document.getElementById('generate-plan-btn').addEventListener('click', () => {
        const startDate = new Date(document.getElementById('start-date').value);
        const endDate = new Date(document.getElementById('end-date').value);
        const createVersions = document.getElementById('create-versions').checked;
        
        // Validate dates
        if (endDate < startDate) {
            alert(currentLanguage === 'de' ? 
                  'Das Enddatum kann nicht vor dem Startdatum liegen' : 
                  'End date cannot be before start date');
            return;
        }
        
        // Generate plan for the selected date range
        generatePlanForDateRange(startDate, endDate, createVersions);
        
        // Hide the modal
        modal.hide();
    });
    
    // Clean up the modal when hidden
    planModal.addEventListener('hidden.bs.modal', () => {
        planModal.remove();
    });
}

// Generate weeks for a specific date range
function generateWeeksForDateRange(startDate, endDate) {
    const weeks = [];
    
    // Adjust startDate to the nearest Monday (first day of week)
    const adjustedStartDate = new Date(startDate);
    const startDayOfWeek = adjustedStartDate.getDay(); // 0 = Sunday, 1 = Monday, ...
    
    // Move to Monday of the same week
    adjustedStartDate.setDate(adjustedStartDate.getDate() - (startDayOfWeek === 0 ? 6 : startDayOfWeek - 1));
    
    // Generate weeks until we cover the end date
    let currentDate = new Date(adjustedStartDate);
    
    while (currentDate <= endDate) {
        const weekDates = [];
        
        // For each day of the week
        for (let day = 0; day < 7; day++) {
            const date = new Date(currentDate);
            date.setDate(currentDate.getDate() + day);
            weekDates.push(formatDateISO(date));
        }
        
        weeks.push(weekDates);
        
        // Move to next Monday
        currentDate.setDate(currentDate.getDate() + 7);
    }
    
    return weeks;
}

// Generate plan for a specific date range
function generatePlanForDateRange(startDate, endDate, createVersions) {
    // Generate weeks for the date range
    const newWeeks = generateWeeksForDateRange(startDate, endDate);
    
    if (newWeeks.length === 0) {
        alert('No valid weeks found in the selected date range');
        return;
    }
    
    // Check if any of these weeks overlap with existing weeks
    const existingWeekIndices = [];
    const weekVersionNames = [];
    
    // For each new week, check if it overlaps with an existing week
    for (let i = 0; i < newWeeks.length; i++) {
        const newWeekStart = new Date(newWeeks[i][0]);
        const newWeekEnd = new Date(newWeeks[i][6]);
        
        for (let j = 0; j < appData.weekDates.length; j++) {
            if (!appData.weekDates[j] || appData.weekDates[j].length === 0) continue;
            
            const existingWeekStart = new Date(appData.weekDates[j][0]);
            const existingWeekEnd = new Date(appData.weekDates[j][6]);
            
            // Check if weeks overlap
            if (
                (newWeekStart >= existingWeekStart && newWeekStart <= existingWeekEnd) ||
                (newWeekEnd >= existingWeekStart && newWeekEnd <= existingWeekEnd) ||
                (newWeekStart <= existingWeekStart && newWeekEnd >= existingWeekEnd)
            ) {
                // Check if the week is locked
                if (isWeekLocked(j)) {
                    console.log(`Week ${j + 1} (${formatDate(existingWeekStart)} - ${formatDate(existingWeekEnd)}) is locked and will not be modified`);
                    continue;
                }
                
                existingWeekIndices.push(j);
                weekVersionNames.push(`Replan ${formatDate(newWeekStart)}`);
                break;
            }
        }
    }
    
    // Store the previous weekDates and schedule
    const previousWeekDates = JSON.parse(JSON.stringify(appData.weekDates));
    const previousSchedule = JSON.parse(JSON.stringify(appData.schedule));
    
    // Update appData.weekDates with the new weeks
    appData.weekDates = newWeeks;
    
    // Generate new schedule
    generateSchedule();
    
    // If we should create versions for existing weeks, create them
    if (createVersions && existingWeekIndices.length > 0) {
        for (let i = 0; i < existingWeekIndices.length; i++) {
            const weekIndex = existingWeekIndices[i];
            
            // Create a new version for this week
            createVersionForWeek(weekIndex, weekVersionNames[i] || `Replan ${new Date().toLocaleDateString()}`);
        }
    }
    
    // Restore previous week dates and schedule for weeks that are locked
    for (let i = 0; i < previousWeekDates.length; i++) {
        if (isWeekLocked(i)) {
            appData.weekDates[i] = previousWeekDates[i];
            appData.schedule[i] = previousSchedule[i];
        }
    }
    
    // Update the UI
    generateWeekOptions();
    renderSchedule();
    saveData();
    
    // Show localized success message
    let successMessage = '';
    if (currentLanguage === 'de') {
        successMessage = `Plan erfolgreich erstellt für ${formatDate(startDate)} - ${formatDate(endDate)}.\n${existingWeekIndices.length} existierende entsperrte Wochen wurden neu geplant.`;
    } else {
        successMessage = `Plan generated successfully for ${formatDate(startDate)} - ${formatDate(endDate)}.\n${existingWeekIndices.length} existing unlocked weeks were replanned.`;
    }
    alert(successMessage);
}

// Create a new version for a specific week
function createVersionForWeek(weekIndex, versionName) {
    // Check if week is locked
    if (isWeekLocked(weekIndex)) {
        console.log(`Cannot create a new version because week ${weekIndex + 1} is locked.`);
        return false;
    }
    
    // Get all existing versions for this week
    const versions = appData.weekVersions[weekIndex] || {};
    
    // Find the highest version number
    const versionNumbers = Object.keys(versions)
        .filter(key => key.startsWith('v'))
        .map(key => parseInt(key.substring(1)))
        .sort((a, b) => b - a);
    
    const newVersionNumber = (versionNumbers[0] || 0) + 1;
    const newVersionKey = `v${newVersionNumber}`;
    
    // Deactivate all other versions
    Object.keys(versions).forEach(key => {
        versions[key].isActive = false;
    });
    
    // Create new version
    versions[newVersionKey] = {
        name: versionName || `Version ${newVersionNumber}`,
        date: new Date().toISOString(),
        isActive: true,
        schedule: JSON.parse(JSON.stringify(appData.schedule[weekIndex]))
    };
    
    appData.weekVersions[weekIndex] = versions;
    
    return true;
}

// Create a new version of the current week's schedule
function createNewVersion(versionName) {
    return createVersionForWeek(appData.currentWeek, versionName);
}

// Check if a week is locked
function isWeekLocked(weekIndex) {
    return appData.lockedWeeks[weekIndex] === true;
}

// Lock or unlock a week
function toggleWeekLock(weekIndex, lock) {
    appData.lockedWeeks[weekIndex] = lock;
    saveData();
    renderSchedule(); // Refresh the UI to show lock status
    updateWeekSelector(); // Update week selector to show lock status
}

// Switch to a different version
function switchVersion(weekIndex, versionKey) {
    // Check if week is locked
    if (isWeekLocked(weekIndex)) {
        alert('Cannot switch versions because this week is locked. Please unlock it first.');
        return false;
    }
    
    const versions = appData.weekVersions[weekIndex];
    if (!versions || !versions[versionKey]) {
        alert('Version not found!');
        return false;
    }
    
    // Deactivate all versions
    Object.keys(versions).forEach(key => {
        versions[key].isActive = (key === versionKey);
    });
    
    // Load the selected version's schedule
    appData.schedule[weekIndex] = JSON.parse(JSON.stringify(versions[versionKey].schedule));
    
    saveData();
    renderSchedule();
    return true;
}

// Get the current active version for a week
function getActiveVersion(weekIndex) {
    const versions = appData.weekVersions[weekIndex];
    if (!versions) return null;
    
    for (const [key, version] of Object.entries(versions)) {
        if (version.isActive) {
            return { key, ...version };
        }
    }
    
    return null;
}

// Get all versions for a week
function getAllVersions(weekIndex) {
    return appData.weekVersions[weekIndex] || {};
}

// Update week selector to show lock status
function updateWeekSelector() {
    const options = elements.weekSelector.querySelectorAll('option');
    
    options.forEach(option => {
        const weekIndex = parseInt(option.value);
        const isLocked = isWeekLocked(weekIndex);
        
        // Update option text to show lock status
        const originalText = option.textContent.replace(' 🔒', '');
        option.textContent = isLocked ? `${originalText} 🔒` : originalText;
        
        // Get active version
        const activeVersion = getActiveVersion(weekIndex);
        if (activeVersion) {
            option.textContent += ` (${activeVersion.name})`;
        }
    });
}

// Render the schedule table
function renderSchedule() {
    const weekIndex = appData.currentWeek;
    const weekSchedule = appData.schedule[weekIndex] || {};
    const isLocked = isWeekLocked(weekIndex);
    const activeVersion = getActiveVersion(weekIndex);
    
    // Update the current week number display
    updateWeekLabel(weekIndex);
    
    // Update the lock button state
    updateLockButton(isLocked);
    
    // Update version controls
    updateVersionControls(weekIndex, activeVersion);
    
    const tbody = elements.scheduleTable.querySelector('tbody');
    tbody.innerHTML = '';
    
    // Show status banner indicating lock and version status
    updateStatusBanner(weekIndex, isLocked, activeVersion);
    
    // Generate skill summary for the entire week
    generateWeeklySkillSummary(weekIndex);
    
    DAYS_OF_WEEK.forEach(day => {
        const daySchedule = weekSchedule[day] || {};
        
        // Create the employees row
        const employeeRow = document.createElement('tr');
        employeeRow.classList.add('employee-skills-row');
        
        // Day column
        const dayCell = document.createElement('td');
        dayCell.textContent = getLocalizedDay(day);
        if (daySchedule.date) {
            const date = new Date(daySchedule.date);
            dayCell.textContent += ` (${formatDate(date)})`;
        }
        employeeRow.appendChild(dayCell);
        
        // Get date for this day
        const dayIndex = DAYS_OF_WEEK.indexOf(day);
        const dateStr = appData.weekDates[weekIndex][dayIndex];
        
        // Create cells for each shift type (employees)
        SHIFT_TYPES.forEach(shiftType => {
            const shiftCell = document.createElement('td');
            shiftCell.classList.add('schedule-cell');
            
            // Set data attributes for editing
            shiftCell.dataset.day = day;
            shiftCell.dataset.shift = shiftType;
            
            const shift = daySchedule[shiftType] || { team: '', employees: [] };
            
            // Add team badge
            const teamBadge = document.createElement('span');
            teamBadge.classList.add('team-badge', `team-${shift.team.toLowerCase()}-badge`);
            teamBadge.textContent = `Team ${shift.team}`;
            shiftCell.appendChild(teamBadge);
            
            if (shift.employees && shift.employees.length > 0) {
                // Check if there are employees from other teams
                const nonTeamEmployees = shift.employees.filter(employeeId => {
                    const employee = appData.employees.find(e => e.id === employeeId);
                    return employee && employee.team !== shift.team;
                });
                
                // Add a badge if mixed teams
                if (nonTeamEmployees.length > 0) {
                    const mixedTeamBadge = document.createElement('span');
                    mixedTeamBadge.classList.add('badge', 'bg-warning', 'text-dark', 'ms-2');
                    mixedTeamBadge.title = 'This shift includes employees from multiple teams';
                    mixedTeamBadge.textContent = 'Mixed Teams';
                    teamBadge.after(mixedTeamBadge);
                }
                
                // Group employees by team
                const employeesByTeam = {};
                const absentEmployees = [];
                
                // Process all employees
                shift.employees.forEach(employeeId => {
                    const employee = appData.employees.find(e => e.id === employeeId);
                    if (!employee) return;
                    
                        const isAbsent = employee.absentDates.includes(dateStr);
                        
                    // Handle absent employees separately
                        if (isAbsent) {
                        absentEmployees.push(employee);
                        return;
                    }
                    
                    // Group by team
                    if (!employeesByTeam[employee.team]) {
                        employeesByTeam[employee.team] = [];
                    }
                    employeesByTeam[employee.team].push(employee);
                });
                
                // Create container for all employee data
                const employeesList = document.createElement('div');
                employeesList.classList.add('mt-2');
                
                // Add a line for each team's employees
                Object.keys(employeesByTeam).sort().forEach(teamId => {
                    const teamEmployees = employeesByTeam[teamId];
                    const employeeNames = teamEmployees.map(emp => emp.name);
                    
                    const teamLine = document.createElement('div');
                    if (teamId !== shift.team) {
                        teamLine.classList.add('text-primary');
                        teamLine.innerHTML = `${employeeNames.join(', ')} <small class="text-muted">(Team ${teamId})</small>`;
                    } else {
                        teamLine.textContent = employeeNames.join(', ');
                    }
                    
                    employeesList.appendChild(teamLine);
                });
                
                // Add absent employees (each on their own line)
                if (absentEmployees.length > 0) {
                    const absentLine = document.createElement('div');
                    absentLine.classList.add('text-danger', 'mt-1');
                    
                    const absentNames = absentEmployees.map(emp => ` ${emp.name} `);
                    absentLine.innerHTML = `${absentNames.join(', ')} <span class="badge bg-danger">Absent</span>`;
                    
                    employeesList.appendChild(absentLine);
                }
                
                shiftCell.appendChild(employeesList);
            } else {
                const emptyMessage = document.createElement('div');
                emptyMessage.classList.add('mt-2');
                emptyMessage.textContent = currentLanguage === 'de' ? 'Keine Mitarbeiter zugewiesen' : 'No employees assigned';
                shiftCell.appendChild(emptyMessage);
            }
            
            // Add click event to edit shift only if week is not locked
            if (!isLocked) {
                shiftCell.addEventListener('click', () => openShiftEditModal(day, shiftType, shift));
                shiftCell.style.cursor = 'pointer';
            } else {
                shiftCell.style.cursor = 'not-allowed';
                shiftCell.title = currentLanguage === 'de' 
                    ? 'Diese Woche ist gesperrt. Entsperren Sie sie, um Änderungen vorzunehmen.' 
                    : 'This week is locked. Unlock it to make changes.';
                shiftCell.classList.add('disabled-cell');
            }
            
            employeeRow.appendChild(shiftCell);
        });
        
        tbody.appendChild(employeeRow);
        
        // Create a separate row for skills
        const skillsRow = document.createElement('tr');
        skillsRow.classList.add('skills-row');
        skillsRow.style.borderTop = "none";
        
        // Empty cell for day column in skills row
        const emptyCellForSkills = document.createElement('td');
        emptyCellForSkills.style.borderTop = "none";
        emptyCellForSkills.style.paddingTop = "0";
        
        // Add daily skill summary to the first column
        const daySkillSummary = generateDailySkillSummary(weekIndex, day);
        if (daySkillSummary.length > 0) {
            const skillsDiv = document.createElement('div');
            skillsDiv.classList.add('shift-skills-summary');
            skillsDiv.style.position = "static";
            skillsDiv.style.width = "100%";
            skillsDiv.style.margin = "0";
            skillsDiv.style.borderTop = "none";
            
            daySkillSummary.forEach(skill => {
                const badge = document.createElement('span');
                badge.classList.add('badge');
                badge.textContent = skill.name;
                badge.title = currentLanguage === 'de' 
                    ? `${skill.count} Mitarbeiter mit dieser Fähigkeit verfügbar` 
                    : `${skill.count} employees with this skill available`;
                skillsDiv.appendChild(badge);
            });
            
            emptyCellForSkills.appendChild(skillsDiv);
        }
        
        skillsRow.appendChild(emptyCellForSkills);
        
        // Add skill cells for each shift type
        SHIFT_TYPES.forEach(shiftType => {
            const shift = daySchedule[shiftType] || { employees: [] };
            
            const skillsCell = document.createElement('td');
            skillsCell.style.borderTop = "none";
            skillsCell.style.paddingTop = "0";
            
            // Get the skills for this shift
            const shiftSkillIds = getShiftSkills(shift.employees, dateStr);
            
            if (shiftSkillIds.length > 0) {
                const skillsDiv = document.createElement('div');
                skillsDiv.classList.add('shift-skills-summary');
                skillsDiv.style.position = "static";
                skillsDiv.style.width = "100%";
                skillsDiv.style.margin = "0";
                skillsDiv.style.borderTop = "none";
                
                shiftSkillIds.forEach(skillId => {
                    const skill = appData.skills.find(s => s.id === skillId);
                    if (skill) {
                        const badge = document.createElement('span');
                        badge.classList.add('badge');
                        badge.textContent = skill.name;
                        badge.title = skill.description;
                        skillsDiv.appendChild(badge);
                    }
                });
                
                skillsCell.appendChild(skillsDiv);
            }
            
            skillsRow.appendChild(skillsCell);
        });
        
        tbody.appendChild(skillsRow);
    });
}

// Get skills for a shift
function getShiftSkills(employeeIds, dateStr) {
    const skillsSet = new Set();
    
    // For each employee in the shift
    employeeIds.forEach(employeeId => {
        const employee = appData.employees.find(e => e.id === employeeId);
        
        // Skip absent employees
        if (employee && !employee.absentDates.includes(dateStr) && employee.skills) {
            // Add each skill to the set
            employee.skills.forEach(skillId => {
                skillsSet.add(skillId);
            });
        }
    });
    
    return Array.from(skillsSet);
}

// Generate daily skill summary
function generateDailySkillSummary(weekIndex, day) {
    const daySchedule = appData.schedule[weekIndex][day] || {};
    const dayIndex = DAYS_OF_WEEK.indexOf(day);
    const dateStr = appData.weekDates[weekIndex][dayIndex];
    
    // Count skills for each shift
    const skillCounts = {};
    
    SHIFT_TYPES.forEach(shiftType => {
        const shift = daySchedule[shiftType] || { employees: [] };
        
        // For each employee in the shift
        shift.employees.forEach(employeeId => {
            const employee = appData.employees.find(e => e.id === employeeId);
            
            // Skip absent employees
            if (employee && !employee.absentDates.includes(dateStr) && employee.skills) {
                // Count each skill
                employee.skills.forEach(skillId => {
                    skillCounts[skillId] = (skillCounts[skillId] || 0) + 1;
                });
            }
        });
    });
    
    // Create summary array with skill details
    const summary = Object.entries(skillCounts).map(([skillId, count]) => {
        const skill = appData.skills.find(s => s.id === parseInt(skillId));
        return {
            id: parseInt(skillId),
            name: skill ? skill.name : 'Unknown',
            count
        };
    });
    
    // Sort by count, descending
    return summary.sort((a, b) => b.count - a.count);
}

// Generate weekly skill summary
function generateWeeklySkillSummary(weekIndex) {
    // Create or get the weekly skill summary container
    let summaryContainer = document.getElementById('weekly-skill-summary');
    if (!summaryContainer) {
        summaryContainer = document.createElement('div');
        summaryContainer.id = 'weekly-skill-summary';
        summaryContainer.classList.add('card', 'mt-3', 'mb-4');
        
        // Insert after the schedule table instead of before it
        const scheduleTable = elements.scheduleTable;
        scheduleTable.parentElement.appendChild(summaryContainer);
    }
    
    // Counter for skills by shift type across the week
    const skillsByShift = {
        AM: {},
        PM: {},
        Night: {}
    };
    
    // For each day in the schedule
    DAYS_OF_WEEK.forEach(day => {
        const daySchedule = appData.schedule[weekIndex][day] || {};
        const dayIndex = DAYS_OF_WEEK.indexOf(day);
        const dateStr = appData.weekDates[weekIndex][dayIndex];
        
        // For each shift type
        SHIFT_TYPES.forEach(shiftType => {
            const shift = daySchedule[shiftType] || { employees: [] };
            
            // For each employee in the shift
            shift.employees.forEach(employeeId => {
                const employee = appData.employees.find(e => e.id === employeeId);
                
                // Skip absent employees
                if (employee && !employee.absentDates.includes(dateStr) && employee.skills) {
                    // Count each skill for this shift type
                    employee.skills.forEach(skillId => {
                        skillsByShift[shiftType][skillId] = (skillsByShift[shiftType][skillId] || 0) + 1;
                    });
                }
            });
        });
    });
    
    // Create the HTML for the summary with updated styling for a more gentle look
    let html = `
        <div class="card-header">
            <h6 class="mb-0">${currentLanguage === 'de' ? 'Wöchentliche Fähigkeiten-Abdeckungsübersicht' : 'Weekly Skill Coverage Summary'}</h6>
        </div>
        <div class="card-body">
            <div class="row">
    `;
    
    // For each shift type, create a column
    SHIFT_TYPES.forEach(shiftType => {
        const shiftSkills = skillsByShift[shiftType];
        const skillSummary = Object.entries(shiftSkills)
            .map(([skillId, count]) => {
                const skill = appData.skills.find(s => s.id === parseInt(skillId));
                return {
                    id: parseInt(skillId),
                    name: skill ? skill.name : 'Unknown',
                    count
                };
            })
            .sort((a, b) => b.count - a.count);
        
        // Get shift type label based on language
        let shiftLabel;
        if (currentLanguage === 'de') {
            if (shiftType === 'AM') shiftLabel = 'Frühschicht';
            else if (shiftType === 'PM') shiftLabel = 'Spätschicht';
            else if (shiftType === 'Night') shiftLabel = 'Nachtschicht';
        } else {
            shiftLabel = `${shiftType} Shifts`;
        }
        
        html += `
            <div class="col-md-4">
                <h6>${shiftLabel}</h6>
                <div class="list-group mb-3">
        `;
        
        if (skillSummary.length > 0) {
            skillSummary.forEach(skill => {
                html += `
                    <div class="list-group-item d-flex justify-content-between align-items-center">
                        ${skill.name}
                        <span class="badge rounded-pill">${skill.count}</span>
                    </div>
                `;
            });
        } else {
            html += `<div class="list-group-item">${currentLanguage === 'de' ? 'Keine Fähigkeiten verfügbar' : 'No skills available'}</div>`;
        }
        
        html += `
                </div>
            </div>
        `;
    });
    
    html += `
            </div>
        </div>
    `;
    
    summaryContainer.innerHTML = html;
}

// Update the status banner
function updateStatusBanner(weekIndex, isLocked, activeVersion) {
    // Get or create status banner
    let statusBanner = document.getElementById('status-banner');
    if (!statusBanner) {
        statusBanner = document.createElement('div');
        statusBanner.id = 'status-banner';
        statusBanner.classList.add('small', 'mb-2', 'border-top', 'pt-1');
        
        // Insert after version controls instead of before schedule table
        const versionControls = document.getElementById('version-controls');
        if (versionControls) {
            versionControls.parentElement.insertBefore(statusBanner, versionControls.nextSibling);
        } else {
            // Fallback to original position if version controls aren't found
            elements.scheduleTable.parentElement.insertBefore(statusBanner, elements.scheduleTable);
        }
    }
    
    // Set appropriate status message and style with more subtle appearance
    if (isLocked) {
        statusBanner.className = 'small mb-2 text-warning border-top pt-1';
        statusBanner.innerHTML = `
            🔒 <span>${currentLanguage === 'de' ? `Woche ${weekIndex + 1} ist gesperrt` : `Week ${weekIndex + 1} is locked`}</span> 
            <span class="ms-1">(${currentLanguage === 'de' ? 'Aktuelle Version' : 'Current version'}: ${activeVersion ? activeVersion.name : 'Unknown'})</span>
        `;
    } else {
        statusBanner.className = 'small mb-2 text-info border-top pt-1';
        statusBanner.innerHTML = `
            ✏️ <span>${currentLanguage === 'de' ? `Woche ${weekIndex + 1} ist bearbeitbar` : `Week ${weekIndex + 1} is editable`}</span>
            <span class="ms-1">(${currentLanguage === 'de' ? 'Aktuelle Version' : 'Current version'}: ${activeVersion ? activeVersion.name : 'Unknown'})</span>
        `;
    }
}

// Update the lock button state
function updateLockButton(isLocked) {
    // Get or create lock button
    let lockButton = document.getElementById('lock-button');
    if (!lockButton) {
        lockButton = document.createElement('button');
        lockButton.id = 'lock-button';
        lockButton.classList.add('btn', 'me-2');
        
        const actionButtons = elements.planButton.parentElement;
        actionButtons.insertBefore(lockButton, elements.planButton);
        
        // Add event listener
        lockButton.addEventListener('click', () => {
            const weekIndex = appData.currentWeek;
            const currentlyLocked = isWeekLocked(weekIndex);
            
            // If currently locked, confirm unlock
            if (currentlyLocked) {
                const confirmMsg = currentLanguage === 'de'
                    ? 'Sind Sie sicher, dass Sie diese Woche entsperren möchten? Dies ermöglicht Änderungen am Zeitplan.'
                    : 'Are you sure you want to unlock this week? This will allow edits to the schedule.';
                
                if (confirm(confirmMsg)) {
                    toggleWeekLock(weekIndex, false);
                }
            } else {
                // If unlocked, confirm lock
                const confirmMsg = currentLanguage === 'de'
                    ? 'Sind Sie sicher, dass Sie diese Woche sperren möchten? Dies verhindert weitere Änderungen, bis die Woche entsperrt wird.'
                    : 'Are you sure you want to lock this week? This will prevent further changes until unlocked.';
                
                if (confirm(confirmMsg)) {
                    toggleWeekLock(weekIndex, true);
                }
            }
        });
    }
    
    // Update button appearance based on lock state
    if (isLocked) {
        lockButton.textContent = currentLanguage === 'de' ? '🔓 Woche entsperren' : '🔓 Unlock Week';
        lockButton.classList.remove('btn-danger');
        lockButton.classList.add('btn-success');
    } else {
        lockButton.textContent = currentLanguage === 'de' ? '🔒 Woche sperren' : '🔒 Lock Week';
        lockButton.classList.remove('btn-success');
        lockButton.classList.add('btn-danger');
    }
}

// Update version controls
function updateVersionControls(weekIndex, activeVersion) {
    // Get or create version controls container
    let versionControls = document.getElementById('version-controls');
    if (!versionControls) {
        versionControls = document.createElement('div');
        versionControls.id = 'version-controls';
        versionControls.classList.add('w-100', 'mb-3', 'mt-2');
        
        // Get the card-header element that contains the week selector
        const cardHeader = elements.weekSelector.closest('.card-header');
        
        // Insert the version controls after the card header (on its own line)
        cardHeader.parentElement.insertBefore(versionControls, cardHeader.nextSibling);
    }
    
    // Get all versions for the current week
    const versions = getAllVersions(weekIndex);
    const isLocked = isWeekLocked(weekIndex);
    
    // Generate version control HTML - removing the Switch Version button since we'll switch automatically
    versionControls.innerHTML = `
        <div class="d-flex align-items-center">
            <label class="me-2"><strong>${t('schedule.versions')}:</strong></label>
            <select id="version-selector" class="form-select me-2" ${isLocked ? 'disabled' : ''}>
                ${Object.entries(versions).map(([key, version]) => `
                    <option value="${key}" ${version.isActive ? 'selected' : ''}>
                        ${version.name} (${new Date(version.date).toLocaleDateString()})
                    </option>
                `).join('')}
            </select>
            <button id="new-version-btn" class="btn btn-outline-primary btn-sm" ${isLocked ? 'disabled' : ''}>
                ${t('schedule.saveAsNewVersion')}
            </button>
        </div>
    `;
    
    // Add event listeners if week is not locked
    if (!isLocked) {
        // New version button
        document.getElementById('new-version-btn').addEventListener('click', () => {
            const versionName = prompt('Enter a name for the new version:', `Version ${Object.keys(versions).length + 1}`);
            if (versionName) {
                if (createNewVersion(versionName)) {
                    alert(`New version "${versionName}" created successfully!`);
                    renderSchedule();
                }
            }
        });
        
        // Add change event listener to version selector for direct switching
        const versionSelector = document.getElementById('version-selector');
        versionSelector.addEventListener('change', () => {
            const selectedVersionKey = versionSelector.value;
            // Skip if the selected version is already active
            if (activeVersion && selectedVersionKey === activeVersion.key) {
                return;
            }
            
            // Switch to the selected version directly without confirmation
            directSwitchVersion(weekIndex, selectedVersionKey);
        });
    }
}

// Function to switch versions directly without confirmation
function directSwitchVersion(weekIndex, versionKey) {
    // Check if week is locked
    if (isWeekLocked(weekIndex)) {
        alert('Cannot switch versions because this week is locked. Please unlock it first.');
        return false;
    }
    
    const versions = appData.weekVersions[weekIndex];
    if (!versions || !versions[versionKey]) {
        alert('Version not found!');
        return false;
    }
    
    // Deactivate all versions
    Object.keys(versions).forEach(key => {
        versions[key].isActive = (key === versionKey);
    });
    
    // Load the selected version's schedule
    appData.schedule[weekIndex] = JSON.parse(JSON.stringify(versions[versionKey].schedule));
    
    saveData();
    renderSchedule();
    
    // Show a non-blocking notification to inform the user that the version has been switched
    console.log(`Switched to version "${versions[versionKey].name}"`);
    
    return true;
}

// Render employees table
function renderEmployees() {
    const employeesTable = document.getElementById('employees-table').getElementsByTagName('tbody')[0];
    employeesTable.innerHTML = '';
    
    appData.employees.sort((a, b) => a.name.localeCompare(b.name)).forEach(employee => {
        const row = document.createElement('tr');
        
        // Name cell
        const nameCell = document.createElement('td');
        nameCell.textContent = employee.name;
        row.appendChild(nameCell);
        
        // Team cell
        const teamCell = document.createElement('td');
        const teamBadge = document.createElement('span');
        teamBadge.classList.add('team-badge', `team-${employee.team.toLowerCase()}-badge`);
        teamBadge.textContent = t(`teams.team${employee.team}`);
        teamCell.appendChild(teamBadge);
        row.appendChild(teamCell);
        
        // Skills cell
        const skillsCell = document.createElement('td');
        if (employee.skills && employee.skills.length > 0) {
            const skillsList = document.createElement('div');
            skillsList.classList.add('skills-summary');
            
            employee.skills.forEach(skillId => {
                const skill = appData.skills.find(s => s.id === skillId);
                if (skill) {
                    const badge = document.createElement('span');
                    badge.classList.add('badge');
                    badge.textContent = skill.name;
                    skillsList.appendChild(badge);
                }
            });
            
            skillsCell.appendChild(skillsList);
        } else {
            skillsCell.textContent = '-';
        }
        row.appendChild(skillsCell);
        
        // Absent dates cell
        const absentCell = document.createElement('td');
        if (employee.absentDates && employee.absentDates.length > 0) {
            employee.absentDates.forEach(date => {
                const datePill = document.createElement('span');
                datePill.classList.add('absent-date-pill');
                datePill.textContent = formatDate(new Date(date));
                absentCell.appendChild(datePill);
            });
        } else {
            absentCell.textContent = '-';
        }
        row.appendChild(absentCell);
        
        // Actions cell
        const actionsCell = document.createElement('td');
        
        // Edit button
        const editBtn = document.createElement('button');
        editBtn.classList.add('btn', 'btn-sm', 'btn-outline-primary', 'me-2');
        editBtn.textContent = t('common.edit');
        editBtn.addEventListener('click', () => openEmployeeModal(employee));
        actionsCell.appendChild(editBtn);
        
        // Delete button
        const deleteBtn = document.createElement('button');
        deleteBtn.classList.add('btn', 'btn-sm', 'btn-outline-danger');
        deleteBtn.textContent = t('common.delete');
        deleteBtn.addEventListener('click', () => {
            if (confirm(t('employees.confirmDelete'))) {
                deleteEmployee(employee.id);
            }
        });
        actionsCell.appendChild(deleteBtn);
        
        row.appendChild(actionsCell);
        
        employeesTable.appendChild(row);
    });
}

// Render teams view
function renderTeams() {
    const teamLists = {
        'A': elements.teamAList,
        'B': elements.teamBList,
        'C': elements.teamCList
    };
    
    // Clear team lists
    Object.values(teamLists).forEach(list => list.innerHTML = '');
    
    // Group employees by team
    const teamEmployees = {
        'A': appData.employees.filter(e => e.team === 'A'),
        'B': appData.employees.filter(e => e.team === 'B'),
        'C': appData.employees.filter(e => e.team === 'C')
    };
    
    // Render employees for each team
    Object.entries(teamEmployees).forEach(([team, employees]) => {
        const teamList = teamLists[team];
        
        if (employees.length === 0) {
            const emptyItem = document.createElement('li');
            emptyItem.classList.add('list-group-item', 'text-muted');
            emptyItem.textContent = 'No employees in this team';
            teamList.appendChild(emptyItem);
        } else {
            employees.forEach(employee => {
                const item = document.createElement('li');
                item.classList.add('list-group-item', 'd-flex', 'justify-content-between', 'align-items-center');
                
                const nameSpan = document.createElement('span');
                nameSpan.textContent = employee.name;
                
                const absentBadge = document.createElement('span');
                absentBadge.classList.add('badge', 'bg-secondary', 'rounded-pill');
                absentBadge.textContent = `${employee.absentDates.length} absent days`;
                
                item.appendChild(nameSpan);
                item.appendChild(absentBadge);
                
                teamList.appendChild(item);
            });
        }
    });
}

// Open employee modal for adding or editing
function openEmployeeModal(employee = null) {
    // Reset form
    elements.employeeForm.reset();
    elements.absentDatesContainer.innerHTML = '';
    
    // If there's a multiselect for skills, initialize it
    if (elements.employeeSkills) {
        // Clear existing options
        elements.employeeSkills.innerHTML = '';
        
        // Add options for each skill
        appData.skills.forEach(skill => {
            const option = document.createElement('option');
            option.value = skill.id;
            option.textContent = skill.name;
            option.title = skill.description;
            elements.employeeSkills.appendChild(option);
        });
    }
    
    if (employee) {
        // Editing existing employee
        elements.employeeId.value = employee.id;
        elements.employeeName.value = employee.name;
        elements.employeeTeam.value = employee.team;
        
        // Select skills if multiselect exists
        if (elements.employeeSkills && employee.skills) {
            Array.from(elements.employeeSkills.options).forEach(option => {
                option.selected = employee.skills.includes(parseInt(option.value));
            });
        }
        
        // Add absent dates
        if (employee.absentDates && employee.absentDates.length > 0) {
            employee.absentDates.forEach(dateStr => {
                addAbsentDateInput(dateStr);
            });
        } else {
            addAbsentDateInput();
        }
    } else {
        // Adding new employee
        elements.employeeId.value = '';
        addAbsentDateInput();
    }
    
    elements.employeeModal.show();
}

// Add an input for absent date
function addAbsentDateInput(dateStr = '') {
    const container = document.createElement('div');
    container.classList.add('input-group', 'mb-2');
    
    const input = document.createElement('input');
    input.type = 'date';
    input.classList.add('form-control', 'absent-date');
    if (dateStr) {
        input.value = dateStr;
    }
    
    const button = document.createElement('button');
    button.type = 'button';
    button.classList.add('btn', 'btn-outline-danger', 'remove-date-btn');
    button.textContent = 'Remove';
    button.addEventListener('click', () => container.remove());
    
    container.appendChild(input);
    container.appendChild(button);
    
    elements.absentDatesContainer.appendChild(container);
}

// Add a date range selector for absences
function addAbsentDateRange() {
    const container = document.createElement('div');
    container.classList.add('card', 'mb-3', 'p-3');
    
    const title = document.createElement('h6');
    title.textContent = currentLanguage === 'de' ? 'Abwesenheitszeitraum hinzufügen' : 'Add Absence Range';
    title.classList.add('mb-3');
    
    const rangeRow = document.createElement('div');
    rangeRow.classList.add('row', 'g-2', 'mb-2');
    
    // Start date column
    const startCol = document.createElement('div');
    startCol.classList.add('col-sm-5');
    
    const startLabel = document.createElement('label');
    startLabel.classList.add('form-label');
    startLabel.textContent = currentLanguage === 'de' ? 'Von' : 'From';
    
    const startInput = document.createElement('input');
    startInput.type = 'date';
    startInput.classList.add('form-control', 'range-start-date');
    
    startCol.appendChild(startLabel);
    startCol.appendChild(startInput);
    
    // End date column
    const endCol = document.createElement('div');
    endCol.classList.add('col-sm-5');
    
    const endLabel = document.createElement('label');
    endLabel.classList.add('form-label');
    endLabel.textContent = currentLanguage === 'de' ? 'Bis' : 'To';
    
    const endInput = document.createElement('input');
    endInput.type = 'date';
    endInput.classList.add('form-control', 'range-end-date');
    
    endCol.appendChild(endLabel);
    endCol.appendChild(endInput);
    
    // Add button column
    const btnCol = document.createElement('div');
    btnCol.classList.add('col-sm-2', 'd-flex', 'align-items-end');
    
    const addBtn = document.createElement('button');
    addBtn.type = 'button';
    addBtn.classList.add('btn', 'btn-primary', 'w-100');
    addBtn.textContent = currentLanguage === 'de' ? 'Hinzufügen' : 'Add';
    addBtn.addEventListener('click', () => {
        const startDate = new Date(startInput.value);
        const endDate = new Date(endInput.value);
        
        if (!startInput.value || !endInput.value) {
            alert(currentLanguage === 'de' ? 'Bitte geben Sie Start- und Enddatum ein' : 'Please enter both start and end dates');
            return;
        }
        
        if (startDate > endDate) {
            alert(currentLanguage === 'de' ? 'Das Startdatum muss vor dem Enddatum liegen' : 'Start date must be before end date');
            return;
        }
        
        // Add each day in the range as an individual absent date
        const dates = getDatesInRange(startDate, endDate);
        dates.forEach(dateStr => {
            addAbsentDateInput(dateStr);
        });
        
        // Clear the inputs
        startInput.value = '';
        endInput.value = '';
    });
    
    btnCol.appendChild(addBtn);
    
    // Close button for the range selector
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.classList.add('btn-close', 'position-absolute', 'top-0', 'end-0', 'm-2');
    closeBtn.addEventListener('click', () => container.remove());
    
    rangeRow.appendChild(startCol);
    rangeRow.appendChild(endCol);
    rangeRow.appendChild(btnCol);
    
    container.appendChild(title);
    container.appendChild(rangeRow);
    container.appendChild(closeBtn);
    
    // Add the date range selector before the individual date inputs
    elements.absentDatesContainer.parentElement.insertBefore(container, elements.absentDatesContainer);
}

// Helper function to get all dates in a range (inclusive)
function getDatesInRange(startDate, endDate) {
    const dates = [];
    const currentDate = new Date(startDate);
    
    // Loop until we reach the end date
    while (currentDate <= endDate) {
        dates.push(formatDateISO(currentDate));
        // Move to the next day
        currentDate.setDate(currentDate.getDate() + 1);
    }
    
    return dates;
}

// Save employee data
function saveEmployee() {
    const id = elements.employeeId.value ? parseInt(elements.employeeId.value) : Date.now();
    const name = elements.employeeName.value;
    const team = elements.employeeTeam.value;
    
    // Get skills if multiselect exists
    let skills = [];
    if (elements.employeeSkills) {
        skills = Array.from(elements.employeeSkills.selectedOptions).map(option => parseInt(option.value));
    }
    
    // Get absent dates
    const absentDates = [];
    elements.absentDatesContainer.querySelectorAll('.absent-date').forEach(input => {
        if (input.value) {
            absentDates.push(input.value);
        }
    });
    
    if (!name) {
        alert('Please enter a name for the employee.');
        return;
    }
    
    // Create or update employee
    const employee = {
        id,
        name,
        team,
        skills,
        absentDates
    };
    
    // If editing, find and update the employee
    const existingIndex = appData.employees.findIndex(e => e.id === id);
    if (existingIndex >= 0) {
        appData.employees[existingIndex] = employee;
    } else {
        // Otherwise add new employee
        appData.employees.push(employee);
    }
    
    // Update schedule if teams have changed
    generateSchedule();
    
    saveData();
    elements.employeeModal.hide();
    
    // Refresh views
    renderEmployees();
    renderTeams();
    renderSkills();
    renderSchedule();
}

// Delete an employee
function deleteEmployee(id) {
    if (confirm('Are you sure you want to delete this employee?')) {
        appData.employees = appData.employees.filter(e => e.id !== id);
        
        // Update schedule
        generateSchedule();
        
        saveData();
        
        // Refresh views
        renderEmployees();
        renderTeams();
        renderSkills();
        renderSchedule();
    }
}

// Open shift edit modal
function openShiftEditModal(day, shiftType, shift) {
    // Check if the week is locked
    if (isWeekLocked(appData.currentWeek)) {
        alert('This week is locked. Unlock it first to make changes.');
        return;
    }
    
    elements.shiftDay.value = day;
    elements.shiftType.value = shiftType;
    elements.shiftInfoLabel.textContent = `${day}, ${shiftType} Shift (Default Team: ${shift.team})`;
    
    // Populate employee selection
    elements.shiftEmployees.innerHTML = '';
    
    // Get current week's date for this day
    const weekIndex = appData.currentWeek;
    const dayIndex = DAYS_OF_WEEK.indexOf(day);
    const dateStr = appData.weekDates[weekIndex][dayIndex];
    
    // Show all employees instead of just those from the assigned team
    appData.employees.forEach(employee => {
        const option = document.createElement('option');
        option.value = employee.id;
        
        // Check if employee is absent
        const isAbsent = employee.absentDates.includes(dateStr);
        
        // Add team info to employee name for clarity
        option.textContent = `${employee.name} (Team ${employee.team})`;
        
        // Highlight if employee is from the default team for this shift
        if (employee.team === shift.team) {
            option.classList.add('text-primary');
            option.textContent += ' ✓';
        }
        
        // Check if employee is absent
        if (isAbsent) {
            // Apply strikethrough styling for absent employees
            option.innerHTML = ` ${option.textContent} (ABSENT) `;
            option.disabled = true;
            option.classList.add('text-danger', 'absent-employee');
            
            // Make sure the option is not selected if the employee is absent
            option.selected = false;
        } else {
            // Check if employee is already assigned (only if not absent)
            if (shift.employees.includes(employee.id)) {
                option.selected = true;
            }
        }
        
        elements.shiftEmployees.appendChild(option);
    });
    
    elements.shiftEditModal.show();
}

// Save shift assignment
function saveShiftAssignment() {
    const day = elements.shiftDay.value;
    const shiftType = elements.shiftType.value;
    const weekIndex = appData.currentWeek;
    
    // Check if the week is locked
    if (isWeekLocked(weekIndex)) {
        alert('This week is locked. Unlock it first to make changes.');
        elements.shiftEditModal.hide();
        return;
    }
    
    // Get current date for this day
    const dayIndex = DAYS_OF_WEEK.indexOf(day);
    const dateStr = appData.weekDates[weekIndex][dayIndex];
    
    // Get selected employees
    const selectedEmployees = Array.from(elements.shiftEmployees.selectedOptions)
        .map(option => parseInt(option.value));
    
    // Check if any selected employees are absent
    const absentEmployees = selectedEmployees.filter(id => {
        const employee = appData.employees.find(e => e.id === id);
        return employee && employee.absentDates.includes(dateStr);
    });
    
    if (absentEmployees.length > 0) {
        const absentNames = absentEmployees.map(id => {
            const employee = appData.employees.find(e => e.id === id);
            return employee ? employee.name : 'Unknown';
        }).join(', ');
        
        alert(`Warning: You've selected absent employees: ${absentNames}. These employees are marked as absent on this date and shouldn't be scheduled.`);
        return; // Prevent saving
    }
    
    // Get the default team for this shift
    const defaultTeam = appData.schedule[weekIndex][day][shiftType].team;
    
    // Count employees from non-default teams
    const nonDefaultTeamEmployees = selectedEmployees.filter(id => {
        const employee = appData.employees.find(e => e.id === id);
        return employee && employee.team !== defaultTeam;
    }).length;
    
    // Provide feedback if employees from other teams are selected
    if (nonDefaultTeamEmployees > 0) {
        const confirmMessage = `You've selected ${nonDefaultTeamEmployees} employee(s) from teams other than the default Team ${defaultTeam} for this shift. Continue with this assignment?`;
        if (!confirm(confirmMessage)) {
            return; // User canceled the assignment
        }
    }
    
    // Update the schedule
    appData.schedule[weekIndex][day][shiftType].employees = selectedEmployees;
    
    // Update the week's active version with the changes
    const versions = appData.weekVersions[weekIndex];
    let activeVersionKey = null;
    for (const [key, version] of Object.entries(versions)) {
        if (version.isActive) {
            activeVersionKey = key;
            break;
        }
    }
    
    if (activeVersionKey) {
        versions[activeVersionKey].schedule = JSON.parse(JSON.stringify(appData.schedule[weekIndex]));
    }
    
    saveData();
    elements.shiftEditModal.hide();
    renderSchedule();
    
    // Provide success feedback
    const totalAssigned = selectedEmployees.length;
    if (totalAssigned === 0) {
        alert('No employees assigned to this shift.');
    } else {
        // Create a message showing how many employees from each team were assigned
        const teamCounts = {};
        selectedEmployees.forEach(id => {
            const employee = appData.employees.find(e => e.id === id);
            if (employee) {
                teamCounts[employee.team] = (teamCounts[employee.team] || 0) + 1;
            }
        });
        
        const teamCountMessage = Object.entries(teamCounts)
            .map(([team, count]) => `${count} from Team ${team}`)
            .join(', ');
        
        const successMessage = `Successfully assigned ${totalAssigned} employee(s) to the ${shiftType} shift on ${day} (${teamCountMessage}).`;
        
        // Use a non-blocking notification instead of an alert
        // For simplicity in this implementation, we'll use console.log
        console.log(successMessage);
    }
}

// Export schedule to Excel using ExcelJS is now in excel-export.js

// Show information about new versioning and locking features
function showFeatureInfo() {
    // Check if we've already shown this info
    if (localStorage.getItem('versioning_info_shown')) {
        return;
    }
    
    // Create message with appropriate language
    let featureTitle, featureParagraph, versionControlLabel, scheduleLockingLabel, lookForLabel, gotItButton;
    
    if (currentLanguage === 'de') {
        featureTitle = "Neue Funktionen hinzugefügt!";
        featureParagraph = "Wir haben zwei neue leistungsstarke Funktionen hinzugefügt, die Ihnen bei der Schichtplanung helfen:";
        versionControlLabel = "Versionskontrolle - Erstellen und wechseln Sie zwischen verschiedenen Versionen Ihres Wochenplans";
        scheduleLockingLabel = "Zeitplansperre - Sperren Sie den Zeitplan einer Woche, um weitere Änderungen zu verhindern";
        lookForLabel = "Suchen Sie nach den neuen Versionskontrollen unter der Wochenauswahl und der Sperren/Entsperren-Schaltfläche oben rechts.";
        gotItButton = "Verstanden!";
    } else {
        featureTitle = "New Features Added!";
        featureParagraph = "We've added two new powerful features to help with your shift planning:";
        versionControlLabel = "Version Control - Create and switch between different versions of your weekly schedule";
        scheduleLockingLabel = "Schedule Locking - Lock a week's schedule to prevent further changes";
        lookForLabel = "Look for the new version controls below the week selector and the lock/unlock button in the top right.";
        gotItButton = "Got it!";
    }
    
    const message = `
        <h5>${featureTitle}</h5>
        <p>${featureParagraph}</p>
        <ul>
            <li><strong>${versionControlLabel}</strong></li>
            <li><strong>${scheduleLockingLabel}</strong></li>
        </ul>
        <p>${lookForLabel}</p>
    `;
    
    // Create info modal dynamically
    const infoModal = document.createElement('div');
    infoModal.className = 'modal fade';
    infoModal.id = 'feature-info-modal';
    infoModal.tabIndex = '-1';
    infoModal.innerHTML = `
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header bg-info text-white">
                    <h5 class="modal-title">${featureTitle}</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    ${message}
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-primary" data-bs-dismiss="modal">${gotItButton}</button>
                </div>
            </div>
        </div>
    `;
    
    document.body.appendChild(infoModal);
    
    // Show the modal
    const modal = new bootstrap.Modal(infoModal);
    modal.show();
    
    // Mark as shown
    localStorage.setItem('versioning_info_shown', 'true');
    
    // Clean up when hidden
    infoModal.addEventListener('hidden.bs.modal', () => {
        document.body.removeChild(infoModal);
    });
}

// Render skills table and management
function renderSkills() {
    const skillsTable = document.getElementById('skills-table').getElementsByTagName('tbody')[0];
    skillsTable.innerHTML = '';
    
    appData.skills.sort((a, b) => a.name.localeCompare(b.name)).forEach(skill => {
        const row = document.createElement('tr');
        
        // Name cell
        const nameCell = document.createElement('td');
        nameCell.textContent = skill.name;
        row.appendChild(nameCell);
        
        // Description cell
        const descCell = document.createElement('td');
        descCell.textContent = skill.description || '-';
        row.appendChild(descCell);
        
        // Employees with skill cell
        const employeesCell = document.createElement('td');
        const employeesWithSkill = appData.employees.filter(emp => emp.skills && emp.skills.includes(skill.id));
        
        if (employeesWithSkill.length > 0) {
            const employeesList = document.createElement('div');
            employeesList.classList.add('skills-summary');
            
            employeesWithSkill.forEach(emp => {
                const badge = document.createElement('span');
                badge.classList.add('badge');
                badge.textContent = emp.name;
                employeesList.appendChild(badge);
            });
            
            employeesCell.appendChild(employeesList);
        } else {
            employeesCell.textContent = '-';
        }
        row.appendChild(employeesCell);
        
        // Actions cell
        const actionsCell = document.createElement('td');
        
        // Edit button
        const editBtn = document.createElement('button');
        editBtn.classList.add('btn', 'btn-sm', 'btn-outline-primary', 'me-2');
        editBtn.textContent = t('common.edit');
        editBtn.addEventListener('click', () => openSkillModal(skill));
        actionsCell.appendChild(editBtn);
        
        // Delete button
        const deleteBtn = document.createElement('button');
        deleteBtn.classList.add('btn', 'btn-sm', 'btn-outline-danger');
        deleteBtn.textContent = t('common.delete');
        deleteBtn.addEventListener('click', () => {
            if (confirm(t('skills.confirmDelete'))) {
                deleteSkill(skill.id);
            }
        });
        actionsCell.appendChild(deleteBtn);
        
        row.appendChild(actionsCell);
        
        skillsTable.appendChild(row);
    });
}

// Open skill modal for adding or editing
function openSkillModal(skill = null) {
    if (!elements.skillForm) {
        console.error('Skill form not found');
        return;
    }
    
    // Reset form
    elements.skillForm.reset();
    
    if (skill) {
        // Editing existing skill
        elements.skillId.value = skill.id;
        elements.skillName.value = skill.name;
        elements.skillDescription.value = skill.description || '';
    } else {
        // Adding new skill
        elements.skillId.value = '';
    }
    
    // Show the modal
    const skillModal = new bootstrap.Modal(document.getElementById('skill-modal'));
    skillModal.show();
}

// Save skill data
function saveSkill() {
    const id = elements.skillId.value ? parseInt(elements.skillId.value) : Date.now();
    const name = elements.skillName.value;
    const description = elements.skillDescription.value;
    
    if (!name) {
        alert('Please enter a name for the skill.');
        return;
    }
    
    // Create or update skill
    const skill = {
        id,
        name,
        description
    };
    
    // If editing, find and update the skill
    const existingIndex = appData.skills.findIndex(s => s.id === id);
    if (existingIndex >= 0) {
        appData.skills[existingIndex] = skill;
    } else {
        // Otherwise add new skill
        appData.skills.push(skill);
    }
    
    saveData();
    
    // Hide the modal
    const skillModal = bootstrap.Modal.getInstance(document.getElementById('skill-modal'));
    skillModal.hide();
    
    // Refresh views
    renderSkills();
    renderEmployees(); // To update skill badges in employee list
}

// Delete a skill
function deleteSkill(id) {
    if (confirm('Are you sure you want to delete this skill? It will be removed from all employees.')) {
        // Remove the skill from all employees
        appData.employees.forEach(employee => {
            if (employee.skills) {
                employee.skills = employee.skills.filter(skillId => skillId !== id);
            }
        });
        
        // Remove the skill
        appData.skills = appData.skills.filter(s => s.id !== id);
        
        saveData();
        
        // Refresh views
        renderSkills();
        renderEmployees();
    }
}

// Get ISO week number for a date
function getWeekNumber(d) {
    // Copy date so don't modify original
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    // Set to nearest Thursday: current date + 4 - current day number
    // Make Sunday's day number 7
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
    // Get first day of year
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    // Calculate full weeks to nearest Thursday
    const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    // Return array of year and week number
    return weekNo;
}

// Update the week label to show current week number
function updateWeekLabel(weekIndex) {
    // Check if the week label wrapper exists
    let weekLabelWrapper = document.getElementById('week-label-wrapper');
    
    if (!weekLabelWrapper) {
        // Create week label wrapper
        weekLabelWrapper = document.createElement('div');
        weekLabelWrapper.id = 'week-label-wrapper';
        weekLabelWrapper.classList.add('d-flex', 'align-items-center');
        
        // Find the input group that contains the week selector
        const inputGroup = elements.weekSelector.closest('.input-group');
        
        // Replace input-group-text with our custom label
        const existingLabel = inputGroup.querySelector('.input-group-text');
        if (existingLabel) {
            existingLabel.remove();
        }
        
        // Add the week label wrapper before the select element
        inputGroup.prepend(weekLabelWrapper);
    }
    
    // Get the first date of the selected week
    const weekDates = appData.weekDates[weekIndex];
    if (!weekDates || weekDates.length === 0) {
        // Fallback to index if no dates available
        weekLabelWrapper.innerHTML = `<label class="me-2" style="margin-bottom: 0;"><strong>${t('schedule.weekLabel')} ${weekIndex + 1}:</strong></label>`;
        return;
    }
    
    // Get the Monday date for this week
    const mondayDate = new Date(weekDates[0]);
    
    // Calculate the ISO week number
    const weekNumber = getWeekNumber(mondayDate);
    
    // Set the content of the week label with the calendar week number
    weekLabelWrapper.innerHTML = `<label class="me-2" style="margin-bottom: 0;"><strong>${t('schedule.weekLabel')} ${weekNumber}:</strong></label>`;
}

// Format date for filename
function formatDateForFilename(date) {
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    return `${year}-${month}-${day}`;
}

// Convert column index to Excel column letter
function getExcelColumn(colIndex) {
    let temp, letter = '';
    while (colIndex > 0) {
        temp = (colIndex - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        colIndex = (colIndex - temp - 1) / 26;
    }
    return letter;
}

// Apply borders to Excel cell
function applyBordersToCell(cell) {
    cell.border = {
        top: { style: 'thin', color: { argb: 'BFBFBF' } },
        left: { style: 'thin', color: { argb: 'BFBFBF' } },
        bottom: { style: 'thin', color: { argb: 'BFBFBF' } },
        right: { style: 'thin', color: { argb: 'BFBFBF' } }
    };
}

// Initialize the application when the DOM is loaded
document.addEventListener('DOMContentLoaded', init); 

// Get team name with language support
function getTeamName(teamId) {
    if (!teamId) {
        return currentLanguage === 'de' ? '(Kein Team)' : '(No Team)';
    }
    
    return `Team ${teamId}`;
}

// Excel helper functions are now in excel-export.js

// Get all skills assigned to a team
function getTeamSkills(team) {
    const teamSkills = new Set();
    
    team.employees.forEach(empId => {
        const emp = employees.find(e => e.id === empId);
        if (emp && emp.skills) {
            emp.skills.forEach(skillId => {
                teamSkills.add(skillId);
            });
        }
    });
    
    // Return array of skill names
    return Array.from(teamSkills).map(skillId => {
        const skill = skills.find(s => s.id === skillId);
        return skill ? skill.name : skillId;
    });
}

// Clear all application data and start fresh
function clearAllData() {
    // Get confirmation message in current language
    const confirmMessage = t('notifications.clearDataConfirm');
    
    // Confirm before clearing data
    if (confirm(confirmMessage)) {
        // Show progress indicator
        showProgressIndicator(t('notifications.validatingData'));
        
        // Reset the app data to initial empty state
        appData = {
            employees: [],
            skills: [],
            schedule: {},
            weekVersions: {},
            lockedWeeks: {},
            currentWeek: 0,
            weekDates: generateNextFourWeeks()
        };
        
        // Save empty data to localStorage
        saveData();
        
        // Generate empty schedule for the next four weeks
        generateSchedule();
        
        // Initialize version tracking for each week
        Object.keys(appData.schedule).forEach(weekIndex => {
            appData.weekVersions[weekIndex] = {
                "v1": {
                    name: "Initial Version",
                    date: new Date().toISOString(),
                    isActive: true,
                    schedule: JSON.parse(JSON.stringify(appData.schedule[weekIndex]))
                }
            };
            appData.lockedWeeks[weekIndex] = false;
        });
        
        // Save to localStorage
        saveData();
        
        // Reload the application views
        generateWeekOptions();
        renderSchedule();
        renderEmployees();
        renderTeams();
        renderSkills();
        
        // Success message
        showProgressIndicator(t('notifications.dataCleared'));
    }
}

// Load demo data for application exploration
function loadDemoData() {
    // Show progress indicator
    showProgressIndicator(t('notifications.validatingData'));
    
    // Initialize with sample skills
    const defaultSkills = [
        { id: 1, name: 'First Aid', description: 'Certified in basic first aid' },
        { id: 2, name: 'Forklift', description: 'Licensed to operate forklifts' },
        { id: 3, name: 'Team Lead', description: 'Can lead a team' },
        { id: 4, name: 'Technical Support', description: 'Can provide technical assistance' },
        { id: 5, name: 'Quality Assurance', description: 'Can perform quality checks' },
        { id: 6, name: 'Inventory Management', description: 'Experienced in inventory control' }
    ];
    
    // Initialize with sample employees data - more comprehensive than the default
    const demoEmployees = [
        { id: 1, name: 'John Doe', team: 'A', absentDates: [], skills: [1, 3, 5] },
        { id: 2, name: 'Jane Smith', team: 'A', absentDates: [], skills: [1, 4, 6] },
        { id: 3, name: 'Bob Johnson', team: 'A', absentDates: [], skills: [2, 5] },
        { id: 4, name: 'Alice Brown', team: 'B', absentDates: [], skills: [2, 3, 4] },
        { id: 5, name: 'Charlie Davis', team: 'B', absentDates: [], skills: [4, 6] },
        { id: 6, name: 'Diana Evans', team: 'B', absentDates: [], skills: [1, 3] },
        { id: 7, name: 'Edward Moore', team: 'C', absentDates: [], skills: [2, 5] },
        { id: 8, name: 'Fiona Wilson', team: 'C', absentDates: [], skills: [3, 6] },
        { id: 9, name: 'George Harris', team: 'C', absentDates: [], skills: [1, 4] }
    ];
    
    // Add an absent date for one employee as an example
    const today = new Date();
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    
    // Format as YYYY-MM-DD
    const tomorrowFormatted = formatDateISO(tomorrow);
    demoEmployees[0].absentDates = [tomorrowFormatted];
    
    // Initialize app data with demo data
    appData = {
        employees: demoEmployees,
        skills: defaultSkills,
        schedule: {},
        weekVersions: {},
        lockedWeeks: {},
        currentWeek: 0,
        weekDates: generateNextFourWeeks()
    };
    
    // Generate demo schedule
    generateSchedule();
    
    // Initialize version tracking for each week
    Object.keys(appData.schedule).forEach(weekIndex => {
        // Create the initial version
        appData.weekVersions[weekIndex] = {
            "v1": {
                name: "Initial Version",
                date: new Date().toISOString(),
                isActive: false,
                schedule: JSON.parse(JSON.stringify(appData.schedule[weekIndex]))
            }
        };
        
        // For demo purposes, add a second version for the first week
        if (weekIndex === '0') {
            appData.weekVersions[weekIndex]["v2"] = {
                name: "Optimized",
                date: new Date().toISOString(),
                isActive: true,
                schedule: JSON.parse(JSON.stringify(appData.schedule[weekIndex]))
            };
            // Make a small change to show the difference
            const firstDay = Object.keys(appData.weekVersions[weekIndex]["v2"].schedule)[0];
            if (firstDay && appData.weekVersions[weekIndex]["v2"].schedule[firstDay].AM.employees.length > 0) {
                // Swap an employee between AM and PM shifts for demonstration
                const employee = appData.weekVersions[weekIndex]["v2"].schedule[firstDay].AM.employees[0];
                appData.weekVersions[weekIndex]["v2"].schedule[firstDay].AM.employees.splice(0, 1);
                appData.weekVersions[weekIndex]["v2"].schedule[firstDay].PM.employees.push(employee);
            }
        } else {
            appData.weekVersions[weekIndex]["v1"].isActive = true;
        }
        
        // Lock second week for demonstration
        appData.lockedWeeks[weekIndex] = (weekIndex === '1');
    });
    
    // Save to localStorage
    saveData();
    
    // Reload the application views
    generateWeekOptions();
    renderSchedule();
    renderEmployees();
    renderTeams();
    renderSkills();
    
    // Success message
    showProgressIndicator(t('notifications.demoDataLoaded'));
}