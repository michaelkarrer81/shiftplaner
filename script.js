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
    
    // Update loading text
    document.querySelector('.spinner-border .visually-hidden').textContent = t('app.loading');
    
    // Update error container
    document.querySelector('#error-container h4').textContent = t('app.error.title');
    
    // Update tooltips
    document.getElementById('export-data-button').title = t('tooltips.exportData');
    document.getElementById('import-data-button').title = t('tooltips.importData');
    
    // Update the current view based on which view is active
    if (currentView === 'schedule') updateScheduleView();
    else if (currentView === 'employees') updateEmployeesView();
    else if (currentView === 'teams') updateTeamsView();
    else if (currentView === 'skills') updateSkillsView();
}

function updateScheduleView() {
    // Update schedule view labels
    document.querySelector('label[for="week-selector"]').textContent = t('schedule.selectWeek');
    document.getElementById('export-button').textContent = t('schedule.exportToExcel');
    document.getElementById('plan-button').textContent = t('schedule.generatePlan');
    
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
const SHIFT_TYPES = ['AM', 'PM', 'Night'];

// Safely initialize Bootstrap modals
function initModals() {
    try {
        // Hide loading indicator first
        const loadingIndicator = document.getElementById('loading-indicator');
        if (loadingIndicator) {
            loadingIndicator.classList.add('d-none');
        }
        
        const employeeModalEl = document.getElementById('employee-modal');
        const shiftEditModalEl = document.getElementById('shift-edit-modal');
        
        if (!employeeModalEl || !shiftEditModalEl) {
            throw new Error('Modal elements not found in the DOM');
        }
        
        elements.employeeModal = new bootstrap.Modal(employeeModalEl);
        elements.shiftEditModal = new bootstrap.Modal(shiftEditModalEl);
        
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
            importDataButton: document.getElementById('import-data-button')
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
        
        option.textContent = `Week ${index + 1}: ${startFormatted} - ${endFormatted}`;
        
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
}

// Format date as DD/MM/YYYY
function formatDate(date) {
    return `${date.getDate().toString().padStart(2, '0')}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getFullYear()}`;
}

// Format date as YYYY-MM-DD
function formatDateISO(date) {
    return `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}-${date.getDate().toString().padStart(2, '0')}`;
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
                    <h5 class="modal-title">Generate Plan for Date Range</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    <form id="plan-form">
                        <div class="mb-3">
                            <label for="start-date" class="form-label">Start Date</label>
                            <input type="date" class="form-control" id="start-date" value="${startDateValue}" required>
                        </div>
                        <div class="mb-3">
                            <label for="end-date" class="form-label">End Date</label>
                            <input type="date" class="form-control" id="end-date" value="${endDateValue}" required>
                        </div>
                        <div class="mb-3 form-check">
                            <input type="checkbox" class="form-check-input" id="create-versions" checked>
                            <label class="form-check-label" for="create-versions">Create new versions for existing weeks</label>
                        </div>
                        <div class="mb-3">
                            <p class="text-info">
                                <i class="bi bi-info-circle"></i> Locked weeks will not be modified.
                            </p>
                        </div>
                    </form>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                    <button type="button" class="btn btn-primary" id="generate-plan-btn">Generate Plan</button>
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
            alert('End date cannot be before start date');
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
    
    alert(`Plan generated successfully for ${formatDate(startDate)} - ${formatDate(endDate)}.\n${existingWeekIndices.length} existing unlocked weeks were replanned.`);
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
        const row = document.createElement('tr');
        
        // Day column
        const dayCell = document.createElement('td');
        dayCell.textContent = day;
        if (daySchedule.date) {
            const date = new Date(daySchedule.date);
            dayCell.textContent += ` (${formatDate(date)})`;
        }
        
        // Add skill summary for the day
        const daySkillSummary = generateDailySkillSummary(weekIndex, day);
        if (daySkillSummary.length > 0) {
            const skillsContainer = document.createElement('div');
            skillsContainer.classList.add('mt-2', 'small', 'skills-summary');
            
            const skillsLabel = document.createElement('strong');
            skillsLabel.textContent = 'Skills available today: ';
            skillsContainer.appendChild(skillsLabel);
            
            daySkillSummary.forEach(skill => {
                const skillBadge = document.createElement('span');
                skillBadge.classList.add('badge', 'bg-success', 'me-1');
                skillBadge.textContent = skill.name;
                skillBadge.title = `${skill.count} employees with this skill available`;
                skillsContainer.appendChild(skillBadge);
            });
            
            dayCell.appendChild(skillsContainer);
        }
        
        row.appendChild(dayCell);
        
        // Shift columns
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
            
            // Add employee names
            const employeesList = document.createElement('div');
            
            // Get date for this day
            const dayIndex = DAYS_OF_WEEK.indexOf(day);
            const dateStr = appData.weekDates[weekIndex][dayIndex];
            
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
                
                // Add employee names with their team indicator
                shift.employees.forEach(employeeId => {
                    const employee = appData.employees.find(e => e.id === employeeId);
                    if (employee) {
                        const employeeSpan = document.createElement('div');
                        const isAbsent = employee.absentDates.includes(dateStr);
                        
                        // Apply strikethrough if employee is absent
                        if (isAbsent) {
                            employeeSpan.innerHTML = `<s>${employee.name}</s> <span class="badge bg-danger">Absent</span>`;
                            employeeSpan.classList.add('text-danger');
                        } else {
                            employeeSpan.textContent = employee.name;
                            
                            // Add team indicator if employee is from a different team
                            if (employee.team !== shift.team) {
                                employeeSpan.innerHTML += ` <small class="text-muted">(Team ${employee.team})</small>`;
                                employeeSpan.classList.add('text-primary');
                            }
                            
                            // Add skills badges if employee has skills
                            if (employee.skills && employee.skills.length > 0) {
                                const skillsSpan = document.createElement('div');
                                skillsSpan.classList.add('employee-skills', 'small');
                                
                                employee.skills.forEach(skillId => {
                                    const skill = appData.skills.find(s => s.id === skillId);
                                    if (skill) {
                                        const badge = document.createElement('span');
                                        badge.classList.add('badge', 'bg-info', 'me-1');
                                        badge.textContent = skill.name;
                                        badge.title = skill.description;
                                        skillsSpan.appendChild(badge);
                                    }
                                });
                                
                                employeeSpan.appendChild(skillsSpan);
                            }
                        }
                        
                        employeesList.appendChild(employeeSpan);
                    }
                });
                
                // Add skill summary for this shift
                const shiftSkills = getShiftSkills(shift.employees, dateStr);
                if (shiftSkills.length > 0) {
                    const skillsDiv = document.createElement('div');
                    skillsDiv.classList.add('mt-2', 'shift-skills-summary', 'small');
                    
                    const skillsLabel = document.createElement('strong');
                    skillsLabel.textContent = 'Skills: ';
                    skillsDiv.appendChild(skillsLabel);
                    
                    shiftSkills.forEach(skillId => {
                        const skill = appData.skills.find(s => s.id === skillId);
                        if (skill) {
                            const badge = document.createElement('span');
                            badge.classList.add('badge', 'bg-success', 'me-1');
                            badge.textContent = skill.name;
                            skillsDiv.appendChild(badge);
                        }
                    });
                    
                    employeesList.appendChild(skillsDiv);
                }
            } else {
                employeesList.textContent = 'No employees assigned';
            }
            
            shiftCell.appendChild(employeesList);
            
            // Add click event to edit shift only if week is not locked
            if (!isLocked) {
                shiftCell.addEventListener('click', () => openShiftEditModal(day, shiftType, shift));
                shiftCell.style.cursor = 'pointer';
            } else {
                shiftCell.style.cursor = 'not-allowed';
                shiftCell.title = 'This week is locked. Unlock it to make changes.';
                shiftCell.classList.add('disabled-cell');
            }
            
            row.appendChild(shiftCell);
        });
        
        tbody.appendChild(row);
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
            <h6 class="mb-0">Weekly Skill Coverage Summary</h6>
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
        
        html += `
            <div class="col-md-4">
                <h6>${shiftType} Shifts</h6>
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
            html += `<div class="list-group-item">No skills available</div>`;
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
            🔒 <span>Week ${weekIndex + 1} is locked</span> 
            <span class="ms-1">(Current version: ${activeVersion ? activeVersion.name : 'Unknown'})</span>
        `;
    } else {
        statusBanner.className = 'small mb-2 text-info border-top pt-1';
        statusBanner.innerHTML = `
            ✏️ <span>Week ${weekIndex + 1} is editable</span>
            <span class="ms-1">(Current version: ${activeVersion ? activeVersion.name : 'Unknown'})</span>
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
                if (confirm('Are you sure you want to unlock this week? This will allow edits to the schedule.')) {
                    toggleWeekLock(weekIndex, false);
                }
            } else {
                // If unlocked, confirm lock
                if (confirm('Are you sure you want to lock this week? This will prevent further changes until unlocked.')) {
                    toggleWeekLock(weekIndex, true);
                }
            }
        });
    }
    
    // Update button appearance based on lock state
    if (isLocked) {
        lockButton.textContent = '🔓 Unlock Week';
        lockButton.classList.remove('btn-danger');
        lockButton.classList.add('btn-success');
    } else {
        lockButton.textContent = '🔒 Lock Week';
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
            <label class="me-2"><strong>Versions:</strong></label>
            <select id="version-selector" class="form-select me-2" ${isLocked ? 'disabled' : ''}>
                ${Object.entries(versions).map(([key, version]) => `
                    <option value="${key}" ${version.isActive ? 'selected' : ''}>
                        ${version.name} (${new Date(version.date).toLocaleDateString()})
                    </option>
                `).join('')}
            </select>
            <button id="new-version-btn" class="btn btn-outline-primary btn-sm" ${isLocked ? 'disabled' : ''}>
                Save as New Version
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
    const tbody = elements.employeesTable.querySelector('tbody');
    tbody.innerHTML = '';
    
    appData.employees.forEach(employee => {
        const row = document.createElement('tr');
        
        // Name column
        const nameCell = document.createElement('td');
        nameCell.textContent = employee.name;
        row.appendChild(nameCell);
        
        // Team column
        const teamCell = document.createElement('td');
        const teamBadge = document.createElement('span');
        teamBadge.classList.add('team-badge', `team-${employee.team.toLowerCase()}-badge`);
        teamBadge.textContent = `Team ${employee.team}`;
        teamCell.appendChild(teamBadge);
        row.appendChild(teamCell);
        
        // Skills column
        const skillsCell = document.createElement('td');
        if (employee.skills && employee.skills.length > 0) {
            employee.skills.forEach(skillId => {
                const skill = appData.skills.find(s => s.id === skillId);
                if (skill) {
                    const skillBadge = document.createElement('span');
                    skillBadge.classList.add('badge', 'bg-info', 'me-1', 'mb-1');
                    skillBadge.textContent = skill.name;
                    skillBadge.title = skill.description;
                    skillsCell.appendChild(skillBadge);
                }
            });
        } else {
            skillsCell.textContent = 'None';
        }
        row.appendChild(skillsCell);
        
        // Absent dates column
        const absentCell = document.createElement('td');
        if (employee.absentDates && employee.absentDates.length > 0) {
            employee.absentDates.forEach(dateStr => {
                const date = new Date(dateStr);
                const datePill = document.createElement('span');
                datePill.classList.add('absent-date-pill');
                datePill.textContent = formatDate(date);
                absentCell.appendChild(datePill);
            });
        } else {
            absentCell.textContent = 'None';
        }
        row.appendChild(absentCell);
        
        // Actions column
        const actionsCell = document.createElement('td');
        
        const editButton = document.createElement('button');
        editButton.classList.add('btn', 'btn-sm', 'btn-outline-primary', 'me-2');
        editButton.textContent = 'Edit';
        editButton.addEventListener('click', () => openEmployeeModal(employee));
        
        const deleteButton = document.createElement('button');
        deleteButton.classList.add('btn', 'btn-sm', 'btn-outline-danger');
        deleteButton.textContent = 'Delete';
        deleteButton.addEventListener('click', () => deleteEmployee(employee.id));
        
        actionsCell.appendChild(editButton);
        actionsCell.appendChild(deleteButton);
        
        row.appendChild(actionsCell);
        
        tbody.appendChild(row);
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
            option.innerHTML = `<s>${option.textContent} (ABSENT)</s>`;
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

// Export schedule to Excel using ExcelJS
function exportToExcel() {
    // Show progress indicator
    showProgressIndicator('Preparing Excel export...');
    
    try {
        // Create a new workbook
        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Shift Planner';
        workbook.lastModifiedBy = 'Shift Planner App';
        workbook.created = new Date();
        workbook.modified = new Date();
        
        // Get the current week data
        const weekIndex = appData.currentWeek;
        const weekSchedule = appData.schedule[weekIndex] || {};
        const weekDates = appData.weekDates[weekIndex] || [];
        
        if (!weekDates.length) {
            alert('No data available for the selected week.');
            return;
        }
        
        const startDate = new Date(weekDates[0]);
        const endDate = new Date(weekDates[weekDates.length - 1]);
        const isLocked = isWeekLocked(weekIndex);
        const activeVersion = getActiveVersion(weekIndex);
        
        const weekNumber = getWeekNumber(startDate);
        const weekTitle = `Week ${weekNumber}: ${formatDate(startDate)} - ${formatDate(endDate)}`;
        const versionInfo = activeVersion ? activeVersion.name : 'Default';
        const statusInfo = isLocked ? 'LOCKED 🔒' : 'EDITABLE';
        
        // Define colors
        const colors = {
            titleBg: '4472C4',
            titleText: 'FFFFFF',
            headerBg: 'E0E0E0',
            subHeaderBg: 'F0F0F0',
            versionInfoBg: 'DEEBF7',
            teamA: 'DAEEF3',
            teamB: 'E4DFEC',
            teamC: 'FDE9D9',
            skillHighlight: 'E2EFDA',
            absentHighlight: 'FFC7CE',
            teamAText: '2F75B5',
            teamBText: '7030A0',
            teamCText: 'C65911',
            absentText: '9C0006'
        };
        
        // 1. Create Weekly Summary Sheet
        createWeeklySummarySheet(workbook, weekIndex, weekTitle, versionInfo, statusInfo, colors);
        
        // 2. Create Detailed Schedule Sheet
        createDetailedScheduleSheet(workbook, weekIndex, weekTitle, versionInfo, statusInfo, colors);
        
        // 3. Create Data Records Sheet
        createDataRecordsSheet(workbook, weekIndex, colors);
        
        // Create filename with week number and date
        const currentDate = new Date();
        const fileName = `ShiftPlanner_Week${weekNumber}_${currentDate.getFullYear()}-${(currentDate.getMonth() + 1).toString().padStart(2, '0')}-${currentDate.getDate().toString().padStart(2, '0')}.xlsx`;
        
        // Generate and save Excel file
        workbook.xlsx.writeBuffer().then(buffer => {
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            saveAs(blob, fileName);
            showProgressIndicator('Excel export completed successfully!');
        });
    } catch (error) {
        console.error('Error during Excel export:', error);
        showError('An error occurred during Excel export. Please try again.');
    }
}

// Create the Weekly Summary Sheet with ExcelJS
function createWeeklySummarySheet(workbook, weekIndex, weekTitle, versionInfo, statusInfo, colors) {
    const weekSchedule = appData.schedule[weekIndex] || {};
    const weekDates = appData.weekDates[weekIndex] || [];
    
    // Create sheet
    const worksheet = workbook.addWorksheet('Weekly Summary', {
        properties: { tabColor: { argb: colors.titleBg } }
    });
    
    // Set column widths
    worksheet.columns = [
        { header: '', width: 20 },  // A - Day/Date
        { header: '', width: 25 },  // B - AM Shift
        { header: '', width: 25 },  // C - PM Shift
        { header: '', width: 25 }   // D - Night Shift
    ];
    
    // Title row
    const titleRow = worksheet.addRow([weekTitle]);
    titleRow.height = 30;
    titleRow.font = { bold: true, size: 16, color: { argb: colors.titleText } };
    titleRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.titleBg } };
    titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells('A1:D1');
    applyCellBorders(titleRow, 1, 4, 'medium');
    
    // Status/Version row
    const statusRow = worksheet.addRow([`Status: ${statusInfo} | Version: ${versionInfo}`]);
    statusRow.height = 24;
    statusRow.font = { italic: true, size: 11 };
    statusRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.versionInfoBg } };
    statusRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells('A2:D2');
    applyCellBorders(statusRow, 1, 4);
    
    // Empty row for spacing
    worksheet.addRow([]);
    
    // ===== TEAM ASSIGNMENT SECTION =====
    
    // Team Assignment header
    const teamHeader = worksheet.addRow(['Weekly Team Assignment Summary']);
    teamHeader.height = 25;
    teamHeader.font = { bold: true, size: 14 };
    teamHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    teamHeader.alignment = { horizontal: 'left', vertical: 'middle' };
    worksheet.mergeCells('A4:D4');
    applyCellBorders(teamHeader, 1, 4);
    
    // Team table headers
    const shiftHeaderRow = worksheet.addRow(['Day/Date', 'AM Shift', 'PM Shift', 'Night Shift']);
    shiftHeaderRow.height = 20;
    shiftHeaderRow.font = { bold: true, size: 12 };
    shiftHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.subHeaderBg } };
    shiftHeaderRow.alignment = { horizontal: 'center', vertical: 'middle' };
    applyCellBorders(shiftHeaderRow, 1, 4);
    
    // Add data for each day
    DAYS_OF_WEEK.forEach((day, dayIndex) => {
        const daySchedule = weekSchedule[day] || {};
        const dateStr = weekDates[dayIndex] || '';
        const date = dateStr ? new Date(dateStr) : null;
        const dateFormatted = date ? formatDate(date) : '';
        
        const rowData = [`${day} (${dateFormatted})`];
        
        SHIFT_TYPES.forEach(shiftType => {
            const shift = daySchedule[shiftType] || { team: '-', employees: [] };
            const teamCode = shift.team;
            const employeeCount = shift.employees ? shift.employees.length : 0;
            
            // Count absent employees
            let absentCount = 0;
            if (shift.employees && dateStr) {
                absentCount = shift.employees.filter(empId => {
                    const employee = appData.employees.find(e => e.id === empId);
                    return employee && employee.absentDates.includes(dateStr);
                }).length;
            }
            
            rowData.push(`Team ${teamCode} (${employeeCount - absentCount}/${employeeCount} available)`);
        });
        
        const dataRow = worksheet.addRow(rowData);
        dataRow.height = 20;
        
        // Style the day cell (first column)
        dataRow.getCell(1).font = { size: 11 };
        dataRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF' } };
        dataRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
        dataRow.getCell(1).border = getBorderStyle();
        
        // Style shift cells based on team
        for (let i = 0; i < SHIFT_TYPES.length; i++) {
            const cell = dataRow.getCell(i + 2);
            const cellValue = rowData[i + 1];
            const teamMatch = cellValue.match(/Team ([A-Z])/);
            const team = teamMatch ? teamMatch[1] : null;
            
            let bgColor = colors.subHeaderBg;
            let textColor = '000000';
            
            if (team === 'A') {
                bgColor = colors.teamA;
                textColor = colors.teamAText;
            } else if (team === 'B') {
                bgColor = colors.teamB;
                textColor = colors.teamBText;
            } else if (team === 'C') {
                bgColor = colors.teamC;
                textColor = colors.teamCText;
            }
            
            cell.font = { size: 11, color: { argb: textColor } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = getBorderStyle();
        }
    });
    
    // Add spacing
    worksheet.addRow([]);
    worksheet.addRow([]);
    
    // ===== SKILL COVERAGE SECTION =====
    
    // Skill Coverage header
    const skillRowIndex = worksheet.rowCount;
    const skillHeader = worksheet.addRow(['Weekly Skill Coverage Summary']);
    skillHeader.height = 25;
    skillHeader.font = { bold: true, size: 14 };
    skillHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    skillHeader.alignment = { horizontal: 'left', vertical: 'middle' };
    worksheet.mergeCells(`A${skillRowIndex}:D${skillRowIndex}`);
    applyCellBorders(skillHeader, 1, 4);
    
    // Skill table headers
    const skillHeaderRow = worksheet.addRow(['Skill', 'AM Shift Coverage', 'PM Shift Coverage', 'Night Shift Coverage']);
    skillHeaderRow.height = 20;
    skillHeaderRow.font = { bold: true, size: 12 };
    skillHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.subHeaderBg } };
    skillHeaderRow.alignment = { horizontal: 'center', vertical: 'middle' };
    applyCellBorders(skillHeaderRow, 1, 4);
    
    // Get all skills used in this week
    const allSkillsMap = {};
    DAYS_OF_WEEK.forEach(day => {
        const daySchedule = weekSchedule[day] || {};
        const dayIndex = DAYS_OF_WEEK.indexOf(day);
        const dateStr = weekDates[dayIndex] || '';
        
        SHIFT_TYPES.forEach(shiftType => {
            const shift = daySchedule[shiftType] || { employees: [] };
            
            shift.employees.forEach(employeeId => {
                const employee = appData.employees.find(e => e.id === employeeId);
                if (employee && !employee.absentDates.includes(dateStr)) {
                    (employee.skills || []).forEach(skillId => {
                        if (!allSkillsMap[skillId]) {
                            const skill = appData.skills.find(s => s.id === skillId);
                            if (skill) {
                                allSkillsMap[skillId] = {
                                    name: skill.name,
                                    shifts: {}
                                };
                                SHIFT_TYPES.forEach(s => {
                                    allSkillsMap[skillId].shifts[s] = 0;
                                });
                            }
                        }
                        
                        if (allSkillsMap[skillId]) {
                            allSkillsMap[skillId].shifts[shiftType]++;
                        }
                    });
                }
            });
        });
    });
    
    // Add skill coverage data
    Object.values(allSkillsMap).forEach(skillData => {
        const rowData = [
            skillData.name,
            `${skillData.shifts.AM} employees`,
            `${skillData.shifts.PM} employees`,
            `${skillData.shifts.Night} employees`
        ];
        
        const dataRow = worksheet.addRow(rowData);
        dataRow.height = 20;
        
        // Style the skill name cell
        dataRow.getCell(1).font = { size: 11 };
        dataRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.skillHighlight } };
        dataRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
        dataRow.getCell(1).border = getBorderStyle();
        
        // Style coverage cells
        for (let i = 2; i <= 4; i++) {
            const cell = dataRow.getCell(i);
            cell.font = { size: 11 };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            cell.border = getBorderStyle();
        }
    });
    
    // Add spacing
    worksheet.addRow([]);
    worksheet.addRow([]);
    
    // ===== EXCEPTIONS SECTION (ABSENCES) =====
    
    // Exceptions header
    const exceptionRowIndex = worksheet.rowCount;
    const exceptionHeader = worksheet.addRow(['Weekly Exceptions']);
    exceptionHeader.height = 25;
    exceptionHeader.font = { bold: true, size: 14 };
    exceptionHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    exceptionHeader.alignment = { horizontal: 'left', vertical: 'middle' };
    worksheet.mergeCells(`A${exceptionRowIndex}:D${exceptionRowIndex}`);
    applyCellBorders(exceptionHeader, 1, 4);
    
    // Exceptions table headers
    const absenceHeaderRow = worksheet.addRow(['Employee', 'Absent Date', 'Team', 'Note']);
    absenceHeaderRow.height = 20;
    absenceHeaderRow.font = { bold: true, size: 12 };
    absenceHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.subHeaderBg } };
    absenceHeaderRow.alignment = { horizontal: 'center', vertical: 'middle' };
    applyCellBorders(absenceHeaderRow, 1, 4);
    
    // List all absences for the week
    const absences = [];
    weekDates.forEach(dateStr => {
        appData.employees.forEach(employee => {
            if (employee.absentDates.includes(dateStr)) {
                absences.push({
                    name: employee.name,
                    date: formatDate(new Date(dateStr)),
                    team: `Team ${employee.team}`,
                    note: 'Absent'
                });
            }
        });
    });
    
    // Sort absences by date
    absences.sort((a, b) => a.date.localeCompare(b.date));
    
    // Add absences to the sheet
    absences.forEach(absence => {
        const rowData = [absence.name, absence.date, absence.team, absence.note];
        const dataRow = worksheet.addRow(rowData);
        dataRow.height = 20;
        
        // Style all cells in the absence row
        for (let i = 1; i <= 4; i++) {
            const cell = dataRow.getCell(i);
            cell.font = { size: 11, color: { argb: i === 4 ? colors.absentText : '000000' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.absentHighlight } };
            cell.alignment = { horizontal: 'left', vertical: 'middle' };
            cell.border = getBorderStyle();
        }
    });
}

// Create the Detailed Schedule Sheet with ExcelJS
function createDetailedScheduleSheet(workbook, weekIndex, weekTitle, versionInfo, statusInfo, colors) {
    const weekSchedule = appData.schedule[weekIndex] || {};
    const weekDates = appData.weekDates[weekIndex] || [];
    
    // Create sheet
    const worksheet = workbook.addWorksheet('Detailed Schedule', {
        properties: { tabColor: { argb: '4472C4' } }
    });
    
    // Set column widths
    worksheet.columns = [
        { header: '', width: 20 },  // A - Day/Date
        { header: '', width: 45 },  // B - AM Shift
        { header: '', width: 45 },  // C - PM Shift
        { header: '', width: 45 }   // D - Night Shift
    ];
    
    // Title row
    const titleRow = worksheet.addRow([weekTitle]);
    titleRow.height = 30;
    titleRow.font = { bold: true, size: 16, color: { argb: colors.titleText } };
    titleRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.titleBg } };
    titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells('A1:D1');
    applyCellBorders(titleRow, 1, 4, 'medium');
    
    // Status/Version row
    const statusRow = worksheet.addRow([`Status: ${statusInfo} | Version: ${versionInfo}`]);
    statusRow.height = 24;
    statusRow.font = { italic: true, size: 11 };
    statusRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.versionInfoBg } };
    statusRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells('A2:D2');
    applyCellBorders(statusRow, 1, 4);
    
    // Empty row for spacing
    worksheet.addRow([]);
    
    // Detailed Schedule header
    const detailedHeader = worksheet.addRow(['Detailed Schedule View']);
    detailedHeader.height = 25;
    detailedHeader.font = { bold: true, size: 14 };
    detailedHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    detailedHeader.alignment = { horizontal: 'left', vertical: 'middle' };
    worksheet.mergeCells('A4:D4');
    applyCellBorders(detailedHeader, 1, 4);
    
    // Schedule table headers
    const headerRow = worksheet.addRow(['Day/Date', 'AM Shift', 'PM Shift', 'Night Shift']);
    headerRow.height = 20;
    headerRow.font = { bold: true, size: 12 };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.subHeaderBg } };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
    applyCellBorders(headerRow, 1, 4);
    
    // Add data for each day
    DAYS_OF_WEEK.forEach((day, dayIndex) => {
        const daySchedule = weekSchedule[day] || {};
        const dateStr = weekDates[dayIndex] || '';
        const date = dateStr ? new Date(dateStr) : null;
        const dateFormatted = date ? formatDate(date) : '';
        
        const rowData = [`${day} (${dateFormatted})`];
        const employeeDetails = [];
        
        SHIFT_TYPES.forEach(shiftType => {
            const shift = daySchedule[shiftType] || { team: '-', employees: [] };
            let cellContent = `Team ${shift.team}\n`;
            let employeeList = [];
            
            if (shift.employees && shift.employees.length > 0) {
                employeeList = shift.employees.map(empId => {
                    const employee = appData.employees.find(e => e.id === empId);
                    if (!employee) return null;
                    
                    const isAbsent = employee.absentDates.includes(dateStr);
                    const differentTeam = employee.team !== shift.team;
                    
                    let employeeText = employee.name;
                    
                    if (isAbsent) {
                        employeeText += ' (ABSENT)';
                    }
                    
                    if (differentTeam) {
                        employeeText += ` (Team ${employee.team})`;
                    }
                    
                    // Add skills
                    if (employee.skills && employee.skills.length > 0 && !isAbsent) {
                        const skillNames = employee.skills
                            .map(skillId => {
                                const skill = appData.skills.find(s => s.id === skillId);
                                return skill ? skill.name : '';
                            })
                            .filter(name => name);
                        
                        if (skillNames.length > 0) {
                            employeeText += ` - Skills: ${skillNames.join(', ')}`;
                        }
                    }
                    
                    return {
                        text: employeeText,
                        isAbsent: isAbsent,
                        differentTeam: differentTeam
                    };
                }).filter(Boolean);
            }
            
            rowData.push(cellContent + employeeList.map(e => e.text).join('\n'));
            employeeDetails.push(employeeList);
        });
        
        const dataRow = worksheet.addRow(rowData);
        dataRow.height = 120; // Set tall row for employee lists
        
        // Style the day cell (first column)
        dataRow.getCell(1).font = { size: 11 };
        dataRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
        dataRow.getCell(1).border = getBorderStyle();
        
        // Style shift cells
        for (let i = 0; i < SHIFT_TYPES.length; i++) {
            const cell = dataRow.getCell(i + 2);
            const cellValue = rowData[i + 1];
            const teamMatch = cellValue.match(/Team ([A-Z])/);
            const team = teamMatch ? teamMatch[1] : null;
            
            let bgColor = colors.subHeaderBg;
            
            if (team === 'A') {
                bgColor = colors.teamA;
            } else if (team === 'B') {
                bgColor = colors.teamB;
            } else if (team === 'C') {
                bgColor = colors.teamC;
            }
            
            cell.font = { size: 11 };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
            cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
            cell.border = getBorderStyle();
        }
    });
}

// Create the Data Records Sheet with ExcelJS
function createDataRecordsSheet(workbook, weekIndex, colors) {
    // Create sheet
    const worksheet = workbook.addWorksheet('Data Records', {
        properties: { tabColor: { argb: '4472C4' } }
    });
    
    // Set column widths
    worksheet.columns = [
        { header: '', width: 10 },  // A - ID
        { header: '', width: 25 },  // B - Name
        { header: '', width: 15 },  // C - Team
        { header: '', width: 40 },  // D - Skills
        { header: '', width: 40 }   // E - Absent Dates
    ];
    
    // Title row for Employee section
    const titleRow = worksheet.addRow(['Employee Data Records']);
    titleRow.height = 30;
    titleRow.font = { bold: true, size: 16, color: { argb: colors.titleText } };
    titleRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.titleBg } };
    titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells('A1:E1');
    applyCellBorders(titleRow, 1, 5, 'medium');
    
    // Empty row for spacing
    worksheet.addRow([]);
    
    // Employee table headers
    const headerRow = worksheet.addRow(['ID', 'Name', 'Team', 'Skills', 'Absent Dates']);
    headerRow.height = 20;
    headerRow.font = { bold: true, size: 12 };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
    applyCellBorders(headerRow, 1, 5);
    
    // Add employee data
    appData.employees.forEach(employee => {
        const skillNames = (employee.skills || [])
            .map(skillId => {
                const skill = appData.skills.find(s => s.id === skillId);
                return skill ? skill.name : '';
            })
            .filter(name => name)
            .join(', ');
        
        const absentDatesFormatted = (employee.absentDates || [])
            .map(dateStr => formatDate(new Date(dateStr)))
            .join(', ');
        
        const rowData = [
            employee.id,
            employee.name,
            `Team ${employee.team}`,
            skillNames || 'None',
            absentDatesFormatted || 'None'
        ];
        
        const dataRow = worksheet.addRow(rowData);
        dataRow.height = 22;
        
        // Style ID and Name cells
        dataRow.getCell(1).font = { size: 11 };
        dataRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
        dataRow.getCell(1).border = getBorderStyle();
        
        dataRow.getCell(2).font = { size: 11 };
        dataRow.getCell(2).alignment = { horizontal: 'left', vertical: 'middle' };
        dataRow.getCell(2).border = getBorderStyle();
        
        // Style team cell with team color
        const team = employee.team;
        let teamBgColor = colors.subHeaderBg;
        let teamTextColor = '000000';
        
        if (team === 'A') {
            teamBgColor = colors.teamA;
            teamTextColor = colors.teamAText;
        } else if (team === 'B') {
            teamBgColor = colors.teamB;
            teamTextColor = colors.teamBText;
        } else if (team === 'C') {
            teamBgColor = colors.teamC;
            teamTextColor = colors.teamCText;
        }
        
        dataRow.getCell(3).font = { size: 11, color: { argb: teamTextColor } };
        dataRow.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: teamBgColor } };
        dataRow.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
        dataRow.getCell(3).border = getBorderStyle();
        
        // Style skills cell
        dataRow.getCell(4).font = { size: 11 };
        dataRow.getCell(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.skillHighlight } };
        dataRow.getCell(4).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        dataRow.getCell(4).border = getBorderStyle();
        
        // Style absent dates cell
        if (absentDatesFormatted !== 'None') {
            dataRow.getCell(5).font = { size: 11, color: { argb: colors.absentText } };
            dataRow.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.absentHighlight } };
        } else {
            dataRow.getCell(5).font = { size: 11 };
        }
        dataRow.getCell(5).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        dataRow.getCell(5).border = getBorderStyle();
    });
    
    // Add spacing
    worksheet.addRow([]);
    worksheet.addRow([]);
    
    // Skills section title
    const skillTitleRowIndex = worksheet.rowCount;
    const skillTitleRow = worksheet.addRow(['Skill Data Records']);
    skillTitleRow.height = 30;
    skillTitleRow.font = { bold: true, size: 16, color: { argb: colors.titleText } };
    skillTitleRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.titleBg } };
    skillTitleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells(`A${skillTitleRowIndex}:D${skillTitleRowIndex}`);
    applyCellBorders(skillTitleRow, 1, 4, 'medium');
    
    // Skills table headers
    const skillHeaderRow = worksheet.addRow(['ID', 'Name', 'Description', 'Employees with this Skill']);
    skillHeaderRow.height = 20;
    skillHeaderRow.font = { bold: true, size: 12 };
    skillHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    skillHeaderRow.alignment = { horizontal: 'center', vertical: 'middle' };
    applyCellBorders(skillHeaderRow, 1, 4);
    
    // Add skills data
    appData.skills.forEach(skill => {
        const employeesWithSkill = appData.employees
            .filter(employee => (employee.skills || []).includes(skill.id))
            .map(employee => employee.name)
            .join(', ');
        
        const rowData = [
            skill.id,
            skill.name,
            skill.description || '',
            employeesWithSkill || 'None'
        ];
        
        const dataRow = worksheet.addRow(rowData);
        dataRow.height = 22;
        
        // Style all cells
        for (let i = 1; i <= 4; i++) {
            const cell = dataRow.getCell(i);
            cell.font = { size: 11 };
            cell.alignment = { 
                horizontal: i === 3 || i === 4 ? 'left' : (i === 1 ? 'left' : 'center'), 
                vertical: 'middle', 
                wrapText: true
            };
            cell.border = getBorderStyle();
        }
    });
}

// Helper function for cell borders
function getBorderStyle(style = 'thin') {
    return {
        top: { style, color: { argb: 'BFBFBF' } },
        left: { style, color: { argb: 'BFBFBF' } },
        bottom: { style, color: { argb: 'BFBFBF' } },
        right: { style, color: { argb: 'BFBFBF' } }
    };
}

// Helper function to apply borders to all cells in a row
function applyCellBorders(row, startCol, endCol, style = 'thin') {
    for (let i = startCol; i <= endCol; i++) {
        row.getCell(i).border = getBorderStyle(style);
    }
}

// Show information about new versioning and locking features
function showFeatureInfo() {
    // Check if we've already shown this info
    if (localStorage.getItem('versioning_info_shown')) {
        return;
    }
    
    const message = `
        <h5>New Features Added!</h5>
        <p>We've added two new powerful features to help with your shift planning:</p>
        <ul>
            <li><strong>Version Control</strong> - Create and switch between different versions of your weekly schedule</li>
            <li><strong>Schedule Locking</strong> - Lock a week's schedule to prevent further changes</li>
        </ul>
        <p>Look for the new version controls below the week selector and the lock/unlock button in the top right.</p>
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
                    <h5 class="modal-title">New Features</h5>
                    <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    ${message}
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-primary" data-bs-dismiss="modal">Got it!</button>
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
    if (!elements.skillsTable) return;
    
    const tbody = elements.skillsTable.querySelector('tbody');
    tbody.innerHTML = '';
    
    appData.skills.forEach(skill => {
        const row = document.createElement('tr');
        
        // Name column
        const nameCell = document.createElement('td');
        nameCell.textContent = skill.name;
        row.appendChild(nameCell);
        
        // Description column
        const descCell = document.createElement('td');
        descCell.textContent = skill.description || 'No description';
        row.appendChild(descCell);
        
        // Employees with this skill column
        const employeesCell = document.createElement('td');
        const employeesWithSkill = appData.employees.filter(e => e.skills && e.skills.includes(skill.id));
        
        if (employeesWithSkill.length > 0) {
            employeesWithSkill.forEach(employee => {
                const badge = document.createElement('span');
                badge.classList.add('badge', 'bg-secondary', 'me-1', 'mb-1');
                badge.textContent = employee.name;
                employeesCell.appendChild(badge);
            });
        } else {
            employeesCell.textContent = 'No employees have this skill';
        }
        row.appendChild(employeesCell);
        
        // Actions column
        const actionsCell = document.createElement('td');
        
        const editButton = document.createElement('button');
        editButton.classList.add('btn', 'btn-sm', 'btn-outline-primary', 'me-2');
        editButton.textContent = 'Edit';
        editButton.addEventListener('click', () => openSkillModal(skill));
        
        const deleteButton = document.createElement('button');
        deleteButton.classList.add('btn', 'btn-sm', 'btn-outline-danger');
        deleteButton.textContent = 'Delete';
        deleteButton.addEventListener('click', () => deleteSkill(skill.id));
        
        actionsCell.appendChild(editButton);
        actionsCell.appendChild(deleteButton);
        
        row.appendChild(actionsCell);
        
        tbody.appendChild(row);
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
        weekLabelWrapper.innerHTML = `<label class="me-2" style="margin-bottom: 0;"><strong>Week ${weekIndex + 1}:</strong></label>`;
        return;
    }
    
    // Get the Monday date for this week
    const mondayDate = new Date(weekDates[0]);
    
    // Calculate the ISO week number
    const weekNumber = getWeekNumber(mondayDate);
    
    // Set the content of the week label with the calendar week number
    weekLabelWrapper.innerHTML = `<label class="me-2" style="margin-bottom: 0;"><strong>Week ${weekNumber}:</strong></label>`;
}

// Initialize the application when the DOM is loaded
document.addEventListener('DOMContentLoaded', init); 