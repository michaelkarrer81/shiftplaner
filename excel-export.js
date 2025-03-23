// Excel Export Functions
// This file contains all functions related to Excel export functionality

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
        const weekTitle = currentLanguage === 'de' 
            ? `Woche ${weekNumber}: ${formatDate(startDate)} - ${formatDate(endDate)}`
            : `Week ${weekNumber}: ${formatDate(startDate)} - ${formatDate(endDate)}`;
        const versionInfo = activeVersion ? activeVersion.name : (currentLanguage === 'de' ? 'Standard' : 'Default');
        const statusInfo = isLocked ? 'LOCKED 🔒' : (currentLanguage === 'de' ? 'BEARBEITBAR' : 'EDITABLE');
        
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
        
        // Define translations for various terms
        const translations = {
            shifts: {
                AM: currentLanguage === 'de' ? 'Frühschicht' : 'AM',
                PM: currentLanguage === 'de' ? 'Spätschicht' : 'PM',
                Night: currentLanguage === 'de' ? 'Nachtschicht' : 'Night'
            },
            team: currentLanguage === 'de' ? 'Team' : 'Team',
            absent: currentLanguage === 'de' ? 'ABWESEND' : 'ABSENT',
            skills: currentLanguage === 'de' ? 'Fähigkeiten' : 'Skills',
            noEmployees: currentLanguage === 'de' ? 'Keine Mitarbeiter' : 'No employees'
        };
        
        // 1. Create Weekly Summary Sheet
        createWeeklySummarySheet(workbook, weekIndex, weekTitle, versionInfo, statusInfo, colors, translations);
        
        // 2. Create Detailed Schedule Sheet
        createDetailedScheduleSheet(workbook, weekIndex, weekTitle, versionInfo, statusInfo, colors, translations);
        
        // 3. Create Data Records Sheet
        createDataRecordsSheet(workbook, weekIndex, colors, translations);
        
        // Generate filename
        const formattedDate = formatDateForFilename(new Date());
        const filename = currentLanguage === 'de' 
            ? `Schichtplaner_Export_${formattedDate}.xlsx`
            : `ShiftPlanner_Export_${formattedDate}.xlsx`;
        
        // Generate the Excel file and trigger download
        workbook.xlsx.writeBuffer().then(buffer => {
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            saveAs(blob, filename);
            
            // Show success message
            const successMsg = currentLanguage === 'de' 
                ? "Excel-Export erfolgreich!"
                : "Excel export successful!";
            showProgressIndicator(successMsg, 2000);
        });
    } catch (error) {
        console.error('Error during Excel export:', error);
        showError('An error occurred during Excel export. Please try again.');
    }
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

// Format date for filename (yyyy-mm-dd)
function formatDateForFilename(date) {
    return `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}-${date.getDate().toString().padStart(2, '0')}`;
}

// Create the Weekly Summary Sheet with ExcelJS
function createWeeklySummarySheet(workbook, weekIndex, weekTitle, versionInfo, statusInfo, colors, translations) {
    const weekSchedule = appData.schedule[weekIndex] || {};
    const weekDates = appData.weekDates[weekIndex] || [];
    
    // Create sheet
    const worksheet = workbook.addWorksheet(currentLanguage === 'de' ? 'Wöchentliche Übersicht' : 'Weekly Summary', {
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
    const statusText = currentLanguage === 'de' ? 'Status' : 'Status';
    const versionText = currentLanguage === 'de' ? 'Version' : 'Version';
    const statusRow = worksheet.addRow([`${statusText}: ${statusInfo} | ${versionText}: ${versionInfo}`]);
    statusRow.height = 24;
    statusRow.font = { italic: true, size: 11 };
    statusRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.versionInfoBg } };
    statusRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells('A2:D2');
    applyCellBorders(statusRow, 1, 4);
    
    // Empty row for spacing
    worksheet.addRow([]);
    
    // Schedule header
    const scheduleHeader = worksheet.addRow([currentLanguage === 'de' ? 'Wöchentlicher Zeitplan' : 'Weekly Schedule']);
    scheduleHeader.height = 25;
    scheduleHeader.font = { bold: true, size: 14 };
    scheduleHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    scheduleHeader.alignment = { horizontal: 'left', vertical: 'middle' };
    worksheet.mergeCells('A4:D4');
    applyCellBorders(scheduleHeader, 1, 4);
    
    // Schedule table headers
    const scheduleTableHeaders = currentLanguage === 'de'
        ? ['Tag/Datum', 'Frühschicht', 'Spätschicht', 'Nachtschicht']
        : ['Day/Date', 'AM Shift', 'PM Shift', 'Night Shift'];
    
    const shiftHeaderRow = worksheet.addRow(scheduleTableHeaders);
    shiftHeaderRow.height = 20;
    shiftHeaderRow.font = { bold: true, size: 12 };
    shiftHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.subHeaderBg } };
    shiftHeaderRow.alignment = { horizontal: 'center', vertical: 'middle' };
    applyCellBorders(shiftHeaderRow, 1, 4);
    
    // German weekday translations
    const germanWeekdays = {
        'Monday': 'Montag',
        'Tuesday': 'Dienstag',
        'Wednesday': 'Mittwoch',
        'Thursday': 'Donnerstag',
        'Friday': 'Freitag',
        'Saturday': 'Samstag',
        'Sunday': 'Sonntag'
    };
    
    // Add data for each day
    DAYS_OF_WEEK.forEach((day, dayIndex) => {
        const daySchedule = weekSchedule[day] || {};
        const dateStr = weekDates[dayIndex] || '';
        const date = dateStr ? new Date(dateStr) : null;
        const dateFormatted = date ? formatDate(date) : '';
        
        // Translate day name if language is German
        const displayDay = currentLanguage === 'de' ? germanWeekdays[day] || day : day;
        const rowData = [`${displayDay} (${dateFormatted})`];
        
        // For each shift type (AM, PM, Night)
        SHIFT_TYPES.forEach(shiftType => {
            const shift = daySchedule[shiftType] || { team: '-', employees: [] };
            const teamCode = shift.team;
            
            // Format team header text - keep this with the original styling
            let cellContent = `${translations.shifts[shiftType]} ${translations.team} ${teamCode}\n`;
            
            // Add employee names to each shift
            if (shift.employees && shift.employees.length > 0) {
                const employeeList = shift.employees.map(empId => {
                    const employee = appData.employees.find(e => e.id === empId);
                    if (!employee) return '';
                    
                    const isAbsent = employee.absentDates.includes(dateStr);
                    const differentTeam = employee.team !== shift.team;
                    
                    let employeeText = employee.name;
                    
                    if (isAbsent) {
                        employeeText += ` (${translations.absent})`;
                    }
                    
                    if (differentTeam) {
                        employeeText += ` (${translations.team} ${employee.team})`;
                    }
                    
                    return employeeText;
                }).filter(Boolean);
                
                cellContent += employeeList.join('\n');
            } else {
                cellContent += translations.noEmployees;
            }
            
            rowData.push(cellContent);
        });
        
        const dataRow = worksheet.addRow(rowData);
        
        // Count the total number of employees across all shifts
        let totalEmployeesInDay = 0;
        SHIFT_TYPES.forEach(shiftType => {
            const shift = daySchedule[shiftType] || { employees: [] };
            totalEmployeesInDay += shift.employees ? shift.employees.length : 0;
        });
        
        // Use a more compact height calculation
        const baseHeight = 25;  // Reduced height for headers
        const heightPerEmployee = 14; // More compact height per employee
        const calculatedHeight = baseHeight + (heightPerEmployee * totalEmployeesInDay);
        
        // Set with a reasonable minimum and maximum
        dataRow.height = Math.max(35, Math.min(250, calculatedHeight));
        
        // Style the day cell (first column)
        dataRow.getCell(1).font = { size: 11 };
        dataRow.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF' } };
        dataRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
        dataRow.getCell(1).border = getBorderStyle();
        
        // Style shift cells based on team
        for (let i = 0; i < SHIFT_TYPES.length; i++) {
            const cell = dataRow.getCell(i + 2);
            const teamMatch = rowData[i + 1].match(/Team ([A-Z])/);
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
            
            // Apply rich text formatting to make just the team header bold and colored
            // while keeping employee names smaller and black
            const shiftText = rowData[i + 1];
            const headerEndIndex = shiftText.indexOf('\n');
            
            if (headerEndIndex !== -1) {
                const header = shiftText.substring(0, headerEndIndex + 1);
                const employees = shiftText.substring(headerEndIndex + 1);
                
                cell.value = {
                    richText: [
                        { 
                            text: header, 
                            font: { size: 11, color: { argb: textColor }, bold: true } 
                        },
                        { 
                            text: employees, 
                            font: { size: 9, color: { argb: '000000' } } 
                        }
                    ]
                };
            } else {
                cell.font = { size: 11, color: { argb: textColor } };
                cell.value = shiftText;
            }
            
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
            cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
            cell.border = getBorderStyle();
        }
    });
    
    // Add spacing
    worksheet.addRow([]);
    worksheet.addRow([]);
    
    // ===== SKILL COVERAGE SECTION =====
    
    // Skill Coverage header
    const skillRowIndex = worksheet.rowCount;
    const skillHeader = worksheet.addRow([currentLanguage === 'de' ? 'Wöchentliche Fähigkeiten-Abdeckungsübersicht' : 'Weekly Skill Coverage Summary']);
    skillHeader.height = 25;
    skillHeader.font = { bold: true, size: 14 };
    skillHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    skillHeader.alignment = { horizontal: 'left', vertical: 'middle' };
    worksheet.mergeCells(`A${skillRowIndex}:D${skillRowIndex}`);
    applyCellBorders(skillHeader, 1, 4);
    
    // Skill table headers
    const skillTableHeaders = currentLanguage === 'de' 
        ? ['Fähigkeit', 'Frühschicht-Abdeckung', 'Spätschicht-Abdeckung', 'Nachtschicht-Abdeckung']
        : ['Skill', 'AM Shift Coverage', 'PM Shift Coverage', 'Night Shift Coverage'];
    
    const skillHeaderRow = worksheet.addRow(skillTableHeaders);
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
        const employeesText = currentLanguage === 'de' ? 'Mitarbeiter' : 'employees';
        const rowData = [
            skillData.name,
            `${skillData.shifts.AM} ${employeesText}`,
            `${skillData.shifts.PM} ${employeesText}`,
            `${skillData.shifts.Night} ${employeesText}`
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
    const exceptionHeader = worksheet.addRow([currentLanguage === 'de' ? 'Wöchentliche Ausnahmen' : 'Weekly Exceptions']);
    exceptionHeader.height = 25;
    exceptionHeader.font = { bold: true, size: 14 };
    exceptionHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    exceptionHeader.alignment = { horizontal: 'left', vertical: 'middle' };
    worksheet.mergeCells(`A${exceptionRowIndex}:D${exceptionRowIndex}`);
    applyCellBorders(exceptionHeader, 1, 4);
    
    // Exceptions table headers
    const exceptionTableHeaders = currentLanguage === 'de'
        ? ['Mitarbeiter', 'Abwesenheitsdatum', 'Team', 'Notiz']
        : ['Employee', 'Absent Date', 'Team', 'Note'];
    
    const absenceHeaderRow = worksheet.addRow(exceptionTableHeaders);
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
function createDetailedScheduleSheet(workbook, weekIndex, weekTitle, versionInfo, statusInfo, colors, translations) {
    const weekSchedule = appData.schedule[weekIndex] || {};
    const weekDates = appData.weekDates[weekIndex] || [];
    
    // Create sheet
    const worksheet = workbook.addWorksheet(currentLanguage === 'de' ? 'Detaillierter Zeitplan' : 'Detailed Schedule', {
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
    const detailedHeader = worksheet.addRow([currentLanguage === 'de' ? 'Detaillierte Zeitplanansicht' : 'Detailed Schedule View']);
    detailedHeader.height = 25;
    detailedHeader.font = { bold: true, size: 14 };
    detailedHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.headerBg } };
    detailedHeader.alignment = { horizontal: 'left', vertical: 'middle' };
    worksheet.mergeCells('A4:D4');
    applyCellBorders(detailedHeader, 1, 4);
    
    // Schedule table headers
    const scheduleTableHeaders = currentLanguage === 'de'
        ? ['Tag/Datum', 'Frühschicht', 'Spätschicht', 'Nachtschicht']
        : ['Day/Date', 'AM Shift', 'PM Shift', 'Night Shift'];
    
    const headerRow = worksheet.addRow(scheduleTableHeaders);
    headerRow.height = 20;
    headerRow.font = { bold: true, size: 12 };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.subHeaderBg } };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
    applyCellBorders(headerRow, 1, 4);
    
    // German weekday translations
    const germanWeekdays = {
        'Monday': 'Montag',
        'Tuesday': 'Dienstag',
        'Wednesday': 'Mittwoch',
        'Thursday': 'Donnerstag',
        'Friday': 'Freitag',
        'Saturday': 'Samstag',
        'Sunday': 'Sonntag'
    };
    
    // Add data for each day
    DAYS_OF_WEEK.forEach((day, dayIndex) => {
        const daySchedule = weekSchedule[day] || {};
        const dateStr = weekDates[dayIndex] || '';
        const date = dateStr ? new Date(dateStr) : null;
        const dateFormatted = date ? formatDate(date) : '';
        
        // Translate day name if language is German
        const displayDay = currentLanguage === 'de' ? germanWeekdays[day] || day : day;
        const rowData = [`${displayDay} (${dateFormatted})`];
        const employeeDetails = [];
        
        SHIFT_TYPES.forEach(shiftType => {
            const shift = daySchedule[shiftType] || { team: '-', employees: [] };
            let cellContent = `${translations.shifts[shiftType]} ${translations.team} ${shift.team}\n`;
            let employeeList = [];
            
            if (shift.employees && shift.employees.length > 0) {
                employeeList = shift.employees.map(empId => {
                    const employee = appData.employees.find(e => e.id === empId);
                    if (!employee) return null;
                    
                    const isAbsent = employee.absentDates.includes(dateStr);
                    const differentTeam = employee.team !== shift.team;
                    
                    let employeeText = employee.name;
                    
                    if (isAbsent) {
                        employeeText += ` (${translations.absent})`;
                    }
                    
                    if (differentTeam) {
                        employeeText += ` (${translations.team} ${employee.team})`;
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
        
        // Count the total number of employees across all shifts
        let totalEmployeesInDay = 0;
        SHIFT_TYPES.forEach(shiftType => {
            const shift = daySchedule[shiftType] || { employees: [] };
            totalEmployeesInDay += shift.employees ? shift.employees.length : 0;
        });
        
        // Use a more compact height calculation
        const baseHeight = 30;  // Reduced height for headers
        const heightPerEmployee = 16; // More compact height per employee
        const calculatedHeight = baseHeight + (heightPerEmployee * totalEmployeesInDay);
        
        // Set with a reasonable minimum and maximum
        dataRow.height = Math.max(45, Math.min(300, calculatedHeight));
        
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
            
            // Apply rich text formatting to make just the team header bold and colored
            // while keeping employee names smaller and black
            const shiftText = rowData[i + 1];
            const headerEndIndex = shiftText.indexOf('\n');
            
            if (headerEndIndex !== -1) {
                const header = shiftText.substring(0, headerEndIndex + 1);
                const employees = shiftText.substring(headerEndIndex + 1);
                
                cell.value = {
                    richText: [
                        { 
                            text: header, 
                            font: { size: 11, bold: true } 
                        },
                        { 
                            text: employees, 
                            font: { size: 9, color: { argb: '000000' } } 
                        }
                    ]
                };
            } else {
                cell.font = { size: 11 };
                cell.value = shiftText;
            }
            
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
            cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
            cell.border = getBorderStyle();
        }
    });
}

// Create the Data Records Sheet with ExcelJS
function createDataRecordsSheet(workbook, weekIndex, colors, translations) {
    // Create sheet
    const worksheet = workbook.addWorksheet(currentLanguage === 'de' ? 'Datensätze' : 'Data Records', {
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
    const titleRow = worksheet.addRow([currentLanguage === 'de' ? 'Mitarbeiterdatensätze' : 'Employee Data Records']);
    titleRow.height = 30;
    titleRow.font = { bold: true, size: 16, color: { argb: colors.titleText } };
    titleRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.titleBg } };
    titleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells('A1:E1');
    applyCellBorders(titleRow, 1, 5, 'medium');
    
    // Empty row for spacing
    worksheet.addRow([]);
    
    // Employee table headers
    const employeeTableHeaders = currentLanguage === 'de'
        ? ['ID', 'Name', 'Team', 'Fähigkeiten', 'Abwesenheitsdaten']
        : ['ID', 'Name', 'Team', 'Skills', 'Absent Dates'];
        
    const headerRow = worksheet.addRow(employeeTableHeaders);
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
        
        const noneText = currentLanguage === 'de' ? 'Keine' : 'None';
        
        const rowData = [
            employee.id,
            employee.name,
            `Team ${employee.team}`,
            skillNames || noneText,
            absentDatesFormatted || noneText
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
        if (absentDatesFormatted !== noneText) {
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
    const skillTitleRow = worksheet.addRow([currentLanguage === 'de' ? 'Fähigkeitendatensätze' : 'Skill Data Records']);
    skillTitleRow.height = 30;
    skillTitleRow.font = { bold: true, size: 16, color: { argb: colors.titleText } };
    skillTitleRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.titleBg } };
    skillTitleRow.alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.mergeCells(`A${skillTitleRowIndex}:D${skillTitleRowIndex}`);
    applyCellBorders(skillTitleRow, 1, 4, 'medium');
    
    // Skills table headers
    const skillTableHeaders = currentLanguage === 'de'
        ? ['ID', 'Name', 'Beschreibung', 'Mitarbeiter mit dieser Fähigkeit']
        : ['ID', 'Name', 'Description', 'Employees with this Skill'];
        
    const skillHeaderRow = worksheet.addRow(skillTableHeaders);
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
        
        const noneText = currentLanguage === 'de' ? 'Keine' : 'None';
        
        const rowData = [
            skill.id,
            skill.name,
            skill.description || '',
            employeesWithSkill || noneText
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
