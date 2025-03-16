# Shift Planner Application

A single-page web application for planning employee work shifts by week.

## Features

- **Three Shifts Per Day**: AM-Shift, PM-Shift, and Night-Shift
- **Team-Based Scheduling**: Teams A and B alternate between AM and PM shifts weekly, Team C always works Night shifts
- **Employee Management**: Add, edit, and delete employees, and assign them to teams
- **Absence Tracking**: Track employee absences (holidays, sick days, etc.)
- **Automatic Scheduling**: Automatically plan the next four weeks of shifts
- **Manual Adjustments**: Ability to manually adjust employee assignments for specific shifts
- **Excel Export**: Export the schedule to Excel with formatted reports

## How to Use

### Getting Started

1. Open `index.html` in a web browser
2. The application will start with sample data that you can modify

### Schedule View

- View the weekly schedule with all shifts
- Use the "Select Week" dropdown to navigate between different weeks
- Click on any shift cell to edit the employee assignments for that shift
- Use the "Generate 4-Week Plan" button to create a fresh schedule for the next four weeks
- Use the "Export to Excel" button to export the schedule to an Excel file

### Employees View

- View, add, edit, and delete employees
- Set employee team assignments
- Manage employee absence dates
- When an employee is marked as absent on a specific date, they will not be available for scheduling on that date

### Teams View

- View the distribution of employees across teams
- See how many absence days each employee has

## Technical Details

- The application uses **localStorage** to store data between sessions
- No server or backend is required - all data is stored in the browser
- The application uses the following libraries:
  - Bootstrap 5 for UI components
  - SheetJS (xlsx) for Excel export
  - FileSaver.js for downloading files

## Browser Compatibility

The application is compatible with all modern browsers:
- Chrome
- Firefox
- Safari
- Edge

## Note

This is a client-side only application. All data is stored in your browser's localStorage and will be lost if you clear your browser data. For production use, consider implementing a server-side component for data persistence. 