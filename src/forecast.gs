// ============================
// TASK COMPLETION FORECASTING CONFIGURATION
// ============================

const FORECAST_CONFIG = {
  // Master sheet name where forecasting dashboard will be created
  MASTER_SHEET_NAME: 'Task Completion Forecast',
  
  // Data source sheet name
  DATA_SHEET_NAME: 'Jira_Data',
  
  // External Flex team sheet configuration
  FLEX_SHEET_ID: 'GOOGLE SHEET ID HERE', 
  FLEX_SHEET_NAME: 'Tasks', 
  
  // Assignee exclusion list - these assignees will be excluded from statistics
  EXCLUDED_ASSIGNEES: [
    'John Doe',
    'Babar Azam'
  ],
  
  // Forecasting parameters
  TASKS_PER_DEV_PER_DAY: 3,
  WORKING_DAYS_PER_WEEK: 5,
  
  // Team configuration
  TEAMS: {
    'Team1': {
      name: 'Team1',
      leadName: 'Lead1',
      developers: 5,
      language: 'Python',
      teamCode: 'PY-Team1'
    },
    'Team2': {
      name: 'Team2',
      leadName: 'Lead2',
      developers: 8,
      language: 'Python',
      teamCode: 'PY-Team2'
    },
    'Team3': {
      name: 'Team3',
      leadName: 'Lead3',
      developers: 11,
      language: 'Rust',
      teamCode: 'RS-Team3'
    },
    'Team4': {
      name: 'Team4',
      leadName: 'Lead4',
      developers: 11,
      language: 'Rust',
      teamCode: 'RS-Team4'
    },
    'Team5': {
      name: 'Team5',
      leadName: 'Lead5',
      developers: 7,
      language: 'JavaScript',
      teamCode: 'JS-Team5'
    },
    'Team6': {
      name: 'Team6',
      leadName: 'Lead6',
      developers: 5,
      language: 'JavaScript',
      teamCode: 'JS-Team6'
    },
    'Flex': {
      name: 'Flex',
      leadName: 'Lead7',
      developers: 0, // Will be calculated dynamically
      language: 'Team6-Language',
      teamCode: 'FLEX'
    }
  },
  
  // Column mappings for Jira_Data tab
  KEY_COLUMN: 'Key',
  ASSIGNEE_COLUMN: 'Assignee',
  TEAM_COLUMN: 'Team',
  STATUS_COLUMN: 'Status',
  SUMMARY_COLUMN: 'Summary',
  CREATED_COLUMN: 'Created',
  UPDATED_COLUMN: 'Updated',
  
  // Column mappings for Flex team data
  FLEX_COLUMNS: {
    TASK_ID: 'task_id',
    WORKER_ID: 'worker_id',
    LANGUAGE: 'Language',
    TASK_START_DATE: 'task_start_date',
    TASK_COMPLETION_STATUS: 'task_completion_status',
    REPO_APPROVAL: 'repo_approval'
  },
  
  // Status values for active tasks (not completed)
  ACTIVE_STATUS_VALUES: {
    JIRA: ['In Progress', 'Repo Approved', 'Interaction in progress', 'Peer review', 'Lead review', 'Blocked', 'To Do', 'Rework Required'],
    FLEX: ['In progress', 'Ready for Review', 'Review in progress', 'Changes Requested', 'Rejected', 'Rework']
  },
  
  // Team to language mapping
  TEAM_LANGUAGE_MAP: {
    'Team1': 'Python',
    'Team2': 'Python',
    'storm': 'Python',
    'Titans': 'Python',
    'Cyborg': 'Python',
    'Team3': 'Rust',
    'Team4': 'Rust',
    'Team5': 'JavaScript',
    'Team6': 'JavaScript',
    'Apex': 'JavaScript',
    'Astral': 'JavaScript',
    'Flex': 'Team6-Language',
    'Vega': 'JavaScript'
  },
  
  // Colors for dashboard
  COLORS: {
    HEADER: '#1a73e8',
    PYTHON: '#3776ab',
    RUST: '#ce422b',
    JAVASCRIPT: '#f7df1e',
    FLEX: '#9c27b0',
    LIGHT_GRAY: '#f8f9fa',
    BORDER: '#dadce0',
    FORECAST_HEADER: '#4285f4',
    WARNING: '#ff9800',
    SUCCESS: '#4caf50',
    DANGER: '#f44336'
  }
};

// ============================
// HELPER FUNCTIONS
// ============================

/**
 * Helper function to check if an assignee should be excluded
 */
function shouldExcludeAssignee(assignee) {
  if (!assignee) return false;
  
  const normalizedAssignee = assignee.toString().trim();
  
  return FORECAST_CONFIG.EXCLUDED_ASSIGNEES.some(excludedName => {
    return normalizedAssignee.toLowerCase().includes(excludedName.toLowerCase()) ||
           excludedName.toLowerCase().includes(normalizedAssignee.toLowerCase());
  });
}

/**
 * Helper function to find column index
 */
function findColumnIndex(headers, columnName) {
  return headers.findIndex(h => h.toString().toLowerCase().trim() === columnName.toLowerCase().trim());
}

/**
 * Helper function to check if task is active (not completed)
 */
function isActiveTask(status, isFlexTeam = false) {
  const activeStatuses = isFlexTeam ? FORECAST_CONFIG.ACTIVE_STATUS_VALUES.FLEX : FORECAST_CONFIG.ACTIVE_STATUS_VALUES.JIRA;
  return activeStatuses.includes(status.toString().trim());
}

/**
 * Helper function to add working days to a date
 */
function addWorkingDays(startDate, workingDays) {
  const result = new Date(startDate);
  let daysAdded = 0;
  
  while (daysAdded < workingDays) {
    result.setDate(result.getDate() + 1);
    // Skip weekends (0 = Sunday, 6 = Saturday)
    if (result.getDay() !== 0 && result.getDay() !== 6) {
      daysAdded++;
    }
  }
  
  return result;
}

// ============================
// DATA COLLECTION FUNCTIONS
// ============================

/**
 * Main function to create the forecasting dashboard
 */
function createTaskForecastingDashboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let forecastSheet = ss.getSheetByName(FORECAST_CONFIG.MASTER_SHEET_NAME);
    
    // Create sheet if it doesn't exist
    if (!forecastSheet) {
      forecastSheet = ss.insertSheet(FORECAST_CONFIG.MASTER_SHEET_NAME);
    }
    
    // Clear existing content
    forecastSheet.clear();
    
    // Collect and analyze data
    const forecastData = collectForecastData();
    
    // Build forecast dashboard
    buildForecastDashboard(forecastSheet, forecastData);
    
    // Show completion message
    SpreadsheetApp.getUi().alert(
      'Forecast Dashboard Created',
      `Task Completion Forecasting Dashboard created successfully!\n\n` +
      `Total active tasks analyzed: ${forecastData.totalActiveTasks}\n` +
      `Languages covered: ${Object.keys(forecastData.languageForecasts).join(', ')}\n` +
      `Teams analyzed: ${forecastData.teamForecasts.length}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error creating forecast dashboard:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Collects all forecast data from Jira and Flex sources
 */
function collectForecastData() {
  const jiraData = collectJiraForecastData();
  const flexData = collectFlexForecastData();
  
  // Combine data
  const allTasks = [...jiraData.tasks, ...flexData.tasks];
  
  // Calculate team-based forecasts
  const teamForecasts = calculateTeamForecasts(allTasks);
  
  // Calculate language-based forecasts
  const languageForecasts = calculateLanguageForecasts(teamForecasts);
  
  // Calculate assignee workloads
  const assigneeWorkloads = calculateAssigneeWorkloads(allTasks);
  
  return {
    totalActiveTasks: allTasks.length,
    jiraTasks: jiraData.tasks.length,
    flexTasks: flexData.tasks.length,
    excludedTasks: jiraData.excludedTasks + flexData.excludedTasks,
    teamForecasts: teamForecasts,
    languageForecasts: languageForecasts,
    assigneeWorkloads: assigneeWorkloads,
    generatedAt: new Date()
  };
}

/**
 * Collects forecast data from Jira_Data sheet
 */
function collectJiraForecastData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(FORECAST_CONFIG.DATA_SHEET_NAME);
  
  if (!dataSheet) {
    throw new Error(`Sheet '${FORECAST_CONFIG.DATA_SHEET_NAME}' not found.`);
  }
  
  const data = dataSheet.getDataRange().getValues();
  if (data.length === 0) {
    throw new Error('No data found in Jira_Data sheet.');
  }
  
  const headers = data[0];
  const columnMap = createColumnMap(headers);
  
  const tasks = [];
  let excludedTasks = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[columnMap.KEY_COLUMN] || row[columnMap.ASSIGNEE_COLUMN]) {
      const assignee = row[columnMap.ASSIGNEE_COLUMN] ? row[columnMap.ASSIGNEE_COLUMN].toString().trim() : '';
      const status = row[columnMap.STATUS_COLUMN] ? row[columnMap.STATUS_COLUMN].toString().trim() : '';
      const team = row[columnMap.TEAM_COLUMN] ? row[columnMap.TEAM_COLUMN].toString().trim() : '';
      
      // Skip if assignee should be excluded
      if (shouldExcludeAssignee(assignee)) {
        excludedTasks++;
        continue;
      }
      
      // Only include active tasks
      if (isActiveTask(status, false)) {
        tasks.push({
          key: row[columnMap.KEY_COLUMN] || '',
          assignee: assignee,
          team: team,
          status: status,
          language: FORECAST_CONFIG.TEAM_LANGUAGE_MAP[team] || 'Unknown',
          summary: row[columnMap.SUMMARY_COLUMN] || '',
          created: row[columnMap.CREATED_COLUMN] || '',
          updated: row[columnMap.UPDATED_COLUMN] || '',
          source: 'Jira',
          isFlexTeam: false
        });
      }
    }
  }
  
  return {
    tasks: tasks,
    excludedTasks: excludedTasks
  };
}

/**
 * Collects forecast data from Flex team external sheet
 */
function collectFlexForecastData() {
  try {
    const flexSheet = SpreadsheetApp.openById(FORECAST_CONFIG.FLEX_SHEET_ID);
    const dataSheet = flexSheet.getSheetByName(FORECAST_CONFIG.FLEX_SHEET_NAME);
    
    if (!dataSheet) {
      console.warn('Flex sheet not found, skipping Flex data');
      return { tasks: [], excludedTasks: 0 };
    }
    
    const data = dataSheet.getDataRange().getValues();
    if (data.length === 0) {
      return { tasks: [], excludedTasks: 0 };
    }
    
    const headers = data[0];
    const flexColumnMap = createFlexColumnMap(headers);
    
    const tasks = [];
    let excludedTasks = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[flexColumnMap.TASK_ID] || row[flexColumnMap.WORKER_ID]) {
        const assignee = row[flexColumnMap.WORKER_ID] ? row[flexColumnMap.WORKER_ID].toString().trim() : '';
        const taskCompletionStatus = row[flexColumnMap.TASK_COMPLETION_STATUS] ? row[flexColumnMap.TASK_COMPLETION_STATUS].toString().trim() : '';
        const repoApproval = row[flexColumnMap.REPO_APPROVAL] ? row[flexColumnMap.REPO_APPROVAL].toString().trim() : '';
        
        // Determine effective status
        const effectiveStatus = taskCompletionStatus || repoApproval;
        
        // Skip if assignee should be excluded
        if (shouldExcludeAssignee(assignee)) {
          excludedTasks++;
          continue;
        }
        
        // Only include active tasks
        if (isActiveTask(effectiveStatus, true)) {
          tasks.push({
            key: row[flexColumnMap.TASK_ID] || '',
            assignee: assignee,
            team: 'Flex',
            status: effectiveStatus,
            language: row[flexColumnMap.LANGUAGE] || 'Team6-Language',
            summary: row[flexColumnMap.TASK_ID] || '',
            created: row[flexColumnMap.TASK_START_DATE] || '',
            updated: '',
            source: 'Flex',
            isFlexTeam: true
          });
        }
      }
    }
    
    return {
      tasks: tasks,
      excludedTasks: excludedTasks
    };
    
  } catch (error) {
    console.error('Error reading Flex data:', error);
    return { tasks: [], excludedTasks: 0 };
  }
}

/**
 * Creates column mapping from headers
 */
function createColumnMap(headers) {
  const columnMap = {};
  
  Object.keys(FORECAST_CONFIG).forEach(key => {
    if (key.endsWith('_COLUMN')) {
      const columnName = FORECAST_CONFIG[key];
      const index = findColumnIndex(headers, columnName);
      columnMap[key] = index;
    }
  });
  
  return columnMap;
}

/**
 * Creates column mapping for Flex team headers
 */
function createFlexColumnMap(headers) {
  const columnMap = {};
  
  Object.keys(FORECAST_CONFIG.FLEX_COLUMNS).forEach(key => {
    const columnName = FORECAST_CONFIG.FLEX_COLUMNS[key];
    const index = findColumnIndex(headers, columnName);
    columnMap[key] = index;
  });
  
  return columnMap;
}

// ============================
// FORECAST CALCULATION FUNCTIONS
// ============================

/**
 * Calculates forecasts for each team
 */
function calculateTeamForecasts(allTasks) {
  const teamForecasts = [];
  
  // Group tasks by team
  const tasksByTeam = {};
  allTasks.forEach(task => {
    if (!tasksByTeam[task.team]) {
      tasksByTeam[task.team] = [];
    }
    tasksByTeam[task.team].push(task);
  });
  
  // Calculate forecast for each team
  Object.keys(tasksByTeam).forEach(teamName => {
    const teamTasks = tasksByTeam[teamName];
    const teamConfig = FORECAST_CONFIG.TEAMS[teamName];
    
    if (!teamConfig) {
      console.warn(`Team configuration not found for: ${teamName}`);
      return;
    }
    
    // Calculate assignee-specific workloads within team
    const assigneeTaskCounts = {};
    teamTasks.forEach(task => {
      if (!assigneeTaskCounts[task.assignee]) {
        assigneeTaskCounts[task.assignee] = 0;
      }
      assigneeTaskCounts[task.assignee]++;
    });
    
    const uniqueAssignees = Object.keys(assigneeTaskCounts);
    const actualDevCount = teamName === 'Flex' ? uniqueAssignees.length : teamConfig.developers;
    
    // Calculate individual completion dates for each assignee
    const assigneeForecasts = [];
    let teamCompletionDate = new Date();
    
    uniqueAssignees.forEach(assignee => {
      const taskCount = assigneeTaskCounts[assignee];
      const workingDaysNeeded = Math.ceil(taskCount / FORECAST_CONFIG.TASKS_PER_DEV_PER_DAY);
      const completionDate = addWorkingDays(new Date(), workingDaysNeeded);
      
      assigneeForecasts.push({
        assignee: assignee,
        taskCount: taskCount,
        workingDaysNeeded: workingDaysNeeded,
        completionDate: completionDate
      });
      
      // Team completion date is when the last assignee finishes
      if (completionDate > teamCompletionDate) {
        teamCompletionDate = completionDate;
      }
    });
    
    // Calculate team-level metrics
    const totalTasks = teamTasks.length;
    const averageTasksPerDev = totalTasks / actualDevCount;
    const averageWorkingDays = Math.ceil(averageTasksPerDev / FORECAST_CONFIG.TASKS_PER_DEV_PER_DAY);

    // START: New logic for 'To Do' task projection
    const toDoTasks = teamTasks.filter(task => task.status === 'To Do');
    const totalToDoTasks = toDoTasks.length;
    let projectedTeamCompletionDate = teamCompletionDate; // Default to current

    if (totalToDoTasks > 0 && uniqueAssignees.length > 1) {
        const targetWorkload = totalTasks / uniqueAssignees.length;
        
        const deficits = uniqueAssignees.map(assignee => {
            const nonToDoCount = teamTasks.filter(t => t.assignee === assignee && t.status !== 'To Do').length;
            return { assignee, nonToDoCount, deficit: targetWorkload - nonToDoCount };
        });

        const positiveDeficits = deficits.filter(d => d.deficit > 0);
        const totalPositiveDeficit = positiveDeficits.reduce((sum, d) => sum + d.deficit, 0);
        
        const projectedToDoCounts = {};

        if (totalPositiveDeficit > 0) {
            let distributedToDoTasks = 0;
            positiveDeficits.forEach((d, index) => {
                const share = d.deficit / totalPositiveDeficit;
                let assignedToDo = Math.round(share * totalToDoTasks);
                
                if (index === positiveDeficits.length - 1) {
                    assignedToDo = totalToDoTasks - distributedToDoTasks;
                }
                projectedToDoCounts[d.assignee] = assignedToDo;
                distributedToDoTasks += assignedToDo;
            });
        }

        assigneeForecasts.forEach(assignee => {
            const currentToDoCount = teamTasks.filter(t => t.assignee === assignee.assignee && t.status === 'To Do').length;
            const nonToDoCount = assignee.taskCount - currentToDoCount;
            
            const projectedToDoCount = projectedToDoCounts[assignee.assignee] || 0;
            const projectedTaskCount = nonToDoCount + projectedToDoCount;
            const projectedWorkingDays = Math.ceil(projectedTaskCount / FORECAST_CONFIG.TASKS_PER_DEV_PER_DAY);
            const projectedCompletionDate = addWorkingDays(new Date(), projectedWorkingDays);

            assignee.currentToDoCount = currentToDoCount;
            assignee.projectedTaskCount = projectedTaskCount;
            assignee.reassignment = projectedToDoCount - currentToDoCount;
            assignee.projectedCompletionDate = projectedCompletionDate;
            assignee.timeSaved = assignee.workingDaysNeeded - projectedWorkingDays;
        });

        const validDates = assigneeForecasts.map(a => a.projectedCompletionDate).filter(d => d);
        if (validDates.length > 0) {
            projectedTeamCompletionDate = new Date(Math.max(...validDates));
        }

    } else {
        assigneeForecasts.forEach(assignee => {
            const currentToDoCount = teamTasks.filter(t => t.assignee === assignee.assignee && t.status === 'To Do').length;
            assignee.currentToDoCount = currentToDoCount;
            assignee.projectedTaskCount = null;
            assignee.reassignment = null;
            assignee.projectedCompletionDate = null;
            assignee.timeSaved = null;
        });
    }
    // END: New logic for 'To Do' task projection
    
    teamForecasts.push({
      teamName: teamName,
      language: teamConfig.language,
      leadName: teamConfig.leadName,
      configuredDevCount: teamConfig.developers,
      actualDevCount: actualDevCount,
      totalTasks: totalTasks,
      averageTasksPerDev: averageTasksPerDev,
      averageWorkingDays: averageWorkingDays,
      teamCompletionDate: teamCompletionDate,
      projectedTeamCompletionDate: projectedTeamCompletionDate,
      assigneeForecasts: assigneeForecasts,
      tasks: teamTasks
    });
  });
  
  return teamForecasts.sort((a, b) => a.language.localeCompare(b.language) || a.teamName.localeCompare(b.teamName));
}

/**
 * Calculates language-level forecasts
 */
function calculateLanguageForecasts(teamForecasts) {
  const languageForecasts = {};
  
  // Group teams by language
  const teamsByLanguage = {};
  teamForecasts.forEach(team => {
    if (!teamsByLanguage[team.language]) {
      teamsByLanguage[team.language] = [];
    }
    teamsByLanguage[team.language].push(team);
  });
  
  // Calculate forecast for each language
  Object.keys(teamsByLanguage).forEach(language => {
    const teamsInLanguage = teamsByLanguage[language];
    
    const totalTasks = teamsInLanguage.reduce((sum, team) => sum + team.totalTasks, 0);
    const totalDevs = teamsInLanguage.reduce((sum, team) => sum + team.actualDevCount, 0);
    const averageTasksPerDev = totalTasks / totalDevs;
    
    // Language completion date is when the last team finishes
    let languageCompletionDate = new Date();
    teamsInLanguage.forEach(team => {
      if (team.teamCompletionDate > languageCompletionDate) {
        languageCompletionDate = team.teamCompletionDate;
      }
    });
    
    const averageWorkingDays = Math.ceil(averageTasksPerDev / FORECAST_CONFIG.TASKS_PER_DEV_PER_DAY);
    
    languageForecasts[language] = {
      language: language,
      totalTasks: totalTasks,
      totalDevs: totalDevs,
      teamCount: teamsInLanguage.length,
      averageTasksPerDev: averageTasksPerDev,
      averageWorkingDays: averageWorkingDays,
      languageCompletionDate: languageCompletionDate,
      teams: teamsInLanguage
    };
  });
  
  return languageForecasts;
}

/**
 * Calculates individual assignee workloads
 */
function calculateAssigneeWorkloads(allTasks) {
  const workloads = {};
  
  allTasks.forEach(task => {
    if (!workloads[task.assignee]) {
      workloads[task.assignee] = {
        assignee: task.assignee,
        team: task.team,
        language: task.language,
        taskCount: 0,
        tasks: []
      };
    }
    
    workloads[task.assignee].taskCount++;
    workloads[task.assignee].tasks.push(task);
  });
  
  // Calculate completion forecasts for each assignee
  Object.keys(workloads).forEach(assignee => {
    const workload = workloads[assignee];
    const workingDaysNeeded = Math.ceil(workload.taskCount / FORECAST_CONFIG.TASKS_PER_DEV_PER_DAY);
    const completionDate = addWorkingDays(new Date(), workingDaysNeeded);
    
    workload.workingDaysNeeded = workingDaysNeeded;
    workload.completionDate = completionDate;
    workload.tasksPerDay = FORECAST_CONFIG.TASKS_PER_DEV_PER_DAY;
  });
  
  return Object.values(workloads).sort((a, b) => b.taskCount - a.taskCount);
}

// ============================
// DASHBOARD BUILDING FUNCTIONS
// ============================

/**
 * Builds the complete forecast dashboard
 */
function buildForecastDashboard(sheet, forecastData) {
  let currentRow = 1;
  
  // 1. Dashboard Header
  currentRow = createForecastHeader(sheet, currentRow, forecastData);
  
  // 2. Language-Level Forecasts
  currentRow = createLanguageForecastSection(sheet, currentRow, forecastData.languageForecasts);
  
  // 3. Team-Level Forecasts
  currentRow = createTeamForecastSection(sheet, currentRow, forecastData.teamForecasts);
  
  // 4. Individual Assignee Workloads
  currentRow = createAssigneeWorkloadSection(sheet, currentRow, forecastData.assigneeWorkloads);
  
  // 5. Detailed Breakdown Tables
  currentRow = createDetailedForecastTables(sheet, currentRow, forecastData);
  
  // Format the sheet
  formatForecastDashboard(sheet);
}

/**
 * Creates the dashboard header
 */
function createForecastHeader(sheet, startRow, forecastData) {
  // Main header
  sheet.getRange(startRow, 1, 1, 12).merge();
  const headerCell = sheet.getRange(startRow, 1);
  headerCell.setValue('TASK COMPLETION FORECASTING DASHBOARD');
  headerCell.setFontSize(22);
  headerCell.setFontWeight('bold');
  headerCell.setBackground(FORECAST_CONFIG.COLORS.FORECAST_HEADER);
  headerCell.setFontColor('#ffffff');
  headerCell.setHorizontalAlignment('center');
  headerCell.setVerticalAlignment('middle');
  sheet.setRowHeight(startRow, 60);
  
  // Forecast parameters
  sheet.getRange(startRow + 1, 1, 1, 12).merge();
  const paramsCell = sheet.getRange(startRow + 1, 1);
  paramsCell.setValue(`Forecast Parameters: ${FORECAST_CONFIG.TASKS_PER_DEV_PER_DAY} tasks per developer per day | ${FORECAST_CONFIG.WORKING_DAYS_PER_WEEK} working days per week`);
  paramsCell.setFontWeight('bold');
  paramsCell.setHorizontalAlignment('center');
  paramsCell.setBackground(FORECAST_CONFIG.COLORS.LIGHT_GRAY);
  
  // Summary info
  sheet.getRange(startRow + 2, 1, 1, 12).merge();
  const summaryCell = sheet.getRange(startRow + 2, 1);
  summaryCell.setValue(`Total Active Tasks: ${forecastData.totalActiveTasks} | Jira Tasks: ${forecastData.jiraTasks} | Flex Tasks: ${forecastData.flexTasks} | Excluded Tasks: ${forecastData.excludedTasks} | Generated: ${forecastData.generatedAt.toLocaleString()}`);
  summaryCell.setFontStyle('italic');
  summaryCell.setHorizontalAlignment('center');
  
  return startRow + 4;
}

/**
 * Creates language-level forecast section
 */
function createLanguageForecastSection(sheet, startRow, languageForecasts) {
  // Section header
  sheet.getRange(startRow, 1, 1, 12).merge();
  const sectionHeader = sheet.getRange(startRow, 1);
  sectionHeader.setValue('LANGUAGE-LEVEL COMPLETION FORECASTS');
  sectionHeader.setFontSize(16);
  sectionHeader.setFontWeight('bold');
  sectionHeader.setBackground(FORECAST_CONFIG.COLORS.LIGHT_GRAY);
  
  let currentRow = startRow + 2;
  
  // Create forecast cards for each language
  Object.keys(languageForecasts).forEach(language => {
    const forecast = languageForecasts[language];
    const completionDateStr = forecast.languageCompletionDate.toLocaleDateString('en-US', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
    
    // Language header
    sheet.getRange(currentRow, 1, 1, 12).merge();
    const langHeader = sheet.getRange(currentRow, 1);
    langHeader.setValue(`${language.toUpperCase()} - Completion Forecast: ${completionDateStr}`);
    langHeader.setFontSize(14);
    langHeader.setFontWeight('bold');
    
    // Set language-specific colors
    if (language === 'Team6-Language') {
      langHeader.setBackground(FORECAST_CONFIG.COLORS.FLEX);
      langHeader.setFontColor('#ffffff');
    } else if (FORECAST_CONFIG.COLORS[language.toUpperCase()]) {
      langHeader.setBackground(FORECAST_CONFIG.COLORS[language.toUpperCase()]);
      if (language === 'JavaScript') {
        langHeader.setFontColor('#000000');
      } else {
        langHeader.setFontColor('#ffffff');
      }
    }
    
    currentRow++;
    
    // Forecast details
    const details = [
      ['Total Active Tasks:', forecast.totalTasks],
      ['Total Developers:', forecast.totalDevs],
      ['Teams Involved:', forecast.teamCount],
      ['Avg Tasks per Dev:', forecast.averageTasksPerDev.toFixed(1)],
      ['Avg Working Days:', forecast.averageWorkingDays],
      ['Completion Date:', completionDateStr]
    ];
    
    details.forEach((detail, index) => {
      const col = (index % 3) * 4 + 1;
      const row = Math.floor(index / 3);
      
      sheet.getRange(currentRow + row, col).setValue(detail[0]);
      sheet.getRange(currentRow + row, col).setFontWeight('bold');
      sheet.getRange(currentRow + row, col + 1).setValue(detail[1]);
      sheet.getRange(currentRow + row, col + 1).setFontWeight('bold');
      
      if (detail[0].includes('Completion Date')) {
        sheet.getRange(currentRow + row, col + 1).setFontColor(FORECAST_CONFIG.COLORS.SUCCESS);
      }
    });
    
    currentRow += 3;
  });
  
  return currentRow + 1;
}

/**
 * Creates team-level forecast section
 */
function createTeamForecastSection(sheet, startRow, teamForecasts) {
  // Section header
  sheet.getRange(startRow, 1, 1, 15).merge();
  const sectionHeader = sheet.getRange(startRow, 1);
  sectionHeader.setValue('TEAM-LEVEL COMPLETION FORECASTS');
  sectionHeader.setFontSize(16);
  sectionHeader.setFontWeight('bold');
  sectionHeader.setBackground(FORECAST_CONFIG.COLORS.LIGHT_GRAY);
  
  // Table headers
  const headerRow = startRow + 2;
  const headers = [
    'Team', 'Language', 'Lead', 'Config Devs', 'Active Devs', 'Total Tasks',
    'Avg Tasks/Dev', 'Avg Working Days', 'Team Completion Date', 'Status',
    'Longest Individual', 'Shortest Individual', 'Team Efficiency', 'Risk Level', 'Notes'
  ];
  
  headers.forEach((header, index) => {
    const cell = sheet.getRange(headerRow, index + 1);
    cell.setValue(header);
    cell.setFontWeight('bold');
    cell.setBackground(FORECAST_CONFIG.COLORS.FORECAST_HEADER);
    cell.setFontColor('#ffffff');
    cell.setBorder(true, true, true, true, false, false);
  });
  
  // Team data rows
  let dataRow = headerRow + 1;
  teamForecasts.forEach(team => {
    const completionDateStr = team.teamCompletionDate.toLocaleDateString();
    
    // Calculate additional metrics
    const longestIndividualDays = Math.max(...team.assigneeForecasts.map(a => a.workingDaysNeeded));
    const shortestIndividualDays = Math.min(...team.assigneeForecasts.map(a => a.workingDaysNeeded));
    
    // Calculate team efficiency (how well distributed the work is)
    const taskDistribution = team.assigneeForecasts.map(a => a.taskCount);
    const maxTasks = Math.max(...taskDistribution);
    const minTasks = Math.min(...taskDistribution);
    const efficiency = minTasks / maxTasks;
    const efficiencyText = efficiency > 0.8 ? 'Excellent' : efficiency > 0.6 ? 'Good' : efficiency > 0.4 ? 'Fair' : 'Poor';
    
    // Determine risk level based on workload and timeline
    let riskLevel = 'Low';
    const avgDays = team.averageWorkingDays;
    if (avgDays > 30) riskLevel = 'High';
    else if (avgDays > 20) riskLevel = 'Medium';
    else if (avgDays > 10) riskLevel = 'Low-Medium';
    
    // Status based on completion timeline
    let status = 'On Track';
    if (avgDays > 30) status = 'Delayed';
    else if (avgDays > 20) status = 'At Risk';
    else if (avgDays <= 7) status = 'Fast Track';
    
    // Notes
    let notes = '';
    if (team.actualDevCount !== team.configuredDevCount) {
      notes += `Config vs Actual devs mismatch; `;
    }
    if (efficiency < 0.6) {
      notes += `Uneven workload distribution; `;
    }
    if (team.totalTasks === 0) {
      notes += `No active tasks; `;
    }
    
    const rowData = [
      team.teamName,
      team.language,
      team.leadName,
      team.configuredDevCount,
      team.actualDevCount,
      team.totalTasks,
      team.averageTasksPerDev.toFixed(1),
      team.averageWorkingDays,
      completionDateStr,
      status,
      `${longestIndividualDays} days`,
      `${shortestIndividualDays} days`,
      efficiencyText,
      riskLevel,
      notes.trim()
    ];
    
    rowData.forEach((value, index) => {
      const cell = sheet.getRange(dataRow, index + 1);
      cell.setValue(value);
      
      // Style specific columns
      if (index === 1) { // Language
        if (team.language === 'Team6-Language') {
          cell.setFontColor(FORECAST_CONFIG.COLORS.FLEX);
        } else if (FORECAST_CONFIG.COLORS[team.language.toUpperCase()]) {
          cell.setFontColor(FORECAST_CONFIG.COLORS[team.language.toUpperCase()]);
        }
        cell.setFontWeight('bold');
      } else if (index === 8) { // Completion date
        cell.setFontColor(FORECAST_CONFIG.COLORS.SUCCESS);
        cell.setFontWeight('bold');
      } else if (index === 9) { // Status
        if (value === 'Delayed') cell.setFontColor(FORECAST_CONFIG.COLORS.DANGER);
        else if (value === 'At Risk') cell.setFontColor(FORECAST_CONFIG.COLORS.WARNING);
        else if (value === 'Fast Track') cell.setFontColor(FORECAST_CONFIG.COLORS.SUCCESS);
        cell.setFontWeight('bold');
      } else if (index === 12) { // Efficiency
        if (value === 'Excellent') cell.setFontColor(FORECAST_CONFIG.COLORS.SUCCESS);
        else if (value === 'Good') cell.setFontColor(FORECAST_CONFIG.COLORS.WARNING);
        else if (value === 'Fair' || value === 'Poor') cell.setFontColor(FORECAST_CONFIG.COLORS.DANGER);
        cell.setFontWeight('bold');
      } else if (index === 13) { // Risk level
        if (value === 'High') cell.setFontColor(FORECAST_CONFIG.COLORS.DANGER);
        else if (value === 'Medium') cell.setFontColor(FORECAST_CONFIG.COLORS.WARNING);
        cell.setFontWeight('bold');
      } else if (index === 14) { // Notes
        cell.setFontSize(9);
        cell.setWrap(true);
      }
      
      cell.setBorder(true, true, true, true, false, false);
    });
    
    // Alternate row coloring
    if (dataRow % 2 === 0) {
      sheet.getRange(dataRow, 1, 1, headers.length).setBackground(FORECAST_CONFIG.COLORS.LIGHT_GRAY);
    }
    
    // Special highlighting for Flex team
    if (team.teamName === 'Flex') {
      sheet.getRange(dataRow, 1, 1, headers.length).setBackground('#f3e5f5');
    }
    
    dataRow++;
  });
  
  return dataRow + 2;
}

/**
 * Creates individual assignee workload section
 */
function createAssigneeWorkloadSection(sheet, startRow, assigneeWorkloads) {
  // Section header
  sheet.getRange(startRow, 1, 1, 12).merge();
  const sectionHeader = sheet.getRange(startRow, 1);
  sectionHeader.setValue('TOP INDIVIDUAL ASSIGNEE WORKLOADS');
  sectionHeader.setFontSize(16);
  sectionHeader.setFontWeight('bold');
  sectionHeader.setBackground(FORECAST_CONFIG.COLORS.LIGHT_GRAY);
  
  // Show top 20 assignees with highest workloads
  const topAssignees = assigneeWorkloads.slice(0, 20);
  
  // Table headers
  const headerRow = startRow + 2;
  const headers = [
    'Assignee', 'Team', 'Language', 'Active Tasks', 'Working Days Needed',
    'Completion Date', 'Tasks per Day Rate', 'Workload Level', 'Status', 'Weeks Needed', 'Notes', 'Priority'
  ];
  
  headers.forEach((header, index) => {
    const cell = sheet.getRange(headerRow, index + 1);
    cell.setValue(header);
    cell.setFontWeight('bold');
    cell.setBackground(FORECAST_CONFIG.COLORS.FORECAST_HEADER);
    cell.setFontColor('#ffffff');
    cell.setBorder(true, true, true, true, false, false);
  });
  
  // Assignee data rows
  let dataRow = headerRow + 1;
  topAssignees.forEach((assignee, index) => {
    const completionDateStr = assignee.completionDate.toLocaleDateString();
    const weeksNeeded = Math.ceil(assignee.workingDaysNeeded / FORECAST_CONFIG.WORKING_DAYS_PER_WEEK);
    
    // Determine workload level
    let workloadLevel = 'Normal';
    let workloadColor = FORECAST_CONFIG.COLORS.SUCCESS;
    if (assignee.taskCount > 30) {
      workloadLevel = 'Very High';
      workloadColor = FORECAST_CONFIG.COLORS.DANGER;
    } else if (assignee.taskCount > 20) {
      workloadLevel = 'High';
      workloadColor = FORECAST_CONFIG.COLORS.WARNING;
    } else if (assignee.taskCount > 10) {
      workloadLevel = 'Moderate';
      workloadColor = FORECAST_CONFIG.COLORS.WARNING;
    }
    
    // Status
    let status = 'Manageable';
    if (assignee.workingDaysNeeded > 30) status = 'Overloaded';
    else if (assignee.workingDaysNeeded > 20) status = 'Heavy Load';
    else if (assignee.workingDaysNeeded > 10) status = 'Moderate Load';
    
    // Notes
    let notes = '';
    if (assignee.taskCount > 25) notes += 'Consider redistribution; ';
    if (assignee.workingDaysNeeded > 30) notes += 'Critical bottleneck; ';
    if (assignee.team === 'Flex') notes += 'External team member; ';
    
    // Priority (higher number = higher priority for attention)
    let priority = 'Low';
    if (assignee.taskCount > 25 || assignee.workingDaysNeeded > 30) priority = 'Critical';
    else if (assignee.taskCount > 15 || assignee.workingDaysNeeded > 20) priority = 'High';
    else if (assignee.taskCount > 10) priority = 'Medium';
    
    const rowData = [
      assignee.assignee,
      assignee.team,
      assignee.language,
      assignee.taskCount,
      assignee.workingDaysNeeded,
      completionDateStr,
      FORECAST_CONFIG.TASKS_PER_DEV_PER_DAY,
      workloadLevel,
      status,
      weeksNeeded,
      notes.trim(),
      priority
    ];
    
    rowData.forEach((value, index) => {
      const cell = sheet.getRange(dataRow, index + 1);
      cell.setValue(value);
      
      // Style specific columns
      if (index === 0) { // Assignee name
        cell.setFontWeight('bold');
        // Highlight top 5 workloads
        if (dataRow <= headerRow + 5) {
          cell.setBackground('#ffecb3');
        }
      } else if (index === 2) { // Language
        if (assignee.language === 'Team6-Language') {
          cell.setFontColor(FORECAST_CONFIG.COLORS.FLEX);
        } else if (FORECAST_CONFIG.COLORS[assignee.language.toUpperCase()]) {
          cell.setFontColor(FORECAST_CONFIG.COLORS[assignee.language.toUpperCase()]);
        }
        cell.setFontWeight('bold');
      } else if (index === 3) { // Task count
        cell.setFontWeight('bold');
        if (assignee.taskCount > 20) cell.setFontColor(FORECAST_CONFIG.COLORS.DANGER);
        else if (assignee.taskCount > 10) cell.setFontColor(FORECAST_CONFIG.COLORS.WARNING);
      } else if (index === 5) { // Completion date
        cell.setFontColor(FORECAST_CONFIG.COLORS.SUCCESS);
      } else if (index === 7) { // Workload level
        cell.setFontColor(workloadColor);
        cell.setFontWeight('bold');
      } else if (index === 8) { // Status
        if (value === 'Overloaded') cell.setFontColor(FORECAST_CONFIG.COLORS.DANGER);
        else if (value === 'Heavy Load') cell.setFontColor(FORECAST_CONFIG.COLORS.WARNING);
        cell.setFontWeight('bold');
      } else if (index === 10) { // Notes
        cell.setFontSize(9);
        cell.setWrap(true);
      } else if (index === 11) { // Priority
        if (value === 'Critical') cell.setFontColor(FORECAST_CONFIG.COLORS.DANGER);
        else if (value === 'High') cell.setFontColor(FORECAST_CONFIG.COLORS.WARNING);
        cell.setFontWeight('bold');
      }
      
      cell.setBorder(true, true, true, true, false, false);
    });
    
    // Alternate row coloring
    if (dataRow % 2 === 0) {
      sheet.getRange(dataRow, 1, 1, headers.length).setBackground(FORECAST_CONFIG.COLORS.LIGHT_GRAY);
    }
    
    // Special highlighting for critical workloads
    if (assignee.taskCount > 25) {
      sheet.getRange(dataRow, 1, 1, headers.length).setBackground('#ffebee');
    }
    
    dataRow++;
  });
  
  return dataRow + 2;
}

/**
 * Creates detailed forecast tables
 */
function createDetailedForecastTables(sheet, startRow, forecastData) {
  let currentRow = startRow;
  
  // Detailed team breakdowns by language
  Object.keys(forecastData.languageForecasts).forEach(language => {
    const languageForecast = forecastData.languageForecasts[language];
    
    // Language section header
    sheet.getRange(currentRow, 1, 1, 12).merge();
    const langHeader = sheet.getRange(currentRow, 1);
    langHeader.setValue(`${language.toUpperCase()} - DETAILED TEAM AND ASSIGNEE BREAKDOWN`);
    langHeader.setFontSize(14);
    langHeader.setFontWeight('bold');
    
    if (language === 'Team6-Language') {
      langHeader.setBackground(FORECAST_CONFIG.COLORS.FLEX);
      langHeader.setFontColor('#ffffff');
    } else if (FORECAST_CONFIG.COLORS[language.toUpperCase()]) {
      langHeader.setBackground(FORECAST_CONFIG.COLORS[language.toUpperCase()]);
      if (language === 'JavaScript') {
        langHeader.setFontColor('#000000');
      } else {
        langHeader.setFontColor('#ffffff');
      }
    }
    
    currentRow += 2;
    
    // For each team in this language
    languageForecast.teams.forEach(team => {
      // Team header
      sheet.getRange(currentRow, 1, 1, 12).merge();
      const teamHeader = sheet.getRange(currentRow, 1);
      const projectedDateStr = team.projectedTeamCompletionDate && team.projectedTeamCompletionDate.getTime() !== team.teamCompletionDate.getTime()
          ? team.projectedTeamCompletionDate.toLocaleDateString()
          : 'N/A';
      teamHeader.setValue(`Team ${team.teamName} - ${team.totalTasks} tasks - Current Completion: ${team.teamCompletionDate.toLocaleDateString()} | Projected Completion: ${projectedDateStr}`);
      teamHeader.setFontWeight('bold');
      teamHeader.setBackground('#e3f2fd');
      
      currentRow++;
      
      // Assignee breakdown table
      const assigneeHeaders = [
          'Assignee', 'Total Tasks', 'Current "To Do"', 'Working Days', 'Completion Date',
          'Projected Tasks', 'Reassignment', 'Projected Completion', 'Time Saved (Days)',
          'Workload', 'Sample Tasks'
      ];
      assigneeHeaders.forEach((header, index) => {
        const cell = sheet.getRange(currentRow, index + 1);
        cell.setValue(header);
        cell.setFontWeight('bold');
        cell.setBackground('#bbdefb');
      });
      
      currentRow++;
      
      // Individual assignees
      team.assigneeForecasts.forEach(assignee => {
        let workload = 'Normal';
        if (assignee.taskCount > 20) workload = 'High';
        else if (assignee.taskCount > 10) workload = 'Moderate';
        
        // Get sample task keys (first 3)
        const assigneeTasks = team.tasks.filter(t => t.assignee === assignee.assignee);
        const sampleTasks = assigneeTasks.slice(0, 3).map(t => t.key).join(', ');

        const projectedCompletionDateStr = assignee.projectedCompletionDate ? assignee.projectedCompletionDate.toLocaleDateString() : 'N/A';
        const reassignmentStr = assignee.reassignment !== null ? (assignee.reassignment > 0 ? `+${assignee.reassignment}` : assignee.reassignment) : 'N/A';
        
        const assigneeData = [
          assignee.assignee,
          assignee.taskCount,
          assignee.currentToDoCount,
          assignee.workingDaysNeeded,
          assignee.completionDate.toLocaleDateString(),
          assignee.projectedTaskCount !== null ? assignee.projectedTaskCount : 'N/A',
          reassignmentStr,
          projectedCompletionDateStr,
          assignee.timeSaved !== null ? assignee.timeSaved : 'N/A',
          workload,
          sampleTasks + (assigneeTasks.length > 3 ? '...' : '')
        ];
        
        assigneeData.forEach((value, index) => {
          const cell = sheet.getRange(currentRow, index + 1);
          cell.setValue(value);
          
          if (index === 0) { // Assignee
              cell.setFontWeight('bold');
          } else if (index === 1 && assignee.taskCount > 15) { // Total Tasks
              cell.setFontColor(FORECAST_CONFIG.COLORS.WARNING).setFontWeight('bold');
          } else if (index === 4) { // Completion Date
              cell.setFontColor(FORECAST_CONFIG.COLORS.SUCCESS);
          } else if (index === 6) { // Reassignment
              if (assignee.reassignment > 0) {
                  cell.setFontColor(FORECAST_CONFIG.COLORS.DANGER).setFontWeight('bold');
              } else if (assignee.reassignment < 0) {
                  cell.setFontColor(FORECAST_CONFIG.COLORS.SUCCESS).setFontWeight('bold');
              }
          } else if (index === 7 && assignee.projectedCompletionDate) { // Projected Completion
              cell.setFontColor(FORECAST_CONFIG.COLORS.SUCCESS);
          } else if (index === 8) { // Time Saved
              if (assignee.timeSaved > 0) {
                  cell.setFontColor(FORECAST_CONFIG.COLORS.SUCCESS).setFontWeight('bold');
              } else if (assignee.timeSaved < 0) {
                  cell.setFontColor(FORECAST_CONFIG.COLORS.DANGER).setFontWeight('bold');
              }
          } else if (index === 9 && workload === 'High') { // Workload
              cell.setFontColor(FORECAST_CONFIG.COLORS.DANGER).setFontWeight('bold');
          } else if (index === 10) { // Sample Tasks
              cell.setFontSize(9);
          }
        });
        
        currentRow++;
      });
      
      currentRow++; // Space between teams
    });
    
    currentRow += 2; // Space between languages
  });
  
  return currentRow;
}

/**
 * Formats the forecast dashboard
 */
function formatForecastDashboard(sheet) {
  // Auto-resize columns
  sheet.autoResizeColumns(1, 15);
  
  // Set specific column widths
  sheet.setColumnWidth(1, 150); // Team/Assignee names
  sheet.setColumnWidth(2, 120); // Language
  sheet.setColumnWidth(9, 150); // Completion dates
  sheet.setColumnWidth(11, 200); // Notes
  sheet.setColumnWidth(15, 200); // Additional notes
  
  // Add borders to the entire data range
  const dataRange = sheet.getDataRange();
  if (dataRange.getNumRows() > 0 && dataRange.getNumColumns() > 0) {
    dataRange.setBorder(true, true, true, true, false, false, FORECAST_CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }
  
  // Freeze header rows
  sheet.setFrozenRows(4);
}

// ============================
// MENU AND UTILITY FUNCTIONS
// ============================



/**
 * Adjusts the tasks per day rate
 */
function adjustTasksPerDayRate() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Adjust Forecast Rate',
    `Current rate: ${FORECAST_CONFIG.TASKS_PER_DEV_PER_DAY} tasks per developer per day\n\nEnter new rate (1-10):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const newRate = parseInt(response.getResponseText());
    if (newRate >= 1 && newRate <= 10) {
      // Note: This changes the rate for the current session only
      // To persist changes, modify the CONFIG object
      ui.alert(
        'Rate Adjustment',
        `Rate temporarily set to ${newRate} tasks per day for this session.\n\n` +
        `To make permanent changes, modify FORECAST_CONFIG.TASKS_PER_DEV_PER_DAY in the script.`,
        ui.ButtonSet.OK
      );
      FORECAST_CONFIG.TASKS_PER_DEV_PER_DAY = newRate;
    } else {
      ui.alert('Invalid rate. Please enter a number between 1 and 10.');
    }
  }
}

/**
 * Views current forecast settings
 */
function viewForecastSettings() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Current Forecast Settings',
    `Tasks per Developer per Day: ${FORECAST_CONFIG.TASKS_PER_DEV_PER_DAY}\n` +
    `Working Days per Week: ${FORECAST_CONFIG.WORKING_DAYS_PER_WEEK}\n` +
    `Data Sheet: ${FORECAST_CONFIG.DATA_SHEET_NAME}\n` +
    `Flex Sheet ID: ${FORECAST_CONFIG.FLEX_SHEET_ID}\n` +
    `Flex Sheet Name: ${FORECAST_CONFIG.FLEX_SHEET_NAME}\n` +
    `Excluded Assignees: ${FORECAST_CONFIG.EXCLUDED_ASSIGNEES.length} configured\n\n` +
    `Teams Configured: ${Object.keys(FORECAST_CONFIG.TEAMS).length}`,
    ui.ButtonSet.OK
  );
}

/**
 * Tests data source connections
 */
function testDataSources() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Test Jira data
    const jiraData = collectJiraForecastData();
    
    // Test Flex data
    const flexData = collectFlexForecastData();
    
    ui.alert(
      'Data Source Test Results',
      `Jira Data Sheet: ✓ Connected\n` +
      `- Active tasks found: ${jiraData.tasks.length}\n` +
      `- Excluded tasks: ${jiraData.excludedTasks}\n\n` +
      `Flex Data Sheet: ${flexData.tasks.length > 0 ? '✓' : '⚠'} ${flexData.tasks.length > 0 ? 'Connected' : 'No data or connection issues'}\n` +
      `- Active tasks found: ${flexData.tasks.length}\n` +
      `- Excluded tasks: ${flexData.excludedTasks}\n\n` +
      `Total Active Tasks: ${jiraData.tasks.length + flexData.tasks.length}`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Data Source Test Failed', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Exports forecast to PDF (placeholder)
 */
function exportForecastToPDF() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Export to PDF',
    'To export the forecast to PDF:\n\n' +
    '1. Go to File > Download > PDF Document (.pdf)\n' +
    '2. Select "Specific sheets" and choose "Task Completion Forecast"\n' +
    '3. Choose "Fit to width" for best formatting\n' +
    '4. Click "Export"',
    ui.ButtonSet.OK
  );
}

/**
 * Sets up automatic forecast generation (placeholder)
 */
function setupAutoForecast() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Auto Forecast Setup',
    'Auto forecast feature requires setting up time-based triggers.\n\n' +
    'This would automatically regenerate the forecast dashboard daily.\n\n' +
    'Implementation requires additional trigger management code.',
    ui.ButtonSet.OK
  );
}

/**
 * Stops automatic forecast generation (placeholder)
 */
function stopAutoForecast() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Stop Auto Forecast',
    'This would stop any existing automatic forecast triggers.\n\n' +
    'Implementation requires trigger management code.',
    ui.ButtonSet.OK
  );
}
