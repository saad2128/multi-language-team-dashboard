const CONFIG = {
  // Master sheet name where dashboard will be created
  MASTER_SHEET_NAME: 'Multi-Language Team Dashboard',
  
  // Language targets
  LANGUAGE_TARGETS: {
    'Python': 10000,
    'Rust': 10000,
    'JavaScript': 10000
  },
  
  // Total project target
  PROJECT_TARGET: 30000,
  
  // Team sheets configuration - organized by language
  LANGUAGE_TEAMS: {
    Python: [
      {
        id: '', // Spreadsheet ID for Python Team 1
        name: 'Python Team Alpha',
        leadName: 'Lead Name 1',
        developers: 5,
        language: 'Python',
        teamCode: 'PY-ALPHA'
      },
      {
        id: '', // Spreadsheet ID for Python Team 2
        name: 'Python Team Beta',
        leadName: 'Lead Name 2',
        developers: 6,
        language: 'Python',
        teamCode: 'PY-BETA'
      }
    ],
    Rust: [
      {
        id: '', // Spreadsheet ID for Rust Team 1
        name: 'Rust Team Gamma',
        leadName: 'Lead Name 3',
        developers: 4,
        language: 'Rust',
        teamCode: 'RS-GAMMA'
      }
    ],
    JavaScript: [
      {
        id: '', // Spreadsheet ID for JS Team 1
        name: 'JavaScript Team Delta',
        leadName: 'Lead Name 4',
        developers: 7,
        language: 'JavaScript',
        teamCode: 'JS-DELTA'
      }
    ]
  },
  
  // Column mappings for new project structure
  TASK_ID_COLUMN: 'task_id',
  WORKER_ID_COLUMN: 'worker_id',
  JIRA_ID_COLUMN: 'jira_id',
  TASK_START_DATE_COLUMN: 'task_start_date',
  REPO_URL_COLUMN: 'repo_url',
  TASK_DESCRIPTION_COLUMN: 'task_description',
  FIRST_PROMPT_COLUMN: 'first_prompt',
  REVIEWER_ID_COLUMN: 'reviewer_id',
  REPO_APPROVAL_COLUMN: 'repo_approval',
  TASK_COMPLETION_STATUS_COLUMN: 'task_completion_status',
  COMMENTS_COLUMN: 'comments',
  
  // Status values mapping
  STATUS_VALUES: {
    NOT_STARTED: ['Not Started', 'TODO', 'Backlog', '', 'Pending', 'New', 'Assigned'],
    IN_PROGRESS: ['In Progress', 'In-Progress', 'Working', 'Started', 'Development', 'Coding', 'Active'],
    REVIEW: ['Ready for Review', 'In Review', 'In-Review', 'Review', 'Under Review', 'Reviewing', 'Peer Review', 'Code Review'],
    COMPLETED: ['Completed', 'Done', 'Finished', 'Closed', 'Deployed', 'Live', 'Approved', 'Merged'],
    REWORK: ['Rework', 'Redo', 'Needs Rework', 'Rework Required', 'Fix Required', 'Revision Needed', 'Changes Requested', 'Blocked', 'On Hold', 'Waiting', 'Paused', 'Stuck'],
    REJECTED: ['Rejected', 'Failed', 'Declined', 'Not Approved']
  },
  
  // Repo approval status mapping
  REPO_APPROVAL_VALUES: {
    PENDING: ['Pending', 'Waiting', 'Submitted', '', 'Under Review'],
    APPROVED: ['Approved', 'Yes', 'Accepted', 'Good', 'Pass'],
    REJECTED: ['Rejected', 'No', 'Failed', 'Denied', 'Fail'],
    NEEDS_CHANGES: ['Needs Changes', 'Revision Required', 'Fix Required', 'Changes Requested']
  },
  
  // Dashboard styling
  COLORS: {
    HEADER: '#1a73e8',
    COMPLETED: '#34a853',
    IN_PROGRESS: '#fbbc04',
    REVIEW: '#4285f4',
    NOT_STARTED: '#9e9e9e',
    REWORK: '#ff6d01',
    REJECTED: '#ea4335',
    APPROVED: '#34a853',
    PENDING: '#fbbc04',
    NEEDS_CHANGES: '#ff6d01',
    BORDER: '#dadce0',
    LIGHT_GRAY: '#f8f9fa',
    PYTHON: '#3776ab',
    RUST: '#ce422b',
    JAVASCRIPT: '#f7df1e',
    JIRA_HEADER: '#0052cc'
  }
};

// ============================
// MAIN FUNCTIONS
// ============================

/**
 * Creates or updates the master tracker dashboard
 */
function createMasterTracker() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let masterSheet = ss.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    
    // Create sheet if it doesn't exist
    if (!masterSheet) {
      masterSheet = ss.insertSheet(CONFIG.MASTER_SHEET_NAME);
    }
    
    // Clear existing content
    masterSheet.clear();
    
    // Collect data from all team sheets
    const allTeamsData = collectAllTeamsData();
    
    // Build dashboard
    buildDashboard(masterSheet, allTeamsData);
    
    // Show completion message
    SpreadsheetApp.getUi().alert('Multi-Language Team Dashboard created successfully!');
    
  } catch (error) {
    console.error('Error creating master tracker:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Exports detailed task and JIRA data to Excel format
 */
function exportDetailedTaskData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const exportSheetName = 'Task Data Export';
    let exportSheet = ss.getSheetByName(exportSheetName);
    
    // Create sheet if it doesn't exist
    if (!exportSheet) {
      exportSheet = ss.insertSheet(exportSheetName);
    }
    
    // Clear existing content
    exportSheet.clear();
    
    // Collect detailed task data
    const taskData = collectDetailedTaskData();
    
    // Build export sheet
    buildTaskExportSheet(exportSheet, taskData);
    
    // Show completion message
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Task Data Export Complete!',
      `Task data has been exported to the "${exportSheetName}" sheet.\n\n` +
      'To download as Excel:\n' +
      '1. Go to File > Download > Microsoft Excel (.xlsx)\n' +
      '2. Or select the export sheet and copy the data.\n\n' +
      `Total tasks found: ${taskData.totalTasks}`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error exporting task data:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Collects data from all team sheets
 */
function collectAllTeamsData() {
  const allData = {
    teams: [],
    languageStats: {
      Python: { totalTasks: 0, completed: 0, inProgress: 0, review: 0, notStarted: 0, rework: 0, rejected: 0 },
      Rust: { totalTasks: 0, completed: 0, inProgress: 0, review: 0, notStarted: 0, rework: 0, rejected: 0 },
      JavaScript: { totalTasks: 0, completed: 0, inProgress: 0, review: 0, notStarted: 0, rework: 0, rejected: 0 }
    },
    approvalStats: {
      Python: { approved: 0, pending: 0, rejected: 0, needsChanges: 0 },
      Rust: { approved: 0, pending: 0, rejected: 0, needsChanges: 0 },
      JavaScript: { approved: 0, pending: 0, rejected: 0, needsChanges: 0 }
    },
    overallStats: {
      totalTasks: 0,
      completed: 0,
      inProgress: 0,
      review: 0,
      notStarted: 0,
      rework: 0,
      rejected: 0
    }
  };
  
  // Process each language and its teams
  Object.entries(CONFIG.LANGUAGE_TEAMS).forEach(([language, teams]) => {
    teams.forEach(team => {
      if (team.id) {
        try {
          const teamData = getTeamData(team);
          allData.teams.push(teamData);
          
          // Update language-specific stats
          const langStats = allData.languageStats[language];
          langStats.totalTasks += teamData.stats.total;
          langStats.completed += teamData.stats.completed;
          langStats.inProgress += teamData.stats.inProgress;
          langStats.review += teamData.stats.review;
          langStats.notStarted += teamData.stats.notStarted;
          langStats.rework += teamData.stats.rework;
          langStats.rejected += teamData.stats.rejected;
          
          // Update approval stats
          const approvalStats = allData.approvalStats[language];
          approvalStats.approved += teamData.approvalStats.approved;
          approvalStats.pending += teamData.approvalStats.pending;
          approvalStats.rejected += teamData.approvalStats.rejected;
          approvalStats.needsChanges += teamData.approvalStats.needsChanges;
          
          // Update overall stats
          allData.overallStats.totalTasks += teamData.stats.total;
          allData.overallStats.completed += teamData.stats.completed;
          allData.overallStats.inProgress += teamData.stats.inProgress;
          allData.overallStats.review += teamData.stats.review;
          allData.overallStats.notStarted += teamData.stats.notStarted;
          allData.overallStats.rework += teamData.stats.rework;
          allData.overallStats.rejected += teamData.stats.rejected;
          
        } catch (error) {
          console.error(`Error processing team ${team.name}:`, error);
        }
      }
    });
  });
  
  return allData;
}

/**
 * Gets data from a single team sheet
 */
function getTeamData(teamConfig) {
  const sheet = SpreadsheetApp.openById(teamConfig.id).getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length === 0) {
    return {
      config: teamConfig,
      stats: {
        total: 0,
        completed: 0,
        inProgress: 0,
        review: 0,
        notStarted: 0,
        rework: 0,
        rejected: 0
      },
      approvalStats: {
        approved: 0,
        pending: 0,
        rejected: 0,
        needsChanges: 0
      },
      tasks: []
    };
  }
  
  // Find column indices
  const headers = data[0];
  const taskIdCol = findColumnIndex(headers, CONFIG.TASK_ID_COLUMN);
  const workerIdCol = findColumnIndex(headers, CONFIG.WORKER_ID_COLUMN);
  const statusCol = findColumnIndex(headers, CONFIG.TASK_COMPLETION_STATUS_COLUMN);
  const approvalCol = findColumnIndex(headers, CONFIG.REPO_APPROVAL_COLUMN);
  const jiraIdCol = findColumnIndex(headers, CONFIG.JIRA_ID_COLUMN);
  const repoUrlCol = findColumnIndex(headers, CONFIG.REPO_URL_COLUMN);
  const startDateCol = findColumnIndex(headers, CONFIG.TASK_START_DATE_COLUMN);
  
  // Process tasks
  const tasks = [];
  const stats = {
    total: 0,
    completed: 0,
    inProgress: 0,
    review: 0,
    notStarted: 0,
    rework: 0,
    rejected: 0
  };
  
  const approvalStats = {
    approved: 0,
    pending: 0,
    rejected: 0,
    needsChanges: 0
  };
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[taskIdCol] || row[workerIdCol]) { // Check if row has data
      const status = row[statusCol] ? row[statusCol].toString() : '';
      const approval = row[approvalCol] ? row[approvalCol].toString() : '';
      const statusCategory = categorizeStatus(status);
      const approvalCategory = categorizeApproval(approval);
      
      tasks.push({
        taskId: row[taskIdCol] || '',
        workerId: row[workerIdCol] || '',
        jiraId: row[jiraIdCol] || '',
        repoUrl: row[repoUrlCol] || '',
        startDate: row[startDateCol] || '',
        status: status,
        approval: approval,
        statusCategory: statusCategory,
        approvalCategory: approvalCategory
      });
      
      stats.total++;
      stats[statusCategory]++;
      approvalStats[approvalCategory]++;
    }
  }
  
  // Ensure all stats are integers
  Object.keys(stats).forEach(key => {
    stats[key] = parseInt(stats[key]) || 0;
  });
  
  Object.keys(approvalStats).forEach(key => {
    approvalStats[key] = parseInt(approvalStats[key]) || 0;
  });
  
  return {
    config: teamConfig,
    stats: stats,
    approvalStats: approvalStats,
    tasks: tasks
  };
}

/**
 * Collects detailed task data for export
 */
function collectDetailedTaskData() {
  const taskData = {
    teams: [],
    allTasks: [],
    totalTasks: 0,
    languages: ['Python', 'Rust', 'JavaScript']
  };
  
  // Process each language and team
  Object.entries(CONFIG.LANGUAGE_TEAMS).forEach(([language, teams]) => {
    teams.forEach(team => {
      if (team.id) {
        try {
          const teamTaskData = getDetailedTeamData(team);
          taskData.teams.push(teamTaskData);
          
          // Add tasks to overall collection
          teamTaskData.tasks.forEach(task => {
            taskData.allTasks.push({
              ...task,
              teamName: team.name,
              leadName: team.leadName,
              language: team.language,
              teamCode: team.teamCode
            });
          });
          
          taskData.totalTasks += teamTaskData.tasks.length;
          
        } catch (error) {
          console.error(`Error processing detailed data for team ${team.name}:`, error);
        }
      }
    });
  });
  
  return taskData;
}

/**
 * Gets detailed task data from a single team sheet
 */
function getDetailedTeamData(teamConfig) {
  const sheet = SpreadsheetApp.openById(teamConfig.id).getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length === 0) {
    return {
      config: teamConfig,
      tasks: []
    };
  }
  
  // Get all headers and their indices
  const headers = data[0];
  const columnMap = {};
  
  Object.keys(CONFIG).forEach(key => {
    if (key.endsWith('_COLUMN')) {
      const columnName = CONFIG[key];
      columnMap[key] = findColumnIndex(headers, columnName);
    }
  });
  
  // Process all tasks with full details
  const tasks = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[columnMap.TASK_ID_COLUMN] || row[columnMap.WORKER_ID_COLUMN]) {
      const status = row[columnMap.TASK_COMPLETION_STATUS_COLUMN] ? 
        row[columnMap.TASK_COMPLETION_STATUS_COLUMN].toString() : '';
      const approval = row[columnMap.REPO_APPROVAL_COLUMN] ? 
        row[columnMap.REPO_APPROVAL_COLUMN].toString() : '';
      
      tasks.push({
        taskId: row[columnMap.TASK_ID_COLUMN] || '',
        workerId: row[columnMap.WORKER_ID_COLUMN] || '',
        jiraId: row[columnMap.JIRA_ID_COLUMN] || '',
        taskStartDate: row[columnMap.TASK_START_DATE_COLUMN] || '',
        repoUrl: row[columnMap.REPO_URL_COLUMN] || '',
        taskDescription: row[columnMap.TASK_DESCRIPTION_COLUMN] || '',
        firstPrompt: row[columnMap.FIRST_PROMPT_COLUMN] || '',
        reviewerId: row[columnMap.REVIEWER_ID_COLUMN] || '',
        repoApproval: approval,
        taskCompletionStatus: status,
        comments: row[columnMap.COMMENTS_COLUMN] || '',
        statusCategory: categorizeStatus(status),
        approvalCategory: categorizeApproval(approval),
        rowNumber: i + 1
      });
    }
  }
  
  return {
    config: teamConfig,
    tasks: tasks
  };
}

/**
 * Helper function to find column index
 */
function findColumnIndex(headers, columnName) {
  return headers.findIndex(h => h.toString().toLowerCase() === columnName.toLowerCase());
}

/**
 * Categorizes task completion status
 */
function categorizeStatus(status) {
  const normalizedStatus = status.toString().toLowerCase().trim();
  
  for (const [category, values] of Object.entries(CONFIG.STATUS_VALUES)) {
    if (values.some(v => v.toLowerCase() === normalizedStatus)) {
      switch(category) {
        case 'NOT_STARTED': return 'notStarted';
        case 'IN_PROGRESS': return 'inProgress';
        case 'REVIEW': return 'review';
        case 'COMPLETED': return 'completed';
        case 'REWORK': return 'rework';
        case 'REJECTED': return 'rejected';
      }
    }
  }
  
  return 'notStarted'; // Default category
}

/**
 * Categorizes repo approval status
 */
function categorizeApproval(approval) {
  const normalizedApproval = approval.toString().toLowerCase().trim();
  
  for (const [category, values] of Object.entries(CONFIG.REPO_APPROVAL_VALUES)) {
    if (values.some(v => v.toLowerCase() === normalizedApproval)) {
      switch(category) {
        case 'PENDING': return 'pending';
        case 'APPROVED': return 'approved';
        case 'REJECTED': return 'rejected';
        case 'NEEDS_CHANGES': return 'needsChanges';
      }
    }
  }
  
  return 'pending'; // Default category
}

/**
 * Builds the dashboard layout
 */
function buildDashboard(sheet, data) {
  let currentRow = 1;
  
  // 1. Dashboard Header
  currentRow = createDashboardHeader(sheet, currentRow);
  
  // 2. Overall Progress Section
  currentRow = createOverallProgress(sheet, data.overallStats, currentRow);
  
  // 3. Language Breakdown Section
  currentRow = createLanguageBreakdown(sheet, data.languageStats, data.approvalStats, currentRow);
  
  // 4. Team Statistics Table
  currentRow = createTeamStatisticsTable(sheet, data.teams, currentRow);
  
  // 5. Detailed Analytics
  currentRow = createDetailedAnalytics(sheet, data, currentRow);
  
  // Format the entire sheet
  formatDashboard(sheet);
}

/**
 * Creates dashboard header
 */
function createDashboardHeader(sheet, startRow) {
  sheet.getRange(startRow, 1, 1, 12).merge();
  const headerCell = sheet.getRange(startRow, 1);
  headerCell.setValue('MULTI-LANGUAGE TEAM DASHBOARD');
  headerCell.setFontSize(22);
  headerCell.setFontWeight('bold');
  headerCell.setBackground(CONFIG.COLORS.HEADER);
  headerCell.setFontColor('#ffffff');
  headerCell.setHorizontalAlignment('center');
  headerCell.setVerticalAlignment('middle');
  sheet.setRowHeight(startRow, 60);
  
  // Project info
  sheet.getRange(startRow + 1, 1, 1, 12).merge();
  const projectInfo = sheet.getRange(startRow + 1, 1);
  projectInfo.setValue('Python (10,000 tasks) | Rust (10,000 tasks) | JavaScript (10,000 tasks) | Total Target: 30,000');
  projectInfo.setFontWeight('bold');
  projectInfo.setHorizontalAlignment('center');
  projectInfo.setBackground(CONFIG.COLORS.LIGHT_GRAY);
  
  // Last updated timestamp
  sheet.getRange(startRow + 2, 1, 1, 12).merge();
  const timestampCell = sheet.getRange(startRow + 2, 1);
  timestampCell.setValue('Last Updated: ' + new Date().toLocaleString());
  timestampCell.setFontStyle('italic');
  timestampCell.setHorizontalAlignment('center');
  
  return startRow + 4;
}

/**
 * Creates overall progress section
 */
function createOverallProgress(sheet, stats, startRow) {
  // Section header
  sheet.getRange(startRow, 1, 1, 12).merge();
  const sectionHeader = sheet.getRange(startRow, 1);
  sectionHeader.setValue('OVERALL PROJECT PROGRESS');
  sectionHeader.setFontSize(16);
  sectionHeader.setFontWeight('bold');
  sectionHeader.setBackground(CONFIG.COLORS.LIGHT_GRAY);
  
  // Progress metrics
  const progressRow = startRow + 2;
  const completionPercentage = (stats.completed / CONFIG.PROJECT_TARGET * 100).toFixed(1);
  
  // Create progress cards
  const cards = [
    {
      label: 'Project Target',
      value: CONFIG.PROJECT_TARGET.toLocaleString(),
      color: CONFIG.COLORS.HEADER
    },
    {
      label: 'Total Tasks',
      value: stats.totalTasks.toLocaleString(),
      color: CONFIG.COLORS.HEADER
    },
    {
      label: 'Completed',
      value: `${stats.completed.toLocaleString()} (${completionPercentage}%)`,
      color: CONFIG.COLORS.COMPLETED
    },
    {
      label: 'In Progress',
      value: stats.inProgress.toLocaleString(),
      color: CONFIG.COLORS.IN_PROGRESS
    },
    {
      label: 'In Review',
      value: stats.review.toLocaleString(),
      color: CONFIG.COLORS.REVIEW
    },
    {
      label: 'Rework',
      value: stats.rework.toLocaleString(),
      color: CONFIG.COLORS.REWORK
    }
  ];
  
  // Display cards
  cards.forEach((card, index) => {
    const col = (index * 2) + 1;
    
    // Label
    sheet.getRange(progressRow, col).setValue(card.label);
    sheet.getRange(progressRow, col).setFontWeight('bold');
    
    // Value
    sheet.getRange(progressRow + 1, col).setValue(card.value);
    sheet.getRange(progressRow + 1, col).setFontSize(14);
    sheet.getRange(progressRow + 1, col).setFontColor(card.color);
    sheet.getRange(progressRow + 1, col).setFontWeight('bold');
  });
  
  // Progress bar
  const progressBarRow = progressRow + 3;
  sheet.getRange(progressBarRow, 1).setValue('Overall Completion:');
  sheet.getRange(progressBarRow, 1).setFontWeight('bold');
  
  sheet.getRange(progressBarRow, 2, 1, 10).merge();
  createProgressBar(sheet, progressBarRow, 2, parseFloat(completionPercentage));
  
  return progressBarRow + 3;
}

/**
 * Creates language breakdown section
 */
function createLanguageBreakdown(sheet, languageStats, approvalStats, startRow) {
  // Section header
  sheet.getRange(startRow, 1, 1, 12).merge();
  const sectionHeader = sheet.getRange(startRow, 1);
  sectionHeader.setValue('LANGUAGE BREAKDOWN');
  sectionHeader.setFontSize(16);
  sectionHeader.setFontWeight('bold');
  sectionHeader.setBackground(CONFIG.COLORS.LIGHT_GRAY);
  
  let currentRow = startRow + 2;
  
  // Create breakdown for each language
  ['Python', 'Rust', 'JavaScript'].forEach(language => {
    const stats = languageStats[language];
    const approval = approvalStats[language];
    const target = CONFIG.LANGUAGE_TARGETS[language];
    const completion = (stats.completed / target * 100).toFixed(1);
    
    // Language header
    sheet.getRange(currentRow, 1, 1, 12).merge();
    const langHeader = sheet.getRange(currentRow, 1);
    langHeader.setValue(`${language.toUpperCase()} - Target: ${target} | Completed: ${stats.completed} (${completion}%)`);
    langHeader.setFontWeight('bold');
    langHeader.setBackground(CONFIG.COLORS[language.toUpperCase()]);
    if (language === 'JavaScript') {
      langHeader.setFontColor('#000000');
    } else {
      langHeader.setFontColor('#ffffff');
    }
    
    currentRow++;
    
    // Task Completion Status label
    sheet.getRange(currentRow, 1, 1, 12).merge();
    const statusLabel = sheet.getRange(currentRow, 1);
    statusLabel.setValue('Task Completion Status:');
    statusLabel.setFontWeight('bold');
    statusLabel.setFontSize(11);
    statusLabel.setFontStyle('italic');
    statusLabel.setBackground('#f0f0f0');
    
    currentRow++;
    
    // Task status breakdown
    const statusData = [
      ['Not Started', stats.notStarted, CONFIG.COLORS.NOT_STARTED],
      ['In Progress', stats.inProgress, CONFIG.COLORS.IN_PROGRESS],
      ['Review', stats.review, CONFIG.COLORS.REVIEW],
      ['Completed', stats.completed, CONFIG.COLORS.COMPLETED],
      ['Rework', stats.rework, CONFIG.COLORS.REWORK],
      ['Rejected', stats.rejected, CONFIG.COLORS.REJECTED]
    ];
    
    statusData.forEach((item, index) => {
      const col = (index * 2) + 1;
      sheet.getRange(currentRow, col).setValue(item[0]);
      sheet.getRange(currentRow, col).setFontSize(10);
      
      const countCell = sheet.getRange(currentRow, col + 1);
      countCell.setValue(item[1]);
      countCell.setFontWeight('bold');
      countCell.setFontColor(item[2]);
      countCell.setHorizontalAlignment('center');
    });
    
    currentRow++;
    
    // Repo Approval Status label
    sheet.getRange(currentRow, 1, 1, 12).merge();
    const approvalLabel = sheet.getRange(currentRow, 1);
    approvalLabel.setValue('Repo Approval Status:');
    approvalLabel.setFontWeight('bold');
    approvalLabel.setFontSize(11);
    approvalLabel.setFontStyle('italic');
    approvalLabel.setBackground('#f0f0f0');
    
    currentRow++;
    
    // Approval status breakdown
    const approvalData = [
      ['Pending Approval', approval.pending, CONFIG.COLORS.PENDING],
      ['Approved', approval.approved, CONFIG.COLORS.APPROVED],
      ['Needs Changes', approval.needsChanges, CONFIG.COLORS.NEEDS_CHANGES],
      ['Rejected', approval.rejected, CONFIG.COLORS.REJECTED]
    ];
    
    approvalData.forEach((item, index) => {
      const col = (index * 2) + 1;
      sheet.getRange(currentRow, col).setValue(item[0]);
      sheet.getRange(currentRow, col).setFontSize(10);
      
      const countCell = sheet.getRange(currentRow, col + 1);
      countCell.setValue(item[1]);
      countCell.setFontWeight('bold');
      countCell.setFontColor(item[2]);
      countCell.setHorizontalAlignment('center');
    });
    
    currentRow += 2;
  });
  
  return currentRow + 1;
}

/**
 * Creates team statistics table
 */
function createTeamStatisticsTable(sheet, teams, startRow) {
  // Section header
  sheet.getRange(startRow, 1, 1, 16).merge();
  const sectionHeader = sheet.getRange(startRow, 1);
  sectionHeader.setValue('TEAM STATISTICS');
  sectionHeader.setFontSize(16);
  sectionHeader.setFontWeight('bold');
  sectionHeader.setBackground(CONFIG.COLORS.LIGHT_GRAY);
  
  // Table headers
  const headerRow = startRow + 2;
  const headers = [
    'Team Name', 'Language', 'Lead', 'Developers', 'Team Code',
    'Total Tasks', 'Completed', 'In Progress', 'Review', 'Rework', 'Rejected',
    'Approval Rate', 'Progress %', 'Performance', 'Sheet Link'
  ];
  
  headers.forEach((header, index) => {
    const cell = sheet.getRange(headerRow, index + 1);
    cell.setValue(header);
    cell.setFontWeight('bold');
    cell.setBackground(CONFIG.COLORS.HEADER);
    cell.setFontColor('#ffffff');
    cell.setBorder(true, true, true, true, false, false);
  });
  
  // Team data rows
  let dataRow = headerRow + 1;
  teams.forEach(team => {
    const stats = team.stats;
    const approval = team.approvalStats;
    const config = team.config;
    
    const progressPercentage = stats.total > 0 
      ? (stats.completed / stats.total * 100).toFixed(1)
      : 0;
    
    const approvalRate = (approval.approved + approval.pending + approval.needsChanges + approval.rejected) > 0
      ? (approval.approved / (approval.approved + approval.pending + approval.needsChanges + approval.rejected) * 100).toFixed(1)
      : 0;
    
    // Performance indicator
    let performance = 'Average';
    const progress = parseFloat(progressPercentage);
    if (progress >= 80) performance = 'Excellent';
    else if (progress >= 60) performance = 'Good';
    else if (progress < 30) performance = 'Needs Attention';
    
    const rowData = [
      config.name,
      config.language,
      config.leadName,
      config.developers,
      config.teamCode,
      stats.total,
      stats.completed,
      stats.inProgress,
      stats.review,
      stats.rework,
      stats.rejected,
      `${approvalRate}%`,
      `${progressPercentage}%`,
      performance,
      `=HYPERLINK("https://docs.google.com/spreadsheets/d/${config.id}", "Open Sheet")`
    ];
    
    rowData.forEach((value, index) => {
      const cell = sheet.getRange(dataRow, index + 1);
      cell.setValue(value);
      
      // Style specific columns
      if (index === 1) { // Language column
        cell.setFontColor(CONFIG.COLORS[config.language.toUpperCase()]);
        cell.setFontWeight('bold');
      } else if (index === 6) { // Completed column
        cell.setFontColor(CONFIG.COLORS.COMPLETED);
        cell.setFontWeight('bold');
      } else if (index === 9) { // Rework column
        if (parseInt(value) > 0) {
          cell.setFontColor(CONFIG.COLORS.REWORK);
          cell.setFontWeight('bold');
        }
      } else if (index === 10) { // Rejected column
        if (parseInt(value) > 0) {
          cell.setFontColor(CONFIG.COLORS.REJECTED);
          cell.setFontWeight('bold');
        }
      } else if (index === 12) { // Progress column
        const progress = parseFloat(value);
        if (progress >= 80) {
          cell.setFontColor(CONFIG.COLORS.COMPLETED);
        } else if (progress >= 50) {
          cell.setFontColor(CONFIG.COLORS.IN_PROGRESS);
        } else {
          cell.setFontColor(CONFIG.COLORS.REJECTED);
        }
        cell.setFontWeight('bold');
      } else if (index === 13) { // Performance column
        if (value === 'Excellent') {
          cell.setFontColor(CONFIG.COLORS.COMPLETED);
        } else if (value === 'Good') {
          cell.setFontColor(CONFIG.COLORS.IN_PROGRESS);
        } else if (value === 'Needs Attention') {
          cell.setFontColor(CONFIG.COLORS.REJECTED);
        }
        cell.setFontWeight('bold');
      }
      
      cell.setBorder(true, true, true, true, false, false);
    });
    
    // Alternate row coloring
    if (dataRow % 2 === 0) {
      sheet.getRange(dataRow, 1, 1, headers.length).setBackground(CONFIG.COLORS.LIGHT_GRAY);
    }
    
    dataRow++;
  });
  
  return dataRow + 2;
}

/**
 * Creates detailed analytics section
 */
function createDetailedAnalytics(sheet, data, startRow) {
  // Section header
  sheet.getRange(startRow, 1, 1, 12).merge();
  const sectionHeader = sheet.getRange(startRow, 1);
  sectionHeader.setValue('DETAILED ANALYTICS & INSIGHTS');
  sectionHeader.setFontSize(16);
  sectionHeader.setFontWeight('bold');
  sectionHeader.setBackground(CONFIG.COLORS.LIGHT_GRAY);
  
  let currentRow = startRow + 2;
  
  // Team performance comparison
  sheet.getRange(currentRow, 1, 1, 12).merge();
  const perfHeader = sheet.getRange(currentRow, 1);
  perfHeader.setValue('TEAM PERFORMANCE COMPARISON');
  perfHeader.setFontWeight('bold');
  perfHeader.setBackground('#e8eaf6');
  
  currentRow += 2;
  
  // Create performance bars for each team
  data.teams.forEach(team => {
    const progressPercentage = team.stats.total > 0 
      ? (team.stats.completed / team.stats.total * 100)
      : 0;
    
    // Team name and language
    sheet.getRange(currentRow, 1, 1, 2).merge();
    const nameCell = sheet.getRange(currentRow, 1);
    nameCell.setValue(`${team.config.name} (${team.config.language})`);
    nameCell.setFontWeight('bold');
    nameCell.setFontColor(CONFIG.COLORS[team.config.language.toUpperCase()]);
    
    // Progress bar
    sheet.getRange(currentRow, 3, 1, 7).merge();
    createProgressBar(sheet, currentRow, 3, progressPercentage);
    
    // Stats
    sheet.getRange(currentRow, 10).setValue(`${progressPercentage.toFixed(1)}%`);
    sheet.getRange(currentRow, 10).setFontWeight('bold');
    
    sheet.getRange(currentRow, 11).setValue(`${team.stats.completed}/${team.stats.total}`);
    sheet.getRange(currentRow, 12).setValue(`Lead: ${team.config.leadName}`);
    
    currentRow++;
  });
  
  // Language targets progress
  currentRow += 2;
  sheet.getRange(currentRow, 1, 1, 12).merge();
  const targetHeader = sheet.getRange(currentRow, 1);
  targetHeader.setValue('LANGUAGE TARGET PROGRESS');
  targetHeader.setFontWeight('bold');
  targetHeader.setBackground('#e8eaf6');
  
  currentRow += 2;
  
  ['Python', 'Rust', 'JavaScript'].forEach(language => {
    const stats = data.languageStats[language];
    const target = CONFIG.LANGUAGE_TARGETS[language];
    const completion = (stats.completed / target * 100);
    
    // Language name
    sheet.getRange(currentRow, 1).setValue(language);
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    sheet.getRange(currentRow, 1).setFontColor(CONFIG.COLORS[language.toUpperCase()]);
    
    // Target info
    sheet.getRange(currentRow, 2).setValue(`Target: ${target}`);
    
    // Progress bar
    sheet.getRange(currentRow, 3, 1, 7).merge();
    createProgressBar(sheet, currentRow, 3, completion);
    
    // Completion stats
    sheet.getRange(currentRow, 10).setValue(`${completion.toFixed(1)}%`);
    sheet.getRange(currentRow, 10).setFontWeight('bold');
    
    sheet.getRange(currentRow, 11).setValue(`${stats.completed}/${target}`);
    sheet.getRange(currentRow, 12).setValue(`Remaining: ${target - stats.completed}`);
    
    currentRow++;
  });
  
  return currentRow + 2;
}

/**
 * Builds the task export sheet
 */
function buildTaskExportSheet(sheet, taskData) {
  let currentRow = 1;
  
  // 1. Header
  currentRow = createTaskExportHeader(sheet, currentRow, taskData);
  
  // 2. Summary Statistics
  currentRow = createTaskExportSummary(sheet, currentRow, taskData);
  
  // 3. All Tasks Table
  currentRow = createTasksTable(sheet, currentRow, taskData);
  
  // Format the sheet
  formatTaskExportSheet(sheet);
}

/**
 * Creates task export header
 */
function createTaskExportHeader(sheet, startRow, taskData) {
  sheet.getRange(startRow, 1, 1, 15).merge();
  const headerCell = sheet.getRange(startRow, 1);
  headerCell.setValue('DETAILED TASK DATA EXPORT');
  headerCell.setFontSize(20);
  headerCell.setFontWeight('bold');
  headerCell.setBackground(CONFIG.COLORS.HEADER);
  headerCell.setFontColor('#ffffff');
  headerCell.setHorizontalAlignment('center');
  headerCell.setVerticalAlignment('middle');
  sheet.setRowHeight(startRow, 50);
  
  // Export timestamp
  sheet.getRange(startRow + 1, 1, 1, 15).merge();
  const timestampCell = sheet.getRange(startRow + 1, 1);
  timestampCell.setValue('Exported: ' + new Date().toLocaleString());
  timestampCell.setFontStyle('italic');
  timestampCell.setHorizontalAlignment('center');
  
  return startRow + 3;
}

/**
 * Creates task export summary
 */
function createTaskExportSummary(sheet, startRow, taskData) {
  // Section header
  sheet.getRange(startRow, 1, 1, 15).merge();
  const sectionHeader = sheet.getRange(startRow, 1);
  sectionHeader.setValue('EXPORT SUMMARY');
  sectionHeader.setFontSize(16);
  sectionHeader.setFontWeight('bold');
  sectionHeader.setBackground(CONFIG.COLORS.LIGHT_GRAY);
  
  // Summary data
  const summaryRow = startRow + 2;
  const languageCounts = {
    Python: taskData.allTasks.filter(t => t.language === 'Python').length,
    Rust: taskData.allTasks.filter(t => t.language === 'Rust').length,
    JavaScript: taskData.allTasks.filter(t => t.language === 'JavaScript').length
  };
  
  const summaryData = [
    ['Total Tasks Found:', taskData.totalTasks],
    ['Python Tasks:', languageCounts.Python],
    ['Rust Tasks:', languageCounts.Rust],
    ['JavaScript Tasks:', languageCounts.JavaScript],
    ['Teams Processed:', taskData.teams.length]
  ];
  
  summaryData.forEach((item, index) => {
    sheet.getRange(summaryRow + index, 1).setValue(item[0]);
    sheet.getRange(summaryRow + index, 1).setFontWeight('bold');
    sheet.getRange(summaryRow + index, 2).setValue(item[1]);
    sheet.getRange(summaryRow + index, 2).setFontWeight('bold');
    sheet.getRange(summaryRow + index, 2).setFontColor(CONFIG.COLORS.HEADER);
  });
  
  return summaryRow + summaryData.length + 2;
}

/**
 * Creates tasks table
 */
function createTasksTable(sheet, startRow, taskData) {
  // Section header
  sheet.getRange(startRow, 1, 1, 15).merge();
  const sectionHeader = sheet.getRange(startRow, 1);
  sectionHeader.setValue('ALL TASKS');
  sectionHeader.setFontSize(16);
  sectionHeader.setFontWeight('bold');
  sectionHeader.setBackground(CONFIG.COLORS.LIGHT_GRAY);
  
  // Table headers
  const headerRow = startRow + 2;
  const headers = [
    'Task ID', 'Worker ID', 'JIRA ID', 'Team', 'Language', 'Lead', 'Team Code',
    'Start Date', 'Repo URL', 'Task Description', 'Reviewer ID', 
    'Completion Status', 'Repo Approval', 'Comments', 'Row #'
  ];
  
  headers.forEach((header, index) => {
    const cell = sheet.getRange(headerRow, index + 1);
    cell.setValue(header);
    cell.setFontWeight('bold');
    cell.setBackground(CONFIG.COLORS.HEADER);
    cell.setFontColor('#ffffff');
    cell.setBorder(true, true, true, true, false, false);
  });
  
  // Data rows
  let dataRow = headerRow + 1;
  taskData.allTasks.forEach(task => {
    const rowData = [
      task.taskId,
      task.workerId,
      task.jiraId,
      task.teamName,
      task.language,
      task.leadName,
      task.teamCode,
      task.taskStartDate,
      task.repoUrl,
      task.taskDescription,
      task.reviewerId,
      task.taskCompletionStatus,
      task.repoApproval,
      task.comments,
      task.rowNumber
    ];
    
    rowData.forEach((value, index) => {
      const cell = sheet.getRange(dataRow, index + 1);
      cell.setValue(value);
      
      // Style based on content
      if (index === 4) { // Language column
        cell.setFontColor(CONFIG.COLORS[task.language.toUpperCase()]);
        cell.setFontWeight('bold');
      } else if (index === 11) { // Status column
        const statusColor = getStatusColor(task.statusCategory);
        cell.setFontColor(statusColor);
        if (task.statusCategory === 'completed') {
          cell.setFontWeight('bold');
        }
      } else if (index === 12) { // Approval column
        const approvalColor = getApprovalColor(task.approvalCategory);
        cell.setFontColor(approvalColor);
        if (task.approvalCategory === 'approved') {
          cell.setFontWeight('bold');
        }
      }
      
      cell.setBorder(true, true, true, true, false, false);
    });
    
    // Alternate row coloring
    if (dataRow % 2 === 0) {
      sheet.getRange(dataRow, 1, 1, headers.length).setBackground(CONFIG.COLORS.LIGHT_GRAY);
    }
    
    dataRow++;
  });
  
  return dataRow + 2;
}

/**
 * Creates a progress bar in a cell
 */
function createProgressBar(sheet, row, col, percentage) {
  const cell = sheet.getRange(row, col);
  const barLength = 30; // Number of characters for the bar
  const filledLength = Math.round(barLength * percentage / 100);
  const emptyLength = barLength - filledLength;
  
  const filledChar = '█';
  const emptyChar = '░';
  
  const progressBar = filledChar.repeat(filledLength) + emptyChar.repeat(emptyLength);
  const progressText = ` ${percentage.toFixed(1)}%`;
  
  cell.setValue(progressBar + progressText);
  cell.setFontFamily('Consolas');
  
  // Color based on percentage
  if (percentage >= 80) {
    cell.setFontColor(CONFIG.COLORS.COMPLETED);
  } else if (percentage >= 60) {
    cell.setFontColor(CONFIG.COLORS.IN_PROGRESS);
  } else if (percentage >= 30) {
    cell.setFontColor(CONFIG.COLORS.REVIEW);
  } else {
    cell.setFontColor(CONFIG.COLORS.REJECTED);
  }
}

/**
 * Gets color based on status category
 */
function getStatusColor(statusCategory) {
  switch(statusCategory) {
    case 'completed': return CONFIG.COLORS.COMPLETED;
    case 'inProgress': return CONFIG.COLORS.IN_PROGRESS;
    case 'review': return CONFIG.COLORS.REVIEW;
    case 'rework': return CONFIG.COLORS.REWORK;
    case 'rejected': return CONFIG.COLORS.REJECTED;
    default: return CONFIG.COLORS.NOT_STARTED;
  }
}

/**
 * Gets color based on approval category
 */
function getApprovalColor(approvalCategory) {
  switch(approvalCategory) {
    case 'approved': return CONFIG.COLORS.APPROVED;
    case 'pending': return CONFIG.COLORS.PENDING;
    case 'needsChanges': return CONFIG.COLORS.NEEDS_CHANGES;
    case 'rejected': return CONFIG.COLORS.REJECTED;
    default: return CONFIG.COLORS.PENDING;
  }
}

/**
 * Formats the dashboard
 */
function formatDashboard(sheet) {
  // Auto-resize columns
  sheet.autoResizeColumns(1, 16);
  
  // Set minimum column widths
  sheet.setColumnWidth(1, 250); // Team name column
  sheet.setColumnWidth(2, 100); // Language column
  sheet.setColumnWidth(3, 150); // Lead column
  
  // Add borders to the entire data range
  const dataRange = sheet.getDataRange();
  dataRange.setBorder(true, true, true, true, false, false, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  
  // Center align numeric columns
  for (let col = 4; col <= 15; col++) {
    sheet.getRange(1, col, sheet.getMaxRows()).setHorizontalAlignment('center');
  }
  
  // Freeze header rows
  sheet.setFrozenRows(4);
}

/**
 * Formats the task export sheet
 */
function formatTaskExportSheet(sheet) {
  // Auto-resize columns
  sheet.autoResizeColumns(1, 15);
  
  // Set specific column widths
  sheet.setColumnWidth(1, 120); // Task ID
  sheet.setColumnWidth(2, 120); // Worker ID
  sheet.setColumnWidth(3, 120); // JIRA ID
  sheet.setColumnWidth(4, 200); // Team name
  sheet.setColumnWidth(9, 300); // Repo URL
  sheet.setColumnWidth(10, 400); // Task description
  sheet.setColumnWidth(14, 300); // Comments
  
  // Add borders to the entire data range
  const dataRange = sheet.getDataRange();
  dataRange.setBorder(true, true, true, true, false, false, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  
  // Freeze header rows
  sheet.setFrozenRows(3);
}

// ============================
// MENU AND TRIGGERS
// ============================

/**
 * Creates custom menu on spreadsheet open
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Multi-Language Dashboard')
    .addItem('Create/Update Dashboard', 'createMasterTracker')
    .addSeparator()
    .addItem('Export Detailed Task Data', 'exportDetailedTaskData')
    .addSeparator()
    .addItem('Configure Team Sheets', 'showConfigurationDialog')
    .addSeparator()
    .addItem('Schedule Auto-Update', 'scheduleAutoUpdate')
    .addItem('Remove Auto-Update', 'removeAutoUpdate')
    .addToUi();
}

/**
 * Shows configuration dialog
 */
function showConfigurationDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ConfigDialog')
    .setWidth(700)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Configure Multi-Language Teams');
}

/**
 * Schedules automatic updates every 2 hours
 */
function scheduleAutoUpdate() {
  // Remove existing triggers
  removeAutoUpdate();
  
  // Create new trigger for every 2 hours
  ScriptApp.newTrigger('createMasterTracker')
    .timeBased()
    .everyHours(2)
    .create();
  
  SpreadsheetApp.getUi().alert('Auto-update scheduled! The dashboard will refresh every 2 hours.');
}

/**
 * Sets up automatic refresh on spreadsheet open
 */
function setupAutoRefresh() {
  // Schedule the dashboard to refresh automatically
  scheduleAutoUpdate();
  
  // Also create the dashboard immediately
  createMasterTracker();
}

/**
 * Removes automatic update trigger
 */
function removeAutoUpdate() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'createMasterTracker') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

// ============================
// HELPER FUNCTIONS
// ============================

/**
 * Gets configuration for use in HTML dialog
 */
function getConfiguration() {
  return CONFIG;
}

/**
 * Updates configuration from HTML dialog
 */
function updateConfiguration(newConfig) {
  console.log('Configuration update requested:', newConfig);
  SpreadsheetApp.getUi().alert('Configuration saved! Please update the CONFIG object in the script with your team sheet IDs.');
}

/**
 * Validates team sheet structure
 */
function validateTeamSheet(sheetId) {
  try {
    const sheet = SpreadsheetApp.openById(sheetId).getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const requiredColumns = [
      CONFIG.TASK_ID_COLUMN,
      CONFIG.WORKER_ID_COLUMN,
      CONFIG.TASK_COMPLETION_STATUS_COLUMN,
      CONFIG.REPO_APPROVAL_COLUMN
    ];
    
    const missingColumns = requiredColumns.filter(col => 
      !headers.some(header => header.toString().toLowerCase() === col.toLowerCase())
    );
    
    if (missingColumns.length > 0) {
      return {
        valid: false,
        message: `Missing required columns: ${missingColumns.join(', ')}`
      };
    }
    
    return {
      valid: true,
      message: 'Sheet structure is valid'
    };
    
  } catch (error) {
    return {
      valid: false,
      message: `Error accessing sheet: ${error.message}`
    };
  }
}

/**
 * Gets sample data for testing
 */
function createSampleData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Create Sample Data',
    'This will create sample team sheets with test data. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    // Implementation for creating sample sheets would go here
    ui.alert('Sample data creation feature coming soon!');
  }
}