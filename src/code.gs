// ============================
// CONFIGURATION - UPDATED WITH EXCLUSION LIST
// ============================

const CONFIG = {
  // Master sheet name where dashboard will be created
  MASTER_SHEET_NAME: 'Dashboard',
  
  // Data source sheet name
  DATA_SHEET_NAME: 'Jira_Data',
  
  // External Flex team sheet configuration
  FLEX_SHEET_ID: 'SHEET ID Here', 
  FLEX_SHEET_NAME: 'Tasks', 
  
  // NEW: Assignee exclusion list - these assignees will be excluded from statistics
  EXCLUDED_ASSIGNEES: [
    'John Doe',
    'David'
  ],
  
  // Language targets
  LANGUAGE_TARGETS: {
    'Python': 10000,
    'Rust': 10000,
    'JavaScript': 10000
  },
  
  // Total project target
  PROJECT_TARGET: 30000,
  
  // Team configuration - includes Flex team
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
      developers: 7,
      language: 'JavaScript',
      teamCode: 'JS-Team2'
    },
    'Team3': {
      name: 'Team3',
      leadName: 'Lead3',
      developers: 11,
      language: 'Rust',
      teamCode: 'RS-Team3'
    }
	,
    // Flex team configuration
    'Flex': {
      name: 'Flex',
      leadName: 'Lead4',
      developers: 0,
      language: 'Multi-Language',
      teamCode: 'FLEX'
    },
  },
  
  // Column mappings for new structure from Jira_Data tab
  KEY_COLUMN: 'Key',
  ASSIGNEE_COLUMN: 'Assignee',
  TEAM_COLUMN: 'Team',
  STATUS_COLUMN: 'Status',
  FOLDER_ID_COLUMN: 'folder/id',
  SUMMARY_COLUMN: 'Summary',
  TASK_DESCRIPTION_COLUMN: 'task_description',
  FIRST_PROMPT_COLUMN: 'first_prompt',
  REPO_URL_COLUMN: 'repo_url',
  GITHUB_LINK_COLUMN: 'github link',
  LT_URL_COLUMN: 'LT_url',
  GDRIVE_LINK_COLUMN: 'gdrive link',
  CREATED_COLUMN: 'Created',
  UPDATED_COLUMN: 'Updated',
  
  // Column mappings for Flex team data
  FLEX_COLUMNS: {
    TASK_ID: 'task_id',
    WORKER_ID: 'worker_id',
    LANGUAGE: 'Language',
    TASK_START_DATE: 'task_start_date',
    REPO_URL: 'repo_url',
    TASK_DESCRIPTION: 'task_description',
    FIRST_PROMPT: 'first_prompt',
    LT_URL: 'LT_url',
    GDRIVE_LINK: 'gdrive_link',
    REVIEWER_ID: 'reviewer_id',
    REPO_APPROVAL: 'repo_approval',
    TASK_COMPLETION_STATUS: 'task_completion_status',
    COMMENTS: 'comments',
    REVIEW_STATUS: 'Review_Status'
  },
  
  // Status values mapping based on your specific requirements
  STATUS_VALUES: {
    IN_PROGRESS: ['In Progress', 'Repo Approved', 'Interaction in progress'],
    REVIEW: ['Repo Review', 'Peer review', 'Lead review'],
    ISSUE: ['Blocked'],
    REJECTED: ['Repo Rejected'],
    REWORK: ['Rework Required'],
    COMPLETED: ['Done']
  },
  
  // Status mapping for Flex team
  FLEX_STATUS_VALUES: {
    IN_PROGRESS: ['In progress'],
    REVIEW: ['Ready for Review', 'Review in progress'],
    ISSUE: ['Changes Requested'],
    REJECTED: ['Rejected'],
    REWORK: ['Rework'],
    COMPLETED: ['Approved']
  },
  
  // Team to language mapping
  TEAM_LANGUAGE_MAP: {
    'Team1': 'Python',
    'Team2': 'JavaScript',
    'Team3': 'Rust',
    'Flex': 'Multi-Language',
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
    JIRA_HEADER: '#0052cc',
    FLEX: '#9c27b0'
  }
};

// ============================
// EXCLUSION HELPER FUNCTIONS
// ============================

/**
 * Helper function to check if an assignee should be excluded
 */
function shouldExcludeAssignee(assignee) {
  if (!assignee) return false;
  
  const normalizedAssignee = assignee.toString().trim();
  
  return CONFIG.EXCLUDED_ASSIGNEES.some(excludedName => {
    // Case-insensitive comparison and handle partial matches
    return normalizedAssignee.toLowerCase().includes(excludedName.toLowerCase()) ||
           excludedName.toLowerCase().includes(normalizedAssignee.toLowerCase());
  });
}

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
    
    // Collect data from both Jira_Data sheet and Flex external sheet
    const allTeamsData = collectAllTeamsData();
    
    // Build dashboard
    buildDashboard(masterSheet, allTeamsData);
    
    // Show completion message
    SpreadsheetApp.getUi().alert('PR Writer V2 Dashboard created successfully!');
    
  } catch (error) {
    console.error('Error creating master tracker:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Function to read data from external Flex team Google Sheet
 */
function getFlexTeamData() {
  try {
    // Open the external Google Sheet
    const flexSheet = SpreadsheetApp.openById(CONFIG.FLEX_SHEET_ID);
    const dataSheet = flexSheet.getSheetByName(CONFIG.FLEX_SHEET_NAME);
    
    if (!dataSheet) {
      console.warn(`Flex sheet '${CONFIG.FLEX_SHEET_NAME}' not found in external document.`);
      return [];
    }
    
    const data = dataSheet.getDataRange().getValues();
    
    if (data.length === 0) {
      console.warn('No data found in Flex team sheet.');
      return [];
    }
    
    // Find column indices for Flex data
    const headers = data[0];
    const flexColumnMap = getFlexColumnMap(headers);
    
    const flexTasks = [];
    
    // Process each row of Flex data
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[flexColumnMap.TASK_ID] || row[flexColumnMap.WORKER_ID]) { // Check if row has data
        const language = row[flexColumnMap.LANGUAGE] ? row[flexColumnMap.LANGUAGE].toString().trim() : 'Unknown';
        const taskCompletionStatus = row[flexColumnMap.TASK_COMPLETION_STATUS] ? row[flexColumnMap.TASK_COMPLETION_STATUS].toString() : '';
        const repoApproval = row[flexColumnMap.REPO_APPROVAL] ? row[flexColumnMap.REPO_APPROVAL].toString() : '';
        
        // Determine the effective status - prioritize task_completion_status, fallback to repo_approval
        let effectiveStatus = taskCompletionStatus;
        if (!effectiveStatus && repoApproval) {
          effectiveStatus = repoApproval;
        }
        
        const statusInfo = categorizeFlexStatus(effectiveStatus);
        
        const task = {
          key: row[flexColumnMap.TASK_ID] || '',
          assignee: row[flexColumnMap.WORKER_ID] || '',
          team: 'Flex',
          status: effectiveStatus,
          folderId: '', // Not available in Flex data
          summary: row[flexColumnMap.TASK_DESCRIPTION] || '',
          taskDescription: row[flexColumnMap.TASK_DESCRIPTION] || '',
          firstPrompt: row[flexColumnMap.FIRST_PROMPT] || '',
          repoUrl: row[flexColumnMap.REPO_URL] || '',
          githubLink: '', // Not available in Flex data
          ltUrl: row[flexColumnMap.LT_URL] || '',
          gdriveLink: row[flexColumnMap.GDRIVE_LINK] || '',
          created: row[flexColumnMap.TASK_START_DATE] || '',
          updated: '', // Not available in Flex data
          statusCategory: statusInfo.category,
          statusSubcategory: statusInfo.subcategory,
          language: language,
          reviewerId: row[flexColumnMap.REVIEWER_ID] || '',
          repoApproval: repoApproval,
          taskCompletionStatus: taskCompletionStatus,
          comments: row[flexColumnMap.COMMENTS] || '',
          reviewStatus: row[flexColumnMap.REVIEW_STATUS] || '',
          rowNumber: i + 1,
          isFlexTeam: true
        };
        
        flexTasks.push(task);
      }
    }
    
    console.log(`Found ${flexTasks.length} tasks from Flex team`);
    return flexTasks;
    
  } catch (error) {
    console.error('Error reading Flex team data:', error);
    SpreadsheetApp.getUi().alert(
      'Flex Team Data Error',
      `Could not read Flex team data: ${error.toString()}\n\nPlease check:\n1. Sheet ID is correct\n2. Sheet has proper permissions\n3. Sheet name is correct`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return [];
  }
}

/**
 * Creates column mapping for Flex team headers
 */
function getFlexColumnMap(headers) {
  const columnMap = {};
  
  // Map each configured Flex column to its index
  Object.keys(CONFIG.FLEX_COLUMNS).forEach(key => {
    const columnName = CONFIG.FLEX_COLUMNS[key];
    const index = findColumnIndex(headers, columnName);
    columnMap[key] = index;
  });
  
  return columnMap;
}

/**
 * Categorizes Flex team status and returns both category and subcategory
 */
function categorizeFlexStatus(status) {
  const normalizedStatus = status.toString().trim();
  
  for (const [category, values] of Object.entries(CONFIG.FLEX_STATUS_VALUES)) {
    if (values.some(v => v === normalizedStatus)) {
      const categoryName = category === 'IN_PROGRESS' ? 'inProgress' :
                          category === 'REVIEW' ? 'review' :
                          category === 'ISSUE' ? 'issue' :
                          category === 'REJECTED' ? 'rejected' :
                          category === 'REWORK' ? 'rework' :
                          category === 'COMPLETED' ? 'completed' : 'inProgress';
      
      return {
        category: categoryName,
        subcategory: normalizedStatus
      };
    }
  }
  
  return {
    category: 'inProgress',
    subcategory: normalizedStatus
  };
}

/**
 * UPDATED: Collects data from both the Jira_Data sheet and Flex external sheet with exclusion filtering
 */
function collectAllTeamsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET_NAME);
  
  if (!dataSheet) {
    throw new Error(`Sheet '${CONFIG.DATA_SHEET_NAME}' not found. Please ensure the Jira_Data sheet exists.`);
  }
  
  const data = dataSheet.getDataRange().getValues();
  
  if (data.length === 0) {
    throw new Error('No data found in Jira_Data sheet.');
  }
  
  // Find column indices
  const headers = data[0];
  const columnMap = getColumnMap(headers);
  
  // Initialize language stats including Multi-Language for Flex
  const allData = {
    teams: [],
    languageStats: {
      Python: { 
        totalTasks: 0, 
        completed: 0, 
        inProgress: 0, 
        review: 0, 
        issue: 0, 
        rejected: 0,
        rework: 0,
        excluded: 0,
        // Sub-category breakdowns
        inProgressSub: {
          'In Progress': 0,
          'Repo Approved': 0,
          'Interaction in progress': 0
        },
        reviewSub: {
          'Repo Review': 0,
          'Peer review': 0,
          'Lead Review': 0
        },
        issueSub: {
          'Blocked': 0
        },
        rejectedSub: {
          'Repo Rejected': 0
        },
        reworkSub: {
          'Rework Required': 0
        },
        completedSub: {
          'Done': 0
        }
      },
      Rust: { 
        totalTasks: 0, 
        completed: 0, 
        inProgress: 0, 
        review: 0, 
        issue: 0, 
        rejected: 0,
        rework: 0,
        excluded: 0,
        // Sub-category breakdowns
        inProgressSub: {
          'In Progress': 0,
          'Repo Approved': 0,
          'Interaction in progress': 0
        },
        reviewSub: {
          'Repo Review': 0,
          'Peer review': 0,
          'Lead Review': 0
        },
        issueSub: {
          'Blocked': 0
        },
        rejectedSub: {
          'Repo Rejected': 0
        },
        reworkSub: {
          'Rework Required': 0
        },
        completedSub: {
          'Done': 0
        }
      },
      JavaScript: { 
        totalTasks: 0, 
        completed: 0, 
        inProgress: 0, 
        review: 0, 
        issue: 0, 
        rejected: 0,
        rework: 0,
        excluded: 0,
        // Sub-category breakdowns
        inProgressSub: {
          'In Progress': 0,
          'Repo Approved': 0,
          'Interaction in progress': 0
        },
        reviewSub: {
          'Repo Review': 0,
          'Peer review': 0,
          'Lead Review': 0
        },
        issueSub: {
          'Blocked': 0
        },
        rejectedSub: {
          'Repo Rejected': 0
        },
        reworkSub: {
          'Rework Required': 0
        },
        completedSub: {
          'Done': 0
        }
      },
      'Multi-Language': {
        totalTasks: 0,
        completed: 0,
        inProgress: 0,
        review: 0,
        issue: 0,
        rejected: 0,
        rework: 0,
        excluded: 0,
        // Flex-specific sub-category breakdowns
        inProgressSub: {
          'In progress': 0
        },
        reviewSub: {
          'Ready for Review': 0,
          'Review in progress': 0
        },
        issueSub: {
          'Changes Requested': 0
        },
        rejectedSub: {
          'Rejected': 0
        },
        reworkSub: {
          'Rework': 0
        },
        completedSub: {
          'Approved': 0
        }
      }
    },
    overallStats: {
      totalTasks: 0,
      completed: 0,
      inProgress: 0,
      review: 0,
      issue: 0,
      rejected: 0,
      rework: 0,
      excluded: 0
    },
    rawTasks: [],
    excludedTasks: []
  };
  
  // Process Jira_Data sheet with exclusion filter
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[columnMap.KEY_COLUMN] || row[columnMap.ASSIGNEE_COLUMN]) {
      const teamName = row[columnMap.TEAM_COLUMN] ? row[columnMap.TEAM_COLUMN].toString().trim() : '';
      const assignee = row[columnMap.ASSIGNEE_COLUMN] ? row[columnMap.ASSIGNEE_COLUMN].toString().trim() : '';
      const status = row[columnMap.STATUS_COLUMN] ? row[columnMap.STATUS_COLUMN].toString() : '';
      const statusInfo = categorizeStatus(status);
      const language = CONFIG.TEAM_LANGUAGE_MAP[teamName] || 'Unknown';
      
      const task = {
        key: row[columnMap.KEY_COLUMN] || '',
        assignee: assignee,
        team: teamName,
        status: status,
        folderId: row[columnMap.FOLDER_ID_COLUMN] || '',
        summary: row[columnMap.SUMMARY_COLUMN] || '',
        taskDescription: row[columnMap.TASK_DESCRIPTION_COLUMN] || '',
        firstPrompt: row[columnMap.FIRST_PROMPT_COLUMN] || '',
        repoUrl: row[columnMap.REPO_URL_COLUMN] || '',
        githubLink: row[columnMap.GITHUB_LINK_COLUMN] || '',
        ltUrl: row[columnMap.LT_URL_COLUMN] || '',
        gdriveLink: row[columnMap.GDRIVE_LINK_COLUMN] || '',
        created: row[columnMap.CREATED_COLUMN] || '',
        updated: row[columnMap.UPDATED_COLUMN] || '',
        statusCategory: statusInfo.category,
        statusSubcategory: statusInfo.subcategory,
        language: language,
        rowNumber: i + 1,
        isFlexTeam: false,
        isExcluded: shouldExcludeAssignee(assignee)
      };
      
      // Check if assignee should be excluded
      if (shouldExcludeAssignee(assignee)) {
        allData.excludedTasks.push(task);
        allData.overallStats.excluded++;
        
        // Track excluded tasks by language
        if (allData.languageStats[language]) {
          allData.languageStats[language].excluded++;
        }
        
        console.log(`Excluding task ${task.key} assigned to ${assignee}`);
      } else {
        // Only include in statistics if not excluded
        allData.rawTasks.push(task);
        
        // Update overall stats
        allData.overallStats.totalTasks++;
        allData.overallStats[statusInfo.category]++;
      }
    }
  }
  
  // Process Flex team data with exclusion filter
  const flexTasks = getFlexTeamData();
  flexTasks.forEach(task => {
    // Check if Flex team assignee should be excluded
    if (shouldExcludeAssignee(task.assignee)) {
      task.isExcluded = true;
      allData.excludedTasks.push(task);
      allData.overallStats.excluded++;
      
      // Track excluded Flex tasks
      if (allData.languageStats['Multi-Language']) {
        allData.languageStats['Multi-Language'].excluded++;
      }
      
      console.log(`Excluding Flex task ${task.key} assigned to ${task.assignee}`);
    } else {
      // Only include in statistics if not excluded
      allData.rawTasks.push(task);
      
      // Update overall stats
      allData.overallStats.totalTasks++;
      allData.overallStats[task.statusCategory]++;
    }
  });
  
  // Create team data structures (only for non-excluded tasks)
  const teamMap = {};
  
  allData.rawTasks.forEach(task => {
    if (!teamMap[task.team]) {
      const teamConfig = CONFIG.TEAMS[task.team] || {
        name: task.team,
        leadName: 'Unknown',
        developers: 0,
        language: task.language,
        teamCode: `${task.language.substring(0, 2).toUpperCase()}-${task.team.toUpperCase()}`
      };
      
      teamMap[task.team] = {
        config: teamConfig,
        stats: {
          total: 0,
          completed: 0,
          inProgress: 0,
          review: 0,
          issue: 0,
          rejected: 0,
          rework: 0,
          excluded: 0
        },
        // Sub-category breakdowns
        inProgressSub: {},
        reviewSub: {},
        issueSub: {},
        rejectedSub: {},
        reworkSub: {},
        completedSub: {},
        tasks: []
      };
    }
    
    teamMap[task.team].tasks.push(task);
    teamMap[task.team].stats.total++;
    teamMap[task.team].stats[task.statusCategory]++;
    
    // Update subcategory stats for teams
    const subCategoryKey = task.statusCategory + 'Sub';
    if (!teamMap[task.team][subCategoryKey][task.statusSubcategory]) {
      teamMap[task.team][subCategoryKey][task.statusSubcategory] = 0;
    }
    teamMap[task.team][subCategoryKey][task.statusSubcategory]++;
  });
  
  // Add excluded task counts to teams
  allData.excludedTasks.forEach(task => {
    if (teamMap[task.team]) {
      teamMap[task.team].stats.excluded++;
    }
  });
  
  // Convert team map to array and sort
  allData.teams = Object.values(teamMap).sort((a, b) => {
    if (a.config.language !== b.config.language) {
      return a.config.language.localeCompare(b.config.language);
    }
    return a.config.name.localeCompare(b.config.name);
  });
  
  // Calculate language stats by aggregating from teams (only non-excluded tasks)
  ['Python', 'Rust', 'JavaScript', 'Multi-Language'].forEach(language => {
    const teamsForLanguage = allData.teams.filter(team => team.config.language === language);
    
    teamsForLanguage.forEach(team => {
      const langStats = allData.languageStats[language];
      
      // Add team stats to language stats
      langStats.totalTasks += team.stats.total;
      langStats.completed += team.stats.completed;
      langStats.inProgress += team.stats.inProgress;
      langStats.review += team.stats.review;
      langStats.issue += team.stats.issue;
      langStats.rejected += team.stats.rejected;
      langStats.rework += team.stats.rework;
      
      // Add subcategory stats
      Object.keys(team.inProgressSub || {}).forEach(subcat => {
        if (!langStats.inProgressSub[subcat]) langStats.inProgressSub[subcat] = 0;
        langStats.inProgressSub[subcat] += team.inProgressSub[subcat];
      });
      
      Object.keys(team.reviewSub || {}).forEach(subcat => {
        if (!langStats.reviewSub[subcat]) langStats.reviewSub[subcat] = 0;
        langStats.reviewSub[subcat] += team.reviewSub[subcat];
      });
      
      Object.keys(team.issueSub || {}).forEach(subcat => {
        if (!langStats.issueSub[subcat]) langStats.issueSub[subcat] = 0;
        langStats.issueSub[subcat] += team.issueSub[subcat];
      });
      
      Object.keys(team.rejectedSub || {}).forEach(subcat => {
        if (!langStats.rejectedSub[subcat]) langStats.rejectedSub[subcat] = 0;
        langStats.rejectedSub[subcat] += team.rejectedSub[subcat];
      });
      
      Object.keys(team.reworkSub || {}).forEach(subcat => {
        if (!langStats.reworkSub[subcat]) langStats.reworkSub[subcat] = 0;
        langStats.reworkSub[subcat] += team.reworkSub[subcat];
      });
      
      Object.keys(team.completedSub || {}).forEach(subcat => {
        if (!langStats.completedSub[subcat]) langStats.completedSub[subcat] = 0;
        langStats.completedSub[subcat] += team.completedSub[subcat];
      });
    });
  });
  
  return allData;
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
    
    // Collect detailed task data including Flex
    const taskData = collectDetailedTaskData();
    
    // Build export sheet
    buildTaskExportSheet(exportSheet, taskData);
    
    // Show completion message
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Task Data Export Complete!',
      `Task data has been exported to the "${exportSheetName}" sheet\n\n` +
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
 * Collects detailed task data for export
 */
function collectDetailedTaskData() {
  const allData = collectAllTeamsData();
  
  return {
    teams: allData.teams,
    allTasks: allData.rawTasks,
    excludedTasks: allData.excludedTasks,
    totalTasks: allData.rawTasks.length,
    languages: ['Python', 'Rust', 'JavaScript', 'Multi-Language']
  };
}

/**
 * Creates column mapping from headers
 */
function getColumnMap(headers) {
  const columnMap = {};
  
  // Map each configured column to its index
  Object.keys(CONFIG).forEach(key => {
    if (key.endsWith('_COLUMN')) {
      const columnName = CONFIG[key];
      const index = findColumnIndex(headers, columnName);
      columnMap[key] = index;
    }
  });
  
  return columnMap;
}

/**
 * Helper function to find column index
 */
function findColumnIndex(headers, columnName) {
  return headers.findIndex(h => h.toString().toLowerCase().trim() === columnName.toLowerCase().trim());
}

/**
 * Categorizes task completion status and returns both category and subcategory
 */
function categorizeStatus(status) {
  const normalizedStatus = status.toString().trim();
  
  for (const [category, values] of Object.entries(CONFIG.STATUS_VALUES)) {
    if (values.some(v => v === normalizedStatus)) {
      const categoryName = category === 'IN_PROGRESS' ? 'inProgress' :
                          category === 'REVIEW' ? 'review' :
                          category === 'ISSUE' ? 'issue' :
                          category === 'REJECTED' ? 'rejected' :
                          category === 'REWORK' ? 'rework' :
                          category === 'COMPLETED' ? 'completed' : 'inProgress';
      
      return {
        category: categoryName,
        subcategory: normalizedStatus
      };
    }
  }
  
  return {
    category: 'inProgress',
    subcategory: normalizedStatus
  };
}

// ============================
// DASHBOARD BUILDING FUNCTIONS
// ============================

/**
 * Builds the dashboard layout
 */
function buildDashboard(sheet, data) {
  let currentRow = 1;
  
  // 1. Dashboard Header
  currentRow = createDashboardHeader(sheet, currentRow);
  
  // 2. Overall Progress Section
  currentRow = createOverallProgress(sheet, data.overallStats, currentRow);
  
  // 3. Team Statistics Table
  currentRow = createTeamStatisticsTable(sheet, data.teams, currentRow);
  
  // 4. Language Breakdown Section
  currentRow = createLanguageBreakdown(sheet, data.languageStats, currentRow);
  
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
  headerCell.setValue('PR Writer V2 DASHBOARD');
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
  projectInfo.setValue('Python (10,000 tasks) | Rust (10,000 tasks) | JavaScript (10,000 tasks) | Flex Team (Multi-Language) | Total Target: 30,000');
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
 * UPDATED: Creates overall progress section with exclusion info
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
  
  // Create progress cards with exclusion info
  const cards = [
    {
      label: 'Project Target',
      value: CONFIG.PROJECT_TARGET.toLocaleString(),
      color: CONFIG.COLORS.HEADER
    },
    {
      label: 'Active Tasks',
      value: stats.totalTasks.toLocaleString(),
      color: CONFIG.COLORS.HEADER
    },
    {
      label: 'Excluded Tasks',
      value: stats.excluded.toLocaleString(),
      color: CONFIG.COLORS.NOT_STARTED
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
      label: 'Rework Required',
      value: stats.rework.toLocaleString(),
      color: CONFIG.COLORS.REWORK
    },
    {
      label: 'Issues/Rejected',
      value: (stats.issue + stats.rejected).toLocaleString(),
      color: CONFIG.COLORS.REJECTED
    }
  ];
  
  // Display cards in two rows
  const cardsPerRow = 4;
  for (let i = 0; i < cards.length; i++) {
    const card = cards[i];
    const row = Math.floor(i / cardsPerRow);
    const col = (i % cardsPerRow) * 3 + 1;
    
    // Label
    sheet.getRange(progressRow + row * 3, col).setValue(card.label);
    sheet.getRange(progressRow + row * 3, col).setFontWeight('bold');
    
    // Value
    sheet.getRange(progressRow + row * 3 + 1, col).setValue(card.value);
    sheet.getRange(progressRow + row * 3 + 1, col).setFontSize(14);
    sheet.getRange(progressRow + row * 3 + 1, col).setFontColor(card.color);
    sheet.getRange(progressRow + row * 3 + 1, col).setFontWeight('bold');
  }
  
  // Progress bar
  const progressBarRow = progressRow + 6;
  sheet.getRange(progressBarRow, 1).setValue('Overall Completion:');
  sheet.getRange(progressBarRow, 1).setFontWeight('bold');
  
  sheet.getRange(progressBarRow, 2, 1, 10).merge();
  createProgressBar(sheet, progressBarRow, 2, parseFloat(completionPercentage));
  
  // Exclusion info
  const exclusionRow = progressBarRow + 2;
  sheet.getRange(exclusionRow, 1, 1, 12).merge();
  const exclusionInfo = sheet.getRange(exclusionRow, 1);
  exclusionInfo.setValue(`Note: ${stats.excluded} tasks excluded from statistics (assigned to team leads and specified exclusions)`);
  exclusionInfo.setFontStyle('italic');
  exclusionInfo.setFontSize(10);
  exclusionInfo.setFontColor('#666666');
  exclusionInfo.setBackground('#fff3e0');
  
  return exclusionRow + 3;
}

/**
 * Creates language breakdown section with stats aggregated from all teams per language
 */
function createLanguageBreakdown(sheet, languageStats, startRow) {
  // Section header
  sheet.getRange(startRow, 1, 1, 12).merge();
  const sectionHeader = sheet.getRange(startRow, 1);
  sectionHeader.setValue('LANGUAGE BREAKDOWN');
  sectionHeader.setFontSize(16);
  sectionHeader.setFontWeight('bold');
  sectionHeader.setBackground(CONFIG.COLORS.LIGHT_GRAY);
  
  let currentRow = startRow + 2;
  
  // Create breakdown for each language including Multi-Language
  ['Python', 'Rust', 'JavaScript', 'Multi-Language'].forEach(language => {
    const stats = languageStats[language];
    const target = CONFIG.LANGUAGE_TARGETS[language] || 0;
    const completion = target > 0 ? (stats.completed / target * 100).toFixed(1) : 'N/A';
    const progressPercentage = stats.totalTasks > 0 ? (stats.completed / stats.totalTasks * 100).toFixed(1) : '0.0';
    
    // Language header with detailed info
    sheet.getRange(currentRow, 1, 1, 12).merge();
    const langHeader = sheet.getRange(currentRow, 1);
    
    let headerText;
    if (language === 'Multi-Language') {
      headerText = `${language.toUpperCase()} (FLEX TEAM) - Active Tasks: ${stats.totalTasks} | Excluded: ${stats.excluded} | Completed: ${stats.completed} | Progress: ${progressPercentage}%`;
    } else {
      headerText = `${language.toUpperCase()} - Target: ${target} | Active Tasks: ${stats.totalTasks} | Excluded: ${stats.excluded} | Completed: ${stats.completed} | Progress: ${progressPercentage}% | Target Completion: ${completion}%`;
    }
    
    langHeader.setValue(headerText);
    langHeader.setFontWeight('bold');
    
    // Set color based on language
    if (language === 'Multi-Language') {
      langHeader.setBackground(CONFIG.COLORS.FLEX);
      langHeader.setFontColor('#ffffff');
    } else {
      langHeader.setBackground(CONFIG.COLORS[language.toUpperCase()]);
      if (language === 'JavaScript') {
        langHeader.setFontColor('#000000');
      } else {
        langHeader.setFontColor('#ffffff');
      }
    }
    
    currentRow++;
    
    // Main categories summary
    sheet.getRange(currentRow, 1, 1, 12).merge();
    const summaryLabel = sheet.getRange(currentRow, 1);
    summaryLabel.setValue(`Status Summary (${language === 'Multi-Language' ? 'Flex Team' : 'All Teams Combined'}):`);
    summaryLabel.setFontWeight('bold');
    summaryLabel.setFontSize(11);
    summaryLabel.setFontStyle('italic');
    summaryLabel.setBackground('#f0f0f0');
    
    currentRow++;
    
    // Main status breakdown
    const statusData = [
      ['In Progress', stats.inProgress, CONFIG.COLORS.IN_PROGRESS],
      ['Review', stats.review, CONFIG.COLORS.REVIEW],
      ['Completed', stats.completed, CONFIG.COLORS.COMPLETED],
      ['Rework Required', stats.rework, CONFIG.COLORS.REWORK],
      ['Issues', stats.issue, CONFIG.COLORS.REWORK],
      ['Rejected', stats.rejected, CONFIG.COLORS.REJECTED]
    ];
    
    statusData.forEach((item, index) => {
      const col = (index * 2) + 1;
      if (col <= 11) {
        sheet.getRange(currentRow, col).setValue(item[0]);
        sheet.getRange(currentRow, col).setFontSize(10);
        sheet.getRange(currentRow, col).setFontWeight('bold');
        
        const countCell = sheet.getRange(currentRow, col + 1);
        countCell.setValue(item[1]);
        countCell.setFontWeight('bold');
        countCell.setFontColor(item[2]);
        countCell.setHorizontalAlignment('center');
        countCell.setFontSize(12);
      }
    });
    
    currentRow++;
    
    // Detailed subcategory breakdowns
    if (stats.inProgress > 0 && stats.inProgressSub) {
      currentRow = createSubcategorySection(sheet, currentRow, `IN PROGRESS DETAILS (${language === 'Multi-Language' ? 'Flex Team' : 'All Teams'}):`, stats.inProgressSub, CONFIG.COLORS.IN_PROGRESS);
    }
    
    if (stats.review > 0 && stats.reviewSub) {
      currentRow = createSubcategorySection(sheet, currentRow, `REVIEW DETAILS (${language === 'Multi-Language' ? 'Flex Team' : 'All Teams'}):`, stats.reviewSub, CONFIG.COLORS.REVIEW);
    }
    
    if (stats.rework > 0 && stats.reworkSub) {
      currentRow = createSubcategorySection(sheet, currentRow, `REWORK DETAILS (${language === 'Multi-Language' ? 'Flex Team' : 'All Teams'}):`, stats.reworkSub, CONFIG.COLORS.REWORK);
    }
    
    if (stats.issue > 0 && stats.issueSub) {
      currentRow = createSubcategorySection(sheet, currentRow, `ISSUES DETAILS (${language === 'Multi-Language' ? 'Flex Team' : 'All Teams'}):`, stats.issueSub, CONFIG.COLORS.REWORK);
    }
    
    if (stats.rejected > 0 && stats.rejectedSub) {
      currentRow = createSubcategorySection(sheet, currentRow, `REJECTED DETAILS (${language === 'Multi-Language' ? 'Flex Team' : 'All Teams'}):`, stats.rejectedSub, CONFIG.COLORS.REJECTED);
    }
    
    if (stats.completed > 0 && stats.completedSub) {
      currentRow = createSubcategorySection(sheet, currentRow, `COMPLETED DETAILS (${language === 'Multi-Language' ? 'Flex Team' : 'All Teams'}):`, stats.completedSub, CONFIG.COLORS.COMPLETED);
    }
    
    currentRow += 2;
  });
  
  return currentRow + 1;
}

/**
 * Creates a subcategory section for detailed status breakdown
 */
function createSubcategorySection(sheet, startRow, title, subcategoryStats, color) {
  // Section title
  sheet.getRange(startRow, 1, 1, 12).merge();
  const sectionTitle = sheet.getRange(startRow, 1);
  sectionTitle.setValue(title);
  sectionTitle.setFontWeight('bold');
  sectionTitle.setFontSize(10);
  sectionTitle.setFontStyle('italic');
  sectionTitle.setBackground('#e8eaf6');
  sectionTitle.setFontColor(color);
  
  let currentRow = startRow + 1;
  
  // Display subcategories
  let col = 1;
  Object.entries(subcategoryStats).forEach(([subcategory, count]) => {
    if (count > 0) {
      // Subcategory name
      sheet.getRange(currentRow, col).setValue(subcategory);
      sheet.getRange(currentRow, col).setFontSize(9);
      sheet.getRange(currentRow, col).setWrap(true);
      
      // Count
      const countCell = sheet.getRange(currentRow, col + 1);
      countCell.setValue(count);
      countCell.setFontWeight('bold');
      countCell.setFontColor(color);
      countCell.setHorizontalAlignment('center');
      
      col += 2;
      if (col > 10) {
        col = 1;
        currentRow++;
      }
    }
  });
  
  return currentRow + 2;
}

/**
 * Creates team statistics table
 */
function createTeamStatisticsTable(sheet, teams, startRow) {
  // Section header
  sheet.getRange(startRow, 1, 1, 18).merge();
  const sectionHeader = sheet.getRange(startRow, 1);
  sectionHeader.setValue('TEAM STATISTICS');
  sectionHeader.setFontSize(16);
  sectionHeader.setFontWeight('bold');
  sectionHeader.setBackground(CONFIG.COLORS.LIGHT_GRAY);
  
  // Table headers
  const headerRow = startRow + 2;
  const headers = [
    'Team Name', 'Language', 'Lead', 'Developers', 'Team Code',
    'Active Tasks', 'Excluded', 'Completed', 'In Progress', 'Review', 'Rework Required',
    'Issues', 'Rejected', 'Progress %', 'Performance', 'Status', 'Details', 'Source'
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
    const config = team.config;
    
    const progressPercentage = stats.total > 0 
      ? (stats.completed / stats.total * 100).toFixed(1)
      : 0;
    
    // Performance indicator
    let performance = 'Average';
    const progress = parseFloat(progressPercentage);
    if (progress >= 80) performance = 'Excellent';
    else if (progress >= 60) performance = 'Good';
    else if (progress < 30) performance = 'Needs Attention';
    
    // Status indicator
    let status = 'Active';
    if (stats.total === 0) status = 'No Tasks';
    else if (progress === 100) status = 'Complete';
    else if (stats.rework > 0) status = 'Rework Required';
    else if (stats.issue > stats.completed) status = 'Issues';
    
    // Create details string for subcategories
    let details = '';
    if (stats.inProgress > 0) {
      const inProgressDetails = Object.entries(team.inProgressSub || {})
        .filter(([key, value]) => value > 0)
        .map(([key, value]) => `${key}: ${value}`)
        .join(', ');
      if (inProgressDetails) details += `In Progress: ${inProgressDetails}; `;
    }
    
    if (stats.review > 0) {
      const reviewDetails = Object.entries(team.reviewSub || {})
        .filter(([key, value]) => value > 0)
        .map(([key, value]) => `${key}: ${value}`)
        .join(', ');
      if (reviewDetails) details += `Review: ${reviewDetails}; `;
    }
    
    if (stats.rework > 0) {
      const reworkDetails = Object.entries(team.reworkSub || {})
        .filter(([key, value]) => value > 0)
        .map(([key, value]) => `${key}: ${value}`)
        .join(', ');
      if (reworkDetails) details += `Rework: ${reworkDetails}; `;
    }
    
    if (stats.issue > 0) {
      const issueDetails = Object.entries(team.issueSub || {})
        .filter(([key, value]) => value > 0)
        .map(([key, value]) => `${key}: ${value}`)
        .join(', ');
      if (issueDetails) details += `Issues: ${issueDetails}; `;
    }
    
    if (stats.rejected > 0) {
      const rejectedDetails = Object.entries(team.rejectedSub || {})
        .filter(([key, value]) => value > 0)
        .map(([key, value]) => `${key}: ${value}`)
        .join(', ');
      if (rejectedDetails) details += `Rejected: ${rejectedDetails}; `;
    }
    
    // Determine data source
    const dataSource = config.name === 'Flex' ? 'External Sheet' : 'Jira_Data';
    
    const rowData = [
      config.name,
      config.language,
      config.leadName,
      config.developers,
      config.teamCode,
      stats.total,
      stats.excluded || 0,
      stats.completed,
      stats.inProgress,
      stats.review,
      stats.rework,
      stats.issue,
      stats.rejected,
      `${progressPercentage}%`,
      performance,
      status,
      details.trim(),
      dataSource
    ];
    
    rowData.forEach((value, index) => {
      const cell = sheet.getRange(dataRow, index + 1);
      cell.setValue(value);
      
      // Style specific columns
      if (index === 1) { // Language column
        if (config.language === 'Multi-Language') {
          cell.setFontColor(CONFIG.COLORS.FLEX);
        } else if (CONFIG.COLORS[config.language.toUpperCase()]) {
          cell.setFontColor(CONFIG.COLORS[config.language.toUpperCase()]);
        }
        cell.setFontWeight('bold');
      } else if (index === 6) { // Excluded column
        if (parseInt(value) > 0) {
          cell.setFontColor(CONFIG.COLORS.NOT_STARTED);
          cell.setFontWeight('bold');
        }
      } else if (index === 7) { // Completed column
        cell.setFontColor(CONFIG.COLORS.COMPLETED);
        cell.setFontWeight('bold');
      } else if (index === 10) { // Rework column
        if (parseInt(value) > 0) {
          cell.setFontColor(CONFIG.COLORS.REWORK);
          cell.setFontWeight('bold');
        }
      } else if (index === 11) { // Issues column
        if (parseInt(value) > 0) {
          cell.setFontColor(CONFIG.COLORS.REWORK);
          cell.setFontWeight('bold');
        }
      } else if (index === 12) { // Rejected column
        if (parseInt(value) > 0) {
          cell.setFontColor(CONFIG.COLORS.REJECTED);
          cell.setFontWeight('bold');
        }
      } else if (index === 13) { // Progress column
        const progress = parseFloat(value);
        if (progress >= 80) {
          cell.setFontColor(CONFIG.COLORS.COMPLETED);
        } else if (progress >= 50) {
          cell.setFontColor(CONFIG.COLORS.IN_PROGRESS);
        } else {
          cell.setFontColor(CONFIG.COLORS.REJECTED);
        }
        cell.setFontWeight('bold');
      } else if (index === 14) { // Performance column
        if (value === 'Excellent') {
          cell.setFontColor(CONFIG.COLORS.COMPLETED);
        } else if (value === 'Good') {
          cell.setFontColor(CONFIG.COLORS.IN_PROGRESS);
        } else if (value === 'Needs Attention') {
          cell.setFontColor(CONFIG.COLORS.REJECTED);
        }
        cell.setFontWeight('bold');
      } else if (index === 15) { // Status column
        if (value === 'Complete') {
          cell.setFontColor(CONFIG.COLORS.COMPLETED);
        } else if (value === 'Rework Required') {
          cell.setFontColor(CONFIG.COLORS.REWORK);
        } else if (value === 'Issues') {
          cell.setFontColor(CONFIG.COLORS.REJECTED);
        } else if (value === 'No Tasks') {
          cell.setFontColor(CONFIG.COLORS.NOT_STARTED);
        }
        cell.setFontWeight('bold');
      } else if (index === 16) { // Details column
        cell.setFontSize(8);
        cell.setWrap(true);
      } else if (index === 17) { // Source column
        if (value === 'External Sheet') {
          cell.setFontColor(CONFIG.COLORS.FLEX);
          cell.setFontWeight('bold');
        }
      }
      
      cell.setBorder(true, true, true, true, false, false);
    });
    
    // Alternate row coloring
    if (dataRow % 2 === 0) {
      sheet.getRange(dataRow, 1, 1, headers.length).setBackground(CONFIG.COLORS.LIGHT_GRAY);
    }
    
    // Special highlighting for Flex team
    if (config.name === 'Flex') {
      sheet.getRange(dataRow, 1, 1, headers.length).setBackground('#f3e5f5');
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
    
    // Color based on language
    if (team.config.language === 'Multi-Language') {
      nameCell.setFontColor(CONFIG.COLORS.FLEX);
    } else if (CONFIG.COLORS[team.config.language.toUpperCase()]) {
      nameCell.setFontColor(CONFIG.COLORS[team.config.language.toUpperCase()]);
    }
    
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
  
  ['Python', 'Rust', 'JavaScript', 'Multi-Language'].forEach(language => {
    const stats = data.languageStats[language];
    const target = CONFIG.LANGUAGE_TARGETS[language] || 0;
    const completion = target > 0 ? (stats.completed / target * 100) : 0;
    
    // Language name
    sheet.getRange(currentRow, 1).setValue(language);
    sheet.getRange(currentRow, 1).setFontWeight('bold');
    
    if (language === 'Multi-Language') {
      sheet.getRange(currentRow, 1).setFontColor(CONFIG.COLORS.FLEX);
    } else if (CONFIG.COLORS[language.toUpperCase()]) {
      sheet.getRange(currentRow, 1).setFontColor(CONFIG.COLORS[language.toUpperCase()]);
    }
    
    // Target info
    if (target > 0) {
      sheet.getRange(currentRow, 2).setValue(`Target: ${target}`);
    } else {
      sheet.getRange(currentRow, 2).setValue('No specific target');
    }
    
    // Progress bar
    sheet.getRange(currentRow, 3, 1, 7).merge();
    if (target > 0) {
      createProgressBar(sheet, currentRow, 3, completion);
    } else {
      // For Multi-Language, show completion percentage based on total tasks
      const taskCompletion = stats.totalTasks > 0 ? (stats.completed / stats.totalTasks * 100) : 0;
      createProgressBar(sheet, currentRow, 3, taskCompletion);
    }
    
    // Completion stats
    if (target > 0) {
      sheet.getRange(currentRow, 10).setValue(`${completion.toFixed(1)}%`);
      sheet.getRange(currentRow, 11).setValue(`${stats.completed}/${target}`);
      sheet.getRange(currentRow, 12).setValue(`Remaining: ${target - stats.completed}`);
    } else {
      const taskCompletion = stats.totalTasks > 0 ? (stats.completed / stats.totalTasks * 100) : 0;
      sheet.getRange(currentRow, 10).setValue(`${taskCompletion.toFixed(1)}%`);
      sheet.getRange(currentRow, 11).setValue(`${stats.completed}/${stats.totalTasks}`);
      sheet.getRange(currentRow, 12).setValue(`Pending: ${stats.totalTasks - stats.completed}`);
    }
    
    sheet.getRange(currentRow, 10).setFontWeight('bold');
    
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
  sheet.getRange(startRow, 1, 1, 19).merge(); // Extended for exclusion info
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
  sheet.getRange(startRow + 1, 1, 1, 19).merge();
  const timestampCell = sheet.getRange(startRow + 1, 1);
  timestampCell.setValue('Exported: ' + new Date().toLocaleString());
  timestampCell.setFontStyle('italic');
  timestampCell.setHorizontalAlignment('center');
  
  return startRow + 3;
}

/**
 * Creates task export summary with exclusion info
 */
function createTaskExportSummary(sheet, startRow, taskData) {
  // Section header
  sheet.getRange(startRow, 1, 1, 19).merge();
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
    JavaScript: taskData.allTasks.filter(t => t.language === 'JavaScript').length,
    'Multi-Language': taskData.allTasks.filter(t => t.language === 'Multi-Language').length
  };
  
  const flexTasks = taskData.allTasks.filter(t => t.isFlexTeam).length;
  const jiraTasks = taskData.allTasks.filter(t => !t.isFlexTeam).length;
  const excludedTasks = taskData.excludedTasks ? taskData.excludedTasks.length : 0;
  
  const summaryData = [
    ['Active Tasks Found:', taskData.totalTasks],
    ['Excluded Tasks:', excludedTasks],
    ['Total Tasks (Active + Excluded):', taskData.totalTasks + excludedTasks],
    ['Jira_Data Tasks:', jiraTasks],
    ['Flex Team Tasks:', flexTasks],
    ['Python Tasks:', languageCounts.Python],
    ['Rust Tasks:', languageCounts.Rust],
    ['JavaScript Tasks:', languageCounts.JavaScript],
    ['Multi-Language Tasks (Flex):', languageCounts['Multi-Language']],
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
 * Creates tasks table with exclusion handling
 */
function createTasksTable(sheet, startRow, taskData) {
  // Section header
  sheet.getRange(startRow, 1, 1, 19).merge();
  const sectionHeader = sheet.getRange(startRow, 1);
  sectionHeader.setValue('ALL ACTIVE TASKS (Jira_Data + Flex Team - Exclusions Applied)');
  sectionHeader.setFontSize(16);
  sectionHeader.setFontWeight('bold');
  sectionHeader.setBackground(CONFIG.COLORS.LIGHT_GRAY);
  
  // Table headers
  const headerRow = startRow + 2;
  const headers = [
    'Key', 'Assignee', 'Team', 'Language', 'Status', 'Summary', 
    'Task Description', 'Repo URL', 'GitHub Link', 'LT URL', 'GDrive Link',
    'Created', 'Updated', 'Folder/ID', 'Row #', 'Source', 'Reviewer ID', 'Repo Approval', 'Exclusion Status'
  ];
  
  headers.forEach((header, index) => {
    const cell = sheet.getRange(headerRow, index + 1);
    cell.setValue(header);
    cell.setFontWeight('bold');
    cell.setBackground(CONFIG.COLORS.HEADER);
    cell.setFontColor('#ffffff');
    cell.setBorder(true, true, true, true, false, false);
  });
  
  // Data rows - only active tasks
  let dataRow = headerRow + 1;
  taskData.allTasks.forEach(task => {
    const rowData = [
      task.key,
      task.assignee,
      task.team,
      task.language,
      task.status,
      task.summary,
      task.taskDescription,
      task.repoUrl,
      task.githubLink || 'N/A',
      task.ltUrl,
      task.gdriveLink,
      task.created,
      task.updated || 'N/A',
      task.folderId || 'N/A',
      task.rowNumber,
      task.isFlexTeam ? 'External Sheet' : 'Jira_Data',
      task.reviewerId || 'N/A',
      task.repoApproval || 'N/A',
      'Active'
    ];
    
    rowData.forEach((value, index) => {
      const cell = sheet.getRange(dataRow, index + 1);
      cell.setValue(value);
      
      // Style based on content
      if (index === 2) { // Team column
        if (task.team === 'Flex') {
          cell.setFontColor(CONFIG.COLORS.FLEX);
          cell.setFontWeight('bold');
        }
      } else if (index === 3) { // Language column
        if (task.language === 'Multi-Language') {
          cell.setFontColor(CONFIG.COLORS.FLEX);
        } else if (task.language !== 'Unknown' && CONFIG.COLORS[task.language.toUpperCase()]) {
          cell.setFontColor(CONFIG.COLORS[task.language.toUpperCase()]);
        }
        cell.setFontWeight('bold');
      } else if (index === 4) { // Status column
        const statusColor = getStatusColor(task.statusCategory);
        cell.setFontColor(statusColor);
        if (task.statusCategory === 'completed') {
          cell.setFontWeight('bold');
        }
      } else if (index === 15) { // Source column
        if (value === 'External Sheet') {
          cell.setFontColor(CONFIG.COLORS.FLEX);
          cell.setFontWeight('bold');
        }
      } else if (index === 18) { // Exclusion Status column
        cell.setFontColor(CONFIG.COLORS.COMPLETED);
        cell.setFontWeight('bold');
      }
      
      cell.setBorder(true, true, true, true, false, false);
    });
    
    // Alternate row coloring
    if (dataRow % 2 === 0) {
      sheet.getRange(dataRow, 1, 1, headers.length).setBackground(CONFIG.COLORS.LIGHT_GRAY);
    }
    
    // Special highlighting for Flex team tasks
    if (task.isFlexTeam) {
      sheet.getRange(dataRow, 1, 1, headers.length).setBackground('#f3e5f5');
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
  const barLength = 30;
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
    case 'issue': return CONFIG.COLORS.REWORK;
    case 'rejected': return CONFIG.COLORS.REJECTED;
    default: return CONFIG.COLORS.IN_PROGRESS;
  }
}

/**
 * Formats the dashboard
 */
function formatDashboard(sheet) {
  // Auto-resize columns
  sheet.autoResizeColumns(1, 19);
  
  // Set minimum column widths
  sheet.setColumnWidth(1, 250); // Team name column
  sheet.setColumnWidth(2, 120); // Language column
  sheet.setColumnWidth(3, 150); // Lead column
  sheet.setColumnWidth(17, 300); // Details column
  sheet.setColumnWidth(18, 120); // Source column
  
  // Add borders to the entire data range
  const dataRange = sheet.getDataRange();
  dataRange.setBorder(true, true, true, true, false, false, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  
  // Center align numeric columns
  for (let col = 4; col <= 18; col++) {
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
  sheet.autoResizeColumns(1, 19);
  
  // Set specific column widths
  sheet.setColumnWidth(1, 120); // Key
  sheet.setColumnWidth(2, 120); // Assignee
  sheet.setColumnWidth(3, 100); // Team
  sheet.setColumnWidth(6, 300); // Summary
  sheet.setColumnWidth(7, 400); // Task description
  sheet.setColumnWidth(8, 300); // Repo URL
  sheet.setColumnWidth(9, 300); // GitHub Link
  sheet.setColumnWidth(16, 120); // Source
  sheet.setColumnWidth(17, 120); // Reviewer ID
  sheet.setColumnWidth(18, 120); // Repo Approval
  sheet.setColumnWidth(19, 120); // Exclusion Status
  
  // Add borders to the entire data range
  const dataRange = sheet.getDataRange();
  dataRange.setBorder(true, true, true, true, false, false, CONFIG.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  
  // Freeze header rows
  sheet.setFrozenRows(3);
}

// ============================
// DEBUG AND UTILITY FUNCTIONS
// ============================

/**
 * Debug function to show excluded assignees
 */
function debugExcludedAssignees() {
  try {
    console.log('Starting excluded assignees debug...');
    const allData = collectAllTeamsData();
    
    console.log('=== EXCLUDED ASSIGNEES DEBUG ===');
    console.log('Configured exclusion list:', CONFIG.EXCLUDED_ASSIGNEES);
    console.log('Total excluded tasks:', allData.excludedTasks.length);
    console.log('Total active tasks:', allData.rawTasks.length);
    
    // Show excluded tasks by assignee
    const excludedByAssignee = {};
    allData.excludedTasks.forEach(task => {
      if (!excludedByAssignee[task.assignee]) {
        excludedByAssignee[task.assignee] = 0;
      }
      excludedByAssignee[task.assignee]++;
    });
    
    console.log('\nExcluded tasks by assignee:');
    Object.entries(excludedByAssignee).forEach(([assignee, count]) => {
      console.log(`  ${assignee}: ${count} tasks`);
    });
    
    // Show excluded tasks by team/language
    const excludedByLanguage = {};
    allData.excludedTasks.forEach(task => {
      if (!excludedByLanguage[task.language]) {
        excludedByLanguage[task.language] = 0;
      }
      excludedByLanguage[task.language]++;
    });
    
    console.log('\nExcluded tasks by language:');
    Object.entries(excludedByLanguage).forEach(([language, count]) => {
      console.log(`  ${language}: ${count} tasks`);
    });
    
    SpreadsheetApp.getUi().alert(
      'Exclusion Debug Complete',
      `Exclusion debug complete!\n\n` +
      `Total excluded tasks: ${allData.excludedTasks.length}\n` +
      `Total active tasks: ${allData.rawTasks.length}\n` +
      `Excluded assignees found: ${Object.keys(excludedByAssignee).length}\n\n` +
      `Check the console logs for detailed breakdown by assignee and language.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Debug error:', error);
    SpreadsheetApp.getUi().alert('Debug Error: ' + error.toString());
  }
}

/**
 * Test function to verify Flex team connection
 */
function testFlexTeamConnection() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Test connection to external sheet
    const flexTasks = getFlexTeamData();
    
    if (flexTasks.length === 0) {
      ui.alert(
        'Flex Team Connection Test',
        'No data found in Flex team sheet. Please check:\n' +
        '1. Sheet ID is correct in CONFIG.FLEX_SHEET_ID\n' +
        '2. Sheet name is correct in CONFIG.FLEX_SHEET_NAME\n' +
        '3. Sheet has proper sharing permissions\n' +
        '4. Sheet contains data with the expected columns',
        ui.ButtonSet.OK
      );
    } else {
      // Show sample of data found
      const sampleTask = flexTasks[0];
      const statusBreakdown = {};
      
      flexTasks.forEach(task => {
        if (!statusBreakdown[task.statusCategory]) {
          statusBreakdown[task.statusCategory] = 0;
        }
        statusBreakdown[task.statusCategory]++;
      });
      
      const statusSummary = Object.entries(statusBreakdown)
        .map(([status, count]) => `${status}: ${count}`)
        .join(', ');
      
      ui.alert(
        'Flex Team Connection Test - SUCCESS!',
        `Found ${flexTasks.length} tasks from Flex team.\n\n` +
        `Status breakdown: ${statusSummary}\n\n` +
        `Sample task:\n` +
        `- ID: ${sampleTask.key}\n` +
        `- Worker: ${sampleTask.assignee}\n` +
        `- Language: ${sampleTask.language}\n` +
        `- Status: ${sampleTask.status}`,
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('Error testing Flex team connection:', error);
    SpreadsheetApp.getUi().alert(
      'Flex Team Connection Test - ERROR',
      `Failed to connect to Flex team sheet:\n${error.toString()}\n\n` +
      'Please check the configuration and try again.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Function to configure Flex team settings
 */
function configureFlexTeamSettings() {
  const ui = SpreadsheetApp.getUi();
  
  // Get current settings
  const currentSheetId = CONFIG.FLEX_SHEET_ID;
  const currentSheetName = CONFIG.FLEX_SHEET_NAME;
  
  ui.alert(
    'Flex Team Configuration',
    `Current settings:\n` +
    `Sheet ID: ${currentSheetId}\n` +
    `Sheet Name: ${currentSheetName}\n\n` +
    `To update these settings:\n` +
    `1. Edit the CONFIG object at the top of the script\n` +
    `2. Update FLEX_SHEET_ID with your Google Sheet ID\n` +
    `3. Update FLEX_SHEET_NAME with your sheet tab name\n` +
    `4. Update the Flex team lead name and developer count in CONFIG.TEAMS.Flex\n\n` +
    `The external sheet should have these columns:\n` +
    `task_id, worker_id, Language, task_start_date, repo_url, task_description, first_prompt, LT_url, gdrive_link, reviewer_id, repo_approval, task_completion_status, comments, Review_Status`,
    ui.ButtonSet.OK
  );
}

/**
 * Validates the Jira_Data sheet structure
 */
function validateJiraDataSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(CONFIG.DATA_SHEET_NAME);
    
    if (!dataSheet) {
      SpreadsheetApp.getUi().alert(
        'Validation Failed',
        `Sheet '${CONFIG.DATA_SHEET_NAME}' not found. Please create a sheet named '${CONFIG.DATA_SHEET_NAME}' with the required columns.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const data = dataSheet.getDataRange().getValues();
    
    if (data.length === 0) {
      SpreadsheetApp.getUi().alert(
        'Validation Failed',
        'The Jira_Data sheet is empty. Please add data with proper headers.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const headers = data[0];
    const requiredColumns = [
      CONFIG.KEY_COLUMN,
      CONFIG.ASSIGNEE_COLUMN,
      CONFIG.TEAM_COLUMN,
      CONFIG.STATUS_COLUMN,
      CONFIG.SUMMARY_COLUMN
    ];
    
    const missingColumns = requiredColumns.filter(col => 
      findColumnIndex(headers, col) === -1
    );
    
    if (missingColumns.length > 0) {
      SpreadsheetApp.getUi().alert(
        'Validation Failed',
        `Missing required columns: ${missingColumns.join(', ')}\n\nFound columns: ${headers.join(', ')}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Count data rows and exclusions
    const allData = collectAllTeamsData();
    const dataRows = data.length - 1;
    
    // Test Flex connection
    const flexTasks = getFlexTeamData();
    
    SpreadsheetApp.getUi().alert(
      'Validation Successful',
      `Jira_Data sheet structure is valid!\n\n` +
      `Found ${dataRows} Jira data rows with all required columns.\n` +
      `Found ${flexTasks.length} Flex team tasks from external sheet.\n` +
      `Active tasks after exclusions: ${allData.rawTasks.length}\n` +
      `Excluded tasks: ${allData.excludedTasks.length}\n\n` +
      `Total tasks: ${dataRows + flexTasks.length}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Error validating data sheets:', error);
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Debug function to test language stats calculation
 */
function debugLanguageStats() {
  try {
    console.log('Starting language stats debug...');
    const allData = collectAllTeamsData();
    
    console.log('=== LANGUAGE STATS DEBUG ===');
    
    // Show raw tasks by language including Multi-Language
    ['Python', 'Rust', 'JavaScript', 'Multi-Language'].forEach(language => {
      const tasksForLanguage = allData.rawTasks.filter(task => task.language === language);
      const excludedForLanguage = allData.excludedTasks.filter(task => task.language === language);
      console.log(`\n${language} Active Tasks (${tasksForLanguage.length} total):`);
      
      tasksForLanguage.forEach(task => {
        console.log(`  - ${task.key}: ${task.status} (${task.statusCategory}/${task.statusSubcategory}) [Team: ${task.team}] [Source: ${task.isFlexTeam ? 'External' : 'Jira'}]`);
      });
      
      console.log(`\n${language} Excluded Tasks (${excludedForLanguage.length} total):`);
      excludedForLanguage.forEach(task => {
        console.log(`  - ${task.key}: ${task.assignee} [EXCLUDED]`);
      });
      
      console.log(`\n${language} Final Stats:`, allData.languageStats[language]);
    });
    
    // Show teams data for comparison
    console.log('\n=== TEAM STATS FOR COMPARISON ===');
    allData.teams.forEach(team => {
      console.log(`\n${team.config.name} (${team.config.language}):`);
      console.log('  Main stats:', team.stats);
      console.log('  Tasks count:', team.tasks.length);
      console.log('  Excluded count:', team.stats.excluded || 0);
    });
    
    SpreadsheetApp.getUi().alert(
      'Debug Complete',
      `Language stats debug complete!\n\n` +
      `Total active tasks: ${allData.rawTasks.length}\n` +
      `Total excluded tasks: ${allData.excludedTasks.length}\n` +
      `Flex team active tasks: ${allData.rawTasks.filter(t => t.isFlexTeam).length}\n` +
      `Jira active tasks: ${allData.rawTasks.filter(t => !t.isFlexTeam).length}\n\n` +
      `Check the console logs for detailed information including exclusion details.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('Debug error:', error);
    SpreadsheetApp.getUi().alert('Debug Error: ' + error.toString());
  }
}

// ============================
// MENU AND TRIGGERS
// ============================

/**
 * Creates custom menu on spreadsheet open
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PR Writer V2 Dashboard')
    .addItem('Create/Update Dashboard', 'createMasterTracker')
    .addSeparator()
    .addItem('Export Detailed Task Data', 'exportDetailedTaskData')
    .addSeparator()
    .addSubMenu(ui.createMenu('Target Adjustments')
      .addItem('Create Adjustments Sheet', 'createTargetAdjustmentsSheet')
      .addItem('Add Quick Adjustment', 'addQuickTargetAdjustment')
      .addItem('View Active Adjustments', 'viewTargetAdjustments')
    )
    .addSeparator()
    .addItem('Setup Auto Refresh (2 hours)', 'setupDeveloperAutomaticRefresh')
    .addItem('Stop Auto Refresh', 'stopDeveloperAutomaticRefresh')
    .addSeparator()
    .addItem('Export Dashboard to PDF', 'exportDeveloperDashboardToPDF')
    .addItem('Test Dashboard', 'testDeveloperDashboard')
    .addSeparator()
    .addItem('Schedule Auto-Update', 'scheduleAutoUpdate')
    .addItem('Remove Auto-Update', 'removeAutoUpdate')
    .addSeparator()
    .addItem('Generate Devs Progress', 'createJiraDashboard')
    .addItem('Setup Auto Refresh (2 hours)', 'setupAutomaticRefresh')
    .addItem('Stop Auto Refresh', 'stopAutomaticRefresh')
    .addSeparator()
    .addItem('Export Dashboard to PDF', 'exportDashboardToPDF')
    .addSeparator()
    .addItem('Create Rework Dev Statistics', 'createDoneItemsDashboard')
    .addItem('Create Rework Team Statistics', 'createTeamStatsSheet')
    .addToUi();
}

// ============================
// PLACEHOLDER FUNCTIONS
// ============================
// These functions are referenced in the menu but not implemented in this excerpt
// You should add your existing implementations for these functions:

function scheduleAutoUpdate() {
  SpreadsheetApp.getUi().alert('Auto-update scheduling not implemented yet.');
}

function removeAutoUpdate() {
  SpreadsheetApp.getUi().alert('Auto-update removal not implemented yet.');
}

function createJiraDashboard() {
  SpreadsheetApp.getUi().alert('Developer progress dashboard not implemented yet.');
}

function setupAutomaticRefresh() {
  SpreadsheetApp.getUi().alert('Auto refresh setup not implemented yet.');
}

function stopAutomaticRefresh() {
  SpreadsheetApp.getUi().alert('Auto refresh stop not implemented yet.');
}

function exportDashboardToPDF() {
  SpreadsheetApp.getUi().alert('PDF export not implemented yet.');
}

function createDoneItemsDashboard() {
  SpreadsheetApp.getUi().alert('Rework dev statistics not implemented yet.');
}

function createTeamStatsSheet() {
  SpreadsheetApp.getUi().alert('Team statistics sheet not implemented yet.');
}

function refreshDashboard() {
  createMasterTracker(); // Simply call the main function
}
