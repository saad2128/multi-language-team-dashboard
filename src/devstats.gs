// Developer Dashboard Generator Script with Tiered Target Tracking, Flex Team Integration, and Individual Target Adjustments
// Unique Configuration to avoid conflicts
const DEV_DASHBOARD_CONFIG = {
  startDate: new Date('2025-08-08'), // Project start date
  initialTasksPerDay: 4, // Expected tasks per developer per day for first 6 working days
  regularTasksPerDay: 3, // Expected tasks per developer per day after first 6 working days
  initialPeriodDays: 6, // Number of working days with higher target
  excludedAssignees: [
    'David',
    'John Doe'
  ],
  
  // Flex team configuration
  FLEX_TEAM_SHEET_ID: 'SHEET ID HERE',
  FLEX_TEAM_SHEET_NAME: 'Tasks',
  
  // Individual target adjustments sheet configuration
  TARGET_ADJUSTMENTS_SHEET_NAME: 'Target_Adjustments',
  
  // Flex team column mappings
  FLEX_TEAM_COLUMNS: {
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
  
  // Flex team status mapping
  FLEX_TEAM_STATUS_VALUES: {
    IN_PROGRESS: ['In progress'],
    REVIEW: ['Ready for Review', 'Review in progress'],
    ISSUE: ['Changes Requested'],
    REJECTED: ['Rejected'],
    REWORK: ['Rework'],
    COMPLETED: ['Approved']
  }
};

// Function to create or update Target Adjustments sheet
function createTargetAdjustmentsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let adjustmentSheet = ss.getSheetByName(DEV_DASHBOARD_CONFIG.TARGET_ADJUSTMENTS_SHEET_NAME);
  
  if (!adjustmentSheet) {
    adjustmentSheet = ss.insertSheet(DEV_DASHBOARD_CONFIG.TARGET_ADJUSTMENTS_SHEET_NAME);
    
    // Set up headers
    const headers = [
      'Developer Name',
      'Start Date (YYYY-MM-DD)',
      'End Date (YYYY-MM-DD)',
      'Adjustment Type',
      'Custom Target Tasks/Day',
      'Reason',
      'Status',
      'Created Date',
      'Notes'
    ];
    
    adjustmentSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    adjustmentSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('white');
    
    // Add some example data and instructions
    const exampleData = [
      ['John Doe', '2025-08-15', '2025-08-17', 'Leave', '0', 'Sick Leave', 'Active', new Date().toISOString().slice(0, 10), 'Doctor appointment'],
      ['Jane Smith', '2025-08-20', '2025-08-20', 'Custom', '2', 'Training Day', 'Active', new Date().toISOString().slice(0, 10), 'Attending workshop'],
      ['Mike Johnson', '2025-08-25', '2025-08-27', 'Leave', '0', 'Vacation', 'Active', new Date().toISOString().slice(0, 10), 'Family trip']
    ];
    
    adjustmentSheet.getRange(2, 1, exampleData.length, headers.length).setValues(exampleData);
    
    // Add instructions
    adjustmentSheet.getRange('A6').setValue('INSTRUCTIONS:');
    adjustmentSheet.getRange('A6').setFontWeight('bold').setFontSize(14);
    
    const instructions = [
      ['1. Developer Name: Enter exact name as it appears in Jira data'],
      ['2. Start/End Date: Use YYYY-MM-DD format (e.g., 2025-08-15)'],
      ['3. Adjustment Type: "Leave" (0 tasks), "Custom" (specify tasks/day), "Partial" (reduced tasks)'],
      ['4. Custom Target: For "Custom" type, specify tasks per day (can be 0, 0.5, 1, 2, etc.)'],
      ['5. Status: "Active" (applied), "Inactive" (ignored), "Expired" (automatically set)'],
      ['6. The system will automatically calculate working days and apply adjustments'],
      ['7. Adjustments override the standard tiered targets for specified periods'],
      ['8. Delete rows or set Status to "Inactive" to remove adjustments']
    ];
    
    adjustmentSheet.getRange(7, 1, instructions.length, 1).setValues(instructions);
    adjustmentSheet.getRange('A7:A14').setFontColor('#666666').setFontSize(10);
    
    // Format columns
    adjustmentSheet.setColumnWidth(1, 150); // Developer Name
    adjustmentSheet.setColumnWidth(2, 120); // Start Date
    adjustmentSheet.setColumnWidth(3, 120); // End Date
    adjustmentSheet.setColumnWidth(4, 120); // Adjustment Type
    adjustmentSheet.setColumnWidth(5, 150); // Custom Target
    adjustmentSheet.setColumnWidth(6, 200); // Reason
    adjustmentSheet.setColumnWidth(7, 100); // Status
    adjustmentSheet.setColumnWidth(8, 120); // Created Date
    adjustmentSheet.setColumnWidth(9, 250); // Notes
    
    // Add data validation for Adjustment Type
    const adjustmentTypeRange = adjustmentSheet.getRange('D2:D1000');
    const adjustmentTypeValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Leave', 'Custom', 'Partial'])
      .setAllowInvalid(false)
      .setHelpText('Select: Leave (0 tasks), Custom (specify), or Partial (reduced)')
      .build();
    adjustmentTypeRange.setDataValidation(adjustmentTypeValidation);
    
    // Add data validation for Status
    const statusRange = adjustmentSheet.getRange('G2:G1000');
    const statusValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active', 'Inactive', 'Expired'])
      .setAllowInvalid(false)
      .setHelpText('Active: Applied, Inactive: Ignored, Expired: Past date')
      .build();
    statusRange.setDataValidation(statusValidation);
    
    SpreadsheetApp.getUi().alert(
      'Target Adjustments Sheet Created!',
      'A new sheet "Target_Adjustments" has been created with example data and instructions.\n\n' +
      'You can now:\n' +
      '• Add individual developer target adjustments\n' +
      '• Handle leaves, training days, partial work days\n' +
      '• Set custom task targets for specific periods\n\n' +
      'The dashboard will automatically apply these adjustments when calculating targets.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
  
  return adjustmentSheet;
}

// Fixed version of getTargetAdjustments function
function getTargetAdjustments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adjustmentSheet = ss.getSheetByName(DEV_DASHBOARD_CONFIG.TARGET_ADJUSTMENTS_SHEET_NAME);
  
  if (!adjustmentSheet) {
    console.log('No target adjustments sheet found');
    return [];
  }
  
  const data = adjustmentSheet.getDataRange().getValues();
  if (data.length <= 1) {
    return []; // No data or just headers
  }
  
  const headers = data[0];
  const adjustments = [];
  
  // Find column indices
  const developerCol = headers.indexOf('Developer Name');
  const startDateCol = headers.indexOf('Start Date (YYYY-MM-DD)');
  const endDateCol = headers.indexOf('End Date (YYYY-MM-DD)');
  const adjustmentTypeCol = headers.indexOf('Adjustment Type');
  const customTargetCol = headers.indexOf('Custom Target Tasks/Day');
  const reasonCol = headers.indexOf('Reason');
  const statusCol = headers.indexOf('Status');
  const notesCol = headers.indexOf('Notes');
  
  const currentDate = new Date();
  currentDate.setHours(0, 0, 0, 0); // Reset time to midnight for accurate comparison
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Skip empty rows
    if (!row[developerCol] || !row[startDateCol] || !row[endDateCol]) {
      continue;
    }
    
    // Skip rows that start with instructions or are example text
    if (row[developerCol].toString().includes('INSTRUCTIONS') || 
        row[developerCol].toString().startsWith('1.') ||
        row[developerCol].toString().startsWith('2.') ||
        row[developerCol].toString().startsWith('3.') ||
        row[developerCol].toString().startsWith('4.') ||
        row[developerCol].toString().startsWith('5.') ||
        row[developerCol].toString().startsWith('6.') ||
        row[developerCol].toString().startsWith('7.') ||
        row[developerCol].toString().startsWith('8.')) {
      continue;
    }
    
    // Parse dates - handle both Date objects and strings
    let startDate, endDate;
    
    // Handle start date
    if (row[startDateCol] instanceof Date) {
      startDate = new Date(row[startDateCol]);
    } else {
      startDate = new Date(row[startDateCol].toString());
    }
    startDate.setHours(0, 0, 0, 0); // Reset time to midnight
    
    // Handle end date
    if (row[endDateCol] instanceof Date) {
      endDate = new Date(row[endDateCol]);
    } else {
      endDate = new Date(row[endDateCol].toString());
    }
    endDate.setHours(0, 0, 0, 0); // Reset time to midnight
    
    const status = row[statusCol] || 'Active';
    // Do NOT skip expired or inactive adjustments; always include all records for calculation
    
    // Validate dates
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      console.warn(`Invalid date format for ${row[developerCol]} in row ${i + 1}`);
      continue;
    }
    
    if (startDate > endDate) {
      console.warn(`Start date after end date for ${row[developerCol]} in row ${i + 1}`);
      continue;
    }
    
    const adjustment = {
      developer: row[developerCol].toString().trim(),
      startDate: startDate,
      endDate: endDate,
      adjustmentType: row[adjustmentTypeCol] || 'Custom',
      customTarget: parseFloat(row[customTargetCol]) || 0,
      reason: row[reasonCol] || '',
      status: status,
      notes: row[notesCol] || '',
      rowIndex: i + 1 // For updating status if needed
    };
    
    adjustments.push(adjustment);
    console.log(`Loaded adjustment for ${adjustment.developer}: ${adjustment.startDate.toISOString().slice(0, 10)} to ${adjustment.endDate.toISOString().slice(0, 10)}, Type: ${adjustment.adjustmentType}, Target: ${adjustment.customTarget}`);
  }
  
  console.log(`Loaded ${adjustments.length} active target adjustments`);
  return adjustments;
}


// Fixed version of calculateAdjustedExpectedTasks function
function calculateAdjustedExpectedTasks(developer, workingDaysElapsed, adjustments) {
  const startDate = new Date(DEV_DASHBOARD_CONFIG.startDate);
  startDate.setHours(0, 0, 0, 0);
  
  // Get adjustments for this developer
  const devAdjustments = adjustments.filter(adj => 
    adj.developer === developer && 
    adj.status === 'Active'
  );
  
  console.log(`Calculating adjusted tasks for ${developer}: Found ${devAdjustments.length} adjustments`);
  
  if (devAdjustments.length === 0) {
    // No adjustments, use standard calculation
    const standardTasks = calculateTieredExpectedTasks(workingDaysElapsed);
    console.log(`${developer}: No adjustments, using standard: ${standardTasks} tasks`);
    return standardTasks;
  }
  
  let totalExpectedTasks = 0;
  let processedDate = new Date(startDate);
  let dayCount = 0;
  let adjustedDays = 0;
  
  // Process each working day individually
  while (dayCount < workingDaysElapsed) {
    const dayOfWeek = processedDate.getDay();
    
    // Only process working days (Monday to Friday)
    if (dayOfWeek >= 1 && dayOfWeek <= 5) {
      dayCount++;
      
      // Create a date for comparison with time set to midnight
      const checkDate = new Date(processedDate);
      checkDate.setHours(0, 0, 0, 0);
      
      // Check if this date has any adjustments
      let dayAdjustment = null;
      for (const adj of devAdjustments) {
        // Check if checkDate is within the adjustment period (inclusive)
        if (checkDate >= adj.startDate && checkDate <= adj.endDate) {
          dayAdjustment = adj;
          adjustedDays++;
          break;
        }
      }
      
      if (dayAdjustment) {
        // Apply adjustment
        console.log(`${developer}: Applying ${dayAdjustment.adjustmentType} adjustment for ${checkDate.toISOString().slice(0, 10)}`);
        
        switch (dayAdjustment.adjustmentType) {
          case 'Leave':
            totalExpectedTasks += 0;
            break;
          case 'Custom':
            totalExpectedTasks += dayAdjustment.customTarget;
            break;
          case 'Partial':
            // For partial, use half of the standard target
            const standardDayTarget = dayCount <= DEV_DASHBOARD_CONFIG.initialPeriodDays 
              ? DEV_DASHBOARD_CONFIG.initialTasksPerDay 
              : DEV_DASHBOARD_CONFIG.regularTasksPerDay;
            totalExpectedTasks += standardDayTarget * 0.5;
            break;
          default:
            // Fallback to standard calculation
            totalExpectedTasks += dayCount <= DEV_DASHBOARD_CONFIG.initialPeriodDays 
              ? DEV_DASHBOARD_CONFIG.initialTasksPerDay 
              : DEV_DASHBOARD_CONFIG.regularTasksPerDay;
        }
      } else {
        // No adjustment, use standard tiered target
        const standardDayTarget = dayCount <= DEV_DASHBOARD_CONFIG.initialPeriodDays 
          ? DEV_DASHBOARD_CONFIG.initialTasksPerDay 
          : DEV_DASHBOARD_CONFIG.regularTasksPerDay;
        totalExpectedTasks += standardDayTarget;
      }
    }
    
    processedDate.setDate(processedDate.getDate() + 1);
  }
  
  const result = Math.round(totalExpectedTasks * 10) / 10; // Round to 1 decimal place
  console.log(`${developer}: Total adjusted expected tasks: ${result} (${adjustedDays} days adjusted out of ${workingDaysElapsed} working days)`);
  return result;
}


// Enhanced version of getDeveloperAdjustmentSummary function with better formatting
function getDeveloperAdjustmentSummary(developer, adjustments) {
  const devAdjustments = adjustments.filter(adj => 
    adj.developer === developer && 
    adj.status === 'Active'
  );
  
  if (devAdjustments.length === 0) {
    return '';
  }
  
  const summaryParts = devAdjustments.map(adj => {
    // Format dates as MM-DD
    const startStr = (adj.startDate.getMonth() + 1).toString().padStart(2, '0') + '-' + 
                     adj.startDate.getDate().toString().padStart(2, '0');
    const endStr = (adj.endDate.getMonth() + 1).toString().padStart(2, '0') + '-' + 
                   adj.endDate.getDate().toString().padStart(2, '0');
    const dateRange = startStr === endStr ? startStr : `${startStr} to ${endStr}`;
    
    let adjustmentDesc = '';
    switch (adj.adjustmentType) {
      case 'Leave':
        adjustmentDesc = 'Leave (0 tasks)';
        break;
      case 'Custom':
        adjustmentDesc = `Custom (${adj.customTarget} tasks/day)`;
        break;
      case 'Partial':
        adjustmentDesc = 'Partial (50% tasks)';
        break;
    }
    
    // Add reason if it's short enough
    const reasonPart = adj.reason && adj.reason.length < 20 ? ` - ${adj.reason}` : '';
    
    return `${dateRange}: ${adjustmentDesc}${reasonPart}`;
  });
  
  return summaryParts.join('; ');
}


// Function to read data from external Flex team Google Sheet
function getFlexTeamTaskData() {
  try {
    const flexSheet = SpreadsheetApp.openById(DEV_DASHBOARD_CONFIG.FLEX_TEAM_SHEET_ID);
    const dataSheet = flexSheet.getSheetByName(DEV_DASHBOARD_CONFIG.FLEX_TEAM_SHEET_NAME);
    
    if (!dataSheet) {
      console.warn(`Flex sheet '${DEV_DASHBOARD_CONFIG.FLEX_TEAM_SHEET_NAME}' not found in external document.`);
      return [];
    }
    
    const data = dataSheet.getDataRange().getValues();
    
    if (data.length === 0) {
      console.warn('No data found in Flex team sheet.');
      return [];
    }
    
    const headers = data[0];
    const flexColumnMap = getFlexTeamColumnMap(headers);
    
    const flexTasks = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[flexColumnMap.TASK_ID] || row[flexColumnMap.WORKER_ID]) {
        const language = row[flexColumnMap.LANGUAGE] ? row[flexColumnMap.LANGUAGE].toString().trim() : 'Multi-Language';
        const taskCompletionStatus = row[flexColumnMap.TASK_COMPLETION_STATUS] ? row[flexColumnMap.TASK_COMPLETION_STATUS].toString() : '';
        const repoApproval = row[flexColumnMap.REPO_APPROVAL] ? row[flexColumnMap.REPO_APPROVAL].toString() : '';
        
        let effectiveStatus = taskCompletionStatus;
        if (!effectiveStatus && repoApproval) {
          effectiveStatus = repoApproval;
        }
        
        const standardStatus = mapFlexStatusToJiraStandard(effectiveStatus);
        
        const task = {
          issueType: 'Task',
          key: row[flexColumnMap.TASK_ID] || '',
          assignee: row[flexColumnMap.WORKER_ID] || '',
          team: 'Flex',
          status: standardStatus,
          folderId: '',
          summary: row[flexColumnMap.TASK_DESCRIPTION] || '',
          taskDescription: row[flexColumnMap.TASK_DESCRIPTION] || '',
          firstPrompt: row[flexColumnMap.FIRST_PROMPT] || '',
          repoUrl: row[flexColumnMap.REPO_URL] || '',
          githubLink: '',
          ltUrl: row[flexColumnMap.LT_URL] || '',
          gdriveLink: row[flexColumnMap.GDRIVE_LINK] || '',
          created: row[flexColumnMap.TASK_START_DATE] || '',
          updated: '',
          numberOfPrompt: 1, // Default value for Flex tasks
          isFlexTeam: true,
          originalFlexStatus: effectiveStatus,
          language: language
        };
        
        flexTasks.push(task);
      }
    }
    
    console.log(`Found ${flexTasks.length} tasks from Flex team`);
    return flexTasks;
    
  } catch (error) {
    console.error('Error reading Flex team data:', error);
    return [];
  }
}

// Creates column mapping for Flex team headers
function getFlexTeamColumnMap(headers) {
  const columnMap = {};
  
  Object.keys(DEV_DASHBOARD_CONFIG.FLEX_TEAM_COLUMNS).forEach(key => {
    const columnName = DEV_DASHBOARD_CONFIG.FLEX_TEAM_COLUMNS[key];
    const index = findFlexColumnIndex(headers, columnName);
    columnMap[key] = index;
  });
  
  return columnMap;
}

// Helper function to find column index
function findFlexColumnIndex(headers, columnName) {
  return headers.findIndex(h => h.toString().toLowerCase().trim() === columnName.toLowerCase().trim());
}

// Maps Flex team status to standard Jira status
function mapFlexStatusToJiraStandard(flexStatus) {
  const normalizedStatus = flexStatus.toString().trim();
  
  for (const [category, values] of Object.entries(DEV_DASHBOARD_CONFIG.FLEX_TEAM_STATUS_VALUES)) {
    if (values.some(v => v === normalizedStatus)) {
      switch (category) {
        case 'IN_PROGRESS': return 'In Progress';
        case 'REVIEW': return 'Repo Review';
        case 'ISSUE': return 'Blocked';
        case 'REJECTED': return 'Repo Rejected';
        case 'REWORK': return 'In Progress'; // Map rework to In Progress
        case 'COMPLETED': return 'Done';
        default: return 'In Progress';
      }
    }
  }
  
  return 'In Progress'; // Default fallback
}

// Test function to verify Flex team connection
function testFlexTeamDataConnection() {
  try {
    const ui = SpreadsheetApp.getUi();
    const flexTasks = getFlexTeamTaskData();
    
    if (flexTasks.length === 0) {
      ui.alert(
        'Flex Team Connection Test',
        'No data found in Flex team sheet. Please check:\n' +
        '1. Sheet ID is correct in DEV_DASHBOARD_CONFIG.FLEX_TEAM_SHEET_ID\n' +
        '2. Sheet name is correct in DEV_DASHBOARD_CONFIG.FLEX_TEAM_SHEET_NAME\n' +
        '3. Sheet has proper sharing permissions\n' +
        '4. Sheet contains data with the expected columns',
        ui.ButtonSet.OK
      );
    } else {
      const sampleTask = flexTasks[0];
      const statusBreakdown = {};
      
      flexTasks.forEach(task => {
        if (!statusBreakdown[task.status]) {
          statusBreakdown[task.status] = 0;
        }
        statusBreakdown[task.status]++;
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
        `- Status: ${sampleTask.status}\n` +
        `- Original Flex Status: ${sampleTask.originalFlexStatus}`,
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

function createJiraDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create dashboard sheet
  let dashboardSheet = ss.getSheetByName('Devs Progress');
  if (dashboardSheet) {
    dashboardSheet.clear();
  } else {
    dashboardSheet = ss.insertSheet('Devs Progress');
  }
  
  // Read data from Jira_Data tab
  const dataSheet = ss.getSheetByName('Jira_Data');
  if (!dataSheet) {
    throw new Error('Jira_Data sheet not found');
  }
  
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  const jiraRows = data.slice(1);
  
  // Find column indices
  const columnIndices = {
    issueType: headers.indexOf('Issue Type'),
    key: headers.indexOf('Key'),
    assignee: headers.indexOf('Assignee'),
    team: headers.indexOf('Team'),
    status: headers.indexOf('Status'),
    folderId: headers.indexOf('folder/id'),
    summary: headers.indexOf('Summary'),
    taskDescription: headers.indexOf('task_description'),
    firstPrompt: headers.indexOf('first_prompt'),
    repoUrl: headers.indexOf('repo_url'),
    githubLink: headers.indexOf('github link'),
    ltUrl: headers.indexOf('LT_url'),
    gdriveLink: headers.indexOf('gdrive link'),
    created: headers.indexOf('Created'),
    updated: headers.indexOf('Updated'),
    numberOfPrompt: headers.indexOf('number of prompt assignee')
  };
  
  // Get Flex team data and target adjustments
  const flexTasks = getFlexTeamTaskData();
  const targetAdjustments = getTargetAdjustments();
  
  // Convert Jira rows to task objects
  const jiraTasks = jiraRows.map(row => ({
    issueType: row[columnIndices.issueType] || 'Unknown',
    key: row[columnIndices.key] || '',
    assignee: row[columnIndices.assignee] || 'Unassigned',
    team: row[columnIndices.team] || 'Unknown',
    status: row[columnIndices.status] || 'Unknown',
    folderId: row[columnIndices.folderId] || '',
    summary: row[columnIndices.summary] || '',
    taskDescription: row[columnIndices.taskDescription] || '',
    firstPrompt: row[columnIndices.firstPrompt] || '',
    repoUrl: row[columnIndices.repoUrl] || '',
    githubLink: row[columnIndices.githubLink] || '',
    ltUrl: row[columnIndices.ltUrl] || '',
    gdriveLink: row[columnIndices.gdriveLink] || '',
    created: row[columnIndices.created] || '',
    updated: row[columnIndices.updated] || '',
    numberOfPrompt: parseInt(row[columnIndices.numberOfPrompt]) || 0,
    isFlexTeam: false
  }));
  
  // Combine Jira and Flex tasks
  const allTasks = [...jiraTasks, ...flexTasks];
  
  // Process data and calculate stats with adjusted targets
  const stats = calculateDeveloperStatsWithAdjustedTargets(allTasks, columnIndices, targetAdjustments);
  
  // Create dashboard
  createDeveloperDashboardLayout(dashboardSheet, stats, flexTasks.length, targetAdjustments.length);
  
  console.log(`Dashboard created successfully with ${targetAdjustments.length} individual target adjustments. Included ${flexTasks.length} Flex team tasks.`);
}

function calculateWorkingDaysElapsed(startDate, endDate) {
  let workingDays = 0;
  const currentDate = new Date(startDate);
  
  while (currentDate <= endDate) {
    const dayOfWeek = currentDate.getDay();
    // Count Monday (1) to Friday (5) as working days
    if (dayOfWeek >= 1 && dayOfWeek <= 5) {
      workingDays++;
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }
  
  return workingDays;
}

function calculateTieredExpectedTasks(workingDaysElapsed) {
  let expectedTasks = 0;
  
  if (workingDaysElapsed <= DEV_DASHBOARD_CONFIG.initialPeriodDays) {
    // All days are in the initial period (4 tasks/day)
    expectedTasks = workingDaysElapsed * DEV_DASHBOARD_CONFIG.initialTasksPerDay;
  } else {
    // Initial period days at 4 tasks/day + remaining days at 3 tasks/day
    const initialPeriodTasks = DEV_DASHBOARD_CONFIG.initialPeriodDays * DEV_DASHBOARD_CONFIG.initialTasksPerDay;
    const remainingDays = workingDaysElapsed - DEV_DASHBOARD_CONFIG.initialPeriodDays;
    const remainingTasks = remainingDays * DEV_DASHBOARD_CONFIG.regularTasksPerDay;
    expectedTasks = initialPeriodTasks + remainingTasks;
  }
  
  return expectedTasks;
}

// Modified to work with task objects and include individual target adjustments
function calculateDeveloperStatsWithAdjustedTargets(allTasks, columnIndices, adjustments) {
  const assigneeStats = {};
  const currentDate = new Date();
  const workingDaysElapsed = calculateWorkingDaysElapsed(DEV_DASHBOARD_CONFIG.startDate, currentDate);
  
  allTasks.forEach(task => {
    const assignee = task.assignee || 'Unassigned';
    
    // Skip excluded assignees
    if (DEV_DASHBOARD_CONFIG.excludedAssignees.includes(assignee)) {
      return;
    }
    
    const status = task.status || 'Unknown';
    const numberOfPrompts = task.numberOfPrompt || 0;
    const issueType = task.issueType || 'Unknown';
    const team = task.team || 'Unknown';
    const created = task.created;
    const updated = task.updated;
    
    if (!assigneeStats[assignee]) {
      // Calculate adjusted expected tasks for this specific developer
      const adjustedExpectedTasks = calculateAdjustedExpectedTasks(assignee, workingDaysElapsed, adjustments);
      const adjustmentSummary = getDeveloperAdjustmentSummary(assignee, adjustments);
      
      assigneeStats[assignee] = {
        totalTasks: 0,
        totalPrompts: 0,
        statusBreakdown: {},
        issueTypeBreakdown: {},
        teamBreakdown: {},
        avgPromptsPerTask: 0,
        completedTasks: 0,
        inProgressTasks: 0,
        todoTasks: 0,
        reviewTasks: 0,
        issueTasks: 0,
        recentActivity: [],
        reviewStats: {
          repoReview: 0,
          peerReview: 0,
          leadReview: 0
        },
        // Target tracking with individual adjustments
        expectedTasks: adjustedExpectedTasks,
        standardExpectedTasks: calculateTieredExpectedTasks(workingDaysElapsed), // Keep standard for comparison
        completedAndReviewTasks: 0,
        taskGap: 0,
        targetAchievement: 0,
        onTrack: false,
        hasAdjustments: adjustmentSummary !== '',
        adjustmentSummary: adjustmentSummary,
        // Flex team tracking
        flexTasks: 0,
        jiraTasks: 0,
        isFlexWorker: false
      };
    }
    
    const stats = assigneeStats[assignee];
    stats.totalTasks++;
    stats.totalPrompts += numberOfPrompts;
    
    // Track task source
    if (task.isFlexTeam) {
      stats.flexTasks++;
      stats.isFlexWorker = true;
    } else {
      stats.jiraTasks++;
    }
    
    // Status breakdown
    stats.statusBreakdown[status] = (stats.statusBreakdown[status] || 0) + 1;
    
    // Issue type breakdown
    stats.issueTypeBreakdown[issueType] = (stats.issueTypeBreakdown[issueType] || 0) + 1;
    
    // Team breakdown
    stats.teamBreakdown[team] = (stats.teamBreakdown[team] || 0) + 1;
    
    // Status categorization based on exact values
    switch (status) {
      case 'To Do':
        stats.todoTasks++;
        break;
      case 'In Progress':
      case 'Repo Approved':
        stats.inProgressTasks++;
        break;
      case 'Repo Review':
        stats.reviewTasks++;
        stats.reviewStats.repoReview++;
        stats.completedAndReviewTasks++; // Count for target achievement
        break;
      case 'Peer review':
        stats.reviewTasks++;
        stats.reviewStats.peerReview++;
        stats.completedAndReviewTasks++; // Count for target achievement
        break;
      case 'Lead review':
        stats.reviewTasks++;
        stats.reviewStats.leadReview++;
        stats.completedAndReviewTasks++; // Count for target achievement
        break;
      case 'Done':
        stats.completedTasks++;
        stats.completedAndReviewTasks++; // Count for target achievement
        break;
      case 'Repo Rejected':
      case 'Blocked':
        stats.issueTasks++;
        break;
      default:
        break;
    }
    
    // Recent activity (last updated)
    if (updated) {
      stats.recentActivity.push({
        date: updated,
        task: task.summary || 'Unknown Task',
        source: task.isFlexTeam ? 'Flex' : 'Jira'
      });
    }
  });
  
  // Calculate averages, targets, and sort recent activity
  Object.keys(assigneeStats).forEach(assignee => {
    const stats = assigneeStats[assignee];
    stats.avgPromptsPerTask = stats.totalTasks > 0 ? (stats.totalPrompts / stats.totalTasks).toFixed(2) : 0;
    
    // Calculate target metrics using completed + review tasks with adjusted targets
    stats.taskGap = stats.completedAndReviewTasks - stats.expectedTasks;
    stats.targetAchievement = stats.expectedTasks > 0 ? ((stats.completedAndReviewTasks / stats.expectedTasks) * 100).toFixed(1) : 0;
    stats.onTrack = stats.completedAndReviewTasks >= stats.expectedTasks;
    
    stats.recentActivity.sort((a, b) => new Date(b.date) - new Date(a.date));
    stats.recentActivity = stats.recentActivity.slice(0, 5); // Keep only last 5
  });
  
  return assigneeStats;
}

// Modified to include target adjustments information
function createDeveloperDashboardLayout(sheet, stats, flexTaskCount = 0, adjustmentCount = 0) {
  // Set up the dashboard header
  sheet.getRange('A1').setValue('DEVELOPERS PROGRESS DASHBOARD (Flex Team + Individual Adjustments)');
  sheet.getRange('A1').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('A1:J1').merge();
  
  // Last updated timestamp
  sheet.getRange('A2').setValue(`Last Updated: ${new Date().toLocaleString()}`);
  sheet.getRange('A2').setFontStyle('italic');
  
  // Add project info with tiered targets, Flex team info, and adjustments
  const currentDate = new Date();
  const workingDaysElapsed = calculateWorkingDaysElapsed(DEV_DASHBOARD_CONFIG.startDate, currentDate);
  const expectedTasks = calculateTieredExpectedTasks(workingDaysElapsed);
  
  sheet.getRange('A3').setValue(`Project Started: ${DEV_DASHBOARD_CONFIG.startDate.toDateString()} | Working Days: ${workingDaysElapsed} | Target: ${DEV_DASHBOARD_CONFIG.initialTasksPerDay} tasks/day (first ${DEV_DASHBOARD_CONFIG.initialPeriodDays} days), then ${DEV_DASHBOARD_CONFIG.regularTasksPerDay} tasks/day`);
  sheet.getRange('A3').setFontSize(10).setFontStyle('italic').setFontColor('#333333');
  
  // Add expected tasks info with Flex team data and adjustments
  sheet.getRange('A4').setValue(`Standard Expected Tasks: ${expectedTasks} | Flex Team Tasks: ${flexTaskCount} | Individual Adjustments: ${adjustmentCount} | Maintained by Saad Aslam`);
  sheet.getRange('A4').setFontSize(8).setFontStyle('italic').setFontColor('#666666');
  
  // Create summary section
  createDeveloperSummarySection(sheet, stats, flexTaskCount, adjustmentCount);
  
  // Create detailed stats table
  createDeveloperDetailedStatsTable(sheet, stats);
  
  // Format the sheet
  formatDeveloperDashboard(sheet);
  // Update low performance comments sheet
  updateLowPerformanceCommentsAfterDashboard(stats);
}

// Modified to include target adjustments metrics
function createDeveloperSummarySection(sheet, stats, flexTaskCount = 0, adjustmentCount = 0) {
  const assignees = Object.keys(stats);
  const totalAssignees = assignees.length;
  const totalTasks = Object.values(stats).reduce((sum, stat) => sum + stat.totalTasks, 0);
  const totalPrompts = Object.values(stats).reduce((sum, stat) => sum + stat.totalPrompts, 0);
  const totalCompleted = Object.values(stats).reduce((sum, stat) => sum + stat.completedTasks, 0);
  const totalInProgress = Object.values(stats).reduce((sum, stat) => sum + stat.inProgressTasks, 0);
  const totalInReview = Object.values(stats).reduce((sum, stat) => sum + stat.reviewTasks, 0);
  const totalIssues = Object.values(stats).reduce((sum, stat) => sum + stat.issueTasks, 0);
  
  // Flex team metrics
  const totalFlexTasks = Object.values(stats).reduce((sum, stat) => sum + stat.flexTasks, 0);
  const totalJiraTasks = Object.values(stats).reduce((sum, stat) => sum + stat.jiraTasks, 0);
  const flexWorkers = Object.values(stats).filter(stat => stat.isFlexWorker).length;
  
  // Target adjustments metrics
  const devsWithAdjustments = Object.values(stats).filter(stat => stat.hasAdjustments).length;
  const totalAdjustedExpected = Object.values(stats).reduce((sum, stat) => sum + stat.expectedTasks, 0);
  const totalStandardExpected = Object.values(stats).reduce((sum, stat) => sum + stat.standardExpectedTasks, 0);
  const adjustmentImpact = totalAdjustedExpected - totalStandardExpected;
  
  // Calculate target metrics using completed + review tasks
  const totalCompletedAndReview = Object.values(stats).reduce((sum, stat) => sum + stat.completedAndReviewTasks, 0);
  const devsOnTrack = Object.values(stats).filter(stat => stat.onTrack).length;
  const devsBehindTarget = totalAssignees - devsOnTrack;
  const overallTargetAchievement = totalAdjustedExpected > 0 ? ((totalCompletedAndReview / totalAdjustedExpected) * 100).toFixed(1) : 0;
  
  // Calculate performance distribution
  const performanceDistribution = calculateDeveloperPerformanceDistribution(stats);
  
  // Summary cards - with Flex team data and adjustments
  sheet.getRange('A6').setValue('SUMMARY OVERVIEW (Jira + Flex Team + Individual Adjustments)');
  sheet.getRange('A6').setFontWeight('bold').setFontSize(14);
  
  const summaryData = [
    ['Total Developers', totalAssignees],
    ['Total Tasks (All Status)', totalTasks],
    ['Total Jira Tasks', totalJiraTasks],
    ['Total Flex Tasks', totalFlexTasks],
    ['Flex Workers', flexWorkers],
    ['Standard Expected Tasks', totalStandardExpected],
    ['Adjusted Expected Tasks', totalAdjustedExpected],
    ['Target Adjustment Impact', adjustmentImpact > 0 ? `+${adjustmentImpact}` : adjustmentImpact],
    ['Developers with Adjustments', devsWithAdjustments],
    ['Active Individual Adjustments', adjustmentCount],
    ['Overall Target Achievement', `${overallTargetAchievement}%`],
    ['Developers On Track', devsOnTrack],
    ['Developers Behind Target', devsBehindTarget],
    ['Total Prompts', totalPrompts],
    ['Completed Tasks', totalCompleted],
    ['In Review Tasks', totalInReview],
    ['In Progress Tasks', totalInProgress],
    ['Issue Tasks (Rejected/Blocked)', totalIssues],
    ['Average Prompts/Task', totalTasks > 0 ? (totalPrompts / totalTasks).toFixed(2) : 0],
    ['🏆 High Performers', performanceDistribution.highPerformers],
    ['🌟 Average Performers', performanceDistribution.averagePerformers],
    ['⚠️ Low Performers', performanceDistribution.lowPerformers]
  ];
  
  sheet.getRange('A7:B28').setValues(summaryData);
  sheet.getRange('A7:A28').setFontWeight('bold').setFontColor('black');
  sheet.getRange('B7:B28').setHorizontalAlignment('center').setFontColor('black');
  
  // Highlight key target metrics with black fonts
  sheet.getRange('A12:B12').setBackground('#f3e5f5').setFontWeight('bold').setFontColor('black'); // Standard expected
  sheet.getRange('A13:B13').setBackground('#fff3e0').setFontWeight('bold').setFontColor('black'); // Adjusted expected
  sheet.getRange('A14:B14').setBackground('#e3f2fd').setFontWeight('bold').setFontColor('black'); // Adjustment impact
  sheet.getRange('A15:B15').setBackground('#f1f8e9').setFontWeight('bold').setFontColor('black'); // Devs with adjustments
  sheet.getRange('A16:B16').setBackground('#fce4ec').setFontWeight('bold').setFontColor('black'); // Active adjustments
  sheet.getRange('A17:B17').setBackground('#e8f5e8').setFontWeight('bold').setFontColor('black'); // Target achievement
  sheet.getRange('A18:B18').setBackground('#c8e6c9').setFontWeight('bold').setFontColor('black'); // On track
  sheet.getRange('A19:B19').setBackground('#ffebee').setFontWeight('bold').setFontColor('black'); // Behind target
  
  // Highlight Flex team metrics
  sheet.getRange('A9:B9').setBackground('#f3e5f5').setFontWeight('bold').setFontColor('black'); // Jira tasks
  sheet.getRange('A10:B10').setBackground('#e1f5fe').setFontWeight('bold').setFontColor('black'); // Flex tasks
  sheet.getRange('A11:B11').setBackground('#e8f5e8').setFontWeight('bold').setFontColor('black'); // Flex workers
  
  // Highlight performance summary rows with distinct colors and black fonts
  sheet.getRange('A26:B26').setBackground('#00c851').setFontColor('black').setFontWeight('bold'); // High performers
  sheet.getRange('A27:B27').setBackground('#ff8800').setFontColor('black').setFontWeight('bold'); // Average performers
  sheet.getRange('A28:B28').setBackground('#ff4444').setFontColor('black').setFontWeight('bold'); // Low performers
}

function calculateDeveloperPerformanceDistribution(stats) {
  let highPerformers = 0;
  let averagePerformers = 0;
  let lowPerformers = 0;
  
  Object.values(stats).forEach(stat => {
    const performance = calculateDeveloperPerformanceScore(stat);
    if (performance.label.includes('High Performer')) {
      highPerformers++;
    } else if (performance.label.includes('Average Performer')) {
      averagePerformers++;
    } else {
      lowPerformers++;
    }
  });
  
  return {
    highPerformers,
    averagePerformers,
    lowPerformers
  };
}

// Modified to include target adjustments information in table
function createDeveloperDetailedStatsTable(sheet, stats) {
  // Detailed stats header
  sheet.getRange('A30').setValue('DETAILED ASSIGNEE STATISTICS (Jira + Flex Team + Individual Adjustments)');
  sheet.getRange('A30').setFontWeight('bold').setFontSize(14).setFontColor('black');
  
  // Table headers - Updated to include adjustment columns
  const headers = [
    'Rank', 'Assignee', 'Completed+Review', 'Adjusted Target', 'Standard Target', 'Target Gap', 'Achievement %', 
    'Total Tasks', 'Jira Tasks', 'Flex Tasks', 'Total Prompts', 'Avg Prompts/Task', 'Completed', 'In Progress', 'To Do', 
    'In Review', 'Issues', 'Completion %', 'Reviews', 'Performance', 'Team', 'Source', 'Adjustments'
  ];
  
  sheet.getRange('A32:W32').setValues([headers]);
  sheet.getRange('A32:W32').setFontWeight('bold').setBackground('#4285f4').setFontColor('black');
  
  // Populate data and calculate performance scores
  const tableData = [];
  Object.keys(stats).forEach(assignee => {
    const stat = stats[assignee];
    const completionRate = stat.totalTasks > 0 ? (stat.completedTasks / stat.totalTasks * 100).toFixed(1) : 0;
    const performanceScore = calculateDeveloperPerformanceScore(stat);
    const reviewSummary = formatDeveloperReviewSummary(stat.reviewStats);
    
    // Determine primary team
    const primaryTeam = Object.keys(stat.teamBreakdown).reduce((a, b) => 
      stat.teamBreakdown[a] > stat.teamBreakdown[b] ? a : b, 'Unknown');
    
    // Determine data source
    let dataSource = 'Jira Only';
    if (stat.flexTasks > 0 && stat.jiraTasks > 0) {
      dataSource = 'Jira + Flex';
    } else if (stat.flexTasks > 0) {
      dataSource = 'Flex Only';
    }
    
    // Format adjustments for display
    const adjustmentDisplay = stat.hasAdjustments ? 
      (stat.adjustmentSummary.length > 30 ? stat.adjustmentSummary.substring(0, 30) + '...' : stat.adjustmentSummary) :
      'None';
    
    tableData.push([
      0, // Rank placeholder
      assignee,
      stat.completedAndReviewTasks, // Completed + Review tasks
      stat.expectedTasks, // Adjusted target
      stat.standardExpectedTasks, // Standard target for comparison
      stat.taskGap,
      `${stat.targetAchievement}%`,
      stat.totalTasks,
      stat.jiraTasks, // Jira tasks count
      stat.flexTasks, // Flex tasks count
      stat.totalPrompts,
      stat.avgPromptsPerTask,
      stat.completedTasks,
      stat.inProgressTasks,
      stat.todoTasks,
      stat.reviewTasks,
      stat.issueTasks,
      `${completionRate}%`,
      reviewSummary,
      performanceScore.label,
      primaryTeam,
      dataSource, // Data source
      adjustmentDisplay, // Adjustments summary
      performanceScore.score, // Hidden score for sorting
      stat.onTrack, // Hidden onTrack for highlighting
      stat.hasAdjustments // Hidden hasAdjustments for highlighting
    ]);
  });
  
  // Sort by target achievement first, then by performance score
  tableData.sort((a, b) => {
    const aAchievement = parseFloat(a[6]);
    const bAchievement = parseFloat(b[6]);
    if (aAchievement !== bAchievement) return bAchievement - aAchievement;
    return b[23] - a[23]; // Then by performance score
  });
  
  // Add rank numbers and remove hidden columns
  const finalTableData = tableData.map((row, index) => {
    row[0] = index + 1; // Add rank
    return row.slice(0, 23); // Remove hidden columns
  });
  
  if (finalTableData.length > 0) {
    sheet.getRange(33, 1, finalTableData.length, headers.length).setValues(finalTableData);
    
    // Apply target-based highlighting with adjustments
    applyDeveloperTargetBasedHighlighting(sheet, tableData, 33);
  }
  
  // Create progress bars using conditional formatting
  const dataRange = sheet.getRange(33, 3, finalTableData.length, 5);
  const rules = sheet.getConditionalFormatRules();
  
  // Add data bars for visual representation
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint('#4285f4')
    .setGradientMinpoint('#ffffff')
    .setRanges([dataRange])
    .build();
  
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}

function formatDeveloperReviewSummary(reviewStats) {
  const parts = [];
  if (reviewStats.repoReview > 0) parts.push(`Repo:${reviewStats.repoReview}`);
  if (reviewStats.peerReview > 0) parts.push(`Peer:${reviewStats.peerReview}`);
  if (reviewStats.leadReview > 0) parts.push(`Lead:${reviewStats.leadReview}`);
  return parts.length > 0 ? parts.join(' | ') : 'None';
}

function calculateDeveloperPerformanceScore(stat) {
  let score = 0;
  let label = '';
  
  // Enhanced scoring that considers target achievement based on completed + review tasks
  const targetAchievement = parseFloat(stat.targetAchievement);
  const completedAndReviewTasks = stat.completedAndReviewTasks;
  const avgPrompts = parseFloat(stat.avgPromptsPerTask);
  
  // Primary factor: Target achievement based on completed + review tasks (60% weight)
  score += Math.min(targetAchievement * 0.6, 60);
  
  // Secondary factor: Completed + Review tasks count (25% weight)
  score += Math.min(completedAndReviewTasks * 2.5, 25);
  
  // Tertiary factor: Prompt engagement (15% weight)
  score += Math.min(avgPrompts * 1.5, 15);
  
  // Determine performance label based on target achievement of completed + review tasks
  if (targetAchievement >= 100 && completedAndReviewTasks >= 8) {
    label = '🏆 High Performer (On Target)';
  } else if (targetAchievement >= 80 || (completedAndReviewTasks >= 6 && targetAchievement >= 70)) {
    label = '🌟 Average Performer';
  } else if (targetAchievement < 70 || completedAndReviewTasks < 4) {
    label = '⚠️ Low Performer (Behind Target)';
  } else {
    label = '🌟 Average Performer';
  }
  
  return { score: score, label: label };
}

// Modified to handle new table structure with adjustments columns
function applyDeveloperTargetBasedHighlighting(sheet, tableData, startRow) {
  tableData.forEach((row, index) => {
    const rowIndex = startRow + index;
    const onTrack = row[24]; // Hidden onTrack value (updated index)
    const hasAdjustments = row[25]; // Hidden hasAdjustments value
    const targetAchievement = parseFloat(row[6]);
    const performance = row[19]; // Performance column (updated index)
    
    // Get all cells for this row - Set all font colors to black
    const entireRow = sheet.getRange(rowIndex, 1, 1, 23); // Updated column count
    entireRow.setFontColor('black'); // Set all fonts to black
    
    const assigneeCell = sheet.getRange(rowIndex, 2);
    const completedReviewCell = sheet.getRange(rowIndex, 3); // Completed + Review column
    const adjustedTargetCell = sheet.getRange(rowIndex, 4); // Adjusted target column
    const standardTargetCell = sheet.getRange(rowIndex, 5); // Standard target column
    const targetGapCell = sheet.getRange(rowIndex, 6);
    const achievementCell = sheet.getRange(rowIndex, 7);
    const performanceCell = sheet.getRange(rowIndex, 20); // Updated index
    const rankCell = sheet.getRange(rowIndex, 1);
    const flexTasksCell = sheet.getRange(rowIndex, 10); // Flex tasks column
    const sourceCell = sheet.getRange(rowIndex, 22); // Source column
    const adjustmentsCell = sheet.getRange(rowIndex, 23); // Adjustments column
    
    // Primary highlighting based on target achievement with complete cell backgrounds
    if (!onTrack || targetAchievement < 70) {
      // Behind target - Red highlighting with complete cell backgrounds
      entireRow.setBackground('#ffcccb'); // Light red background for entire row
      assigneeCell.setBackground('#ff4444').setFontColor('black').setFontWeight('bold');
      completedReviewCell.setBackground('#ff4444').setFontColor('black').setFontWeight('bold');
      adjustedTargetCell.setBackground('#ff4444').setFontColor('black').setFontWeight('bold');
      targetGapCell.setBackground('#ff4444').setFontColor('black').setFontWeight('bold');
      achievementCell.setBackground('#ff4444').setFontColor('black').setFontWeight('bold');
      performanceCell.setBackground('#ff4444').setFontColor('black').setFontWeight('bold');
      rankCell.setBackground('#ff6b6b').setFontWeight('bold').setFontColor('black');
      
      // Add thick red border
      entireRow.setBorder(true, true, true, true, false, false, '#cc0000', SpreadsheetApp.BorderStyle.SOLID_THICK);
      
    } else if (targetAchievement >= 100) {
      // Meeting or exceeding target - Green highlighting with complete cell backgrounds
      entireRow.setBackground('#c8e6c9'); // Light green background for entire row
      assigneeCell.setBackground('#00c851').setFontColor('black').setFontWeight('bold');
      completedReviewCell.setBackground('#00c851').setFontColor('black').setFontWeight('bold');
      adjustedTargetCell.setBackground('#00c851').setFontColor('black').setFontWeight('bold');
      targetGapCell.setBackground('#00c851').setFontColor('black').setFontWeight('bold');
      achievementCell.setBackground('#00c851').setFontColor('black').setFontWeight('bold');
      performanceCell.setBackground('#00c851').setFontColor('black').setFontWeight('bold');
      rankCell.setBackground('#ffd700').setFontWeight('bold').setFontColor('black'); // Gold
      
      // Add thick green border
      entireRow.setBorder(true, true, true, true, false, false, '#00c851', SpreadsheetApp.BorderStyle.SOLID_THICK);
      
    } else if (targetAchievement >= 80) {
      // Close to target - Orange highlighting with complete cell backgrounds
      entireRow.setBackground('#ffe0b2'); // Light orange background for entire row
      assigneeCell.setBackground('#ff8800').setFontColor('black').setFontWeight('bold');
      completedReviewCell.setBackground('#ff8800').setFontColor('black').setFontWeight('bold');
      adjustedTargetCell.setBackground('#ff8800').setFontColor('black').setFontWeight('bold');
      targetGapCell.setBackground('#ff8800').setFontColor('black').setFontWeight('bold');
      achievementCell.setBackground('#ff8800').setFontColor('black').setFontWeight('bold');
      performanceCell.setBackground('#ff8800').setFontColor('black').setFontWeight('bold');
      rankCell.setBackground('#ffcc02').setFontWeight('bold').setFontColor('black');
      
      // Add orange border
      entireRow.setBorder(true, true, true, true, false, false, '#ff8800', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
    
    // Special highlighting for developers with adjustments
    if (hasAdjustments) {
      adjustmentsCell.setBackground('#fff3e0').setFontWeight('bold').setFontColor('#e65100');
      // Highlight the adjusted target differently from standard target
      adjustedTargetCell.setFontWeight('bold').setFontColor('#1976d2');
      standardTargetCell.setFontStyle('italic').setFontColor('#757575');
    }
    
    // Special highlighting for Flex team tasks
    if (parseInt(row[9]) > 0) { // If has Flex tasks
      flexTasksCell.setBackground('#e1f5fe').setFontWeight('bold').setFontColor('#1976d2');
    }
    
    // Color code the source column
    if (row[21] === 'Flex Only') {
      sourceCell.setBackground('#e1f5fe').setFontWeight('bold').setFontColor('#1976d2');
    } else if (row[21] === 'Jira + Flex') {
      sourceCell.setBackground('#f3e5f5').setFontWeight('bold').setFontColor('#7b1fa2');
    }
    
    // Additional highlighting for target gap column with complete cell background
    const gapValue = parseFloat(row[5]);
    if (gapValue < 0) {
      targetGapCell.setFontWeight('bold').setFontColor('black');
    } else if (gapValue > 0) {
      targetGapCell.setBackground('#e8f5e8').setFontWeight('bold').setFontColor('black');
    }
  });
}

// Modified to accommodate new table structure with adjustments
function formatDeveloperDashboard(sheet) {
  // Auto-resize columns
  sheet.autoResizeColumns(1, 23); // Updated column count
  
  // Add borders to tables with black fonts
  const summaryRange = sheet.getRange('A7:B28'); // Updated range
  summaryRange.setBorder(true, true, true, true, true, true);
  summaryRange.setFontColor('black'); // Ensure black fonts
  
  // Freeze header rows
  sheet.setFrozenRows(5);
  
  // Set alternating row colors for the main table
  const dataRange = sheet.getDataRange();
  const lastRow = dataRange.getLastRow();
  if (lastRow > 32) { // Updated row number
    const tableRange = sheet.getRange(33, 1, lastRow - 32, 23); // Updated range
    tableRange.applyRowBanding();
    tableRange.setFontColor('black'); // Ensure black fonts for table data
  }
  
  // Set column widths for better display - UPDATED for new columns including adjustments
  sheet.setColumnWidth(1, 60);  // Rank
  sheet.setColumnWidth(2, 150); // Assignee
  sheet.setColumnWidth(3, 120); // Completed+Review
  sheet.setColumnWidth(4, 110); // Adjusted Target
  sheet.setColumnWidth(5, 110); // Standard Target
  sheet.setColumnWidth(6, 90);  // Target Gap
  sheet.setColumnWidth(7, 100); // Achievement %
  sheet.setColumnWidth(8, 90);  // Total Tasks
  sheet.setColumnWidth(9, 80);  // Jira Tasks
  sheet.setColumnWidth(10, 80); // Flex Tasks
  sheet.setColumnWidth(11, 100); // Total Prompts
  sheet.setColumnWidth(12, 110); // Avg Prompts/Task
  sheet.setColumnWidth(17, 80); // In Review
  sheet.setColumnWidth(18, 90); // Completion %
  sheet.setColumnWidth(19, 120); // Reviews
  sheet.setColumnWidth(20, 180); // Performance
  sheet.setColumnWidth(21, 120); // Team
  sheet.setColumnWidth(22, 100); // Source
  sheet.setColumnWidth(23, 200); // Adjustments
  
  // Add performance legend with adjustments info
  addDeveloperPerformanceLegend(sheet, lastRow + 2);
}

// Modified to include adjustments information in legend
function addDeveloperPerformanceLegend(sheet, startRow) {
  sheet.getRange(startRow, 1).setValue('PERFORMANCE & TARGET LEGEND (Including Flex Team + Individual Adjustments):');
  sheet.getRange(startRow, 1).setFontWeight('bold').setFontSize(12).setFontColor('black');
  
  const legendData = [
    ['🏆 High Performer (On Target)', '≥100% target achievement (Completed + Review tasks)'],
    ['🌟 Average Performer', '80-99% target achievement or good progress'],
    ['⚠️ Low Performer (Behind Target)', '<80% target achievement (Completed + Review tasks)'],
    ['', ''],
    ['HIGHLIGHTING SYSTEM:', ''],
    ['Green Background', 'Meeting or exceeding target (≥100% achievement)'],
    ['Orange Background', 'Close to target (80-99% achievement)'],
    ['Red Background', 'Behind target (<80% achievement)'],
    ['Blue Background (Flex Tasks)', 'Indicates tasks from Flex team'],
    ['Purple Background (Source)', 'Mixed Jira + Flex tasks'],
    ['Orange Background (Adjustments)', 'Developer has individual target adjustments'],
    ['', ''],
    ['Target Calculation:', `${DEV_DASHBOARD_CONFIG.initialTasksPerDay} tasks/dev/day (first ${DEV_DASHBOARD_CONFIG.initialPeriodDays} working days), then ${DEV_DASHBOARD_CONFIG.regularTasksPerDay} tasks/dev/day`],
    ['Individual Adjustments:', 'Custom targets for leaves, training, or special circumstances'],
    ['Target Metric:', 'Completed Tasks + In Review Tasks vs Adjusted Expected Tasks'],
    ['Data Sources:', 'Jira_Data sheet + External Flex team sheet + Target_Adjustments sheet'],
    ['Flex Team Sheet ID:', DEV_DASHBOARD_CONFIG.FLEX_TEAM_SHEET_ID]
  ];
  
  sheet.getRange(startRow + 1, 1, 17, 2).setValues(legendData);
  sheet.getRange(startRow + 1, 1, 17, 2).setFontColor('black'); // Set all legend text to black
  sheet.getRange(startRow + 1, 1, 17, 1).setFontWeight('bold');
  
  // Color the legend items with black fonts
  sheet.getRange(startRow + 1, 1).setBackground('#00c851').setFontColor('black').setFontWeight('bold'); // High performers
  sheet.getRange(startRow + 2, 1).setBackground('#ff8800').setFontColor('black').setFontWeight('bold'); // Average performers
  sheet.getRange(startRow + 3, 1).setBackground('#ff4444').setFontColor('black').setFontWeight('bold'); // Low performers
  
  // Color the highlighting examples with black fonts
  sheet.getRange(startRow + 6, 1).setBackground('#c8e6c9').setFontWeight('bold').setFontColor('black'); // Green example
  sheet.getRange(startRow + 7, 1).setBackground('#ffe0b2').setFontWeight('bold').setFontColor('black'); // Orange example
  sheet.getRange(startRow + 8, 1).setBackground('#ffcccb').setFontWeight('bold').setFontColor('black'); // Red example
  sheet.getRange(startRow + 9, 1).setBackground('#e1f5fe').setFontWeight('bold').setFontColor('black'); // Flex example
  sheet.getRange(startRow + 10, 1).setBackground('#f3e5f5').setFontWeight('bold').setFontColor('black'); // Mixed example
  sheet.getRange(startRow + 11, 1).setBackground('#fff3e0').setFontWeight('bold').setFontColor('black'); // Adjustments example
  
  // Add borders to legend for better visibility
  sheet.getRange(startRow + 1, 1, 17, 2).setBorder(true, true, true, true, true, true);
}

// Function to set up automatic refresh
function setupDeveloperAutomaticRefresh() {
  // Delete existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'createJiraDashboard') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger to run every 2 hours
  ScriptApp.newTrigger('createJiraDashboard')
    .timeBased()
    .everyHours(2)
    .create();
  
  console.log('Automatic refresh set up for every 2 hours');
}

function stopDeveloperAutomaticRefresh() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'createJiraDashboard') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  SpreadsheetApp.getUi().alert('Automatic refresh stopped');
}

// Bonus feature: Export to PDF
function exportDeveloperDashboardToPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName('Devs Progress');
  
  if (!dashboardSheet) {
    SpreadsheetApp.getUi().alert('Please generate dashboard first');
    return;
  }
  
  const fileName = `Devs_Progress_with_Adjustments_${new Date().toISOString().slice(0, 10)}.pdf`;
  
  // Export only the 'Devs Progress' sheet as PDF
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' +
    'exportFormat=pdf&' +
    'format=pdf&' +
    'size=A4&' +
    'portrait=true&' +
    'fitw=true&' +
    'sheetnames=false&' +
    'printtitle=false&' +
    'pagenumbers=false&' +
    'gridlines=false&' +
    'fzr=false&' +
    'gid=' + dashboardSheet.getSheetId();
  
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  });
  
  const pdfBlob = response.getBlob().setName(fileName);
  const pdfFile = DriveApp.createFile(pdfBlob);
  
  const fileUrl = pdfFile.getUrl();
  
  // Show alert with clickable link
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h3>Dashboard exported successfully!</h3>
      <p><strong>File Name:</strong> ${fileName}</p>
      <p><strong>Link:</strong> <a href="${fileUrl}" target="_blank">${fileUrl}</a></p>
      <p>Dashboard includes Jira data, Flex team data, and individual target adjustments.</p>
      <br>
      <button onclick="google.script.host.close()">Close</button>
    </div>
  `).setWidth(500).setHeight(200);
  
  ui.showModalDialog(htmlOutput, 'PDF Export Complete');
  
  console.log(`PDF exported: ${fileName}`);
  console.log(`PDF URL: ${fileUrl}`);
}

// Helper function to test the script
function testDeveloperDashboard() {
  try {
    createJiraDashboard();
    console.log('Dashboard test completed successfully with Flex team integration and individual adjustments');
  } catch (error) {
    console.error('Dashboard test failed:', error);
  }
}

// Additional utility functions for target management with individual adjustments
function updateDeveloperProjectConfig(startDate, initialTasksPerDay, regularTasksPerDay, initialPeriodDays) {
  // This function allows you to update the project configuration with tiered targets
  DEV_DASHBOARD_CONFIG.startDate = new Date(startDate);
  DEV_DASHBOARD_CONFIG.initialTasksPerDay = initialTasksPerDay;
  DEV_DASHBOARD_CONFIG.regularTasksPerDay = regularTasksPerDay;
  DEV_DASHBOARD_CONFIG.initialPeriodDays = initialPeriodDays;
  
  // Recreate dashboard with new config
  createJiraDashboard();
  
  SpreadsheetApp.getUi().alert(`Project config updated:\nStart Date: ${startDate}\nInitial Target: ${initialTasksPerDay} tasks/dev/day (first ${initialPeriodDays} working days)\nRegular Target: ${regularTasksPerDay} tasks/dev/day\n\nDashboard refreshed with new targets including Flex team data and individual adjustments.`);
}

// Modified to include individual adjustments
function getDevelopersNeedingAttention() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Jira_Data');
  
  if (!dataSheet) {
    throw new Error('Jira_Data sheet not found');
  }
  
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  const jiraRows = data.slice(1);
  
  const columnIndices = {
    assignee: headers.indexOf('Assignee'),
    status: headers.indexOf('Status'),
    numberOfPrompt: headers.indexOf('number of prompt assignee')
  };
  
  // Get Flex team data, target adjustments, and combine
  const flexTasks = getFlexTeamTaskData();
  const targetAdjustments = getTargetAdjustments();
  const jiraTasks = jiraRows.map(row => ({
    assignee: row[columnIndices.assignee] || 'Unassigned',
    status: row[columnIndices.status] || 'Unknown',
    numberOfPrompt: parseInt(row[columnIndices.numberOfPrompt]) || 0,
    isFlexTeam: false
  }));
  
  const allTasks = [...jiraTasks, ...flexTasks];
  const stats = calculateDeveloperStatsWithAdjustedTargets(allTasks, columnIndices, targetAdjustments);
  
  const devsNeedingAttention = [];
  Object.keys(stats).forEach(assignee => {
    const stat = stats[assignee];
    if (!stat.onTrack || parseFloat(stat.targetAchievement) < 80) {
      devsNeedingAttention.push({
        name: assignee,
        completedAndReviewTasks: stat.completedAndReviewTasks,
        expectedTasks: stat.expectedTasks,
        standardExpectedTasks: stat.standardExpectedTasks,
        gap: stat.taskGap,
        achievement: stat.targetAchievement,
        flexTasks: stat.flexTasks,
        jiraTasks: stat.jiraTasks,
        hasAdjustments: stat.hasAdjustments,
        adjustmentSummary: stat.adjustmentSummary,
        source: stat.flexTasks > 0 ? (stat.jiraTasks > 0 ? 'Jira + Flex' : 'Flex Only') : 'Jira Only'
      });
    }
  });
  
  return devsNeedingAttention.sort((a, b) => parseFloat(a.achievement) - parseFloat(b.achievement));
}

// Function to generate target achievement report with individual adjustments
function generateDeveloperTargetReport() {
  const devsNeedingAttention = getDevelopersNeedingAttention();
  const currentDate = new Date();
  const workingDaysElapsed = calculateWorkingDaysElapsed(DEV_DASHBOARD_CONFIG.startDate, currentDate);
  const expectedTasks = calculateTieredExpectedTasks(workingDaysElapsed);
  
  // Get adjustment stats
  const targetAdjustments = getTargetAdjustments();
  const flexTasks = getFlexTeamTaskData();
  
  let report = `TARGET ACHIEVEMENT REPORT (INDIVIDUAL ADJUSTMENTS + FLEX TEAM)\n`;
  report += `Generated: ${currentDate.toLocaleString()}\n`;
  report += `Project Days Elapsed: ${workingDaysElapsed} working days\n`;
  report += `Initial Target: ${DEV_DASHBOARD_CONFIG.initialTasksPerDay} tasks/dev/day (first ${DEV_DASHBOARD_CONFIG.initialPeriodDays} working days)\n`;
  report += `Regular Target: ${DEV_DASHBOARD_CONFIG.regularTasksPerDay} tasks/dev/day (after day ${DEV_DASHBOARD_CONFIG.initialPeriodDays})\n`;
  report += `Standard Expected Tasks per Dev: ${expectedTasks}\n`;
  report += `Active Individual Adjustments: ${targetAdjustments.length}\n`;
  report += `Flex Team Tasks Found: ${flexTasks.length}\n`;
  report += `Metric: Completed Tasks + In Review Tasks vs Adjusted Expected Tasks\n\n`;
  
  // Show breakdown of expected tasks calculation
  if (workingDaysElapsed <= DEV_DASHBOARD_CONFIG.initialPeriodDays) {
    report += `Standard Target Breakdown: ${workingDaysElapsed} days × ${DEV_DASHBOARD_CONFIG.initialTasksPerDay} tasks/day = ${expectedTasks} tasks\n\n`;
  } else {
    const initialTasks = DEV_DASHBOARD_CONFIG.initialPeriodDays * DEV_DASHBOARD_CONFIG.initialTasksPerDay;
    const remainingDays = workingDaysElapsed - DEV_DASHBOARD_CONFIG.initialPeriodDays;
    const remainingTasks = remainingDays * DEV_DASHBOARD_CONFIG.regularTasksPerDay;
    report += `Standard Target Breakdown:\n`;
    report += `  • Initial period: ${DEV_DASHBOARD_CONFIG.initialPeriodDays} days × ${DEV_DASHBOARD_CONFIG.initialTasksPerDay} tasks/day = ${initialTasks} tasks\n`;
    report += `  • Regular period: ${remainingDays} days × ${DEV_DASHBOARD_CONFIG.regularTasksPerDay} tasks/day = ${remainingTasks} tasks\n`;
    report += `  • Total standard expected: ${initialTasks} + ${remainingTasks} = ${expectedTasks} tasks\n\n`;
  }
  
  // Show adjustment summary
  if (targetAdjustments.length > 0) {
    report += `ACTIVE TARGET ADJUSTMENTS:\n`;
    targetAdjustments.forEach((adj, index) => {
      const startStr = adj.startDate.toISOString().slice(0, 10);
      const endStr = adj.endDate.toISOString().slice(0, 10);
      report += `${index + 1}. ${adj.developer}: ${adj.adjustmentType} (${startStr} to ${endStr})\n`;
      report += `   • Reason: ${adj.reason}\n`;
      if (adj.adjustmentType === 'Custom') {
        report += `   • Custom target: ${adj.customTarget} tasks/day\n`;
      }
      report += `\n`;
    });
  }
  
  if (devsNeedingAttention.length === 0) {
    report += `🎉 EXCELLENT! All developers are meeting or exceeding their adjusted targets!\n`;
  } else {
    report += `⚠️ ATTENTION NEEDED: ${devsNeedingAttention.length} developer(s) behind adjusted target:\n\n`;
    
    devsNeedingAttention.forEach((dev, index) => {
      report += `${index + 1}. ${dev.name} (${dev.source})\n`;
      report += `   • Completed + Review Tasks: ${dev.completedAndReviewTasks}\n`;
      report += `   • Adjusted Expected Tasks: ${dev.expectedTasks}\n`;
      if (dev.hasAdjustments) {
        report += `   • Standard Expected Tasks: ${dev.standardExpectedTasks}\n`;
        report += `   • Individual Adjustments: ${dev.adjustmentSummary}\n`;
      }
      report += `   • Gap: ${dev.gap} tasks\n`;
      report += `   • Achievement: ${dev.achievement}%\n`;
      if (dev.flexTasks > 0) {
        report += `   • Flex Tasks: ${dev.flexTasks} | Jira Tasks: ${dev.jiraTasks}\n`;
      }
      report += `\n`;
    });
    
    report += `RECOMMENDATIONS:\n`;
    report += `• Schedule 1-on-1 meetings with developers showing <80% achievement\n`;
    report += `• Investigate blockers and provide additional support\n`;
    report += `• Consider workload redistribution if needed\n`;
    report += `• Review task complexity and estimation accuracy\n`;
    report += `• Focus on moving tasks from 'In Progress' to 'Review' or 'Done'\n`;
    report += `• Check if individual target adjustments need updating\n`;
    report += `• Note: Targets are individually adjusted for leaves and special circumstances\n`;
    report += `• Consider Flex team capacity for workload balancing\n`;
  }
  
  // Display report in a dialog
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px; color: black;">
      <pre style="white-space: pre-wrap
      <br>
      <button onclick="google.script.host.close()">Close</button>
    </div>
  `).setWidth(800).setHeight(700);
  
  ui.showModalDialog(htmlOutput, 'Target Achievement Report with Individual Adjustments');
  
  return report;
}

// Function to calculate expected tasks for a specific number of working days (utility)
function getDeveloperExpectedTasksForDays(days) {
  return calculateTieredExpectedTasks(days);
}

// Function to show current target information including individual adjustments
function showDeveloperCurrentTargetInfo() {
  const currentDate = new Date();
  const workingDaysElapsed = calculateWorkingDaysElapsed(DEV_DASHBOARD_CONFIG.startDate, currentDate);
  const expectedTasks = calculateTieredExpectedTasks(workingDaysElapsed);
  const flexTasks = getFlexTeamTaskData();
  const targetAdjustments = getTargetAdjustments();
  
  let info = `CURRENT TARGET INFORMATION (Including Individual Adjustments)\n`;
  info += `Project Start Date: ${DEV_DASHBOARD_CONFIG.startDate.toDateString()}\n`;
  info += `Working Days Elapsed: ${workingDaysElapsed}\n`;
  info += `Standard Expected Tasks per Developer: ${expectedTasks}\n`;
  info += `Active Individual Adjustments: ${targetAdjustments.length}\n`;
  info += `Flex Team Tasks Available: ${flexTasks.length}\n`;
  info += `Flex Sheet ID: ${DEV_DASHBOARD_CONFIG.FLEX_TEAM_SHEET_ID}\n\n`;
  info += `TARGET STRUCTURE:\n`;
  info += `• Days 1-${DEV_DASHBOARD_CONFIG.initialPeriodDays}: ${DEV_DASHBOARD_CONFIG.initialTasksPerDay} tasks/day\n`;
  info += `• Day ${DEV_DASHBOARD_CONFIG.initialPeriodDays + 1}+: ${DEV_DASHBOARD_CONFIG.regularTasksPerDay} tasks/day`;
  
  SpreadsheetApp.getUi().alert(info);
}

// Function to add a quick adjustment for a developer
function addQuickTargetAdjustment() {
  const ui = SpreadsheetApp.getUi();
  
  // Get developer name
  const developerResponse = ui.prompt(
    'Add Target Adjustment',
    'Enter developer name (exactly as it appears in Jira data):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (developerResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const developer = developerResponse.getResponseText().trim();
  if (!developer) {
    ui.alert('Developer name cannot be empty');
    return;
  }
  
  // Get start date
  const startDateResponse = ui.prompt(
    'Add Target Adjustment',
    'Enter start date (YYYY-MM-DD format):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (startDateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const startDateStr = startDateResponse.getResponseText().trim();
  const startDate = new Date(startDateStr);
  if (isNaN(startDate.getTime())) {
    ui.alert('Invalid start date format. Please use YYYY-MM-DD');
    return;
  }
  
  // Get end date
  const endDateResponse = ui.prompt(
    'Add Target Adjustment',
    'Enter end date (YYYY-MM-DD format):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (endDateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const endDateStr = endDateResponse.getResponseText().trim();
  const endDate = new Date(endDateStr);
  if (isNaN(endDate.getTime())) {
    ui.alert('Invalid end date format. Please use YYYY-MM-DD');
    return;
  }
  
  if (startDate > endDate) {
    ui.alert('Start date cannot be after end date');
    return;
  }
  
  // Get adjustment type
  const adjustmentTypeResponse = ui.prompt(
    'Add Target Adjustment',
    'Enter adjustment type (Leave, Custom, or Partial):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (adjustmentTypeResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const adjustmentType = adjustmentTypeResponse.getResponseText().trim();
  if (!['Leave', 'Custom', 'Partial'].includes(adjustmentType)) {
    ui.alert('Invalid adjustment type. Please use: Leave, Custom, or Partial');
    return;
  }
  
  let customTarget = 0;
  if (adjustmentType === 'Custom') {
    const customTargetResponse = ui.prompt(
      'Add Target Adjustment',
      'Enter custom tasks per day (e.g., 0, 0.5, 1, 2):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (customTargetResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    customTarget = parseFloat(customTargetResponse.getResponseText().trim());
    if (isNaN(customTarget) || customTarget < 0) {
      ui.alert('Invalid custom target. Please enter a number >= 0');
      return;
    }
  }
  
  // Get reason
  const reasonResponse = ui.prompt(
    'Add Target Adjustment',
    'Enter reason for adjustment:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (reasonResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const reason = reasonResponse.getResponseText().trim();
  
  // Create or get the Target Adjustments sheet
  const adjustmentSheet = createTargetAdjustmentsSheet();
  
  // Find the next empty row
  const lastRow = adjustmentSheet.getLastRow();
  const newRow = lastRow + 1;
  
  // Add the new adjustment
  const newData = [
    developer,
    startDateStr,
    endDateStr,
    adjustmentType,
    adjustmentType === 'Custom' ? customTarget : (adjustmentType === 'Leave' ? 0 : ''),
    reason,
    'Active',
    new Date().toISOString().slice(0, 10),
    'Added via quick adjustment'
  ];
  
  adjustmentSheet.getRange(newRow, 1, 1, newData.length).setValues([newData]);
  
  // Refresh dashboard
  createJiraDashboard();
  
  ui.alert(
    'Target Adjustment Added!',
    `Added adjustment for ${developer}\n` +
    `Period: ${startDateStr} to ${endDateStr}\n` +
    `Type: ${adjustmentType}\n` +
    `${adjustmentType === 'Custom' ? `Custom Target: ${customTarget} tasks/day\n` : ''}` +
    `Reason: ${reason}\n\n` +
    'Dashboard has been refreshed with the new adjustment.',
    ui.ButtonSet.OK
  );
}

// Function to view and manage existing adjustments
function viewTargetAdjustments() {
  const adjustments = getTargetAdjustments();
  
  if (adjustments.length === 0) {
    SpreadsheetApp.getUi().alert(
      'No Target Adjustments',
      'No active target adjustments found.\n\n' +
      'Use "Add Target Adjustment" to create one, or check the Target_Adjustments sheet.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  let report = `ACTIVE TARGET ADJUSTMENTS (${adjustments.length} total)\n\n`;
  
  adjustments.forEach((adj, index) => {
    const startStr = adj.startDate.toISOString().slice(0, 10);
    const endStr = adj.endDate.toISOString().slice(0, 10);
    const dateRange = startStr === endStr ? startStr : `${startStr} to ${endStr}`;
    
    report += `${index + 1}. ${adj.developer}\n`;
    report += `   • Period: ${dateRange}\n`;
    report += `   • Type: ${adj.adjustmentType}`;
    if (adj.adjustmentType === 'Custom') {
      report += ` (${adj.customTarget} tasks/day)`;
    }
    report += `\n`;
    report += `   • Reason: ${adj.reason}\n`;
    report += `   • Status: ${adj.status}\n`;
    if (adj.notes) {
      report += `   • Notes: ${adj.notes}\n`;
    }
    report += `\n`;
  });
  
  report += `\nTo modify adjustments, edit the Target_Adjustments sheet directly.`;
  
  const ui = SpreadsheetApp.getUi();
  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px; color: black;">
      <pre style="white-space: pre-wrap; font-family: Arial, sans-serif; color: black;">${report}</pre>
      <br>
      <button onclick="google.script.host.close()">Close</button>
    </div>
  `).setWidth(600).setHeight(500);
  
  ui.showModalDialog(htmlOutput, 'Active Target Adjustments');
}

// Function to verify all adjustments are being loaded correctly
function verifyAllAdjustments() {
  const adjustments = getTargetAdjustments();
  
  console.log(`\n=== Verifying All Target Adjustments ===`);
  console.log(`Total adjustments loaded: ${adjustments.length}`);
  
  if (adjustments.length > 0) {
    console.log('\nAdjustments by developer:');
    const byDeveloper = {};
    adjustments.forEach(adj => {
      if (!byDeveloper[adj.developer]) {
        byDeveloper[adj.developer] = [];
      }
      byDeveloper[adj.developer].push(adj);
    });
    
    Object.keys(byDeveloper).forEach(dev => {
      console.log(`\n${dev}: ${byDeveloper[dev].length} adjustment(s)`);
      byDeveloper[dev].forEach(adj => {
        console.log(`  - ${adj.startDate.toISOString().slice(0, 10)} to ${adj.endDate.toISOString().slice(0, 10)}: ${adj.adjustmentType} (${adj.customTarget} tasks/day)`);
      });
    });
  }
  
  return adjustments;
}

// Function to manually test adjustment calculation for debugging
function manualTestAdjustmentCalculation() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Test Adjustments', 'Enter developer name to test:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
const developerName = response.getResponseText().trim();
    // Get all adjustments for this developer
    const allAdjustments = getTargetAdjustments().filter(adj => adj.developer === developerName);
    // Calculate standard expected tasks
    const currentDate = new Date();
    const workingDaysElapsed = calculateWorkingDaysElapsed(DEV_DASHBOARD_CONFIG.startDate, currentDate);
    const standardExpected = calculateTieredExpectedTasks(workingDaysElapsed);
    // Calculate adjusted expected tasks by applying all adjustments
    let adjustedExpected = 0;
    let adjustedDays = 0;
    let processedDate = new Date(DEV_DASHBOARD_CONFIG.startDate);
    let dayCount = 0;
    while (dayCount < workingDaysElapsed) {
      // For each day, check if any adjustment applies
      let dayTarget = null;
      for (let adj of allAdjustments) {
        if (processedDate >= adj.startDate && processedDate <= adj.endDate) {
          if (adj.adjustmentType === 'Leave') {
            dayTarget = 0;
            break;
          } else if (adj.adjustmentType === 'Custom') {
            dayTarget = adj.customTarget;
            break;
          } else if (adj.adjustmentType === 'Partial') {
            dayTarget = DEV_DASHBOARD_CONFIG.regularTasksPerDay / 2;
            break;
          }
        }
      }
      if (dayTarget === null) {
        // Use tiered logic
        if (dayCount < DEV_DASHBOARD_CONFIG.initialPeriodDays) {
          dayTarget = DEV_DASHBOARD_CONFIG.initialTasksPerDay;
        } else {
          dayTarget = DEV_DASHBOARD_CONFIG.regularTasksPerDay;
        }
      } else {
        adjustedDays++;
      }
      adjustedExpected += dayTarget;
      processedDate.setDate(processedDate.getDate() + 1);
      dayCount++;
    }
    adjustedExpected = Math.round(adjustedExpected * 10) / 10;
    const difference = Math.round((adjustedExpected - standardExpected) * 10) / 10;
    const message = `Developer: ${developerName}\n` +
                    `Standard Expected: ${standardExpected} tasks\n` +
                    `Adjusted Expected: ${adjustedExpected} tasks\n` +
                    `Difference: ${difference} tasks\n` +
                    `Number of Adjustments: ${allAdjustments.length}\n\n` +
                    'Check the console logs for detailed information.';
    ui.alert('Adjustment Test Results', message, ui.ButtonSet.OK);
  }
}

// === Low Performance Comments Feature ===
// Create or get the Low_Performance_Comments sheet
function getLowPerformanceCommentsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Low_Performance_Comments');
  if (!sheet) {
    sheet = ss.insertSheet('Low_Performance_Comments');
    sheet.appendRow(['Developer', 'Comment', 'Last Updated']);
  }
  return sheet;
}

// Add or update a comment for a developer
function addOrUpdateLowPerformanceComment(developer, comment) {
  const sheet = getLowPerformanceCommentsSheet();
  const data = sheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === developer) {
      sheet.getRange(i + 1, 2).setValue(comment);
      sheet.getRange(i + 1, 3).setValue(new Date().toLocaleString());
      found = true;
      break;
    }
  }
  if (!found) {
    sheet.appendRow([developer, comment, new Date().toLocaleString()]);
  }
}

// Remove comment for a developer (if no longer behind target)
function removeLowPerformanceComment(developer) {
  const sheet = getLowPerformanceCommentsSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === developer) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

// Get comment for a developer
function getLowPerformanceComment(developer) {
  const sheet = getLowPerformanceCommentsSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === developer) {
      return data[i][1];
    }
  }
  return '';
}

// Show UI to add/update comment for a developer
function showAddLowPerformanceCommentUI() {
  const ui = SpreadsheetApp.getUi();
  const devResponse = ui.prompt('Low Performance Comment', 'Enter developer name:', ui.ButtonSet.OK_CANCEL);
  if (devResponse.getSelectedButton() !== ui.Button.OK) return;
  const developer = devResponse.getResponseText().trim();
  if (!developer) {
    ui.alert('Developer name cannot be empty');
    return;
  }
  const commentResponse = ui.prompt('Low Performance Comment', 'Enter comment/reason:', ui.ButtonSet.OK_CANCEL);
  if (commentResponse.getSelectedButton() !== ui.Button.OK) return;
  const comment = commentResponse.getResponseText().trim();
  addOrUpdateLowPerformanceComment(developer, comment);
  ui.alert('Comment saved for ' + developer);
}

// Show all low performance comments
function showLowPerformanceCommentsUI() {
  const sheet = getLowPerformanceCommentsSheet();
  const data = sheet.getDataRange().getValues();
  let html = '<div style="font-family:Arial;padding:20px;color:black;">';
  html += '<h3>Low Performance Comments</h3>';
  if (data.length <= 1) {
    html += '<p>No comments found.</p>';
  } else {
    html += '<table border="1" cellpadding="5" style="border-collapse:collapse;color:black;">';
    html += '<tr><th>Developer</th><th>Comment</th><th>Last Updated</th></tr>';
    for (let i = 1; i < data.length; i++) {
      html += `<tr><td>${data[i][0]}</td><td>${data[i][1]}</td><td>${data[i][2]}</td></tr>`;
    }
    html += '</table>';
  }
  html += '<br><button onclick="google.script.host.close()">Close</button></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(600).setHeight(400), 'Low Performance Comments');
}

function autoPopulateLowPerformanceCommentsSheet(stats) {
  const sheet = getLowPerformanceCommentsSheet();
  // Update header to include new columns
  const header = ['Developer', 'Completed+Review', 'Adjusted Target', 'Team', 'Comment', 'Last Updated'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);

  // Read existing comments into a map
  const lastRow = sheet.getLastRow();
  const existingComments = {};
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, header.length).getValues();
    data.forEach(row => {
      const dev = row[0];
      const comment = row[4];
      if (dev) {
        existingComments[dev] = comment;
      }
    });
    // Clear all except header
    sheet.getRange(2, 1, lastRow - 1, header.length).clearContent();
  }

  // Collect data for developers behind target (less than 75% of adjusted target)
  const devsBehind = Object.keys(stats).filter(dev => parseFloat(stats[dev].targetAchievement) < 75);
  const now = new Date().toLocaleString();
  const rows = [];

  devsBehind.forEach(dev => {
    const stat = stats[dev];
    // Determine primary team
    const primaryTeam = stat && stat.teamBreakdown ? Object.keys(stat.teamBreakdown).reduce((a, b) => stat.teamBreakdown[a] > stat.teamBreakdown[b] ? a : b, 'Unknown') : 'Unknown';
    // Use existing comment if present
    const comment = existingComments[dev] || getLowPerformanceComment(dev) || '';
    
    // Use the adjusted target directly from stats (same logic as dev progress)
    const adjustedTarget = stat.expectedTasks;
    
    rows.push([
      dev,
      stat.completedAndReviewTasks,
      adjustedTarget,
      primaryTeam,
      comment,
      now
    ]);
  });

  // Sort rows by team
  rows.sort((a, b) => {
    if (a[3] < b[3]) return -1;
    if (a[3] > b[3]) return 1;
    return 0;
  });

  // Write sorted rows to sheet
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
}

// Update comments sheet after dashboard refresh (auto-populate low performers)
function updateLowPerformanceCommentsAfterDashboard(stats) {
  autoPopulateLowPerformanceCommentsSheet(stats);
}