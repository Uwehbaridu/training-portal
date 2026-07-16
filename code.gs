// ── CORS headers for GitHub Pages requests ───────────────
function setCORSHeaders(output) {
  return output
    .setHeader('Access-Control-Allow-Origin', '*')
    .setHeader('Access-Control-Allow-Methods', 'GET, POST')
    .setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

// ── Cache helpers ────────────────────────────────────────
function getCache() {
  return CacheService.getScriptCache();
}

function getCacheKey(baseKey) {
  const ssId       = SpreadsheetApp.getActiveSpreadsheet().getId();
  const lastUpdate = DriveApp.getFileById(ssId).getLastUpdated().getTime();
  return baseKey + '_' + lastUpdate;
}

// ── Route all GET requests ───────────────────────────────
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'getCourses') {
    return setCORSHeaders(
      ContentService.createTextOutput(
        JSON.stringify(getCourses())
      ).setMimeType(ContentService.MimeType.JSON)
    );
  }

  if (action === 'getStaff') {
    const id = e.parameter.id || '';
    return setCORSHeaders(
      ContentService.createTextOutput(
        JSON.stringify(getStaffDetails(id))
      ).setMimeType(ContentService.MimeType.JSON)
    );
  }

  if (action === 'getCourse') {
    const title = e.parameter.title || '';
    return setCORSHeaders(
      ContentService.createTextOutput(
        JSON.stringify(getCourseDetails(title))
      ).setMimeType(ContentService.MimeType.JSON)
    );
  }

  // Default — return empty JSON
  return setCORSHeaders(
    ContentService.createTextOutput('{}')
      .setMimeType(ContentService.MimeType.JSON)
  );
}

// ── Route all POST requests ──────────────────────────────
function doPost(e) {
  try {
    const record = JSON.parse(e.postData.contents);
    const result = saveAttendance(record);
    return setCORSHeaders(
      ContentService.createTextOutput(
        JSON.stringify({ status: result })
      ).setMimeType(ContentService.MimeType.JSON)
    );
  } catch(err) {
    return setCORSHeaders(
      ContentService.createTextOutput(
        JSON.stringify({ error: err.message })
      ).setMimeType(ContentService.MimeType.JSON)
    );
  }
}

// ── Get course list ──────────────────────────────────────
function getCourses() {
  try {
    const cache      = getCache();
    const dynamicKey = getCacheKey('course_list');
    const cached     = cache.get(dynamicKey);
    if (cached) return JSON.parse(cached);

    const sheet = SpreadsheetApp.getActiveSpreadsheet()
                    .getSheetByName('Course Details');
    if (!sheet) return { error: 'Course Details sheet not found' };

    const data    = sheet.getDataRange().getValues();
    const courses = [];

    for (let i = 1; i < data.length; i++) {
      if (data[i][1]) courses.push(data[i][1].toString().trim());
    }

    courses.sort();
    cache.put(dynamicKey, JSON.stringify(courses), 21600);
    return courses;

  } catch(e) {
    return { error: e.message };
  }
}

// ── Get staff details ────────────────────────────────────
function getStaffDetails(EmpID) {
  try {
    const searchId   = EmpID.toString().trim().toUpperCase();
    const cache      = getCache();
    const dynamicKey = getCacheKey('emp_' + searchId);
    const cached     = cache.get(dynamicKey);
    if (cached) return JSON.parse(cached);

    const sheet = SpreadsheetApp.getActiveSpreadsheet()
                    .getSheetByName('Emp Data');
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toUpperCase() === searchId) {
        const result = {
          name:     data[i][1],
          gender:   data[i][2],
          dept:     data[i][3],
          company:  data[i][4],
          empType:  data[i][5],
          staffCat: data[i][6]
        };
        cache.put(dynamicKey, JSON.stringify(result), 3600);
        return result;
      }
    }
    return null;

  } catch(e) {
    return null;
  }
}

// ── Get course details ───────────────────────────────────
function getCourseDetails(title) {
  try {
    const cache      = getCache();
    const dynamicKey = getCacheKey('course_' + title);
    const cached     = cache.get(dynamicKey);
    if (cached) return JSON.parse(cached);

    const sheet = SpreadsheetApp.getActiveSpreadsheet()
                    .getSheetByName('Course Details');
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === title.trim()) {
        const result = {
          type:     data[i][2],
          faculty:  data[i][3],
          category: data[i][4],
          location: data[i][5],
          hours:    data[i][6],
          days:     data[i][7],
          manHours: data[i][8]
        };
        cache.put(dynamicKey, JSON.stringify(result), 21600);
        return result;
      }
    }
    return null;

  } catch(e) {
    return null;
  }
}

// ── Save attendance ──────────────────────────────────────
function saveAttendance(record) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Attendance');

  sheet.appendRow([
    record.empId,
    record.name,
    record.gender,
    record.empType,
    record.staffCat,
    record.dept,
    record.company,
    record.courseTitle,
    record.type,
    record.faculty,
    record.category,
    record.location,
    record.startDate,
    record.endDate,
    record.hours,
    record.days,
    record.manHours
  ]);

  return 'Success';
}
