// ═══════════════════════════════════════════════════════════════
//  DATABASE CONFIGURATION
// ═══════════════════════════════════════════════════════════════
const MASTER_ID = "1pH1eYliOI5R4gU--AwOmE3fYDbRkRewzIEE5noMOM9A";
const PROFILE_FOLDER_ID = "141YogcsGeeOshBYA_jAwa4hca5n2lFhN";
const REVIEWER_FOLDER_ID = "1EFX9aylAeSRmDWU7OeFIPa0llAMD8wC5";

// Dynamic Subjects Configuration
function getSubjectsMap() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("SUBJECTS_MAP");
  if (cached) return JSON.parse(cached);

  const map = {};
  try {
    const ss = SpreadsheetApp.openById(MASTER_ID);
    let sheet = ss.getSheetByName("Subject Enrollment");

    // Automatically create the enrollment sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet("Subject Enrollment");
      sheet.appendRow(["Subject Title", "Subject Code", "Sheet ID", "Status"]);
      sheet.appendRow([
        "Introduction to Philosophy of Human Person",
        "PHILO",
        "1PPfA1h9PHxQgiBoozKBUdQLD8AuyobBQy-jGXUExfYs",
        "Active",
      ]);
      sheet.appendRow([
        "Pag-aaral ng Kasaysayan at Lipunang Pilipino",
        "PKLP",
        "159RZLsZOB2miwF3uLd-N7p9HD6pYqA3VvoHrFuqi4TI",
        "Active",
      ]);
      sheet.appendRow([
        "Philippine Governance and Politics",
        "PGP",
        "1hTTlEMCbZo_Ff2C3w-ZyCxdLG-KOZpgTCNdnm03rOZ0",
        "Active",
      ]);

      // Setup Data Validation for the Status column
      let rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Active", "Inactive"])
        .build();
      sheet.getRange("D2:D100").setDataValidation(rule);

      // Ensure the switch tab knows about these
      try {
        syncSwitchTab();
      } catch (e) {}
    }

    // Automatically create the Class Schedule sheet if it doesn't exist
    let schedSheet = ss.getSheetByName("Class Schedule");
    if (!schedSheet) {
      schedSheet = ss.insertSheet("Class Schedule");
      schedSheet.appendRow([
        "Subject Code",
        "Day",
        "Time In (24H)",
        "Time Out (24H)",
      ]);

      // Data validation for Day
      let dayRule = SpreadsheetApp.newDataValidation()
        .requireValueInList([
          "Monday",
          "Tuesday",
          "Wednesday",
          "Thursday",
          "Friday",
          "Saturday",
          "Sunday",
        ])
        .build();
      schedSheet.getRange("B2:B100").setDataValidation(dayRule);

      // Data validation for Subject Code (using Subject Enrollment Sheet)
      let subjRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(sheet.getRange("B2:B100"))
        .build();
      schedSheet.getRange("A2:A100").setDataValidation(subjRule);
    }

    const data = sheet.getDataRange().getValues();
    // Assume row 1 is headers: [Subject Title, Subject Code, Sheet ID, Status]
    for (let i = 1; i < data.length; i++) {
      let title = String(data[i][0]).trim();
      let code = String(data[i][1]).trim();
      let id = String(data[i][2]).trim();
      let status =
        data[i].length > 3 ? String(data[i][3]).trim().toUpperCase() : "ACTIVE"; // Default to active if missing

      if (code && id && status !== "INACTIVE") {
        map[code] = { id: id, name: title };
      }
    }

    // Fallback if empty (for backward compatibility during transition)
    if (Object.keys(map).length === 0) {
      map["PHILO"] = {
        id: "1PPfA1h9PHxQgiBoozKBUdQLD8AuyobBQy-jGXUExfYs",
        name: "Introduction to Philosophy of Human Person",
      };
      map["PKLP"] = {
        id: "159RZLsZOB2miwF3uLd-N7p9HD6pYqA3VvoHrFuqi4TI",
        name: "Pag-aaral ng Kasaysayan at Lipunang Pilipino",
      };
      map["PGP"] = {
        id: "1hTTlEMCbZo_Ff2C3w-ZyCxdLG-KOZpgTCNdnm03rOZ0",
        name: "Philippine Governance and Politics",
      };
    }
    cache.put("SUBJECTS_MAP", JSON.stringify(map), 300); // cache for 5 minutes
  } catch (e) {
    Logger.log("Error loading subjects: " + e.message);
  }
  return map;
}

function syncSwitchTab() {
  const ss = SpreadsheetApp.openById(MASTER_ID);
  const enrollSheet = ss.getSheetByName("Subject Enrollment");
  let switchSheet = ss.getSheetByName("Switch");
  if (!enrollSheet || !switchSheet) return;

  const enrollData = enrollSheet.getDataRange().getValues();
  const switchData = switchSheet.getDataRange().getValues();

  // Collect existing codes in Switch tab
  let existingCodes = new Set();
  for (let i = 1; i < switchData.length; i++) {
    let code = String(switchData[i][2]).trim(); // Subejct Code column (C)
    if (code) existingCodes.add(code);
  }

  let changesMade = false;
  let nextNum = existingCodes.size + 1;
  // Check enrollment sheet for new codes
  for (let i = 1; i < enrollData.length; i++) {
    let title = String(enrollData[i][0]).trim();
    let code = String(enrollData[i][1]).trim();
    if (code && !existingCodes.has(code)) {
      // Add to Switch tab (default to Unrelease)
      switchSheet.appendRow([nextNum++, title, code, "Unrelease"]);
      changesMade = true;
    }
  }
  return changesMade;
}

function doGet(e) {
  return HtmlService.createTemplateFromFile("Index").evaluate()
    .setTitle("Records Portal")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ═══════════════════════════════════════════════════════════════
//  SYSTEM & STUDENT AUTHENTICATION
// ═══════════════════════════════════════════════════════════════
function getSystemStatus() {
  try {
    const ss = SpreadsheetApp.openById(MASTER_ID);
    const data = ss.getSheetByName("Switch").getDataRange().getValues();
    const masterStatus = String(data[1][5]).trim() || "OPEN";

    let releasedSubjects = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][1]) break;
      if (String(data[i][3]).trim().toUpperCase() === "RELEASE")
        releasedSubjects.push(String(data[i][2]).trim());
    }

    const annSheet = ss.getSheetByName("Announcements");
    let announcements = [];
    if (annSheet) {
      const annData = annSheet.getDataRange().getValues();
      for (let i = 1; i < annData.length; i++) {
        if (annData[i][1]) {
          let dStr =
            annData[i][0] instanceof Date
              ? annData[i][0].toLocaleDateString("en-US", {
                  month: "short",
                  day: "numeric",
                })
              : annData[i][0];
          announcements.push({
            date: dStr || "Notice",
            message: annData[i][1],
            degree: String(annData[i][2] || "CASUAL")
              .trim()
              .toUpperCase(),
          });
        }
      }
    }
    return {
      status: masterStatus,
      releasedSubjects: releasedSubjects,
      announcements: announcements.reverse(),
    };
  } catch (e) {
    return { status: "ERROR", message: e.message };
  }
}

function registerStudent(id, bdate, password) {
  try {
    id = String(id).trim();
    bdate = String(bdate).trim();
    password = String(password).trim();
    if (!id || !bdate || !password)
      return { success: false, message: "All fields are required." };

    const sys = getSystemStatus();
    if (sys.status === "MAINTENANCE" || sys.status === "CLOSED")
      return { success: false, message: "Portal is " + sys.status };

    let isEnrolled = false;
    const subjectsMap = getSubjectsMap();
    for (const code in subjectsMap) {
      if (!sys.releasedSubjects.includes(code)) continue;
      const sheet = SpreadsheetApp.openById(
        subjectsMap[code].id,
      ).getSheetByName("Input Tab");
      if (!sheet) continue;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]).trim() === id) {
          isEnrolled = true;
          break;
        }
      }
      if (isEnrolled) break;
    }
    if (!isEnrolled)
      return { success: false, message: "ID not found in active rosters." };

    const ss = SpreadsheetApp.openById(MASTER_ID);
    let credSheet = ss.getSheetByName("Credentials");
    if (!credSheet) {
      credSheet = ss.insertSheet("Credentials");
      credSheet.appendRow([
        "Student ID",
        "Birthdate",
        "Password",
        "Registered On",
      ]);
    }

    const credData = credSheet.getDataRange().getValues();
    for (let i = 1; i < credData.length; i++) {
      if (String(credData[i][0]).trim() === id)
        return {
          success: false,
          message: "ID already registered. Please log in.",
        };
    }

    const hashedPassword = hashPassword(password);

    credSheet.appendRow([
      id,
      bdate,
      hashedPassword,
      new Date().toLocaleDateString("en-US"),
    ]);
    return { success: true, message: "Registration successful!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function saveStudentProfile(id, profileData, token) {
  const payload = verifyToken(token);
  if (!payload || payload.id !== String(id)) {
    return { success: false, message: "Unauthorized access." };
  }

  try {
    id = String(id).trim();
    if (!id || !profileData)
      return { success: false, message: "Invalid data." };

    const credSheet =
      SpreadsheetApp.openById(MASTER_ID).getSheetByName("Credentials");
    if (!credSheet)
      return { success: false, message: "System Error: Database missing." };

    // Ensure headers exist for profile data
    const headers = credSheet
      .getRange(1, 1, 1, credSheet.getLastColumn())
      .getValues()[0];
    const expectedHeaders = [
      "Student ID",
      "Birthdate",
      "Password",
      "Registered On",
      "Avatar URL",
      "Role",
      "Father's Name",
      "Father's Contact",
      "Mother's Name",
      "Mother's Contact",
      "Address",
      "Email",
      "Last Password Change"
    ];

    // Add missing headers if they don't exist
    for (let i = headers.length; i < expectedHeaders.length; i++) {
      credSheet.getRange(1, i + 1).setValue(expectedHeaders[i]);
    }

    const credData = credSheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < credData.length; i++) {
      if (String(credData[i][0]).trim() === id) {
        rowIndex = i + 1; // 1-based index for Google Sheets
        break;
      }
    }

    if (rowIndex === -1)
      return { success: false, message: "Student not found in Credentials." };

    // Update Columns G to L (Columns 7 to 12)
    credSheet.getRange(rowIndex, 7).setValue(profileData.fatherName || "");
    credSheet.getRange(rowIndex, 8).setValue(profileData.fatherContact || "");
    credSheet.getRange(rowIndex, 9).setValue(profileData.motherName || "");
    credSheet.getRange(rowIndex, 10).setValue(profileData.motherContact || "");
    credSheet.getRange(rowIndex, 11).setValue(profileData.address || "");
    credSheet.getRange(rowIndex, 12).setValue(profileData.email || "");

    return { success: true, message: "Profile saved successfully." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function loginStudent(id, password) {
  try {
    id = String(id).trim();
    password = String(password).trim();
    if (!id || !password)
      return { success: false, message: "Required fields missing." };

    const sys = getSystemStatus();
    if (sys.status === "MAINTENANCE" || sys.status === "CLOSED")
      return { success: false, message: "Portal is " + sys.status };

    const credSheet =
      SpreadsheetApp.openById(MASTER_ID).getSheetByName("Credentials");
    if (!credSheet)
      return { success: false, message: "System Error: Database missing." };

    const credData = credSheet.getDataRange().getValues();
    let valid = false;
    let profileData = null;
    let needsPasswordReset = false;

    const hashedPassword = hashPassword(password);

    for (let i = 1; i < credData.length; i++) {
      if (String(credData[i][0]).trim() === id) {
        let storedPassword = String(credData[i][2]).trim();

        if (storedPassword === hashedPassword) {
          valid = true;
        } else if (storedPassword === password && storedPassword.length !== 64) {
          // If it matches the plain text and isn't a SHA-256 hash length, it's a legacy password
          valid = true;
          needsPasswordReset = true;
        }

        if (valid) {
          // Fetch existing profile data (Columns G-L are indices 6-11)
          profileData = {
            fatherName: credData[i][6] || "",
            fatherContact: credData[i][7] || "",
            motherName: credData[i][8] || "",
            motherContact: credData[i][9] || "",
            address: credData[i][10] || "",
            email: credData[i][11] || "",
          };
          break;
        }
      }
    }

    if (!valid) return { success: false, message: "Invalid ID or Password." };

    if (needsPasswordReset) {
      return { success: true, needsPasswordReset: true, message: "Security Update: Please reset your password." };
    }

    let studentName = "",
      enrolledSubjects = [];
    const subjectsMap = getSubjectsMap();
    for (const code in subjectsMap) {
      if (!sys.releasedSubjects.includes(code)) continue;
      const data = SpreadsheetApp.openById(subjectsMap[code].id)
        .getSheetByName("Input Tab")
        .getDataRange()
        .getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]).trim() === id) {
          studentName = String(data[i][2]).trim();
          enrolledSubjects.push({ code: code, name: subjectsMap[code].name });
          break;
        }
      }
    }
    if (enrolledSubjects.length === 0)
      return {
        success: false,
        message:
          "Your records are currently updating and will be back online shortly.",
      };

    // Fetch Schedule
    let schedule = [];
    try {
      const ss = SpreadsheetApp.openById(MASTER_ID);
      const schedSheet = ss.getSheetByName("Class Schedule");
      if (schedSheet) {
        // Use getValues to correctly handle native Date objects Google Sheets uses for time
        const sData = schedSheet.getDataRange().getValues();
        let enrolledCodes = enrolledSubjects.map((s) => s.code);
        for (let i = 1; i < sData.length; i++) {
          let code = String(sData[i][0]).trim();
          if (code && enrolledCodes.includes(code)) {
            // Robust parsing for time formats
            let formatTime = (val) => {
              if (val instanceof Date) {
                return (
                  val.getHours().toString().padStart(2, "0") +
                  ":" +
                  val.getMinutes().toString().padStart(2, "0")
                );
              } else if (typeof val === "string") {
                // If it's a string, try to parse basic format like "14:30:00" -> "14:30"
                let parts = val.split(":");
                if (parts.length >= 2) {
                  let h = parts[0].trim();
                  let m = parts[1].trim();
                  // handle edge case like "2:30 PM"
                  if (m.includes("PM") && parseInt(h) < 12)
                    h = (parseInt(h) + 12).toString();
                  if (m.includes("AM") && parseInt(h) === 12) h = "00";
                  return (
                    h.padStart(2, "0") +
                    ":" +
                    m.replace(/[^0-9]/g, "").padStart(2, "0")
                  );
                }
              }
              return String(val).trim();
            };

            schedule.push({
              subjectCode: code,
              day: String(sData[i][1]).trim(),
              timeIn: formatTime(sData[i][2]),
              timeOut: formatTime(sData[i][3]),
            });
          }
        }
      }
    } catch (e) {}

    const token = generateToken({
      id: id,
      role: 'student',
      exp: Date.now() + 24 * 60 * 60 * 1000 // 24 hours
    });

    return {
      success: true,
      token: token,
      studentName: studentName,
      subjects: enrolledSubjects,
      schedule: schedule,
      profile: profileData,
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  FACULTY AUTHENTICATION & OPERATIONS
// ═══════════════════════════════════════════════════════════════
function loginFaculty(id, password) {
  try {
    id = String(id).trim();
    password = String(password).trim();
    if (!id || !password)
      return { success: false, message: "Required fields missing." };

    const facSheet =
      SpreadsheetApp.openById(MASTER_ID).getSheetByName("Faculty");
    if (!facSheet)
      return { success: false, message: "Faculty database not configured." };

    const data = facSheet.getDataRange().getValues();
    const subjectsMap = getSubjectsMap();
    const hashedPassword = hashPassword(password);
    for (let i = 1; i < data.length; i++) {
      let storedPassword = String(data[i][1]).trim();
      let isValid = false;
      if (storedPassword === hashedPassword) {
        isValid = true;
      } else if (storedPassword === password && storedPassword.length !== 64) {
        // Legacy plain-text support for faculty
        isValid = true;
      }

      if (
        String(data[i][0]).trim() === id && isValid
      ) {
        let allSubjects = Object.keys(subjectsMap).map((k) => ({
          code: k,
          name: subjectsMap[k].name,
        }));

        const token = generateToken({
          id: id,
          role: 'faculty',
          exp: Date.now() + 24 * 60 * 60 * 1000 // 24 hours
        });

        return {
          success: true,
          token: token,
          facName: String(data[i][2]).trim(),
          subjects: allSubjects,
        };
      }
    }
    return { success: false, message: "Invalid Faculty Credentials." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function fetchSubjectRoster(subjectCode, token) {
  const payload = verifyToken(token);
  if (!payload || payload.role !== 'faculty') {
    return { success: false, message: "Unauthorized access." };
  }

  try {
    const subjectsMap = getSubjectsMap();
    const ss = SpreadsheetApp.openById(subjectsMap[subjectCode].id);
    const inputSheet = ss.getSheetByName("Input Tab");
    const data = inputSheet.getDataRange().getValues();

    const filteredData = data.filter((row, index) => {
      if (index === 0) return true;
      return String(row[1]).trim() !== "";
    });

    return {
      success: true,
      sheetData: filteredData,
      subjectCode: subjectCode,
      subjectName: subjectsMap[subjectCode].name,
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function updateCellData(subjectCode, rowIndex, colIndex, newValue, token) {
  const payload = verifyToken(token);
  if (!payload || payload.role !== 'faculty') {
    return { success: false, message: "Unauthorized access." };
  }
  try {
    const subjectsMap = getSubjectsMap();
    const ss = SpreadsheetApp.openById(subjectsMap[subjectCode].id);
    const sheet = ss.getSheetByName("Input Tab");
    sheet.getRange(rowIndex + 1, colIndex + 1).setValue(newValue);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  REVIEWER MODULE
// ═══════════════════════════════════════════════════════════════
function fetchReviewerFiles(subjectCode, token) {
  const payload = verifyToken(token);
  if (!payload) {
    return { success: false, message: "Unauthorized access." };
  }
  try {
    const mainFolder = DriveApp.getFolderById(REVIEWER_FOLDER_ID);
    const subfolders = mainFolder.searchFolders(
      'title = "' + subjectCode + '"',
    );

    if (!subfolders.hasNext()) {
      return { success: true, files: [] };
    }

    const subjectFolder = subfolders.next();
    const filesIter = subjectFolder.getFiles();
    const files = [];

    while (filesIter.hasNext()) {
      const file = filesIter.next();
      files.push({
        id: file.getId(),
        name: file.getName(),
        mimeType: file.getMimeType(),
        // Direct download link
        url: "https://drive.google.com/uc?export=download&id=" + file.getId(),
      });
    }

    // Sort files alphabetically by name
    files.sort((a, b) => a.name.localeCompare(b.name));

    return { success: true, files: files };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  GLOBAL ATTENDANCE REPORT (PNG GENERATOR BACKEND)
// ═══════════════════════════════════════════════════════════════
function fetchDailyAttendanceReport(dateStr, token) {
  const payload = verifyToken(token);
  if (!payload || payload.role !== 'faculty') {
    return { success: false, message: "Unauthorized access." };
  }
  try {
    const subjectsMap = getSubjectsMap();
    let report = {
      date: dateStr,
      morning: { present: [], absent: [] }, // PKLP
      afternoon: { present: [], absent: [] }, // PGP
    };

    // Helper to extract attendance from a specific subject sheet
    const processSheet = (subjectCode, timeOfDay) => {
      if (!subjectsMap[subjectCode]) return;

      const ss = SpreadsheetApp.openById(subjectsMap[subjectCode].id);

      // We need names from "Input Tab"
      const inputSheet = ss.getSheetByName("Input Tab");
      if (!inputSheet) return;
      const inputData = inputSheet.getDataRange().getValues();
      let nameMap = {}; // ID -> Name
      for (let i = 1; i < inputData.length; i++) {
        let id = String(inputData[i][1]).trim();
        let name = String(inputData[i][2]).trim();
        if (id && name) nameMap[id] = name;
      }

      // We need attendance from "Attendance Sheet"
      const attSheet = ss.getSheetByName("Attendance Sheet");
      if (!attSheet) return;
      const attData = attSheet.getDataRange().getValues();
      const headers = attData[0];

      // Find target column index matching the requested date
      let targetCol = -1;
      for (let c = 2; c < headers.length; c++) {
        let dh = headers[c];
        if (!dh) continue;

        let headerDateStr = "";
        if (dh instanceof Date) {
          // Try to format it as YYYY-MM-DD to compare with dateStr from HTML <input type="date">
          let y = dh.getFullYear();
          let m = (dh.getMonth() + 1).toString().padStart(2, "0");
          let d = dh.getDate().toString().padStart(2, "0");
          headerDateStr = `${y}-${m}-${d}`;
        } else {
          // Attempt parsing if it's a raw string
          let parsed = new Date(dh);
          if (!isNaN(parsed.getTime())) {
            let y = parsed.getFullYear();
            let m = (parsed.getMonth() + 1).toString().padStart(2, "0");
            let d = parsed.getDate().toString().padStart(2, "0");
            headerDateStr = `${y}-${m}-${d}`;
          }
        }

        if (headerDateStr === dateStr) {
          targetCol = c;
          break;
        }
      }

      if (targetCol === -1) return; // Date not found for this subject

      // Process student rows
      for (let i = 1; i < attData.length; i++) {
        let studentId = String(attData[i][0]).trim();
        if (!studentId || studentId === "-") continue;

        let studentName = nameMap[studentId] || studentId;
        let stat = String(attData[i][targetCol]).trim().toUpperCase();

        // Exclude blank or unrecognized statuses to avoid polluting the list
        if (stat === "P") report[timeOfDay].present.push(studentName);
        else if (stat === "A") report[timeOfDay].absent.push(studentName);
      }

      // Sort alphabetically
      report[timeOfDay].present.sort();
      report[timeOfDay].absent.sort();
    };

    processSheet("PKLP", "morning");
    processSheet("PGP", "afternoon");

    return { success: true, data: report };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  STUDENT DATA FETCHING & AVATARS
// ═══════════════════════════════════════════════════════════════
function fetchSubjectData(id, subjectCode, isFaculty, token) {
  const payload = verifyToken(token);
  if (!payload || (payload.role !== 'faculty' && payload.id !== String(id))) {
    return { success: false, message: "Unauthorized access." };
  }

  try {
    const subjectsMap = getSubjectsMap();
    const subj = subjectsMap[subjectCode];
    const ss = SpreadsheetApp.openById(subj.id);
    const inData = ss.getSheetByName("Input Tab").getDataRange().getValues();
    const inHeaders = inData[0];

    let maxScores = {};
    try {
      const configSheet = ss.getSheetByName("Config");
      if(configSheet) {
        const configData = configSheet.getDataRange().getValues();
        for (let r = 1; r < configData.length; r++) {
          let activityTitle = String(configData[r][0]).trim().toUpperCase();
          let maxScore = configData[r][2]; // Column C
          if (activityTitle) {
            maxScores[activityTitle] = maxScore;
          }
        }
      }
    } catch (e) {}

    let studentRow = null,
      studentName = "";
    let classScores = [];
    let currentStudentScore = 0;

    // First pass: Calculate total score for each student to compute class statistics
    for (let i = 1; i < inData.length; i++) {
      let rId = String(inData[i][1]).trim();
      if (!rId) continue;

      let total = 0;
      for (let c = 4; c < inHeaders.length; c++) {
        let val = parseFloat(inData[i][c]);
        if (!isNaN(val)) total += val;
      }
      classScores.push(total);

      if (rId === String(id).trim()) {
        studentRow = inData[i];
        studentName = String(inData[i][2]).trim();
        currentStudentScore = total;
      }
    }

    if (!studentRow) return { success: false, message: "Records not found." };

    let recitations = 0,
      dailyProgress = [],
      majors = { SA1: null, SA2: null, PE: null },
      hasMajors = false;
    for (let c = 4; c < inHeaders.length; c++) {
      let header = String(inHeaders[c]).trim().toUpperCase();
      let val = studentRow[c] === "" ? null : studentRow[c];
      if (!header) continue;

      if (header === "RECITATIONS") {
        recitations =
          typeof val === "number" ? parseFloat(val.toFixed(2)) : val || 0;
      } else if (header === "SA1") {
        majors.SA1 = { score: val, maxScore: maxScores["SA1"] !== undefined && maxScores["SA1"] !== "" ? maxScores["SA1"] : null };
        hasMajors = true;
      } else if (header === "SA2") {
        majors.SA2 = { score: val, maxScore: maxScores["SA2"] !== undefined && maxScores["SA2"] !== "" ? maxScores["SA2"] : null };
        hasMajors = true;
      } else if (header === "PE") {
        majors.PE = { score: val, maxScore: maxScores["PE"] !== undefined && maxScores["PE"] !== "" ? maxScores["PE"] : null };
        hasMajors = true;
      } else {
        let match = header.match(/W(?:EEK)?\s*(\d+)/i);
        let week = match ? "Week " + match[1] : "Other";
        let maxScore = maxScores[header.toUpperCase()] !== undefined && maxScores[header.toUpperCase()] !== "" ? maxScores[header.toUpperCase()] : null;
        dailyProgress.push({ name: inHeaders[c], score: val, week: week, maxScore: maxScore });
      }
    }

    // Comparative Evaluation logic using Z-Score (ONLY if requested by Faculty)
    let evaluation = null;
    if (isFaculty) {
      evaluation = { status: "Not Evaluated", details: "Insufficient data" };
      if (classScores.length > 0) {
        let sum = classScores.reduce((a, b) => a + b, 0);
        let mean = sum / classScores.length;
        let variance =
          classScores.reduce((a, b) => a + Math.pow(b - mean, 2), 0) /
          classScores.length;
        let sd = Math.sqrt(variance);

        let zScore = sd === 0 ? 0 : (currentStudentScore - mean) / sd;

        if (zScore >= 1.0) {
          evaluation = {
            status: "Exceeds Expectations",
            details: "Demonstrates advanced mastery.",
          };
        } else if (zScore >= -0.5) {
          evaluation = {
            status: "Proficient / Meeting",
            details: "Meets grade-level standards.",
          };
        } else if (zScore >= -1.5) {
          evaluation = {
            status: "Developing / Approaching",
            details: "Understands concepts partially.",
          };
        } else {
          evaluation = {
            status: "Below Basic",
            details: "Requires significant intervention.",
          };
        }
      }
    }

    // Fetch Alternatives Data
    let alternatives = {};
    try {
      const altSheet = ss.getSheetByName("Alternatives");
      if (altSheet) {
        const altData = altSheet.getDataRange().getValues();
        for (let r = 0; r < altData.length; r++) {
          let weekStr = String(altData[r][0]).trim().toUpperCase();
          let link = String(altData[r][1]).trim();
          if (weekStr && link) {
            alternatives[weekStr] = link;
          }
        }
      }
    } catch (e) {}

    const attSheet = ss.getSheetByName("Attendance Sheet");
    let attendance = { totalPresent: 0, totalAbsent: 0, records: [] };
    if (attSheet) {
      const attData = attSheet.getDataRange().getValues();
      let attRow = null;
      for (let i = 1; i < attData.length; i++) {
        if (String(attData[i][0]).trim() === String(id).trim()) {
          attRow = attData[i];
          break;
        }
      }
      if (attRow) {
        for (let c = 2; c < attData[0].length; c++) {
          let dh = attData[0][c];
          if (!dh) continue;
          if (dh instanceof Date)
            dh = dh.toLocaleDateString("en-US", {
              month: "short",
              day: "numeric",
            });
          let stat = String(attRow[c]).trim().toUpperCase();
          if (!stat) continue;
          let isP = stat === "P";
          let isA = stat === "A";
          if (isP) attendance.totalPresent++;
          if (isA) attendance.totalAbsent++;
          if (isP || isA)
            attendance.records.push({
              date: dh,
              status: isP ? "present" : "absent",
            });
        }
      }
    }
    return {
      success: true,
      studentName: studentName,
      subjectName: subj.name,
      subjectCode: subjectCode,
      recitations: recitations,
      dailyProgress: dailyProgress,
      hasMajors: hasMajors,
      majors: majors,
      attendance: attendance,
      evaluation: evaluation,
      alternatives: alternatives,
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getStudentQRCode(id, token) {
  const payload = verifyToken(token);
  if (!payload || payload.id !== String(id)) {
    return { success: false, message: "Unauthorized access." };
  }

  try {
    const folderId = "11AZRnI6LkJXKaSx3yBu44pqSUNHetx7d";
    const files = DriveApp.getFolderById(folderId).searchFiles(
      'title contains "' + id + '"'
    );
    if (files.hasNext()) {
      const blob = files.next().getBlob();
      return {
        success: true,
        data: "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes())
      };
    }
    return { success: false, message: "QR not found" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function getStudentAvatar(id, token) {
  const payload = verifyToken(token);
  // Faculty can fetch any avatar, students can only fetch their own
  if (!payload || (payload.role !== 'faculty' && payload.id !== String(id))) {
    return null;
  }

  try {
    const files = DriveApp.getFolderById(PROFILE_FOLDER_ID).searchFiles(
      'title contains "' + id + '"',
    );
    if (files.hasNext()) {
      const blob = files.next().getBlob();
      return (
        "data:" +
        blob.getContentType() +
        ";base64," +
        Utilities.base64Encode(blob.getBytes())
      );
    }
    return null;
  } catch (err) {
    return null;
  }
}

// BATCH AVATAR FETCHING FOR FACULTY ROSTER
function getMultipleAvatars(ids, token) {
  let avatars = {};

  const payload = verifyToken(token);
  if (!payload || payload.role !== 'faculty') {
    return avatars;
  }

  try {
    // 1. Clean the incoming array to prevent false positives from empty rows
    let validIds = ids.filter((id) => id && id !== "-");
    if (validIds.length === 0) return avatars;

    const folder = DriveApp.getFolderById(PROFILE_FOLDER_ID);
    const files = folder.getFiles(); // 2. Grab all files once (Much faster than looping searches)

    while (files.hasNext()) {
      let file = files.next();
      let name = file.getName();

      for (let i = 0; i < validIds.length; i++) {
        // If the file name matches a valid Student ID (using exact match/regex to prevent collision like "123" matching "1234")
        // Word boundary \b fails with underscores (e.g. 123_avatar.jpg), so we check for non-digits instead.
        let regex = new RegExp("(^|[^0-9])" + validIds[i] + "([^0-9]|$)");
        if (regex.test(name)) {
          let blob = file.getBlob();
          avatars[validIds[i]] =
            "data:" +
            blob.getContentType() +
            ";base64," +
            Utilities.base64Encode(blob.getBytes());
          // Remove the ID from our search list so we don't waste time looking for it again
          validIds.splice(i, 1);
          break;
        }
      }
      if (validIds.length === 0) break; // Stop entirely if we found all students to save memory!
    }
  } catch (e) {
    Logger.log("Error in getMultipleAvatars: " + e.message);
  }
  return avatars;
}

function uploadStudentAvatar(id, dataUrl, token) {
  const payload = verifyToken(token);
  if (!payload || payload.id !== String(id)) {
    return { success: false, message: "Unauthorized access." };
  }

  try {
    const folder = DriveApp.getFolderById(PROFILE_FOLDER_ID);
    const oldFiles = folder.searchFiles('title contains "' + id + '"');
    while (oldFiles.hasNext()) oldFiles.next().setTrashed(true);
    const split = dataUrl.split(",");
    folder.createFile(
      Utilities.newBlob(
        Utilities.base64Decode(split[1]),
        split[0].split(";")[0].replace("data:", ""),
        id + "_avatar",
      ),
    );
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function removeStudentAvatar(id, token) {
  const payload = verifyToken(token);
  if (!payload || payload.id !== String(id)) {
    return { success: false, message: "Unauthorized access." };
  }

  try {
    const files = DriveApp.getFolderById(PROFILE_FOLDER_ID).searchFiles(
      'title contains "' + id + '"',
    );
    while (files.hasNext()) files.next().setTrashed(true);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  PASSWORD RESET & OTP
// NOTE: Profile data (including Email and Last Password Change)
// is stored in the "Credentials" sheet, columns G through M.
// ═══════════════════════════════════════════════════════════════

function generateAndSendOTP(id) {
  try {
    id = String(id).trim();
    if (!id) return { success: false, message: "Invalid ID." };

    const ss = SpreadsheetApp.openById(MASTER_ID);
    const credSheet = ss.getSheetByName("Credentials");
    if (!credSheet) return { success: false, message: "System Error: Database missing." };

    const credData = credSheet.getDataRange().getValues();
    const headers = credData[0];

    // Dynamically find indices
    const emailIndex = headers.indexOf("Email");
    const lastChangeIndex = headers.indexOf("Last Password Change");

    if (emailIndex === -1) return { success: false, message: "Email column not found in database." };

    let rowIndex = -1;
    let email = "";
    let lastChange = "";

    for (let i = 1; i < credData.length; i++) {
      if (String(credData[i][0]).trim() === id) {
        rowIndex = i + 1;
        email = credData[i][emailIndex] || "";
        if (lastChangeIndex !== -1) {
            lastChange = credData[i][lastChangeIndex] || "";
        }
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "Student ID not found." };
    if (!email) return { success: false, message: "No email registered for this account. Please contact administrator." };

    // Check if changed in the last 7 days
    if (lastChange) {
      let lastChangeDate = new Date(lastChange);
      let now = new Date();
      let diffTime = Math.abs(now - lastChangeDate);
      let diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      if (diffDays <= 7) {
        return { success: false, message: "You can only change your password once per week." };
      }
    }

    // Generate 6-digit OTP
    const otp = Math.floor(100000 + Math.random() * 900000).toString();

    // Cache OTP for 10 minutes
    const cache = CacheService.getScriptCache();
    cache.put("OTP_" + id, otp, 600);

    // Send Email
    MailApp.sendEmail({
      to: email,
      subject: "Portal Password Reset OTP",
      htmlBody: `<p>Your OTP for password reset is: <b>${otp}</b></p><p>This OTP will expire in 10 minutes.</p>`
    });

    return { success: true, message: "OTP sent to your registered email." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function verifyOTPAndResetPassword(id, otp, newPassword) {
  try {
    id = String(id).trim();
    otp = String(otp).trim();
    newPassword = String(newPassword).trim();

    if (!id || !otp || !newPassword) return { success: false, message: "All fields are required." };

    const cache = CacheService.getScriptCache();
    const cachedOtp = cache.get("OTP_" + id);

    if (!cachedOtp) return { success: false, message: "OTP expired or invalid." };
    if (cachedOtp !== otp) return { success: false, message: "Incorrect OTP." };

    const ss = SpreadsheetApp.openById(MASTER_ID);
    const credSheet = ss.getSheetByName("Credentials");
    const credData = credSheet.getDataRange().getValues();
    const headers = credData[0];

    // Dynamically find indices
    const passIndex = headers.indexOf("Password");
    const lastChangeIndex = headers.indexOf("Last Password Change");

    let rowIndex = -1;

    for (let i = 1; i < credData.length; i++) {
      if (String(credData[i][0]).trim() === id) {
        rowIndex = i + 1;
        break;
      }
    }


    if (rowIndex === -1) return { success: false, message: "Student ID not found." };
    if (passIndex === -1) return { success: false, message: "Database Error: Password column not found." };

    const hashedNewPassword = hashPassword(newPassword);

    // Update password
    credSheet.getRange(rowIndex, passIndex + 1).setValue(hashedNewPassword);

    // Update last change date
    if (lastChangeIndex !== -1) {
        credSheet.getRange(rowIndex, lastChangeIndex + 1).setValue(new Date().toLocaleDateString("en-US"));
    }

    // Clear OTP from cache
    cache.remove("OTP_" + id);

    return { success: true, message: "Password updated successfully!" };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function fetchRawHtmlContent(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    return { success: true, content: file.getBlob().getDataAsString() };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ═══════════════════════════════════════════════════════════════
//  SECURITY: JWT / HMAC TOKENS & HASHING
// ═══════════════════════════════════════════════════════════════

function getSecretKey() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('JWT_SECRET');
  if (!secret) {
    secret = Utilities.getUuid(); // generate a random secret if none exists
    props.setProperty('JWT_SECRET', secret);
  }
  return secret;
}

function base64UrlEncode(str) {
  const bytes = Utilities.newBlob(str).getBytes();
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function base64UrlEncodeArray(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

function generateToken(payload) {
  const header = { alg: "HS256", typ: "JWT" };
  const encodedHeader = base64UrlEncode(JSON.stringify(header));
  const encodedPayload = base64UrlEncode(JSON.stringify(payload));
  const signatureInput = encodedHeader + "." + encodedPayload;

  const signatureBytes = Utilities.computeHmacSha256Signature(
    signatureInput,
    getSecretKey()
  );
  const encodedSignature = base64UrlEncodeArray(signatureBytes);

  return signatureInput + "." + encodedSignature;
}

function verifyToken(token) {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;

    const signatureInput = parts[0] + "." + parts[1];
    const signatureBytes = Utilities.computeHmacSha256Signature(
      signatureInput,
      getSecretKey()
    );
    const expectedSignature = base64UrlEncodeArray(signatureBytes);

    if (parts[2] !== expectedSignature) return null;

    const payloadBlob = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[1]));
    const payload = JSON.parse(payloadBlob.getDataAsString());

    // Check expiration (if we set one, e.g., 24 hours)
    if (payload.exp && Date.now() > payload.exp) return null;

    return payload;
  } catch (e) {
    return null;
  }
}

function hashPassword(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function checkStudentRegistrationForReset(id, bdate) {
  try {
    id = String(id).trim();
    bdate = String(bdate).trim();

    if (!id || !bdate) return { success: false, message: "ID and Birthdate are required." };

    const ss = SpreadsheetApp.openById(MASTER_ID);
    const credSheet = ss.getSheetByName("Credentials");
    if (!credSheet) return { success: false, message: "Database Error." };

    const credData = credSheet.getDataRange().getValues();
    const headers = credData[0];
    const emailIndex = headers.indexOf("Email");

    for (let i = 1; i < credData.length; i++) {
      if (String(credData[i][0]).trim() === id) {
        let storedBdate = String(credData[i][1]).trim();
        if (storedBdate === bdate) {
          let email = emailIndex !== -1 ? credData[i][emailIndex] : null;
          return { success: true, email: email, message: "Verified" };
        } else {
          return { success: false, message: "Incorrect Birthdate." };
        }
      }
    }

    return { success: false, message: "Student ID not found." };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function registerEmailAndSendOTP(id, bdate, email) {
  try {
    id = String(id).trim();
    bdate = String(bdate).trim();
    email = String(email).trim();

    if (!id || !bdate || !email) return { success: false, message: "All fields are required." };

    const verifyRes = checkStudentRegistrationForReset(id, bdate);
    if (!verifyRes.success) return verifyRes;

    const ss = SpreadsheetApp.openById(MASTER_ID);
    const credSheet = ss.getSheetByName("Credentials");
    const credData = credSheet.getDataRange().getValues();
    const headers = credData[0];
    const emailIndex = headers.indexOf("Email");

    if (emailIndex === -1) return { success: false, message: "Email column missing." };

    let rowIndex = -1;
    for (let i = 1; i < credData.length; i++) {
      if (String(credData[i][0]).trim() === id) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: "Student ID not found." };

    // Save email
    credSheet.getRange(rowIndex, emailIndex + 1).setValue(email);

    // Now trigger send OTP using existing function
    return generateAndSendOTP(id);

  } catch(err) {
    return { success: false, message: err.message };
  }
}
