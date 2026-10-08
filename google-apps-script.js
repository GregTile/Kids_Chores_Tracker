/**
 * Google Apps Script for Kid Activity / Chore Tracker (Multi-Device Sync)
 *
 * HOW IT WORKS:
 * 1. "Logs" Tab  -> Stores all completed chores with timestamps.
 * 2. "Kids" Tab  -> Defines the children profiles (Name, Avatar, Color).
 * 3. "Tasks" Tab -> Defines the chore list (Title, Icon, Style).
 *
 * All iPads fetch the "Kids" and "Tasks" tabs when opening the app,
 * keeping every iPad 100% in sync without manual re-configuration!
 */

function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Logs Sheet
  var logsSheet = ss.getSheetByName("Logs") || ss.insertSheet("Logs");
  if (logsSheet.getLastRow() === 0) {
    logsSheet.appendRow(["Timestamp", "Date", "Time", "Child", "Activity", "Notes", "Device"]);
    logsSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#4F46E5").setFontColor("#FFFFFF");
    logsSheet.setFrozenRows(1);
  }

  // 2. Setup Kids Sheet (Default profiles if empty)
  var kidsSheet = ss.getSheetByName("Kids") || ss.insertSheet("Kids");
  if (kidsSheet.getLastRow() === 0) {
    kidsSheet.appendRow(["ID", "Name", "Avatar", "Color"]);
    kidsSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#059669").setFontColor("#FFFFFF");
    kidsSheet.setFrozenRows(1);

    // Initial default kids
    kidsSheet.appendRow(["child-1", "Mike", "🦁", "from-blue-500 to-indigo-600"]);
    kidsSheet.appendRow(["child-2", "John", "🚀", "from-emerald-500 to-teal-600"]);
    kidsSheet.appendRow(["child-3", "Emma", "🦄", "from-fuchsia-500 to-pink-600"]);
  }

  // 3. Setup Tasks Sheet (Default chores if empty)
  var tasksSheet = ss.getSheetByName("Tasks") || ss.insertSheet("Tasks");
  if (tasksSheet.getLastRow() === 0) {
    tasksSheet.appendRow(["ID", "Title", "Icon", "Color"]);
    tasksSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#D97706").setFontColor("#FFFFFF");
    tasksSheet.setFrozenRows(1);

    // Initial default chores
    tasksSheet.appendRow(["task-1", "Done Homework", "📚", "border-blue-400 bg-blue-50 text-blue-900"]);
    tasksSheet.appendRow(["task-2", "Practiced Piano", "🎹", "border-purple-400 bg-purple-50 text-purple-900"]);
    tasksSheet.appendRow(["task-3", "Cleaned Bedroom", "🛏️", "border-emerald-400 bg-emerald-50 text-emerald-900"]);
    tasksSheet.appendRow(["task-4", "Helped Clean Basement", "🧹", "border-amber-400 bg-amber-50 text-amber-900"]);
    tasksSheet.appendRow(["task-5", "Emptied Dishwasher", "🍽️", "border-cyan-400 bg-cyan-50 text-cyan-900"]);
    tasksSheet.appendRow(["task-6", "Read for 30 Mins", "📖", "border-rose-400 bg-rose-50 text-rose-900"]);
  }
}

/**
 * GET request: Returns the shared Kids and Tasks configuration to all iPads.
 */
function doGet(e) {
  setupSpreadsheet();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Read Kids tab
  var kidsSheet = ss.getSheetByName("Kids");
  var kidsData = kidsSheet.getDataRange().getValues();
  var kids = [];
  for (var i = 1; i < kidsData.length; i++) {
    var row = kidsData[i];
    if (row[1] && row[1].toString().trim() !== "") {
      kids.push({
        id: row[0] ? row[0].toString() : "child-" + i,
        name: row[1].toString().trim(),
        avatar: row[2] ? row[2].toString().trim() : "⭐",
        color: row[3] ? row[3].toString().trim() : "from-blue-500 to-indigo-600"
      });
    }
  }

  // Read Tasks tab
  var tasksSheet = ss.getSheetByName("Tasks");
  var tasksData = tasksSheet.getDataRange().getValues();
  var tasks = [];
  for (var j = 1; j < tasksData.length; j++) {
    var tRow = tasksData[j];
    if (tRow[1] && tRow[1].toString().trim() !== "") {
      tasks.push({
        id: tRow[0] ? tRow[0].toString() : "task-" + j,
        title: tRow[1].toString().trim(),
        icon: tRow[2] ? tRow[2].toString().trim() : "✅",
        color: tRow[3] ? tRow[3].toString().trim() : "border-slate-300 bg-white text-slate-800"
      });
    }
  }

  // Read Logs tab (up to 300 recent entries from the past 7-8 days)
  var logsSheet = ss.getSheetByName("Logs");
  var logs = [];
  if (logsSheet && logsSheet.getLastRow() > 1) {
    var logsData = logsSheet.getDataRange().getValues();
    var timezone = Session.getScriptTimeZone() || "America/New_York";

    // 8-day cutoff buffer to cover full 7 calendar days across all time zones
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 8);
    cutoff.setHours(0, 0, 0, 0);

    // Read backwards so the newest logs appear first
    for (var k = logsData.length - 1; k >= 1; k--) {
      var lRow = logsData[k];
      var child = lRow[3] ? lRow[3].toString().trim() : "";
      var activity = lRow[4] ? lRow[4].toString().trim() : "";
      if (!child || !activity) continue;

      var entryDate = null;
      if (lRow[0] instanceof Date) {
        entryDate = lRow[0];
      } else if (lRow[0] && !isNaN(new Date(lRow[0]).getTime())) {
        entryDate = new Date(lRow[0]);
      } else if (lRow[1] instanceof Date) {
        entryDate = lRow[1];
      } else if (lRow[1] && !isNaN(new Date(lRow[1]).getTime())) {
        entryDate = new Date(lRow[1]);
      }

      if (entryDate && entryDate < cutoff) {
        continue;
      }

      var dateStr = "";
      if (lRow[1] instanceof Date) {
        dateStr = Utilities.formatDate(lRow[1], timezone, "yyyy-MM-dd");
      } else if (lRow[1]) {
        var rawD = lRow[1].toString().trim();
        if (/^\d{4}-\d{2}-\d{2}$/.test(rawD)) {
          dateStr = rawD;
        } else if (entryDate) {
          dateStr = Utilities.formatDate(entryDate, timezone, "yyyy-MM-dd");
        } else {
          dateStr = rawD;
        }
      } else if (entryDate) {
        dateStr = Utilities.formatDate(entryDate, timezone, "yyyy-MM-dd");
      }

      var timeStr = "";
      if (lRow[2] instanceof Date) {
        timeStr = Utilities.formatDate(lRow[2], timezone, "hh:mm a");
      } else if (lRow[2]) {
        timeStr = lRow[2].toString().trim();
      } else if (entryDate) {
        timeStr = Utilities.formatDate(entryDate, timezone, "hh:mm a");
      }

      logs.push({
        id: "sheet-" + k + "-" + (entryDate ? entryDate.getTime() : k),
        timestamp: entryDate ? entryDate.toISOString() : (lRow[0] ? lRow[0].toString() : ""),
        date: dateStr,
        time: timeStr,
        child: child,
        activity: activity,
        notes: lRow[5] ? lRow[5].toString().trim() : "",
        device: lRow[6] ? lRow[6].toString().trim() : ""
      });

      if (logs.length >= 300) break;
    }
  }

  var response = {
    status: "success",
    timestamp: new Date().toISOString(),
    kids: kids,
    tasks: tasks,
    logs: logs
  };

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * POST request: Logs activity completions OR saves configuration changes.
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    setupSpreadsheet();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Parse incoming JSON
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    // Action 1: Save updated configuration from Parent Settings (Kids and Tasks)
    if (data.action === "save_config") {
      if (Array.isArray(data.kids)) {
        var kidsSheet = ss.getSheetByName("Kids");
        kidsSheet.clearContents();
        kidsSheet.appendRow(["ID", "Name", "Avatar", "Color"]);
        kidsSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#059669").setFontColor("#FFFFFF");
        data.kids.forEach(function(k) {
          kidsSheet.appendRow([k.id || "child-" + Date.now(), k.name, k.avatar || "⭐", k.color || "from-blue-500 to-indigo-600"]);
        });
      }

      if (Array.isArray(data.tasks)) {
        var tasksSheet = ss.getSheetByName("Tasks");
        tasksSheet.clearContents();
        tasksSheet.appendRow(["ID", "Title", "Icon", "Color"]);
        tasksSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#D97706").setFontColor("#FFFFFF");
        data.tasks.forEach(function(t) {
          tasksSheet.appendRow([t.id || "task-" + Date.now(), t.title, t.icon || "✅", t.color || "border-slate-300 bg-white text-slate-800"]);
        });
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Configuration updated successfully across all iPads"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Action 2: Default -> Log a completed activity
    var logsSheet = ss.getSheetByName("Logs");
    var now = new Date();
    var timezone = Session.getScriptTimeZone() || "America/New_York";
    var formattedDate = data.date || Utilities.formatDate(now, timezone, "yyyy-MM-dd");
    var formattedTime = data.time || Utilities.formatDate(now, timezone, "hh:mm:ss a");

    var child = data.child || "Unknown";
    var activity = data.activity || data.task || "Unknown Activity";
    var notes = data.notes || "";
    var device = data.device || "iPad";

    logsSheet.appendRow([
      now.toISOString(),
      formattedDate,
      formattedTime,
      child,
      activity,
      notes,
      device
    ]);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Activity logged successfully",
      recorded: { date: formattedDate, time: formattedTime, child: child, activity: activity }
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}
