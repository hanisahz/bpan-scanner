
const CFG = {
  SHEETS: {
    RSVP: "RSVPAttendance",
    ATTENDEES: "Attendees",
    TOKENS: "QR_Tokens",
    LOG: "Attendance_Log"
  },
  TOKEN_TTL_SECONDS: 75,
  TOKEN_REFRESH_SECONDS: 45
};

/**
 * Run setupDinnerSystem() manually from the spreadsheet-bound
 * Apps Script project before deploying the web app.
 */
function setupDinnerSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      "Open Apps Script from your Google Sheet using Extensions > Apps Script, then run setupDinnerSystem()."
    );
  }

  // Save the spreadsheet ID so web requests can open it explicitly.
  PropertiesService.getScriptProperties()
    .setProperty("SPREADSHEET_ID", ss.getId());

  const source = ss.getSheetByName(CFG.SHEETS.RSVP);

  if (!source) {
    throw new Error(
      'Required sheet "RSVPAttendance" was not found.'
    );
  }

  const expected = [
    "Employee Number",
    "RSVP_ID",
    "Name",
    "Spouse Name",
    "Table Number"
  ];

  const headers = source
    .getRange(1, 1, 1, 5)
    .getDisplayValues()[0]
    .map(v => String(v).trim());

  if (expected.some((h, i) => headers[i] !== h)) {
    throw new Error(
      "Check row 1 of RSVPAttendance. Required headers: " +
      expected.join(" | ")
    );
  }

  ensureSheet_(ss, CFG.SHEETS.ATTENDEES, [
    "Attendee ID",
    "RSVP_ID",
    "Employee Number",
    "Name",
    "Type",
    "Table Number",
    "Status",
    "Check-in At",
    "Checked-in By"
  ]);

  ensureSheet_(ss, CFG.SHEETS.TOKENS, [
    "Token Hash",
    "RSVP_ID",
    "Expires At",
    "Created At"
  ]);

  ensureSheet_(ss, CFG.SHEETS.LOG, [
    "Timestamp",
    "Action",
    "RSVP_ID",
    "Attendee ID",
    "Employee Number",
    "Attendee Name",
    "Officer Staff ID",
    "Details"
  ]);

  rebuildAttendees_();

  Logger.log(
    "Setup complete. Spreadsheet ID saved; attendee records rebuilt."
  );
}

/**
 * Open the configured spreadsheet for web requests.
 */
function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties()
    .getProperty("SPREADSHEET_ID");

  if (!id) {
    throw new Error(
      "Spreadsheet ID is not configured. Run setupDinnerSystem() from the bound script first."
    );
  }

  return SpreadsheetApp.openById(id);
}

function doGet(e) {
  const page = e && e.parameter && e.parameter.page
    ? e.parameter.page
    : "pass";

  const file = page === "scanner" ? "Scanner" : "Pass";

  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle(
      page === "scanner"
        ? "Annual Dinner Check-in"
        : "Annual Dinner Pass"
    )
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
  }

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sh.setFrozenRows(1);
  return sh;
}

/**
 * Build one attendee row for each employee and, where applicable,
 * one additional row for their spouse.
 *
 * Existing check-in statuses are preserved when rebuilding.
 */
function rebuildAttendees_() {
  const ss = getSpreadsheet_();
  const rsvp = ss.getSheetByName(CFG.SHEETS.RSVP);
  const sh = ss.getSheetByName(CFG.SHEETS.ATTENDEES);

  if (!rsvp || !sh) {
    throw new Error(
      'Required sheet "RSVPAttendance" or "Attendees" is missing. Run setupDinnerSystem().'
    );
  }

  const last = sh.getLastRow();

  const existing = last > 1
    ? sh.getRange(2, 1, last - 1, 9).getValues()
    : [];

  const byKey = {};

  existing.forEach(r => {
    byKey[String(r[1]) + "|" + String(r[4])] = r;
  });

  const sourceLast = rsvp.getLastRow();

  const data = sourceLast > 1
    ? rsvp.getRange(2, 1, sourceLast - 1, 5).getDisplayValues()
    : [];

  const rows = [];

  data.forEach(r => {
    const employeeNo = String(r[0] || "").trim();
    const rsvpId = String(r[1] || "").trim();
    const employeeName = String(r[2] || "").trim();
    const spouseName = String(r[3] || "").trim();
    const table = r[4];

    if (!employeeNo || !rsvpId || !employeeName) return;

    const oldEmployee = byKey[rsvpId + "|Employee"];

    rows.push(oldEmployee
      ? [
          oldEmployee[0], rsvpId, employeeNo, employeeName,
          "Employee", table, oldEmployee[6] || "PENDING",
          oldEmployee[7] || "", oldEmployee[8] || ""
        ]
      : [
          "ATT-" + rsvpId + "-E", rsvpId, employeeNo,
          employeeName, "Employee", table, "PENDING", "", ""
        ]
    );

    const hasSpouse =
      spouseName &&
      !/^n\/?a$/i.test(spouseName) &&
      spouseName.toLowerCase() !== "none";

    if (hasSpouse) {
      const oldSpouse = byKey[rsvpId + "|Spouse"];

      rows.push(oldSpouse
        ? [
            oldSpouse[0], rsvpId, employeeNo, spouseName,
            "Spouse", table, oldSpouse[6] || "PENDING",
            oldSpouse[7] || "", oldSpouse[8] || ""
          ]
        : [
            "ATT-" + rsvpId + "-S", rsvpId, employeeNo,
            spouseName, "Spouse", table, "PENDING", "", ""
          ]
      );
    }
  });

  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 9).clearContent();
  }

  if (rows.length) {
    sh.getRange(2, 1, rows.length, 9).setValues(rows);
  }
}

/**
 * Employee page: look up a reservation using Employee Number.
 */
function getPass(employeeNumber) {
  const record = findReservation_(employeeNumber);

  if (!record) {
    throw new Error(
      "Employee number was not found. Please check it and try again."
    );
  }

  const attendees = getReservationAttendees_(record.rsvpId);

  if (!attendees.length) {
    throw new Error(
      "No attendee records were found. Please contact the event team."
    );
  }

  return {
    reservationId: record.rsvpId,
    employeeNumber: record.employeeNumber,
    employeeName: record.name,
    spouseName: record.spouseName,
    tableNumber: record.tableNumber,
    attendees: attendees,
    token: issueToken_(record.rsvpId),
    refreshSeconds: CFG.TOKEN_REFRESH_SECONDS
  };
}

function findReservation_(employeeNumber) {
  const sh = getSpreadsheet_().getSheetByName(CFG.SHEETS.RSVP);

  if (!sh) {
    throw new Error('Sheet "RSVPAttendance" was not found.');
  }

  const last = sh.getLastRow();

  const values = last > 1
    ? sh.getRange(2, 1, last - 1, 5).getDisplayValues()
    : [];

  const wanted = String(employeeNumber || "")
    .trim()
    .toUpperCase();

  if (!wanted) return null;

  for (let i = 0; i < values.length; i++) {
    if (
      String(values[i][0] || "").trim().toUpperCase() === wanted
    ) {
      return {
        employeeNumber: String(values[i][0]).trim(),
        rsvpId: String(values[i][1]).trim(),
        name: String(values[i][2]).trim(),
        spouseName: String(values[i][3] || "").trim(),
        tableNumber: values[i][4]
      };
    }
  }

  return null;
}

function getReservationAttendees_(reservationId) {
  const sh = getSpreadsheet_().getSheetByName(CFG.SHEETS.ATTENDEES);

  if (!sh) {
    throw new Error('Sheet "Attendees" was not found.');
  }

  const last = sh.getLastRow();
  if (last < 2) return [];

  return sh.getRange(2, 1, last - 1, 9)
    .getValues()
    .filter(r => String(r[1]) === String(reservationId))
    .map(r => ({
      attendeeId: String(r[0]),
      name: String(r[3]),
      type: String(r[4]),
      tableNumber: r[5],
      status: String(r[6] || "PENDING")
    }));
}

/**
 * Create a random QR token and store only its SHA-256 hash.
 */
function issueToken_(reservationId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const token =
      Utilities.getUuid().replace(/-/g, "") +
      Utilities.getUuid().replace(/-/g, "");

    const now = new Date();
    const expires = new Date(
      now.getTime() + CFG.TOKEN_TTL_SECONDS * 1000
    );

    const sh = getSpreadsheet_().getSheetByName(CFG.SHEETS.TOKENS);

    if (!sh) {
      throw new Error('Sheet "QR_Tokens" was not found.');
    }

    sh.appendRow([
      sha256_(token),
      reservationId,
      expires,
      now
    ]);

    if (sh.getLastRow() > 2500) {
      pruneExpiredTokens_();
    }

    return {
      value: token,
      expiresAt: expires.toISOString()
    };
  } finally {
    lock.releaseLock();
  }
}

function refreshPass(employeeNumber) {
  const record = findReservation_(employeeNumber);

  if (!record) {
    throw new Error(
      "Employee number was not found. Please reopen your pass."
    );
  }

  return {
    token: issueToken_(record.rsvpId),
    refreshSeconds: CFG.TOKEN_REFRESH_SECONDS
  };
}

function pruneExpiredTokens_() {
  const sh = getSpreadsheet_().getSheetByName(CFG.SHEETS.TOKENS);

  if (!sh || sh.getLastRow() < 2) return;

  const last = sh.getLastRow();
  const rows = sh.getRange(2, 1, last - 1, 4).getValues();
  const now = Date.now();

  const keep = rows.filter(r =>
    r[2] && new Date(r[2]).getTime() > now
  );

  sh.getRange(2, 1, last - 1, 4).clearContent();

  if (keep.length) {
    sh.getRange(2, 1, keep.length, 4).setValues(keep);
  }
}

function verifyToken_(token) {
  if (!token) {
    throw new Error("No QR token was received.");
  }

  const sh = getSpreadsheet_().getSheetByName(CFG.SHEETS.TOKENS);

  if (!sh) {
    throw new Error('Sheet "QR_Tokens" was not found.');
  }

  const last = sh.getLastRow();

  if (last < 2) {
    throw new Error("This QR pass is invalid or expired.");
  }

  const hash = sha256_(String(token));
  const rows = sh.getRange(2, 1, last - 1, 4).getValues();
  const now = Date.now();

  for (let i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]) !== hash) continue;

    const expiresAt = new Date(rows[i][2]).getTime();

    if (!expiresAt || expiresAt < now) {
      throw new Error(
        "This QR pass has expired. Ask the attendee to refresh it."
      );
    }

    return String(rows[i][1]);
  }

  throw new Error(
    "This QR pass is invalid. Ask the attendee to refresh it."
  );
}

/**
 * Staff ID is recorded as the officer identifier.
 * This validates that an ID was entered; it does not authenticate it.
 */
function verifyOfficer_(staffId) {
  const id = String(staffId || "").trim();

  if (!id) {
    throw new Error("Please enter your Staff ID.");
  }

  if (id.length > 80) {
    throw new Error("Please enter a valid Staff ID.");
  }

  return id;
}

function inspectToken(token, staffId) {
  verifyOfficer_(staffId);

  const reservationId = verifyToken_(token);
  const rsvpSh = getSpreadsheet_().getSheetByName(CFG.SHEETS.RSVP);

  if (!rsvpSh) {
    throw new Error('Sheet "RSVPAttendance" was not found.');
  }

  const last = rsvpSh.getLastRow();

  const values = last > 1
    ? rsvpSh.getRange(2, 1, last - 1, 5).getDisplayValues()
    : [];

  let reservation = null;

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][1]).trim() === reservationId) {
      reservation = {
        reservationId: String(values[i][1]),
        employeeNumber: String(values[i][0]),
        employeeName: String(values[i][2]),
        spouseName: String(values[i][3]),
        tableNumber: values[i][4]
      };
      break;
    }
  }

  if (!reservation) {
    throw new Error("Reservation was not found in RSVPAttendance.");
  }

  reservation.attendees = getReservationAttendees_(reservationId);
  return reservation;
}

function checkIn(token, attendeeIds, staffId) {
  const officer = verifyOfficer_(staffId);

  const ids = Array.isArray(attendeeIds)
    ? [...new Set(attendeeIds.map(String))]
    : [];

  if (!ids.length) {
    throw new Error("Select at least one attendee to check in.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const reservationId = verifyToken_(token);
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(CFG.SHEETS.ATTENDEES);
    const log = ss.getSheetByName(CFG.SHEETS.LOG);

    if (!sh || !log) {
      throw new Error(
        'Sheet "Attendees" or "Attendance_Log" was not found.'
      );
    }

    const values = sh.getDataRange().getValues();

    if (!values.length) {
      throw new Error("The Attendees sheet is empty.");
    }

    const col = {};
    values[0].forEach((h, i) => {
      col[String(h).trim()] = i;
    });

    const required = [
      "RSVP_ID",
      "Attendee ID",
      "Status",
      "Name",
      "Employee Number"
    ];

    required.forEach(h => {
      if (col[h] === undefined) {
        throw new Error('Missing column in Attendees: "' + h + '".');
      }
    });

    const matched = {};
    const result = [];
    const selectedRows = [];

    for (let i = 1; i < values.length; i++) {
      const r = values[i];
      const id = String(r[col["Attendee ID"]]);

      if (
        String(r[col["RSVP_ID"]]) !== reservationId ||
        !ids.includes(id)
      ) {
        continue;
      }

      matched[id] = true;

      if (String(r[col["Status"]]).toUpperCase() === "PRESENT") {
        result.push({
          attendeeId: id,
          name: String(r[col["Name"]]),
          status: "ALREADY_PRESENT"
        });
      } else {
        selectedRows.push({
          rowNumber: i + 1,
          row: r,
          attendeeId: id
        });
      }
    }

    ids.forEach(id => {
      if (!matched[id]) {
        throw new Error(
          "A selected attendee does not belong to this reservation."
        );
      }
    });

    const now = new Date();

    selectedRows.forEach(x => {
      const statusCol = col["Status"] + 1;

      // Status, Check-in At and Checked-in By are adjacent columns.
      sh.getRange(x.rowNumber, statusCol, 1, 3).setValues([[
        "PRESENT",
        now,
        officer
      ]]);

      const empNo = x.row[col["Employee Number"]];
      const name = x.row[col["Name"]];

      result.push({
        attendeeId: x.attendeeId,
        name: String(name),
        status: "CHECKED_IN"
      });

      log.appendRow([
        now,
        "CHECK_IN",
        reservationId,
        x.attendeeId,
        empNo,
        name,
        officer,
        "Successful check-in"
      ]);
    });

    return {
      reservationId: reservationId,
      results: result,
      attendees: getReservationAttendees_(reservationId)
    };
  } finally {
    lock.releaseLock();
  }
}

function getDashboardSummary(staffId) {
  verifyOfficer_(staffId);

  const sh = getSpreadsheet_().getSheetByName(CFG.SHEETS.ATTENDEES);

  if (!sh) {
    throw new Error('Sheet "Attendees" was not found.');
  }

  const last = sh.getLastRow();
  const rows = last > 1
    ? sh.getRange(2, 1, last - 1, 9).getValues()
    : [];

  return {
    totalEligible: rows.length,
    checkedIn: rows.filter(
      r => String(r[6]).toUpperCase() === "PRESENT"
    ).length,
    employeesPresent: rows.filter(
      r => String(r[4]) === "Employee" &&
           String(r[6]).toUpperCase() === "PRESENT"
    ).length,
    spousesPresent: rows.filter(
      r => String(r[4]) === "Spouse" &&
           String(r[6]).toUpperCase() === "PRESENT"
    ).length,
    pending: rows.filter(
      r => String(r[6]).toUpperCase() !== "PRESENT"
    ).length
  };
}

function sha256_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    value,
    Utilities.Charset.UTF_8
  );

  return bytes.map(b =>
    (b < 0 ? b + 256 : b).toString(16).padStart(2, "0")
  ).join("");
}
