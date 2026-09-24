/* The member app and the admin console both read the same spreadsheet. It is
   opened by ID so the script works whether or not it is bound to the sheet. */
const SHEET_ID = '1XhVmwTimVD1VaPkNVCpriNn_g7wcJXZjPGC7wJ-1vaI';
var SS_CACHE_ = null;
function ss_() {
  return SS_CACHE_ || (SS_CACHE_ = SpreadsheetApp.openById(SHEET_ID));
}

// Global session state holder for the active request execution context
var CURRENT_SESSION_EMAIL = "";

function setSessionEmail(email) {
  if (email) {
    CURRENT_SESSION_EMAIL = String(email).trim().toLowerCase();
  }
}

/* ---- Remembered sign-in ------------------------------------------------
   The client keeps only an opaque token in localStorage; the token -> email
   mapping lives here in the script's property store, so a token can be
   revoked (logout, or the row going INACTIVE) without touching a password.
   Sessions have no expiry — they end when the user logs out.
   ---------------------------------------------------------------------- */

var SESSION_PREFIX = "sess_";

function createSession_(email) {
  const token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "").slice(0, 8);
  PropertiesService.getScriptProperties().setProperty(
    SESSION_PREFIX + token,
    JSON.stringify({ email: String(email).trim().toLowerCase(), created: Date.now() })
  );
  return token;
}

function readSession_(token) {
  if (!token) return null;
  try {
    const raw = PropertiesService.getScriptProperties().getProperty(SESSION_PREFIX + String(token));
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return (parsed && parsed.email) ? parsed : null;
  } catch (err) {
    return null;
  }
}

function deleteSession_(token) {
  if (!token) return;
  try {
    PropertiesService.getScriptProperties().deleteProperty(SESSION_PREFIX + String(token));
  } catch (err) { /* nothing to clean up */ }
}

/* Called by the client on logout. Always reports success — a token that is
   already gone is the outcome the caller wanted. */
function endSession(token) {
  deleteSession_(token);
  return { success: true };
}

/* Rehydrates a session at page load. Re-reads the Users row rather than
   trusting anything cached, so sheet edits (name, role, INACTIVE) take effect
   on the next refresh. Returns the same envelope as verifyLogin so the client
   can share one success path. */
function resumeSession(token) {
  try {
    const session = readSession_(token);
    if (!session) return { success: false };

    const users = readSheetAsMap('Users');
    const currentUser = users.find(u => String(u.email || '').trim().toLowerCase() === session.email);

    if (!currentUser) {
      deleteSession_(token);
      return { success: false };
    }

    if (String(currentUser.status || '').toUpperCase() === 'INACTIVE') {
      deleteSession_(token);
      return { success: false, message: "System update in progress. Access temporarily restricted for database maintenance." };
    }

    CURRENT_SESSION_EMAIL = session.email;

    return {
      success: true,
      user: buildUserObj_(currentUser, session.email),
      data: getData(session.email, currentUser)
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/* Sheets hands back a date-formatted cell as a real Date object, and
   String(dateObj) prints its full JS toString (with GMT offset + timezone
   name) rather than a clean date. Normalize to ISO yyyy-MM-dd — the client
   reformats that for display and it's what <input type="date"> expects. */
function formatBirthday_(val) {
  if (!val) return "";
  if (val instanceof Date && !isNaN(val)) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(val);
}

/* The user shape both verifyLogin and resumeSession hand back to the client. */
function buildUserObj_(currentUser, cleanEmail) {
  const name = String(currentUser.name || "");
  return {
    userId: String(currentUser.userId || ""),
    email: cleanEmail || String(currentUser.email || "").trim().toLowerCase(),
    name: name,
    initials: getInitials(name),
    groupId: String(currentUser.groupId || ""),
    groupName: String(currentUser.groupName || currentUser.groupId || "Small Group"),
    role: String(currentUser.role || "Member"),
    phone: String(currentUser.phone || ""),
    birthday: formatBirthday_(currentUser.birthday),
    church: String(currentUser.church || ""),
    occupation: String(currentUser.occupation || ""),
    status: String(currentUser.status || ""),
    greetingDate: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "EEEE, d MMMM")
  };
}

function getCurrentUserSession() {
  const users = readSheetAsMap('Users');
  let currentUser = null;
  
  if (CURRENT_SESSION_EMAIL) {
    currentUser = users.find(u => String(u.email || '').trim().toLowerCase() === CURRENT_SESSION_EMAIL);
  }
  
  if (!currentUser) {
    const email = Session.getActiveUser().getEmail();
    if (email) {
      currentUser = users.find(u => String(u.email || '').trim().toLowerCase() === String(email).trim().toLowerCase());
    }
  }
  
  if (!currentUser && users.length > 0) {
    currentUser = users[0]; // fallback safely
  }
  
  if (!currentUser) {
    throw new Error("Access denied. No valid user session found.");
  }
  return currentUser;
}

function getInitials(name) {
  if (!name) return "US";
  const parts = name.trim().split(/\s+/);
  if (parts.length === 1) {
    return (parts[0][0] + parts[0][0]).toUpperCase();
  }
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function verifyLogin(email, password) {
  try {
    const cleanEmail = String(email || '').trim().toLowerCase();
    const cleanPassword = String(password || '').trim();

    const users = readSheetAsMap('Users');
    const currentUser = users.find(u => String(u.email || '').trim().toLowerCase() === cleanEmail);
    
    if (!currentUser || String(currentUser.password || "").trim() !== cleanPassword) {
      return { success: false, message: "Invalid email or password." };
    }
    
    if (String(currentUser.status || '').toUpperCase() === 'INACTIVE') {
      return { success: false, message: "System update in progress. Access temporarily restricted for database maintenance." };
    }

    CURRENT_SESSION_EMAIL = cleanEmail;

    return {
      success: true,
      user: buildUserObj_(currentUser, cleanEmail),
      token: createSession_(cleanEmail),
      data: getData(cleanEmail, currentUser)
    };
  } catch (error) {
    throw new Error("Login verification failed: " + error.message);
  }
}

/* Column letters vary by sheet, so writes go by header name (same convention
   readSheetAsMap already uses for reads) rather than hardcoded positions. */
function findUsersColumns_(headerRow) {
  const idx = {};
  headerRow.forEach((h, i) => { idx[String(h).trim().toLowerCase()] = i; });
  return idx;
}

function updateUserProfile(payload) {
  try {
    if (payload.oldEmail) setSessionEmail(payload.oldEmail);
    const ss = ss_();
    const sheet = ss.getSheetByName("Users");
    if (!sheet) return { success: false, message: "Users sheet not found" };

    const data = sheet.getDataRange().getValues();
    const col = findUsersColumns_(data[0]);
    const emailCol = col.email;
    if (emailCol === undefined) return { success: false, message: "Users sheet has no email column" };

    for (let i = 1; i < data.length; i++) {
      const rowEmail = String(data[i][emailCol] || "").trim().toLowerCase();
      if (rowEmail === String(payload.oldEmail).trim().toLowerCase()) {
        const setIfPresent = (key, val) => {
          if (col[key] === undefined || val === undefined || val === null) return;
          sheet.getRange(i + 1, col[key] + 1).setValue(val);
        };
        if (payload.name) setIfPresent('name', payload.name);
        if (payload.email) setIfPresent('email', payload.email);
        if (payload.password && payload.password.trim() !== "") setIfPresent('password', payload.password);
        setIfPresent('phone', payload.phone || "");
        setIfPresent('birthday', payload.birthday || "");
        setIfPresent('church', payload.church || "");
        setIfPresent('occupation', payload.occupation || "");

        // The remembered session points at the old address — repoint it, or the
        // next refresh would look up an email that no longer exists.
        if (payload.token && payload.email &&
            String(payload.email).trim().toLowerCase() !== String(payload.oldEmail).trim().toLowerCase()) {
          const session = readSession_(payload.token);
          if (session) {
            PropertiesService.getScriptProperties().setProperty(
              SESSION_PREFIX + String(payload.token),
              JSON.stringify({ email: String(payload.email).trim().toLowerCase(), created: session.created })
            );
          }
        }
        return { success: true };
      }
    }
    return { success: false, message: "User not found in spreadsheet" };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/* getData(email) returns the whole payload and is what refreshData() on the
   client calls, so its shape must stay complete. Sign-in and session resume
   also call it directly (full, not partial) so Home never renders with a
   notification count that then jumps once a background fetch lands. */
function getData(userEmail, currentUserRow) {
  try {
    if (userEmail) setSessionEmail(userEmail);
    const users = readSheetAsMap('Users');
    let currentUser = null;
    
    if (userEmail) {
      currentUser = users.find(u => String(u.email || '').trim().toLowerCase() === String(userEmail).trim().toLowerCase());
    }
    if (!currentUser) {
      currentUser = users[0] || { userId: 1, name: "Kalaka Rangga", initials: "KR", groupId: "G1", groupName: "The Bekasi", role: "Member" };
    }

    const groups = readSheetAsMap('Groups');
    const suData = readSheetAsMap('ServiceUpdate')[0] || {};
    const verseData = readSheetAsMap('Verse')[0] || {};
    const events = readSheetAsMap('Events');
    const rawResources = readSheetAsMap('Resources');
    const pulseConfig = activePulse_() || {};
    const posts = readSheetAsMap('Posts');
    // readSheetAsMap returns [] for a tab that does not exist yet, so a Reactions
    // sheet that has never been written to reads as "nobody has reacted".
    const reactions = readSheetAsMap(REACTIONS_SHEET);
    const rawAnnouncements = readSheetAsMap('Announcements');
    const pulseResponses = readSheetAsMap('PulseResponses');

    const reactionsByPost = aggregateReactions_(reactions, currentUser.userId);

    const pulseRows = readPulses_();
    const hasSubmittedPulse = readPulseResponses_(pulseRows).some(r =>
      r.pulseId === String(pulseConfig.pulse_id || '').trim() &&
      r.userId === String(currentUser.userId || '').trim()
    );

    const ss = ss_();

    const userMap = {};
    users.forEach(u => {
      if (u.userId) userMap[String(u.userId)] = u;
    });
    
    // Fetch Reflection Replies with dynamic user lookup based on userId
    const repliesSheet = ss.getSheetByName('ReflectionReplies');
    const repliesData = repliesSheet ? repliesSheet.getDataRange().getValues() : [];
    const allReplies = [];
    if (repliesData.length > 1) {
      const headers = repliesData[0].map(h => String(h).trim().toLowerCase());
      const idxReplyId = headers.indexOf('replyid');
      const idxPostId = headers.indexOf('postid') !== -1 ? headers.indexOf('postid') : headers.indexOf('refid');
      const idxUserId = headers.indexOf('userid');
      const idxMessage = headers.indexOf('message') !== -1 ? headers.indexOf('message') : headers.indexOf('text');
      const idxQuotedText = headers.indexOf('quotedtext');
      const idxQuotedAuthor = headers.indexOf('quotedauthor');
      const idxTimestamp = headers.indexOf('timestamp') !== -1 ? headers.indexOf('timestamp') : headers.indexOf('time');

      for (let i = 1; i < repliesData.length; i++) {
        const row = repliesData[i];
        const refId = String(idxPostId !== -1 ? row[idxPostId] : row[1]);
        const uId = String(idxUserId !== -1 ? row[idxUserId] : row[2]);
        const text = String(idxMessage !== -1 ? row[idxMessage] : row[3]);
        const quotedText = String(idxQuotedText !== -1 ? row[idxQuotedText] : row[4] || "");
        const quotedAuthor = String(idxQuotedAuthor !== -1 ? row[idxQuotedAuthor] : row[5] || "");
        const time = String(idxTimestamp !== -1 ? row[idxTimestamp] : row[6] || new Date().toISOString());

        const authorUser = userMap[uId] || users.find(u => String(u.userId) === uId) || {};
        const authorName = String(authorUser.name || "Unknown");

        // A stable id per reply is what the notification read-state keys on. Sheets made
        // by saveReflectionReply carry ReplyID; older hand-made tabs may not, so fall
        // back to the row position.
        const replyId = String(idxReplyId !== -1 ? (row[idxReplyId] || "") : "") || ("r_row" + i);

        allReplies.push({
          id: replyId,
          userId: uId,
          refId: refId,
          author: authorName,
          initials: getInitials(authorName),
          text: text,
          time: time,
          quotedText: quotedText,
          quotedAuthor: quotedAuthor
        });
      }
    }

    // Fetch Reflections
    const refSheet = ss.getSheetByName('Reflections');
    const refData = refSheet ? refSheet.getDataRange().getValues() : [];
    const reflections = [];
    
    if (refData.length > 1) {
      const headers = refData[0].map(h => String(h).trim().toLowerCase());
      const idxUserId = headers.indexOf('userid');
      const idxTitle = headers.indexOf('title');
      const idxText = headers.indexOf('text');
      const idxTime = headers.indexOf('time');
      const idxVis = headers.indexOf('visibility');
      const idxFile = headers.indexOf('filename');
      const idxUrl = headers.indexOf('fileurl');
      const idxStatus = headers.indexOf('status');

      for (let i = 1; i < refData.length; i++) {
        const row = refData[i];
        const titleVal = idxTitle !== -1 ? row[idxTitle] : row[1];
        if (!titleVal) continue;

        // Deleted reflections keep their row so that "ref_" + i stays stable for every
        // row below them — physically removing one would re-point stored reply PostIDs.
        if (idxStatus !== -1 && String(row[idxStatus]).trim().toLowerCase() === 'deleted') continue;

        const uId = idxUserId !== -1 ? row[idxUserId] : row[0];
        const tit = titleVal;
        const txt = idxText !== -1 ? row[idxText] : row[2];
        const tim = idxTime !== -1 ? row[idxTime] : row[3];
        const vis = idxVis !== -1 ? row[idxVis] : row[4];
        const fName = idxFile !== -1 ? row[idxFile] : row[5];
        const fUrl = idxUrl !== -1 ? row[idxUrl] : row[6];

        const authorUser = userMap[String(uId)] || users.find(u => String(u.userId) === String(uId)) || {};
        const rawText = String(txt || "");
        const truncatedText = rawText.length > 250 ? rawText.substring(0, 250) + "..." : rawText;
        const authorName = String(authorUser.name || "Unknown");

        const refId = "ref_" + i;
        const refReplies = allReplies.filter(rp => rp.refId === refId);

        reflections.push({
          id: refId,
          ownerId: String(uId),
          author: authorName,
          session: "Romans 8",
          title: String(tit || ""),
          text: String(truncatedText),
          fullText: String(rawText),
          time: String(tim || new Date().toISOString()),
          initials: getInitials(authorName),
          visibility: String(vis || "Public"),
          mine: String(uId) === String(currentUser.userId),
          fileName: String(fName || ""),
          fileUrl: String(fUrl || ""),
          replies: refReplies
        });
      }
    }

    // Built here, while `reflections` is still in feed order — the return statement below
    // reverses that array in place.
    const notifData = buildNotifications(currentUser, reflections, posts, reactions, userMap);

    const isMentor = String(currentUser.role || "").trim().toLowerCase() === "mentor";
    const currentGroup = groups.find(g => String(g.groupId) === String(currentUser.groupId) || String(g.name) === String(currentUser.groupName)) || {};
    
    let groupMembers = [];
    let allGroupsData = [];

    if (isMentor) {
      const groupMap = {};
      users.forEach(u => {
        const gId = String(u.groupId || "G1").trim();
        if (!groupMap[gId]) {
          const matchGrp = groups.find(g => String(g.groupId) === gId) || { name: "Group " + gId };
          groupMap[gId] = {
            groupId: gId,
            name: matchGrp.name || gId,
            members: []
          };
        }
        const mName = String(u.name || "Member");
        const memberObj = {
          userId: String(u.userId || ""),
          name: mName,
          role: String(u.role || "Member"),
          phone: String(u.phone || "No phone"),
          initials: getInitials(mName),
          present: true,
          lastSeen: String(u.lastSeen || "")
        };
        groupMap[gId].members.push(memberObj);
        groupMembers.push(memberObj); // Populates groupMembers so mentors can use mentions too!
      });
      allGroupsData = Object.values(groupMap);
    } else {
      const userGroupId = String(currentUser.groupId || "").trim();
      const userGroupName = String(currentUser.groupName || "").trim();

      groupMembers = users.filter(u => {
        const mGroupId = String(u.groupId || "").trim();
        const mGroupName = String(u.groupName || "").trim();
        return (userGroupId && mGroupId === userGroupId) || (userGroupName && mGroupName === userGroupName);
      }).map(m => {
        const mName = String(m.name || "Member");
        return {
          userId: String(m.userId || ""),
          name: mName,
          role: String(m.role || "Member"),
          phone: String(m.phone || "No phone"),
          initials: getInitials(mName),
          present: true,
          lastSeen: String(m.lastSeen || "")
        };
      });
    }

    const resources = rawResources.map(r => ({
      title: String(r.title || r.Title || r['Resource Title'] || ''),
      type_size: String(r.type_size || r.type || r['type/size'] || r['Type/Size'] || ''),
      fileUrl: String(r.fileUrl || r.fileurl || r['file url'] || ''),
      isFeatured: String(r.isFeatured || r.isfeatured || r['is featured'] || '').toUpperCase() === 'TRUE'
    }));

    const groupEvents = events.filter(e => {
      const gId = e['Group ID'] || e.GroupID || e.groupId;
      // "ALL" rows come from the admin console's "All groups" / "Mentors only".
      if (String(gId || "").trim().toUpperCase() === "ALL") {
        return !isTruthyCell_(e.is_mentor_only) || isMentor;
      }
      if (gId !== undefined && String(gId).trim() !== "") {
        return String(gId).trim() === String(currentUser.groupId).trim();
      }
      return true; 
    });

    const attendanceByEvent = readAttendanceSummaries_(String(currentUser.groupId || ""),
                                                       String(currentUser.userId || ""));

    const now = new Date();
    const parsedEvents = groupEvents.map(e => {
      let d = e.date instanceof Date ? e.date : new Date(String(e.date || "") + " " + new Date().getFullYear());
      const slot = buildCalendarSlot(isNaN(d.getTime()) ? null : d, e.time);
      const eventId = eventKey_(e, currentUser.groupId, slot.startMs);
      return {
        eventId: eventId,
        attendance: attendanceByEvent[eventId] || null,
        title: String(e.title || ""),
        dateObj: isNaN(d.getTime()) ? new Date() : d,
        dateStr: !isNaN(d.getTime()) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "EEE, d MMM") : String(e.date || ""),
        monthStr: !isNaN(d.getTime()) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM").toUpperCase() : "AUG",
        dayNum: !isNaN(d.getTime()) ? Utilities.formatDate(d, Session.getScriptTimeZone(), "d") : "1",
        time: String(e.time || ""),
        location: String(e.location || ""),
        description: String(e.description || ""),
        badge: !isNaN(d.getTime()) ? getNextEventLabel(d) : "UPCOMING",
        calStart: slot.start,
        calEnd: slot.end,
        calAllDay: slot.allDay,
        calTz: slot.tz,
        startMs: slot.startMs
      };
    }).sort((a, b) => a.dateObj - b.dateObj);

    const futureEvents = parsedEvents.filter(e => e.dateObj >= new Date(now.getFullYear(), now.getMonth(), now.getDate()));
    const nextRaw = futureEvents.length > 0 ? futureEvents[0] : (parsedEvents.length > 0 ? parsedEvents[0] : null);
    const upcomingList = futureEvents.length > 0 ? futureEvents.slice(1) : parsedEvents;

    let nextEvent = null;
    if (nextRaw) {
      nextEvent = {
        eventId: String(nextRaw.eventId || ""),
        attendance: nextRaw.attendance || null,
        title: String(nextRaw.title),
        date: String(nextRaw.dateStr),
        time: String(nextRaw.time),
        location: String(nextRaw.location),
        summary: String(nextRaw.description),
        badge: String(nextRaw.badge),
        month: String(nextRaw.monthStr),
        day: String(nextRaw.dayNum),
        calStart: String(nextRaw.calStart),
        calEnd: String(nextRaw.calEnd),
        calAllDay: nextRaw.calAllDay === true,
        calTz: String(nextRaw.calTz),
        startMs: Number(nextRaw.startMs || 0)
      };
    }

    const questionsList = pulseConfig.questions ? String(pulseConfig.questions).split('|').filter(q => q.trim() !== "") : [];
    const openQuestionsList = pulseConfig.open_questions ? String(pulseConfig.open_questions).split('|').filter(q => q.trim() !== "") : [];
    
    const qCount = questionsList.length;
    const oqCount = openQuestionsList.length;
    const totalQuestions = qCount + oqCount;
    const calculatedMinutes = Math.max(1, Math.ceil((totalQuestions * 30) / 60));

    let closesDateObj = pulseConfig['closes date'] instanceof Date ? pulseConfig['closes date'] : new Date(String(pulseConfig['closes date'] || ''));
    let diffDays = 3;
    let closesText = "Closes soon";
    let isUrgent = false;

    if (!isNaN(closesDateObj.getTime())) {
      diffDays = Math.ceil((closesDateObj - now) / (1000 * 60 * 60 * 24));
      isUrgent = diffDays <= 2;
      closesText = `Closes ${Utilities.formatDate(closesDateObj, Session.getScriptTimeZone(), "EEEE, d MMM · h:mm a")}`;
    }

    let daysLeftText = `${diffDays} days left to respond — take your time`;
    if (diffDays === 0) daysLeftText = `Closes today — take your time`;
    if (diffDays < 0) daysLeftText = `Pulse check closed`;

    const currentUserName = String(currentUser.name || "");

    return {
      user: {
        userId: String(currentUser.userId || ""),
        name: currentUserName,
        initials: getInitials(currentUserName),
        email: String(currentUser.email || ""),
        phone: String(currentUser.phone || ""),
        birthday: formatBirthday_(currentUser.birthday),
        church: String(currentUser.church || ""),
        occupation: String(currentUser.occupation || ""),
        role: String(currentUser.role || "Member"),
        // A mentor's `members`/`group.memberCount` below cover every group, so the
        // client needs this to pick their own group back out of allGroups.
        groupId: String(currentUser.groupId || ""),
        greetingDate: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "EEEE, d MMMM")
      },
      serviceUpdate: {
        enabled: suData.enabled === true || String(suData.enabled).toUpperCase() === "TRUE",
        theme: String(suData.theme || ""),
        module: String(suData.module || ""),
        note: String(suData.note || "")
      },
      verse: {
        label: String(verseData.label || "Verse of the month"),
        text: String(verseData.text || ""),
        ref: String(verseData.reference || "")
      },
      announcements: rawAnnouncements.map((a, i) => ({
        id: "ann_" + i,
        title: String(a.title || a.Title || ""),
        tag: String(a.tag || a.Tag || ""),
        time: String(a.time || a.Time || ""),
        imgUrl: String(a.imgUrl || a.imageUrl || a.image || a.imageLabel || a['image url'] || a['Image URL'] || ""),
        detail: String(a.detail || a.Detail || a.description || a.Description || a.body || a.Body || a['full text'] || a['Full Text'] || "")
      })),
      nextEvent: nextEvent,
      events: upcomingList.map(e => ({
        eventId: String(e.eventId || ""),
        attendance: e.attendance || null,
        title: String(e.title),
        date: String(e.dateStr),
        time: String(e.time),
        location: String(e.location),
        summary: String(e.description),
        month: String(e.monthStr),
        day: String(e.dayNum),
        calStart: String(e.calStart),
        calEnd: String(e.calEnd),
        calAllDay: e.calAllDay === true,
        calTz: String(e.calTz),
        startMs: Number(e.startMs || 0)
      })),
      group: {
        name: String(currentUser.groupName || currentGroup.name || "Small Group"),
        memberCount: groupMembers.length,
        meets: String(currentGroup.meetsText || ""),
        mentor: { 
          name: String(currentGroup.mentorName || ""), 
          title: String(currentGroup.mentorTitle || "Mentor"), 
          phone: String(currentGroup.mentorPhone || ""), 
          initials: getInitials(String(currentGroup.mentorName || "")) 
        }
      },
      members: groupMembers,
      allGroups: allGroupsData,
      resources: {
        featured: resources.find(r => r.isFeatured === true) ? {
          title: resources.find(r => r.isFeatured === true).title,
          meta: resources.find(r => r.isFeatured === true).type_size,
          fileUrl: resources.find(r => r.isFeatured === true).fileUrl
        } : null,
        list: resources.filter(r => r.isFeatured !== true).map(r => ({
          title: String(r.title), 
          type: String(r.type_size), 
          fileUrl: String(r.fileUrl)
        }))
      },
      pulse: {
        windowTitle: String(pulseConfig.windowTitle || "Pulse Check"),
        closes: closesText,
        daysLeftNote: daysLeftText,
        isUrgent: isUrgent,
        submitted: hasSubmittedPulse,
        questionCount: qCount,
        openQuestionCount: oqCount,
        estimatedMinutes: calculatedMinutes,
        questions: questionsList.map(q => String(q)),
        openQuestions: openQuestionsList.map(q => String(q))
      },
      /* readSheetAsMap trims headers but does not lower-case them, so accept
         either spelling of the visibility column. */
      posts: posts.filter(p => {
        const authorUser = userMap[p.userId] || {};
        const authorGroup = groups.find(g => String(g.groupId) === String(authorUser.groupId)) || {};
        return canSeePost_(p.visibility || p.Visibility, authorUser, authorGroup, currentUser);
      }).map(p => {
        const authorUser = userMap[p.userId] || {};
        const authorGroup = groups.find(g => String(g.groupId) === String(authorUser.groupId)) || {};
        const authorName = String(authorUser.name || "Unknown");
        const key = postKey_(p);
        const agg = reactionsByPost[key] || { counts: { amen: 0, love: 0, praise: 0, hope: 0 }, mine: [] };
        return {
          id: key,
          author: authorName,
          group: String(authorGroup.name || "Small Group"),
          type: String(p.type || "Praise"),
          text: String(p.text || ""),
          time: String(p.time || new Date().toISOString()),
          initials: getInitials(authorName),
          visibility: String(p.visibility || p.Visibility || "Public"),
          mine: String(p.userId) === String(currentUser.userId),
          reactions: agg.counts,
          myReactions: agg.mine
        };
      }).reverse(),
      reflections: reflections.reverse(),
      notifications: notifData.list,
      notifUnread: notifData.unreadCount,
      /* The server's clock at the moment this payload was built. The client keeps
         the difference against its own clock (noteServerClock in Core) so anything
         timing-gated — the mentor attendance card in Events — is judged against
         script time rather than whatever the device happens to be set to. */
      serverNow: new Date().getTime()
    };
  } catch (err) {
    throw new Error("getData failed: " + err.message);
  }
}

/* ==========================================================================
   NOTIFICATIONS
   Derived at read time from Reflections + ReflectionReplies rather than stored
   as their own rows, so the feed works over pre-existing data and can never
   drift out of sync with the threads it points at. Only read/unread state is
   persisted, in the NotificationReads sheet.
   ========================================================================== */

const NOTIF_READS_SHEET = 'NotificationReads';
const NOTIF_LIMIT = 50;
const NOTIF_READ_IDS_LIMIT = 200;

function ensureNotificationReadsSheet() {
  const ss = ss_();
  let sheet = ss.getSheetByName(NOTIF_READS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(NOTIF_READS_SHEET);
    sheet.appendRow(['userId', 'lastReadAt', 'readIds']);
  }
  return sheet;
}

function getNotificationReadState(userId) {
  const sheet = ss_().getSheetByName(NOTIF_READS_SHEET);
  const empty = { lastReadAt: "", readIds: [] };
  if (!sheet) return empty;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim() === String(userId).trim()) {
      const rawDate = data[i][1];
      return {
        lastReadAt: rawDate instanceof Date ? rawDate.toISOString() : String(rawDate || ""),
        readIds: String(data[i][2] || "").split(',').map(s => s.trim()).filter(s => s !== "")
      };
    }
  }
  return empty;
}

function writeNotificationReadState(userId, lastReadAt, readIds) {
  const sheet = ensureNotificationReadsSheet();
  const trimmed = readIds.slice(-NOTIF_READ_IDS_LIMIT);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim() === String(userId).trim()) {
      sheet.getRange(i + 1, 2).setValue(lastReadAt);
      sheet.getRange(i + 1, 3).setValue(trimmed.join(','));
      return;
    }
  }
  sheet.appendRow([userId, lastReadAt, trimmed.join(',')]);
}

function isNotificationUnread(notif, readState) {
  if (readState.readIds.indexOf(String(notif.id)) !== -1) return false;
  if (!readState.lastReadAt) return true;

  const t = new Date(notif.time);
  const cutoff = new Date(readState.lastReadAt);
  if (isNaN(t.getTime()) || isNaN(cutoff.getTime())) return true;
  return t > cutoff;
}

function buildNotifications(currentUser, reflections, posts, reactions, userMap) {
  const myId = String(currentUser.userId || "").trim();
  const myName = String(currentUser.name || "").trim();
  const mentionToken = ("@" + myName).toLowerCase();
  const readState = getNotificationReadState(myId);
  const list = [];

  reflections.forEach(r => {
    const isMine = String(r.ownerId) === myId;

    // Never surface a thread the reader cannot open (mirrors the client-side
    // visibility filter in Reflection.html).
    if (String(r.visibility || "Public").toLowerCase() === "private" && !isMine) return;

    (r.replies || []).forEach(rp => {
      if (String(rp.userId || "").trim() === myId) return; // your own reply

      const mentionedMe = myName !== "" && String(rp.text || "").toLowerCase().indexOf(mentionToken) !== -1;
      const quotedMe = myName !== "" &&
        String(rp.quotedAuthor || "").trim().toLowerCase() === myName.toLowerCase();

      let type = "", actionText = "";
      if (mentionedMe) {
        type = "mention";
        actionText = "mentioned you in a reply";
      } else if (quotedMe) {
        type = "reply";
        actionText = "replied to your comment";
      } else if (isMine) {
        type = "reply";
        actionText = "replied to your reflection";
      } else {
        return;
      }

      const entry = {
        id: String(rp.id),
        type: type,
        actor: String(rp.author || "Someone"),
        initials: getInitials(String(rp.author || "")),
        actionText: actionText,
        context: String(r.title || "Reflection"),
        refId: String(r.id),
        time: String(rp.time || "")
      };
      entry.read = !isNotificationUnread(entry, readState);
      list.push(entry);
    });
  });

  /* Reactions on my own posts. Keyed on (post, actor, type) rather than on the
     write time — a member can hold several different reactions on the same
     post at once, so each type they pick is its own notification rather than
     one row getting rewritten. */
  const myPosts = {};
  (posts || []).forEach(p => {
    if (String(p.userId || "").trim() !== myId) return;
    myPosts[postKey_(p)] = p;
  });

  (reactions || []).forEach(rx => {
    const type = String(rx.type || "").trim().toLowerCase();
    if (REACTION_TYPES.indexOf(type) === -1) return;

    const key = String(rx.postId || rx.postid || "").trim();
    const post = myPosts[key];
    if (!post) return;

    const actorId = String(rx.userId || rx.userid || "").trim();
    if (!actorId || actorId === myId) return; // your own reaction

    const actor = String(((userMap || {})[actorId] || {}).name || "Someone");
    const isPraise = String(post.type || "").trim().toLowerCase() === "praise";

    const entry = {
      id: "rx" + key + "_" + actorId.replace(/[^A-Za-z0-9_]/g, "") + "_" + type,
      type: "reaction",
      actor: actor,
      initials: getInitials(actor),
      actionText: 'reacted "' + (REACTION_LABELS[type] || type) + '" to your ' + (isPraise ? "praise" : "prayer"),
      context: String(post.text || "").slice(0, 80),
      refId: "",
      postId: key,
      time: String(rx.timestamp || rx.Timestamp || "")
    };
    entry.read = !isNotificationUnread(entry, readState);
    list.push(entry);
  });

  list.sort((a, b) => {
    const ta = new Date(a.time), tb = new Date(b.time);
    return (isNaN(tb.getTime()) ? 0 : tb) - (isNaN(ta.getTime()) ? 0 : ta);
  });

  const capped = list.slice(0, NOTIF_LIMIT);
  return {
    list: capped,
    unreadCount: capped.filter(n => !n.read).length
  };
}

function markNotificationRead(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();
  if (!payload || !payload.notifId) return false;

  const state = getNotificationReadState(user.userId);
  if (state.readIds.indexOf(String(payload.notifId)) === -1) {
    state.readIds.push(String(payload.notifId));
    writeNotificationReadState(user.userId, state.lastReadAt, state.readIds);
  }
  return true;
}

function markAllNotificationsRead(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();

  // lastReadAt alone covers everything older than now, so the id list can be
  // dropped — that keeps the cell from growing without bound.
  writeNotificationReadState(user.userId, new Date().toISOString(), []);
  return true;
}

/* ==========================================================================
   REACTIONS
   A member can hold several different reactions on the same post at once
   (Amen *and* Love, say) — tapping one again removes just that one. Stored as
   its own sheet, one row per (postId, userId, type) triple that is currently
   active: toggling on appends a row, toggling off deletes it.
   ========================================================================== */

const REACTIONS_SHEET = 'Reactions';
const REACTION_TYPES = ['amen', 'love', 'praise', 'hope'];

/* Shown in the notification line; the client owns the on-screen labels and must
   agree with these. "praise" reads as Agree — the key is what is already stored
   in the sheet, so it stays put while the label changes. */
const REACTION_LABELS = { amen: "Amen", love: "Love", praise: "Agree", hope: "Hope" };

function ensureReactionsSheet() {
  const ss = ss_();
  let sheet = ss.getSheetByName(REACTIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(REACTIONS_SHEET);
    sheet.appendRow(['postId', 'userId', 'type', 'timestamp']);
  }
  return sheet;
}

/* The address a reaction is filed under. Posts rows written before the postId
   column was added carry a blank cell, so fall back to a hash of the row's own
   content — content-derived rather than positional, so it survives inserts and
   deletes above the row. readSheetAsMap trims header names but does not
   lower-case them, hence the spellings. */
function postKey_(p) {
  const explicit = String(p.postId || p.postid || p.PostID || "").trim();
  if (explicit) return explicit.replace(/[^A-Za-z0-9_]/g, "");

  const seed = [p.userId, p.time, String(p.text || "").slice(0, 60)].join("|");
  let h = 5381;
  for (let i = 0; i < seed.length; i++) h = ((h * 33) ^ seed.charCodeAt(i)) >>> 0;
  return "p" + h.toString(36);
}

/* Counts per post plus every type this viewer has picked, from the raw
   Reactions rows. Returns { postId: { counts: {amen,love,praise,hope}, mine: [] } }. */
function aggregateReactions_(reactions, viewerId) {
  const byPost = {};
  const me = String(viewerId || "").trim();

  (reactions || []).forEach(r => {
    const type = String(r.type || "").trim().toLowerCase();
    if (REACTION_TYPES.indexOf(type) === -1) return;

    const key = String(r.postId || r.postid || "").trim();
    if (!key) return;

    if (!byPost[key]) byPost[key] = { counts: { amen: 0, love: 0, praise: 0, hope: 0 }, mine: [] };
    byPost[key].counts[type]++;
    if (String(r.userId || r.userid || "").trim() === me) byPost[key].mine.push(type);
  });

  return byPost;
}

/* payload: { userEmail, postId, type } — toggles that one reaction: added if
   the member had not picked it yet on this post, removed if they had. Only
   ever touches the single (post, member, type) row, so any other reactions
   the member already left on the post are untouched. */
function saveReaction(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();

  const postId = String((payload && payload.postId) || "").trim();
  if (!postId) throw new Error("Missing postId.");

  const type = String((payload && payload.type) || "").trim().toLowerCase();
  if (REACTION_TYPES.indexOf(type) === -1) {
    throw new Error("Unknown reaction: " + type);
  }

  const sheet = ensureReactionsSheet();
  const data = sheet.getDataRange().getValues();
  const now = new Date().toISOString();
  const myId = String(user.userId).trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim() === postId &&
        String(data[i][1] || "").trim() === myId &&
        String(data[i][2] || "").trim().toLowerCase() === type) {
      sheet.deleteRow(i + 1);
      return false; // now off
    }
  }

  sheet.appendRow([postId, user.userId, type, now]);
  return true; // now on
}

function saveReflectionReply(payload) {
  if (payload && payload.userEmail) {
    setSessionEmail(payload.userEmail);
  }
  const user = getCurrentUserSession();
  const ss = ss_();
  let sheet = ss.getSheetByName('ReflectionReplies');
  
  if (!sheet) {
    sheet = ss.insertSheet('ReflectionReplies');
    sheet.appendRow(['ReplyID', 'PostID', 'UserID', 'Message', 'QuotedText', 'QuotedAuthor', 'Timestamp']);
  }
  
  const replyId = "r_" + Date.now();
  sheet.appendRow([
    replyId,
    payload.refId,
    user.userId, // Storing the authenticated session's exact userId
    payload.text,
    payload.quotedText || "",
    payload.quotedAuthor || "",
    payload.timestamp || new Date().toISOString()
  ]);

  // The client needs this id to address the reply for edit/delete before the next refresh.
  return replyId;
}

/* ---- Reflection & reply editing ----------------------------------------
   Ownership is always checked against the UserID stored in the sheet, never
   against anything the client sends. */

function assertOwner_(rowUserId, sessionUserId) {
  if (String(rowUserId).trim() !== String(sessionUserId).trim()) {
    throw new Error("Not authorized: you can only change your own posts.");
  }
}

/* Who may read a Posts row, keyed on its `visibility` cell:
     "mentor" — the author, plus the mentor of the author's group
     "group"  — the author, plus that group's members and its mentor
     anything else, including blank — everyone
   A mentor of some *other* group is not privileged here. Applied server-side in
   getData so a hidden post is never sent to the browser at all. */
function canSeePost_(visibility, authorUser, authorGroup, viewer) {
  const vis = String(visibility || "").trim().toLowerCase();
  if (vis !== "mentor" && vis !== "group") return true;

  if (String(authorUser.userId || "").trim() === String(viewer.userId || "").trim()) return true;

  // Same id-then-name matching as the member roster below; ids first, because
  // group display names are not unique and get renamed.
  const aId = String(authorUser.groupId || "").trim();
  const vId = String(viewer.groupId || "").trim();
  const aName = String(authorUser.groupName || "").trim().toLowerCase();
  const vName = String(viewer.groupName || "").trim().toLowerCase();
  const sameGroup = (aId && aId === vId) || (aName && aName === vName);

  // A mentor's Users row may carry no groupId, so also accept the Groups sheet's mentorName.
  const viewerIsMentor = String(viewer.role || "").trim().toLowerCase() === "mentor";
  const namedMentor = String(authorGroup.mentorName || "").trim().toLowerCase() ===
                      String(viewer.name || "").trim().toLowerCase();
  const isGroupMentor = viewerIsMentor && (sameGroup || namedMentor);

  if (vis === "mentor") return isGroupMentor;
  return sameGroup || isGroupMentor;
}

/* refId is "ref_" + the row's index in the values array, so the 1-based sheet
   row is that index + 1. */
function getReflectionRowInfo_(refId) {
  const rowIndex = parseInt(String(refId || "").replace("ref_", ""), 10);
  if (!rowIndex || rowIndex < 1) throw new Error("Invalid reflection id: " + refId);

  const sheet = ss_().getSheetByName('Reflections');
  if (!sheet) throw new Error("Spreadsheet error: Sheet tab named 'Reflections' not found.");

  const rowNum = rowIndex + 1;
  if (rowNum > sheet.getLastRow()) throw new Error("Reflection not found: " + refId);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h).trim().toLowerCase());
  const values = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];

  const idx = {
    userId: headers.indexOf('userid') !== -1 ? headers.indexOf('userid') : 0,
    title: headers.indexOf('title') !== -1 ? headers.indexOf('title') : 1,
    text: headers.indexOf('text') !== -1 ? headers.indexOf('text') : 2,
    visibility: headers.indexOf('visibility') !== -1 ? headers.indexOf('visibility') : 4,
    status: headers.indexOf('status')
  };

  return { sheet: sheet, rowNum: rowNum, headers: headers, idx: idx, values: values };
}

/* Replies are addressed by their ReplyID column. Sheets without that column fall
   back to positional "r_row" ids in getData, and deleting rows there would shift
   every id below — so those sheets are refused rather than corrupted. */
function getReplyRowInfo_(replyId) {
  const sheet = ss_().getSheetByName('ReflectionReplies');
  if (!sheet) throw new Error("Spreadsheet error: Sheet tab named 'ReflectionReplies' not found.");

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error("Reply not found: " + replyId);

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const idxReplyId = headers.indexOf('replyid');
  if (idxReplyId === -1) {
    throw new Error("This ReflectionReplies sheet has no ReplyID column, so replies cannot be edited or deleted safely.");
  }
  const idxUserId = headers.indexOf('userid') !== -1 ? headers.indexOf('userid') : 2;
  const idxMessage = headers.indexOf('message') !== -1 ? headers.indexOf('message') : headers.indexOf('text');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idxReplyId]).trim() === String(replyId).trim()) {
      return {
        sheet: sheet,
        rowNum: i + 1,
        idxUserId: idxUserId,
        idxMessage: idxMessage !== -1 ? idxMessage : 3,
        values: data[i]
      };
    }
  }
  throw new Error("Reply not found: " + replyId);
}

function updateReflection(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();
  const info = getReflectionRowInfo_(payload.refId);

  assertOwner_(info.values[info.idx.userId], user.userId);

  info.sheet.getRange(info.rowNum, info.idx.title + 1).setValue(payload.title);
  info.sheet.getRange(info.rowNum, info.idx.text + 1).setValue(payload.text);
  info.sheet.getRange(info.rowNum, info.idx.visibility + 1).setValue(payload.visibility || "Public");

  return true;
}

function deleteReflection(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();
  const info = getReflectionRowInfo_(payload.refId);

  assertOwner_(info.values[info.idx.userId], user.userId);

  if (info.idx.status === -1) {
    throw new Error("The Reflections sheet needs a 'Status' column header before reflections can be deleted.");
  }
  info.sheet.getRange(info.rowNum, info.idx.status + 1).setValue("Deleted");

  // Replies to a deleted reflection go with it. Bottom-up so each deleteRow does
  // not shift rows still to be visited.
  const repliesSheet = ss_().getSheetByName('ReflectionReplies');
  if (repliesSheet && repliesSheet.getLastRow() > 1) {
    const data = repliesSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const idxPostId = headers.indexOf('postid') !== -1 ? headers.indexOf('postid') : headers.indexOf('refid');
    if (idxPostId !== -1) {
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][idxPostId]).trim() === String(payload.refId).trim()) {
          repliesSheet.deleteRow(i + 1);
        }
      }
    }
  }

  return true;
}

function updateReply(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();
  const info = getReplyRowInfo_(payload.replyId);

  assertOwner_(info.values[info.idxUserId], user.userId);
  info.sheet.getRange(info.rowNum, info.idxMessage + 1).setValue(payload.text);

  return true;
}

function deleteReply(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();
  const info = getReplyRowInfo_(payload.replyId);

  assertOwner_(info.values[info.idxUserId], user.userId);
  info.sheet.deleteRow(info.rowNum);

  return true;
}

function saveReflectionWithFile(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();
  const ss = ss_();
  let fileUrl = "", fileName = "";
  
  if (payload.fileData && payload.fileName) {
    try {
      const folderId = "10Z7VWwJt6KuCL15CZ68ETsgex4cnls5R";
      const folder = DriveApp.getFolderById(folderId);
      const splitData = payload.fileData.split(",");
      const contentType = splitData[0].match(/:(.*?);/)[1];
      const decodedBytes = Utilities.base64Decode(splitData[1]);
      const blob = Utilities.newBlob(decodedBytes, contentType, payload.fileName);
      
      const file = folder.createFile(blob);
      fileUrl = file.getUrl();
      fileName = file.getName();
    } catch (err) {
      console.error("File upload error: " + err.toString());
      throw new Error("Failed to upload file to Drive: " + err.toString());
    }
  }
  
  const sheet = ss.getSheetByName('Reflections');
  if (!sheet) {
    throw new Error("Spreadsheet error: Sheet tab named 'Reflections' not found.");
  }
  
  sheet.appendRow([
    user.userId, 
    payload.title, 
    payload.text, 
    new Date(), 
    payload.visibility || "Public", 
    fileName, 
    fileUrl
  ]);
  
  return { fileName: fileName, fileUrl: fileUrl, newId: "ref_" + sheet.getLastRow() };
}

function savePulse(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();
  const pulseConfig = activePulse_() || {};
  const questionsList = pulseConfig.questions ? String(pulseConfig.questions).split('|').filter(q => q.trim() !== "") : [];
  const openQuestionsList = pulseConfig.open_questions ? String(pulseConfig.open_questions).split('|').filter(q => q.trim() !== "") : [];
  
  const qLen = questionsList.length;
  const oqLen = openQuestionsList.length;
  
  let parts = [];
  for (let i = 0; i < qLen; i++) {
    parts.push(payload.answers[i] !== undefined ? payload.answers[i] : "");
  }
  for (let i = 0; i < oqLen; i++) {
    parts.push(`"${payload.answers['oq_' + i] || ""}"`);
  }
  const straightforwardString = "{" + parts.join(";") + "}";

  // Written by header so the pulse id lands in its own column; the admin
  // console groups responses by it.
  const sheet = sheetWithHeaders_('PulseResponses', ['userId', 'pulse id', 'timestamp', 'answers', 'note']);
  appendByHeader_(sheet, {
    'userId': user.userId,
    'pulse id': String(pulseConfig.pulse_id || ''),
    'timestamp': new Date().toISOString(),
    'answers': straightforwardString,
    'note': payload.note || ""
  });
}

/* Written by header name rather than position: the Posts sheet is edited by
   hand, so a visibility column may sit anywhere. Falls back to the historic
   postId | userId | type | text | time | visibility order if the header row is
   blank. */
function savePost(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();
  const sheet = ss_().getSheetByName('Posts');

  // Whitelisted here, never taken verbatim — the client cannot invent a level.
  const raw = String(payload.visibility || "").trim().toLowerCase();
  const vis = raw === "mentor" ? "Mentor" : (raw === "group" ? "Group" : "Public");

  // Random suffix as well as the clock: two members can post in the same
  // millisecond. Alphanumeric because the client interpolates this id into an
  // onclick attribute and esc() does not escape quotes.
  const postId = "p" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

  const values = {
    postid: postId,
    userid: user.userId,
    type: payload.type,
    text: payload.text,
    time: payload.timestamp || new Date().toISOString(),
    visibility: vis
  };
  const order = ['postid', 'userid', 'type', 'text', 'time', 'visibility'];
  const width = Math.max(sheet.getLastColumn(), order.length);
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0]
    .map(h => String(h).trim().toLowerCase());

  const named = headers.some(h => values.hasOwnProperty(h));
  sheet.appendRow(named ? headers.map(h => values.hasOwnProperty(h) ? values[h] : "")
                        : order.map(k => values[k]));

  // The client needs this to react to its own post before the next refresh.
  return postId;
}

/* Locates a Posts row by its postId. Older rows written before the postId
   column existed carry a blank cell there, so postKey_ — the same fallback
   getData already uses to hand those rows an id — is reused here to find
   them again. Header-name lookup, same reasoning as savePost. */
function getPostRowInfo_(postId) {
  const sheet = ss_().getSheetByName('Posts');
  if (!sheet) throw new Error("Spreadsheet error: Sheet tab named 'Posts' not found.");

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error("Post not found: " + postId);

  const headers = data[0].map(h => String(h).trim());
  const lower = headers.map(h => h.toLowerCase());
  const idx = {
    userId: lower.indexOf('userid') !== -1 ? lower.indexOf('userid') : 1,
    type: lower.indexOf('type') !== -1 ? lower.indexOf('type') : 2,
    text: lower.indexOf('text') !== -1 ? lower.indexOf('text') : 3,
    visibility: lower.indexOf('visibility') !== -1 ? lower.indexOf('visibility') : 5
  };

  for (let i = 1; i < data.length; i++) {
    const rowObj = {};
    headers.forEach((h, ci) => { if (h) rowObj[h] = data[i][ci]; });
    if (postKey_(rowObj) === String(postId).trim()) {
      return { sheet: sheet, rowNum: i + 1, idx: idx, values: data[i] };
    }
  }
  throw new Error("Post not found: " + postId);
}

function updatePost(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();
  const info = getPostRowInfo_(payload.postId);

  assertOwner_(info.values[info.idx.userId], user.userId);

  const raw = String(payload.visibility || "").trim().toLowerCase();
  const vis = raw === "mentor" ? "Mentor" : (raw === "group" ? "Group" : "Public");

  info.sheet.getRange(info.rowNum, info.idx.type + 1).setValue(payload.type || "Praise");
  info.sheet.getRange(info.rowNum, info.idx.text + 1).setValue(payload.text);
  info.sheet.getRange(info.rowNum, info.idx.visibility + 1).setValue(vis);

  return true;
}

/* Unlike Reflections, a post's id is not tied to its row position (postid is
   an explicit column, or a content hash for legacy rows), so a hard delete
   here cannot invalidate any other post's id the way it would there. */
function deletePost(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();
  const info = getPostRowInfo_(payload.postId);

  assertOwner_(info.values[info.idx.userId], user.userId);
  info.sheet.deleteRow(info.rowNum);

  // Reactions left on a deleted post go with it. Bottom-up so each deleteRow
  // does not shift rows still to be visited.
  const reactionsSheet = ss_().getSheetByName(REACTIONS_SHEET);
  if (reactionsSheet && reactionsSheet.getLastRow() > 1) {
    const data = reactionsSheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0] || "").trim() === String(payload.postId).trim()) {
        reactionsSheet.deleteRow(i + 1);
      }
    }
  }

  return true;
}

/* ==========================================================================
   EVENT ATTENDANCE — one row per member per event, written from the mentor
   attendance screen (see Attendance.html).

   The Events tab has no id column, so rows are keyed by a derived eventId:
   the sheet's own eventId cell if that column is ever added, otherwise the
   group plus the event's start instant. Stable as long as the event's date
   and time cells stay put.
   ========================================================================== */

const ATTENDANCE_SHEET = 'Attendance';
const ATTENDANCE_HEADERS = ['eventId', 'groupId', 'userId', 'status', 'note', 'markedBy', 'timestamp'];

function eventKey_(rawRow, groupId, startMs) {
  const explicit = String((rawRow && (rawRow.eventId || rawRow.EventID || rawRow['Event ID'])) || "").trim();
  if (explicit) return explicit;
  if (!startMs) return "";      // unreadable time cell — nothing stable to key on
  const stamp = Utilities.formatDate(new Date(Number(startMs)), Session.getScriptTimeZone(), "yyyyMMdd'T'HHmm");
  return String(groupId || "").trim() + "|" + stamp;
}

/* Resolves the Attendance tab and its header positions.
   Creates the tab, or writes the header row onto an empty one. A tab that
   still holds rows under headers we cannot read is left alone and reported —
   appending under a mismatched header would scatter data across the wrong
   columns (the tab's previous shape was userId | date | TRUE). */
function attendanceSheet_() {
  const ss = ss_();
  let sheet = ss.getSheetByName(ATTENDANCE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ATTENDANCE_SHEET);
    sheet.appendRow(ATTENDANCE_HEADERS);
  }

  const seed = () => { sheet.clear(); sheet.appendRow(ATTENDANCE_HEADERS); };
  if (sheet.getLastRow() === 0) seed();

  let headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0]
                     .map(h => String(h).trim().toLowerCase());
  let idx = {};
  ATTENDANCE_HEADERS.forEach(h => { idx[h] = headers.indexOf(h.toLowerCase()); });

  const missing = ATTENDANCE_HEADERS.filter(h => idx[h] === -1);
  if (missing.length) {
    if (sheet.getLastRow() > 1) {
      throw new Error("The Attendance tab needs these columns in row 1: " +
                      ATTENDANCE_HEADERS.join(", ") + " (missing: " + missing.join(", ") + ").");
    }
    seed();
    headers = ATTENDANCE_HEADERS.map(h => h.toLowerCase());
    idx = {};
    ATTENDANCE_HEADERS.forEach(h => { idx[h] = headers.indexOf(h.toLowerCase()); });
  }

  return { sheet: sheet, idx: idx, width: headers.length };
}

/* Writes a handful of cells on one row, batched.
   `updates` is [{ col, value }] with col 0-based; contiguous columns collapse
   into a single setValues, so a member whose status, note and audit stamp all
   move costs one range write under the default header order rather than four. */
function writeRowCells_(sheet, rowNum, updates) {
  if (!updates.length) return;
  const sorted = updates.slice().sort((a, b) => a.col - b.col);

  let run = [sorted[0]];
  const flush = () => {
    sheet.getRange(rowNum, run[0].col + 1, 1, run.length)
         .setValues([run.map(u => u.value)]);
  };

  for (let i = 1; i < sorted.length; i++) {
    if (sorted[i].col === run[run.length - 1].col + 1) run.push(sorted[i]);
    else { flush(); run = [sorted[i]]; }
  }
  flush();
}

/* { eventId: { recorded, present, absent, total, myStatus } } for one group.
   myStatus is how this caller themselves was marked, which is all a member is
   shown of an event's attendance (see memberAttendanceCard in Events); it stays
   empty when no row in that event belongs to them.

   Anything unreadable — no tab, legacy headers — reads as "nobody has taken
   attendance yet", which is exactly how the client should behave in that case. */
function readAttendanceSummaries_(groupId, userId) {
  const rows = readSheetAsMap(ATTENDANCE_SHEET);
  const me = String(userId || "").trim();
  const byEvent = {};
  rows.forEach(r => {
    const eventId = String(r.eventId || "").trim();
    if (!eventId) return;
    if (groupId && String(r.groupId || "").trim() !== String(groupId).trim()) return;

    const bucket = byEvent[eventId] ||
                   (byEvent[eventId] = { recorded: true, present: 0, absent: 0, total: 0, myStatus: "" });
    const status = String(r.status || "").trim().toLowerCase() === "present" ? "present" : "absent";
    status === "present" ? bucket.present++ : bucket.absent++;
    bucket.total++;

    if (me && String(r.userId || "").trim() === me) bucket.myStatus = status;
  });
  return byEvent;
}

/* The saved record for one event, member by member — what the edit mode of the
   attendance screen prefills itself from. Payload: { userEmail, eventId }.

   Fetched on demand rather than carried on every getData: only a mentor opening
   this one screen ever needs it, and getData runs on every navigation.

   savedAt is the newest timestamp across the event's rows; one save stamps all
   of its rows alike, so it reads as "when this record was written". */
function getEventAttendance(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();

  if (String(user.role || "").trim().toLowerCase() !== "mentor") {
    throw new Error("Only a mentor can open a saved attendance record.");
  }

  const eventId = String((payload && payload.eventId) || "").trim();
  const groupId = String(user.groupId || "").trim();
  const empty = { recorded: false, savedAt: "", marks: [] };
  if (!eventId) return empty;

  const marks = [];
  let savedAt = "";
  readSheetAsMap(ATTENDANCE_SHEET).forEach(r => {
    if (String(r.eventId || "").trim() !== eventId) return;
    if (groupId && String(r.groupId || "").trim() !== groupId) return;

    marks.push({
      userId: String(r.userId || ""),
      status: String(r.status || "").trim().toLowerCase() === "present" ? "present" : "absent",
      note: String(r.note || "")
    });

    const stamp = r.timestamp instanceof Date ? r.timestamp.toISOString() : String(r.timestamp || "");
    if (stamp > savedAt) savedAt = stamp;
  });

  if (!marks.length) return empty;
  return { recorded: true, savedAt: savedAt, marks: marks };
}

/* Reconciles the event's rows against the roster it is handed: a member whose
   status and note both still match is left completely alone, one who moved has
   only those cells rewritten, a new one is appended, and a row belonging to
   nobody on the roster any more is dropped. An edit that changed one member
   therefore touches one row, and a save with nothing to do writes nothing —
   the previous delete-every-row-and-re-append restamped the whole group.

   `marks` must be the full roster, not just what changed: rows are dropped by
   their absence from it. Payload:
   { userEmail, eventId, marks: [{ userId, status, note }] } */
function saveEventAttendance(payload) {
  if (payload && payload.userEmail) setSessionEmail(payload.userEmail);
  const user = getCurrentUserSession();

  if (String(user.role || "").trim().toLowerCase() !== "mentor") {
    throw new Error("Only a mentor can record attendance.");
  }

  const eventId = String((payload && payload.eventId) || "").trim();
  if (!eventId) throw new Error("This event has no usable date and time, so attendance cannot be saved against it.");

  const marks = (payload && payload.marks) || [];
  if (!marks.length) throw new Error("No members to record.");

  const groupId = String(user.groupId || "").trim();
  const { sheet, idx, width } = attendanceSheet_();

  /* The event's rows as they stand, by member. Anything past the first row for
     one member is a leftover from an older double-write and goes on the drop
     list. Another group's rows are never in scope, even where the two share an
     explicit eventId. */
  const byUser = {};
  const drop = [];
  if (sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (String(row[idx.eventId] || "").trim() !== eventId) continue;
      if (groupId && String(row[idx.groupId] || "").trim() !== groupId) continue;

      const uid = String(row[idx.userId] || "").trim();
      if (byUser[uid]) { drop.push(i + 1); continue; }
      byUser[uid] = {
        rowNum: i + 1,
        status: String(row[idx.status] || "").trim().toLowerCase() === "present" ? "present" : "absent",
        note: String(row[idx.note] || "")
      };
    }
  }

  const stamp = new Date().toISOString();
  const seen = {};
  const append = [];
  let touched = 0, present = 0, absent = 0;

  marks.forEach(m => {
    const uid    = String(m.userId || "");
    const status = String(m.status || "").trim().toLowerCase() === "present" ? "present" : "absent";
    const note   = String(m.note || "");
    status === "present" ? present++ : absent++;
    seen[uid.trim()] = true;

    const prev = byUser[uid.trim()];
    if (!prev) {
      const row = new Array(width).fill("");
      row[idx.eventId]   = eventId;
      row[idx.groupId]   = groupId;
      row[idx.userId]    = uid;
      row[idx.status]    = status;
      row[idx.note]      = note;
      row[idx.markedBy]  = String(user.userId || "");
      row[idx.timestamp] = stamp;
      append.push(row);
      return;
    }

    const updates = [];
    if (prev.status !== status) updates.push({ col: idx.status, value: status });
    if (prev.note !== note)     updates.push({ col: idx.note,   value: note });
    if (!updates.length) return;          // unchanged — leave the row untouched

    // The row did move, so it carries who moved it and when.
    updates.push({ col: idx.markedBy,  value: String(user.userId || "") });
    updates.push({ col: idx.timestamp, value: stamp });
    writeRowCells_(sheet, prev.rowNum, updates);
    touched++;
  });

  // Members who have since left the group, plus the duplicates found above.
  Object.keys(byUser).forEach(uid => { if (!seen[uid]) drop.push(byUser[uid].rowNum); });

  // Bottom-up so each deleteRow does not shift rows still to be visited.
  drop.sort((a, b) => b - a).forEach(rowNum => sheet.deleteRow(rowNum));

  if (append.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, append.length, width).setValues(append);
  }

  return { recorded: true, present: present, absent: absent, total: present + absent,
           changed: touched + append.length + drop.length };
}

function readSheetAsMap(sheetName) {
  const sheet = ss_().getSheetByName(sheetName);
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      if (h) obj[h] = row[i];
    });
    return obj;
  });
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.page === 'admin') return adminDoGet_();
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Small Group')
      .setFaviconUrl('https://lh3.googleusercontent.com/d/17-3NVNtoneKapt_mJMzMXEK7EVuRC_RE#.png')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput(`<div style="padding:40px;text-align:center;"><h2>System Notice</h2><p>${error.message}</p></div>`);
  }
}

/**
 * Turns an event's date cell + time cell into the start/end strings the
 * Google Calendar TEMPLATE url expects (see app.addEventToCalendar in Events).
 * Emitted as local wall-clock time, with the script timezone alongside it so
 * the client can pass ctz and skip any UTC conversion. A time cell that can't
 * be read falls back to an all-day slot on the event's date.
 *
 * startMs is the same instant as `start`, but as epoch milliseconds — the only
 * form the client can compare against a clock (the strings above are wall-clock
 * text with no timezone in them). 0 means "no usable start", which is how the
 * mentor attendance card in Events decides to stay hidden.
 */
function buildCalendarSlot(dateObj, timeRaw) {
  const tz = Session.getScriptTimeZone();
  if (!dateObj || isNaN(dateObj.getTime())) return { start: "", end: "", allDay: false, tz: tz, startMs: 0 };

  let h = null, m = 0;
  if (timeRaw instanceof Date) {                 // sheet cell formatted as a time
    h = timeRaw.getHours(); m = timeRaw.getMinutes();
  } else {
    const match = String(timeRaw || "").trim().match(/(\d{1,2})(?::(\d{2}))?\s*([AaPp][Mm])?/);
    if (match) {
      h = parseInt(match[1], 10);
      m = match[2] ? parseInt(match[2], 10) : 0;
      const mer = (match[3] || "").toLowerCase();
      if (mer === "pm" && h < 12) h += 12;
      if (mer === "am" && h === 12) h = 0;
      if (h > 23 || m > 59) h = null;
    }
  }

  const start = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
  if (h === null) {
    const endDay = new Date(start.getTime() + 24 * 60 * 60 * 1000);
    return {
      start: Utilities.formatDate(start, tz, "yyyyMMdd"),
      end: Utilities.formatDate(endDay, tz, "yyyyMMdd"),
      allDay: true, tz: tz, startMs: start.getTime()
    };
  }
  start.setHours(h, m, 0, 0);
  const end = new Date(start.getTime() + 60 * 60 * 1000);   // 1 hour default
  return {
    start: Utilities.formatDate(start, tz, "yyyyMMdd'T'HHmmss"),
    end: Utilities.formatDate(end, tz, "yyyyMMdd'T'HHmmss"),
    allDay: false, tz: tz, startMs: start.getTime()
  };
}

function getNextEventLabel(eventDateObj) {
  const now = new Date();
  const diffTime = eventDateObj - now;
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  
  if (diffDays < 0) return "PAST SESSION";
  if (diffDays === 0) return "TODAY";
  if (diffDays <= 7) {
    return "THIS " + Utilities.formatDate(eventDateObj, Session.getScriptTimeZone(), "EEEE").toUpperCase();
  }
  if (diffDays <= 14) return "NEXT WEEK";
  if (diffDays <= 30) return "NEXT MONTH";
  return `COMING IN ${Math.round(diffDays / 30)} MONTHS`;
}

/**
 * Pulls the raw content of another HTML file in this project so it can be
 * inlined into Index.html. Function declarations and top-level const/let in
 * each partial share one global scope once concatenated, so the modules can
 * reference each other freely.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
//check push again