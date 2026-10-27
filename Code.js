/**
 * Southwest Asia · Physical Geography · M3 — Google Apps Script Web App
 * ------------------------------------------------------------
 * Serves BOTH the student lesson app AND the teacher analytics dashboard.
 *
 * Routing:
 *   ?view=teacher  → teacher dashboard (PIN-gated)
 *   (no parameter) → student lesson app (unchanged behavior)
 *
 * Sheet routing (per section):
 *   Quiz scores     → "M3/10"  or  "M3/11"
 *   Grammar tagging → "M3/10 - Grammar Tagging"  or  "M3/11 - Grammar Tagging"
 *   Reteach log     → "Reteach Log" (single shared sheet)
 */

var SPREADSHEET_ID = '1ffriu1qebQL5xjsnfn7FGBnGKyNlynw-928wK1kXmw0';
var SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1ffriu1qebQL5xjsnfn7FGBnGKyNlynw-928wK1kXmw0';
var VALID_SECTIONS = ['M3/10', 'M3/11'];
var TEACHER_PIN = '2025';

var QUIZ_HEADERS = [
  'Timestamp', 'Student ID', 'Full Name', 'Section',
  'Quiz Key', 'Quiz Title', 'Score', 'Total', 'Percentage', 'Responses (JSON)'
];

var GRAMMAR_HEADERS = [
  'Timestamp', 'Student ID', 'Full Name', 'Section',
  'Lesson Section', 'Correct', 'Wrong', 'Unattempted', 'Total Words',
  'Percentage', 'Tag Detail (JSON)'
];

var RETeach_HEADERS = [
  'Timestamp', 'Section Filter', 'Quiz Key', 'Item Type', 'Item Identifier',
  'Question Text', 'Status', 'Date Planned', 'Date Completed', 'Method',
  'Students Targeted', 'Notes', 'Post-Check Result'
];

/**
 * Serves the web app. Routes to teacher dashboard if ?view=teacher,
 * otherwise serves the student lesson app.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.view === 'teacher') {
    return HtmlService.createTemplateFromFile('teacher')
      .evaluate()
      .setTitle('M3 SW Asia — Teacher Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Southwest Asia · Physical Geography · M3')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Lets HTML templates pull in stylesheet.html / script.html / teacher-*.html. */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ============================================================
   SHEET NAME RESOLVERS — derive the correct sheet per section
   ============================================================ */

function _quizSheetName(section) {
  return String(section || '').trim();
}

function _grammarSheetName(section) {
  return String(section || '').trim() + ' - Grammar Tagging';
}

function _reteachSheetName() {
  return 'Reteach Log';
}

/* ============================================================
   VALIDATION
   ============================================================ */
function _isValidStudentId(id) {
  return /^\d{5}$/.test(String(id || '').trim());
}

function _isValidSection(section) {
  return VALID_SECTIONS.indexOf(String(section || '').trim()) !== -1;
}

function _validateEntry(payload) {
  var id = String((payload && payload.studentId) || '').trim();
  var name = String((payload && payload.fullName) || '').trim();
  var section = String((payload && payload.section) || '').trim();

  if (!_isValidStudentId(id)) {
    return { valid: false, message: 'Student ID must be exactly 5 digits.' };
  }
  if (!name) {
    return { valid: false, message: 'Full name is required.' };
  }
  if (!_isValidSection(section)) {
    return { valid: false, message: 'Section must be one of: ' + VALID_SECTIONS.join(', ') + '.' };
  }
  return { valid: true, id: id, name: name, section: section };
}

/* ============================================================
   SHEET HELPERS
   ============================================================ */

function _getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  var firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var headersMatch = true;
  for (var i = 0; i < headers.length; i++) {
    if (firstRow[i] !== headers[i]) {
      headersMatch = false;
      break;
    }
  }
  if (!headersMatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _readSheetRows(sheetName, expectedHeaders) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    return sheet.getRange(2, 1, lastRow - 1, expectedHeaders.length).getValues();
  } catch (err) {
    return [];
  }
}

/* ============================================================
   PUBLIC ENDPOINTS — called from script.html via google.script.run
   ============================================================ */

function submitQuizScore(payload) {
  var check = _validateEntry(payload);
  if (!check.valid) {
    return { success: false, message: check.message };
  }
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var targetSheet = _quizSheetName(check.section);
      var sheet = _getOrCreateSheet(targetSheet, QUIZ_HEADERS);
      sheet.appendRow([
        new Date(),
        check.id,
        check.name,
        check.section,
        (payload && payload.quizKey) || '',
        (payload && payload.quizTitle) || '',
        (payload && payload.score) != null ? payload.score : '',
        (payload && payload.total) != null ? payload.total : '',
        (payload && payload.percentage) != null ? payload.percentage : '',
        JSON.stringify((payload && payload.responses) || [])
      ]);
    } finally {
      lock.releaseLock();
    }
    return { success: true };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

function submitGrammarResult(payload) {
  var check = _validateEntry(payload);
  if (!check.valid) {
    return { success: false, message: check.message };
  }
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var targetSheet = _grammarSheetName(check.section);
      var sheet = _getOrCreateSheet(targetSheet, GRAMMAR_HEADERS);
      sheet.appendRow([
        new Date(),
        check.id,
        check.name,
        check.section,
        (payload && payload.lessonSection) || '',
        (payload && payload.correct) != null ? payload.correct : '',
        (payload && payload.wrong) != null ? payload.wrong : '',
        (payload && payload.unattempted) != null ? payload.unattempted : '',
        (payload && payload.total) != null ? payload.total : '',
        (payload && payload.percentage) != null ? payload.percentage : '',
        JSON.stringify((payload && payload.detail) || [])
      ]);
    } finally {
      lock.releaseLock();
    }
    return { success: true };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/* ============================================================
   TEACHER DASHBOARD — analytics endpoints (READ-ONLY)
   ============================================================ */

function teacherLogin(pin) {
  if (String(pin || '').trim() === TEACHER_PIN) {
    return {
      success: true,
      spreadsheetUrl: SPREADSHEET_URL,
      spreadsheetId: SPREADSHEET_ID
    };
  }
  return { success: false, message: 'Incorrect PIN. Try again.' };
}

function getQuizAnalytics(sectionFilter) {
  try {
    var filter = String(sectionFilter || 'ALL').trim();
    var sheetNames = (filter === 'ALL') ? ['M3/10', 'M3/11'] : [filter];

    var sectionMap = {};

    sheetNames.forEach(function(sheetName) {
      var rows = _readSheetRows(sheetName, QUIZ_HEADERS);
      rows.forEach(function(row) {
        var quizKey = String(row[4] || '').trim();
        var quizTitle = String(row[5] || '').trim();
        var jsonStr = String(row[9] || '').trim();
        if (!jsonStr) return;

        var responses;
        try { responses = JSON.parse(jsonStr); } catch (e) { return; }
        if (!Array.isArray(responses) || !responses.length) return;

        if (!sectionMap[quizKey]) {
          sectionMap[quizKey] = {
            quizKey: quizKey,
            quizTitle: quizTitle,
            questions: {}
          };
        }
        var sectionData = sectionMap[quizKey];

        responses.forEach(function(r, idx) {
          var qText = (r.question || '').toString();
          if (!qText) return;
          var qKey = 'Q' + (idx + 1) + '::' + qText;
          if (!sectionData.questions[qKey]) {
            sectionData.questions[qKey] = {
              order: idx + 1,
              question: qText,
              type: r.type || 'unknown',
              correct: 0,
              wrong: 0,
              total: 0
            };
          }
          var q = sectionData.questions[qKey];
          q.total++;
          if (r.isCorrect === true) q.correct++;
          else q.wrong++;
        });
      });
    });

    var sectionsArr = [];
    var grandCorrect = 0;
    var grandWrong = 0;

    Object.keys(sectionMap).forEach(function(key) {
      var sd = sectionMap[key];
      var qArr = [];
      Object.keys(sd.questions).forEach(function(qk) {
        var q = sd.questions[qk];
        var pct = q.total > 0 ? Math.round((q.correct / q.total) * 100) : 0;
        qArr.push({
          order: q.order,
          question: q.question,
          type: q.type,
          correct: q.correct,
          wrong: q.wrong,
          total: q.total,
          percentage: pct,
          status: pct >= 80 ? 'strong' : (pct >= 60 ? 'review' : 'reteach')
        });
        grandCorrect += q.correct;
        grandWrong += q.wrong;
      });
      qArr.sort(function(a, b) { return a.order - b.order; });

      var sectionTotal = qArr.reduce(function(sum, q) { return sum + q.total; }, 0);
      var sectionCorrect = qArr.reduce(function(sum, q) { return sum + q.correct; }, 0);
      var sectionAvg = sectionTotal > 0 ? Math.round((sectionCorrect / sectionTotal) * 100) : 0;

      var hardest = null;
      qArr.forEach(function(q) {
        if (q.total >= 1) {
          if (!hardest || q.percentage < hardest.percentage) hardest = q;
        }
      });

      sectionsArr.push({
        quizKey: sd.quizKey,
        quizTitle: sd.quizTitle,
        questions: qArr,
        totalAttempts: sectionTotal,
        avgPercentage: sectionAvg,
        hardestQuestion: hardest
      });
    });

    var order = ['landforms','plains','deserts','coasts','climate','vegetation','resources','landuse'];
    sectionsArr.sort(function(a, b) {
      var ai = order.indexOf(a.quizKey);
      var bi = order.indexOf(b.quizKey);
      if (ai === -1) ai = 999;
      if (bi === -1) bi = 999;
      return ai - bi;
    });

    var grandTotal = grandCorrect + grandWrong;
    var grandPct = grandTotal > 0 ? Math.round((grandCorrect / grandTotal) * 100) : 0;

    return {
      success: true,
      sections: sectionsArr,
      grandTotal: grandTotal,
      grandCorrect: grandCorrect,
      grandWrong: grandWrong,
      grandPercentage: grandPct
    };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

function getGrammarAnalytics(sectionFilter) {
  try {
    var filter = String(sectionFilter || 'ALL').trim();
    var baseSections = (filter === 'ALL') ? ['M3/10', 'M3/11'] : [filter];

    var sectionMap = {};

    baseSections.forEach(function(baseSection) {
      var grammarSheetName = baseSection + ' - Grammar Tagging';
      var rows = _readSheetRows(grammarSheetName, GRAMMAR_HEADERS);
      rows.forEach(function(row) {
        var lessonSection = String(row[4] || '').trim();
        var jsonStr = String(row[10] || '').trim();
        if (!jsonStr || !lessonSection) return;

        var detail;
        try { detail = JSON.parse(jsonStr); } catch (e) { return; }
        if (!Array.isArray(detail) || !detail.length) return;

        if (!sectionMap[lessonSection]) {
          sectionMap[lessonSection] = { lessonSection: lessonSection, words: {} };
        }
        var sd = sectionMap[lessonSection];

        detail.forEach(function(d) {
          var word = (d.word || '').toString().trim();
          if (!word) return;
          var correctTag = parseInt(d.correct, 10) || 0;
          var wKey = word + '::' + correctTag;

          if (!sd.words[wKey]) {
            sd.words[wKey] = {
              word: word,
              correctTag: correctTag,
              correct: 0,
              wrong: 0,
              unattempted: 0,
              total: 0,
              mislabels: {}
            };
          }
          var w = sd.words[wKey];
          w.total++;
          if (d.tagged === null || d.tagged === undefined || d.tagged === 0) {
            w.unattempted++;
          } else if (d.isCorrect === true) {
            w.correct++;
          } else {
            w.wrong++;
            var tag = parseInt(d.tagged, 10) || 0;
            w.mislabels[tag] = (w.mislabels[tag] || 0) + 1;
          }
        });
      });
    });

    var catLabel = { 1: 'Noun', 2: 'Verb', 3: 'Adjective', 4: 'Adverb' };

    var sectionsArr = [];
    Object.keys(sectionMap).forEach(function(key) {
      var sd = sectionMap[key];
      var wordArr = [];
      Object.keys(sd.words).forEach(function(wk) {
        var w = sd.words[wk];
        var attempted = w.correct + w.wrong;
        var pct = attempted > 0 ? Math.round((w.correct / attempted) * 100) : 0;

        var mislabeledArr = [];
        Object.keys(w.mislabels).forEach(function(k) {
          mislabeledArr.push({
            tag: parseInt(k, 10),
            label: catLabel[k] || 'Unknown',
            count: w.mislabels[k]
          });
        });
        mislabeledArr.sort(function(a, b) { return b.count - a.count; });

        wordArr.push({
          word: w.word,
          correctTag: w.correctTag,
          correctTagLabel: catLabel[w.correctTag] || 'Unknown',
          correct: w.correct,
          wrong: w.wrong,
          unattempted: w.unattempted,
          total: w.total,
          attempted: attempted,
          percentage: pct,
          mislabels: mislabeledArr,
          status: pct >= 80 ? 'strong' : (pct >= 60 ? 'review' : 'reteach')
        });
      });
      wordArr.sort(function(a, b) { return a.percentage - b.percentage; });

      var sectionTotalAttempts = wordArr.reduce(function(s, w) { return s + w.attempted; }, 0);
      var sectionCorrect = wordArr.reduce(function(s, w) { return s + w.correct; }, 0);
      var sectionAvg = sectionTotalAttempts > 0 ? Math.round((sectionCorrect / sectionTotalAttempts) * 100) : 0;

      var hardestWord = wordArr.length ? wordArr[0] : null;

      sectionsArr.push({
        lessonSection: sd.lessonSection,
        words: wordArr,
        totalAttempts: sectionTotalAttempts,
        avgPercentage: sectionAvg,
        hardestWord: hardestWord
      });
    });

    var order = ['landforms','plains','deserts','coasts','climate','vegetation','resources','landuse'];
    sectionsArr.sort(function(a, b) {
      var ai = order.indexOf(a.lessonSection);
      var bi = order.indexOf(b.lessonSection);
      if (ai === -1) ai = 999;
      if (bi === -1) bi = 999;
      return ai - bi;
    });

    return { success: true, sections: sectionsArr };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

function getQuizAttemptsBySection(sectionFilter, quizKey) {
  try {
    var filter = String(sectionFilter || 'ALL').trim();
    var targetQuiz = String(quizKey || '').trim();
    var sheetNames = (filter === 'ALL') ? ['M3/10', 'M3/11'] : [filter];

    var attempts = [];

    sheetNames.forEach(function(sheetName) {
      var rows = _readSheetRows(sheetName, QUIZ_HEADERS);
      rows.forEach(function(row) {
        var qKey = String(row[4] || '').trim();
        if (qKey !== targetQuiz) return;
        var jsonStr = String(row[9] || '').trim();
        var responses = [];
        try { responses = JSON.parse(jsonStr); } catch (e) { responses = []; }

        attempts.push({
          timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0] || ''),
          studentId: String(row[1] || ''),
          fullName: String(row[2] || ''),
          section: String(row[3] || ''),
          quizTitle: String(row[5] || ''),
          score: row[6],
          total: row[7],
          percentage: row[8],
          responses: responses
        });
      });
    });

    attempts.sort(function(a, b) {
      return String(b.timestamp).localeCompare(String(a.timestamp));
    });

    return { success: true, attempts: attempts };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

function getQuestionStudentBreakdown(sectionFilter, quizKey, questionOrder) {
  try {
    var filter = String(sectionFilter || 'ALL').trim();
    var targetQuiz = String(quizKey || '').trim();
    var qOrder = parseInt(questionOrder, 10) || 0;
    var sheetNames = (filter === 'ALL') ? ['M3/10', 'M3/11'] : [filter];

    var allAttempts = [];

    sheetNames.forEach(function(sheetName) {
      var rows = _readSheetRows(sheetName, QUIZ_HEADERS);
      rows.forEach(function(row) {
        var qKey = String(row[4] || '').trim();
        if (qKey !== targetQuiz) return;
        var jsonStr = String(row[9] || '').trim();
        if (!jsonStr) return;
        var responses;
        try { responses = JSON.parse(jsonStr); } catch (e) { return; }
        if (!Array.isArray(responses) || responses.length < qOrder) return;

        var ts = row[0] instanceof Date ? row[0].getTime() : new Date(row[0]).getTime();
        allAttempts.push({
          timestamp: ts,
          studentId: String(row[1] || '').trim(),
          fullName: String(row[2] || '').trim(),
          section: String(row[3] || '').trim(),
          quizTitle: String(row[5] || '').trim(),
          responses: responses
        });
      });
    });

    var latestByStudent = {};
    allAttempts.forEach(function(a) {
      var key = a.studentId + '::' + a.section;
      if (!latestByStudent[key] || a.timestamp > latestByStudent[key].timestamp) {
        latestByStudent[key] = a;
      }
    });

    var correctList = [];
    var wrongList = [];
    var wrongAnswerCounts = {};
    var questionText = '';
    var questionType = '';

    Object.keys(latestByStudent).forEach(function(key) {
      var attempt = latestByStudent[key];
      var r = attempt.responses[qOrder - 1];
      if (!r) return;

      if (!questionText && r.question) {
        questionText = String(r.question);
        questionType = r.type || 'unknown';
      }

      var entry = {
        studentId: attempt.studentId,
        fullName: attempt.fullName,
        section: attempt.section,
        chosen: r.type === 'match' ? 'See attempt detail' : (r.chosen != null ? r.chosen : '—'),
        correctAnswer: r.correctAnswer != null ? r.correctAnswer : '—',
        attemptedAt: new Date(attempt.timestamp).toISOString()
      };

      if (r.isCorrect === true) {
        correctList.push(entry);
      } else {
        wrongList.push(entry);
        var ans = String(entry.chosen || '—').trim();
        wrongAnswerCounts[ans] = (wrongAnswerCounts[ans] || 0) + 1;
      }
    });

    correctList.sort(function(a, b) { return a.fullName.localeCompare(b.fullName); });
    wrongList.sort(function(a, b) { return a.fullName.localeCompare(b.fullName); });

    var topWrong = null;
    Object.keys(wrongAnswerCounts).forEach(function(k) {
      if (!topWrong || wrongAnswerCounts[k] > topWrong.count) {
        topWrong = { answer: k, count: wrongAnswerCounts[k] };
      }
    });

    return {
      success: true,
      question: questionText,
      questionType: questionType,
      questionOrder: qOrder,
      correct: correctList,
      wrong: wrongList,
      topWrongAnswer: topWrong,
      totalStudentsConsidered: correctList.length + wrongList.length
    };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

function getGrammarStudentBreakdown(sectionFilter, lessonSection, wordText, correctTag) {
  try {
    var filter = String(sectionFilter || 'ALL').trim();
    var targetSection = String(lessonSection || '').trim();
    var targetWord = String(wordText || '').trim();
    var targetTag = parseInt(correctTag, 10) || 0;
    var baseSections = (filter === 'ALL') ? ['M3/10', 'M3/11'] : [filter];

    var catLabel = { 1: 'Noun', 2: 'Verb', 3: 'Adjective', 4: 'Adverb' };

    var allAttempts = [];

    baseSections.forEach(function(baseSection) {
      var grammarSheetName = baseSection + ' - Grammar Tagging';
      var rows = _readSheetRows(grammarSheetName, GRAMMAR_HEADERS);
      rows.forEach(function(row) {
        var ls = String(row[4] || '').trim();
        if (ls !== targetSection) return;
        var jsonStr = String(row[10] || '').trim();
        if (!jsonStr) return;
        var detail;
        try { detail = JSON.parse(jsonStr); } catch (e) { return; }
        if (!Array.isArray(detail)) return;

        var ts = row[0] instanceof Date ? row[0].getTime() : new Date(row[0]).getTime();
        allAttempts.push({
          timestamp: ts,
          studentId: String(row[1] || '').trim(),
          fullName: String(row[2] || '').trim(),
          section: String(row[3] || '').trim(),
          detail: detail
        });
      });
    });

    var latestByStudent = {};
    allAttempts.forEach(function(a) {
      var key = a.studentId + '::' + a.section;
      if (!latestByStudent[key] || a.timestamp > latestByStudent[key].timestamp) {
        latestByStudent[key] = a;
      }
    });

    var correctList = [];
    var wrongList = [];
    var mislabelCounts = {};

    Object.keys(latestByStudent).forEach(function(key) {
      var attempt = latestByStudent[key];
      var match = null;
      for (var i = 0; i < attempt.detail.length; i++) {
        var d = attempt.detail[i];
        var w = (d.word || '').toString().trim();
        var t = parseInt(d.correct, 10) || 0;
        if (w === targetWord && t === targetTag) {
          match = d;
          break;
        }
      }
      if (!match) return;

      var entry = {
        studentId: attempt.studentId,
        fullName: attempt.fullName,
        section: attempt.section,
        tagged: match.tagged,
        taggedLabel: (match.tagged && catLabel[match.tagged]) ? catLabel[match.tagged] : '—',
        attemptedAt: new Date(attempt.timestamp).toISOString()
      };

      if (match.isCorrect === true) {
        correctList.push(entry);
      } else if (match.tagged === null || match.tagged === undefined || match.tagged === 0) {
        wrongList.push({
          studentId: entry.studentId,
          fullName: entry.fullName,
          section: entry.section,
          tagged: 0,
          taggedLabel: '— (not tagged)',
          attemptedAt: entry.attemptedAt
        });
      } else {
        wrongList.push(entry);
        var lbl = entry.taggedLabel;
        mislabelCounts[lbl] = (mislabelCounts[lbl] || 0) + 1;
      }
    });

    correctList.sort(function(a, b) { return a.fullName.localeCompare(b.fullName); });
    wrongList.sort(function(a, b) { return a.fullName.localeCompare(b.fullName); });

    var mislabelArr = [];
    Object.keys(mislabelCounts).forEach(function(lbl) {
      mislabelArr.push({ label: lbl, count: mislabelCounts[lbl] });
    });
    mislabelArr.sort(function(a, b) { return b.count - a.count; });

    return {
      success: true,
      word: targetWord,
      correctTag: targetTag,
      correctTagLabel: catLabel[targetTag] || 'Unknown',
      correct: correctList,
      wrong: wrongList,
      mislabelBreakdown: mislabelArr,
      totalStudentsConsidered: correctList.length + wrongList.length
    };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/* ============================================================
   RETeach LOG — endpoints
   ============================================================ */

/**
 * Saves a reteach log entry. Appends a new row to the Reteach Log sheet.
 *
 * @param {object} payload - {
 *   sectionFilter, quizKey, itemType, itemIdentifier,
 *   questionText, status, datePlanned, dateCompleted,
 *   method, studentsTargeted, notes, postCheckResult
 * }
 * @return {object} { success, message }
 */
function saveReteachLog(payload) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sheet = _getOrCreateSheet(_reteachSheetName(), RETeach_HEADERS);
      sheet.appendRow([
        new Date(),
        (payload && payload.sectionFilter) || 'ALL',
        (payload && payload.quizKey) || '',
        (payload && payload.itemType) || 'quiz',
        (payload && payload.itemIdentifier) || '',
        (payload && payload.questionText) || '',
        (payload && payload.status) || 'Needs Reteach',
        (payload && payload.datePlanned) || '',
        (payload && payload.dateCompleted) || '',
        (payload && payload.method) || '',
        (payload && payload.studentsTargeted) || '',
        (payload && payload.notes) || '',
        (payload && payload.postCheckResult) || ''
      ]);
    } finally {
      lock.releaseLock();
    }
    return { success: true };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Returns the full log history for a specific quiz question.
 *
 * @param {string} quizKey - e.g. 'landforms'
 * @param {number} questionOrder - 1-based question number
 * @return {object} { success, logs: [...] }
 */
function getReteachLogsForQuestion(quizKey, questionOrder) {
  try {
    var targetQuiz = String(quizKey || '').trim();
    var targetOrder = parseInt(questionOrder, 10) || 0;
    var identifier = 'Q' + targetOrder;

    var rows = _readSheetRows(_reteachSheetName(), RETeach_HEADERS);
    var logs = [];

    rows.forEach(function(row) {
      var itemType = String(row[3] || '').trim();
      var itemId = String(row[4] || '').trim();
      var qKey = String(row[2] || '').trim();
      if (itemType !== 'quiz') return;
      if (qKey !== targetQuiz) return;
      if (itemId !== identifier) return;

      logs.push({
        timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0] || ''),
        sectionFilter: String(row[1] || ''),
        quizKey: qKey,
        itemType: itemType,
        itemIdentifier: itemId,
        questionText: String(row[5] || ''),
        status: String(row[6] || ''),
        datePlanned: row[7] ? (row[7] instanceof Date ? _fmtDate(row[7]) : String(row[7])) : '',
        dateCompleted: row[8] ? (row[8] instanceof Date ? _fmtDate(row[8]) : String(row[8])) : '',
        method: String(row[9] || ''),
        studentsTargeted: String(row[10] || ''),
        notes: String(row[11] || ''),
        postCheckResult: String(row[12] || '')
      });
    });

    logs.sort(function(a, b) {
      return String(b.timestamp).localeCompare(String(a.timestamp));
    });

    return { success: true, logs: logs };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Returns the full log history for a specific grammar word.
 *
 * @param {string} lessonSection - e.g. 'landforms'
 * @param {string} wordText - the word
 * @param {number} correctTag - the tag code
 * @return {object} { success, logs: [...] }
 */
function getReteachLogsForGrammar(lessonSection, wordText, correctTag) {
  try {
    var targetSection = String(lessonSection || '').trim();
    var targetWord = String(wordText || '').trim();
    var targetTag = parseInt(correctTag, 10) || 0;
    var identifier = targetWord + '::' + targetTag;

    var rows = _readSheetRows(_reteachSheetName(), RETeach_HEADERS);
    var logs = [];

    rows.forEach(function(row) {
      var itemType = String(row[3] || '').trim();
      var itemId = String(row[4] || '').trim();
      var qKey = String(row[2] || '').trim();
      if (itemType !== 'grammar') return;
      if (qKey !== targetSection) return;
      if (itemId !== identifier) return;

      logs.push({
        timestamp: row[0] instanceof Date ? row[0].toISOString() : String(row[0] || ''),
        sectionFilter: String(row[1] || ''),
        quizKey: qKey,
        itemType: itemType,
        itemIdentifier: itemId,
        questionText: String(row[5] || ''),
        status: String(row[6] || ''),
        datePlanned: row[7] ? (row[7] instanceof Date ? _fmtDate(row[7]) : String(row[7])) : '',
        dateCompleted: row[8] ? (row[8] instanceof Date ? _fmtDate(row[8]) : String(row[8])) : '',
        method: String(row[9] || ''),
        studentsTargeted: String(row[10] || ''),
        notes: String(row[11] || ''),
        postCheckResult: String(row[12] || '')
      });
    });

    logs.sort(function(a, b) {
      return String(b.timestamp).localeCompare(String(a.timestamp));
    });

    return { success: true, logs: logs };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Returns a map of latest reteach status for every logged item.
 * Used by the dashboard to show status badges on question rows.
 *
 * @return {object} { success, statuses: { 'quiz::landforms::Q5': {...}, ... } }
 */
function getReteachStatusMap() {
  try {
    var rows = _readSheetRows(_reteachSheetName(), RETeach_HEADERS);
    var statuses = {};

    rows.forEach(function(row) {
      var itemType = String(row[3] || '').trim();
      var itemId = String(row[4] || '').trim();
      var qKey = String(row[2] || '').trim();
      var key = itemType + '::' + qKey + '::' + itemId;
      var ts = row[0] instanceof Date ? row[0].getTime() : new Date(row[0]).getTime();

      if (!statuses[key] || ts > statuses[key]._ts) {
        statuses[key] = {
          _ts: ts,
          status: String(row[6] || ''),
          datePlanned: row[7] ? (row[7] instanceof Date ? _fmtDate(row[7]) : String(row[7])) : '',
          dateCompleted: row[8] ? (row[8] instanceof Date ? _fmtDate(row[8]) : String(row[8])) : '',
          method: String(row[9] || ''),
          postCheckResult: String(row[12] || '')
        };
      }
    });

    Object.keys(statuses).forEach(function(k) {
      delete statuses[k]._ts;
    });

    return { success: true, statuses: statuses };
  } catch (err) {
    return { success: false, message: 'Server error: ' + err.message };
  }
}

/**
 * Helper: format a Date as YYYY-MM-DD.
 */
function _fmtDate(d) {
  try {
    var y = d.getFullYear();
    var m = String(d.getMonth() + 1).padStart(2, '0');
    var dd = String(d.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + dd;
  } catch (err) {
    return '';
  }
}