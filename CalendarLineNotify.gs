/**
 * Googleカレンダーの予定を取得し、自分のLINEへ通知するGoogle Apps Scriptスクリプト。
 *
 * calendar_line_notify.py (Python版) と同じ表示ルールを、Googleのサーバー上で
 * 毎朝自動実行するためのスクリプトです。権限はGoogleカレンダーの読み取りのみで、
 * 書き込みは行いません。
 *
 * セットアップ手順は README.md の「10. 毎朝の自動送信 (Google Apps Script)」を参照してください。
 *
 * 使い方:
 *   - previewOnly        … LINEへは送信せず、ログシートに結果を記録する(動作確認用)
 *   - sendDailyNotification … 実際にLINEへ送信する(トリガーで毎朝自動実行する関数)
 *
 * 必要なスクリプトプロパティ(ファイル > プロジェクトの設定 から設定する):
 *   - LINE_CHANNEL_ACCESS_TOKEN (必須)
 *   - LINE_USER_ID              (必須)
 *   - GOOGLE_CALENDAR_ID        (任意。省略時は "primary")
 */

var DEADLINE_PREFIX = '【締切】';
var DEADLINE_DAYS_AHEAD = 3; // 今日から3日後の終わりまで
var PROJECT_PREFIX = '【案件】';
var PROJECT_WARNING_WITHIN_DAYS = 3; // 残りこの日数以内で ⚠️ を付ける
var LINE_MAX_TEXT_LENGTH = 4500; // LINEのテキスト上限(5000文字)に余裕を持たせた値
var MEETING_LINK_MARK = '🔗 '; // 今日の予定欄で会議の参加URLの前に付ける印
var TIME_ZONE = 'Asia/Tokyo';
var LOG_SHEET_NAME = 'ログ';
var WEEKDAY_JA = ['日', '月', '火', '水', '木', '金', '土'];

// ---- トリガーから呼び出す本体 -------------------------------------------

function sendDailyNotification() {
  var now = new Date();
  var result;

  try {
    result = buildNotificationText_(now);
  } catch (e) {
    appendLog_('失敗(取得エラー)', String(e));
    return;
  }

  if (result.error) {
    appendLog_('失敗(取得エラー)', result.error);
    return;
  }

  var sendResult;
  try {
    sendResult = sendLineMessage_(result.text);
  } catch (e) {
    appendLog_('失敗(送信エラー)', String(e) + '\n---本文---\n' + result.text);
    return;
  }

  if (sendResult.ok) {
    appendLog_('送信成功', result.text);
  } else {
    appendLog_('送信失敗', sendResult.error + '\n---本文---\n' + result.text);
  }
}

/** 動作確認用: LINEへは送信せず、組み立てた内容をログシートに記録するだけ。 */
function previewOnly() {
  var now = new Date();
  var result;
  try {
    result = buildNotificationText_(now);
  } catch (e) {
    Logger.log('エラー: ' + e);
    appendLog_('プレビュー失敗', String(e));
    return;
  }

  if (result.error) {
    Logger.log('エラー: ' + result.error);
    appendLog_('プレビュー失敗', result.error);
    return;
  }

  Logger.log(result.text);
  appendLog_('プレビュー', result.text);
}

// ---- 通知本文の組み立て ----------------------------------------------

function buildNotificationText_(now) {
  var calendarId = getCalendarId_();
  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    return {
      error:
        'カレンダーが見つかりません。スクリプトプロパティ GOOGLE_CALENDAR_ID の値を確認してください。(id=' +
        calendarId +
        ')',
    };
  }

  var todayRange = getTodayRange_(now);
  var deadlineRange = getDeadlineRange_(now);

  var todayRaw, deadlineRaw;
  try {
    // getEvents は指定期間と「重なる」予定を返すため、着手日が今日より前の
    // 進行中の案件もここから拾える(ページ分割はCalendarApp側で自動処理される)
    todayRaw = calendar.getEvents(todayRange.start, todayRange.end);
    deadlineRaw = calendar.getEvents(deadlineRange.start, deadlineRange.end);
  } catch (e) {
    return { error: 'カレンダーの取得に失敗しました: ' + e };
  }

  var todayEvents = filterTodayEvents_(todayRaw, todayRange).filter(function (e) {
    return !isProjectEvent_(e); // 案件は別欄に出すため、今日の予定欄には重複表示しない
  });
  var ongoingProjects = filterOngoingProjects_(todayRaw);
  var deadlineEvents = filterDeadlineEvents_(deadlineRaw, deadlineRange);

  var todayDate = dateOnly_(now);
  var todayLines = todayEvents.map(function (e) {
    return formatTodayLine_(e, getMeetingLinksForEvent_(calendarId, e));
  });
  var projectLines = ongoingProjects.map(function (e) {
    return formatProjectLine_(e, todayDate);
  });
  var deadlineLines = deadlineEvents.map(function (e) {
    return formatEventLine_(e, true);
  });

  var headerToday = '📅 今日の予定（' + formatDateWithWeekday_(now) + '）';
  var headerProjects = '📁 進行中の案件';
  var deadlineLastDay = formatMd_(new Date(deadlineRange.end.getTime() - 24 * 60 * 60 * 1000));
  var headerDeadline = '⏰ 締切（本日〜' + deadlineLastDay + '）';

  var sections = [
    { header: headerToday, lines: todayLines, empty: '今日の予定はありません' },
    { header: headerProjects, lines: projectLines, empty: '進行中の案件はありません' },
    { header: headerDeadline, lines: deadlineLines, empty: '締切の予定はありません' },
  ];

  return { text: buildSectionsText_(sections) };
}

// ---- 日付範囲 (JST) ----------------------------------------------------

/** 指定日時のJSTでの「0時」を表すDateを返す(appsscript.jsonのtimeZone設定が前提)。 */
function dateOnly_(d) {
  var ymd = Utilities.formatDate(d, TIME_ZONE, 'yyyy-MM-dd').split('-').map(Number);
  return new Date(ymd[0], ymd[1] - 1, ymd[2]);
}

function getTodayRange_(now) {
  var start = dateOnly_(now);
  var end = new Date(start.getTime() + 24 * 60 * 60 * 1000);
  return { start: start, end: end };
}

function getDeadlineRange_(now) {
  var start = dateOnly_(now);
  var end = new Date(start.getTime() + (DEADLINE_DAYS_AHEAD + 1) * 24 * 60 * 60 * 1000);
  return { start: start, end: end };
}

// ---- 予定の判定・整形 ---------------------------------------------------

function isAllDayEvent_(event) {
  return event.isAllDayEvent();
}

function isDeadlineEvent_(event) {
  return (event.getTitle() || '').indexOf(DEADLINE_PREFIX) === 0;
}

function isProjectEvent_(event) {
  return isAllDayEvent_(event) && (event.getTitle() || '').indexOf(PROJECT_PREFIX) === 0;
  // 【完了】に名前が変わると先頭が【案件】ではなくなるため、自動的にfalseになる
}

function getEventStart_(event) {
  return isAllDayEvent_(event) ? event.getAllDayStartDate() : event.getStartTime();
}

function getEventEnd_(event) {
  // 終日予定の終了日はCalendarAppでも排他的(最終日の翌日)として返る
  return isAllDayEvent_(event) ? event.getAllDayEndDate() : event.getEndTime();
}

function getProjectDueDate_(event) {
  var end = getEventEnd_(event);
  return new Date(end.getTime() - 24 * 60 * 60 * 1000); // 締切日 = 終了日の前日
}

function byStartTime_(a, b) {
  return getEventStart_(a).getTime() - getEventStart_(b).getTime();
}

function filterTodayEvents_(events, range) {
  // 開始日だけでなく、今日の範囲と重なっている予定(=昨日以前から始まっている
  // 複数日の終日予定なども含む)を対象にする。events自体もgetEventsで
  // 「今日と重なる」もの済みで取得されているため、ここでは終了時刻側の
  // 判定(end > range.start)を加えるだけでよい。
  return events
    .filter(function (e) {
      var start = getEventStart_(e).getTime();
      var end = getEventEnd_(e).getTime();
      return start < range.end.getTime() && end > range.start.getTime();
    })
    .sort(byStartTime_);
}

function filterDeadlineEvents_(events, range) {
  return events
    .filter(function (e) {
      if (!isDeadlineEvent_(e)) return false;
      var start = getEventStart_(e).getTime();
      return start >= range.start.getTime() && start < range.end.getTime();
    })
    .sort(byStartTime_);
}

function filterOngoingProjects_(events) {
  // events は「今日と重なる期間」で取得済みの一覧を想定(着手日が今日より前でもよい)
  return events.filter(isProjectEvent_).sort(function (a, b) {
    return getProjectDueDate_(a).getTime() - getProjectDueDate_(b).getTime();
  });
}

function formatEventLine_(event, showDate) {
  var title = event.getTitle() || '(無題の予定)';
  var start = getEventStart_(event);
  var end = getEventEnd_(event);
  var timePart;

  if (isAllDayEvent_(event)) {
    var lastDay = new Date(end.getTime() - 24 * 60 * 60 * 1000);
    if (formatYmd_(lastDay) !== formatYmd_(start)) {
      timePart = '終日(' + formatMd_(start) + '〜' + formatMd_(lastDay) + ')';
    } else {
      timePart = '終日';
    }
  } else if (formatYmd_(end) !== formatYmd_(start)) {
    timePart = formatMdHm_(start) + '〜' + formatMdHm_(end);
  } else {
    timePart = formatHm_(start) + '-' + formatHm_(end);
  }

  if (showDate) {
    return formatMd_(start) + ' ' + timePart + ' ' + title;
  }
  return timePart + ' ' + title;
}

/** 今日の予定欄向け: 終日予定は予定名だけ、時間がある予定は開始時刻+予定名。
 *  会議の参加URLがある場合は、予定名の下に1件ずつ追加する。 */
function formatTodayLine_(event, meetingUrls) {
  var title = event.getTitle() || '(無題の予定)';
  var line = isAllDayEvent_(event) ? title : formatHm_(getEventStart_(event)) + ' ' + title;

  (meetingUrls || []).forEach(function (url) {
    line += '\n' + MEETING_LINK_MARK + url;
  });

  return line;
}

// ---- 会議の参加URL抽出 ---------------------------------------------------

/**
 * 1件の予定について、会議情報(hangoutLink/conferenceData)・説明欄・場所欄から
 * Zoom/Google MeetのURLを抽出し、重複を除いて返す。
 * 会議以外のURLや説明欄の本文そのものは返さない。
 */
function getMeetingLinksForEvent_(calendarId, event) {
  try {
    var urls = [];
    urls = urls.concat(getConferenceMeetingUrls_(calendarId, event));
    urls = urls.concat(extractMeetingUrls_(event.getDescription()));
    urls = urls.concat(extractMeetingUrls_(event.getLocation()));
    return dedupeUrls_(urls);
  } catch (e) {
    // 参加リンクの取得に失敗しても、予定自体の通知は継続する
    Logger.log('会議URLの取得に失敗しました(' + (event.getTitle() || '') + '): ' + e);
    return [];
  }
}

/**
 * Calendar詳細取得APIのhangoutLink / conferenceData.entryPointsから会議URLを取得する。
 * Advanced Calendar Service(appsscript.jsonのenabledAdvancedServices)が未設定の場合は
 * 何もせず空配列を返す(説明欄・場所欄からの抽出のみにフォールバックする)。
 */
function getConferenceMeetingUrls_(calendarId, event) {
  if (typeof Calendar === 'undefined' || !Calendar.Events) {
    return [];
  }

  var apiEventId = String(event.getId()).replace(/@google\.com$/, '');
  var apiEvent;
  try {
    apiEvent = Calendar.Events.get(calendarId, apiEventId);
  } catch (e) {
    return []; // 取得できない予定(権限・一時的なエラーなど)は無視する
  }

  var urls = [];
  if (apiEvent.hangoutLink) {
    urls.push(apiEvent.hangoutLink);
  }
  if (apiEvent.conferenceData && apiEvent.conferenceData.entryPoints) {
    apiEvent.conferenceData.entryPoints.forEach(function (ep) {
      if (ep.entryPointType === 'video' && ep.uri) {
        urls.push(ep.uri);
      }
    });
  }
  return urls.filter(isMeetingUrl_);
}

/** テキスト中のURLのうち、Zoom/Google MeetのURLだけを抽出する(本文は返さない)。 */
function extractMeetingUrls_(text) {
  if (!text) return [];
  var matches = text.match(/https?:\/\/[^\s<>"']+/g) || [];
  return matches.map(cleanMeetingUrl_).filter(isMeetingUrl_);
}

/** URLの末尾に付いた句読点・括弧などを取り除く。 */
function cleanMeetingUrl_(url) {
  return url.replace(/[)\]」』、。,.;:]+$/, '');
}

/** Zoom/Google Meetの「参加用」URLかどうかを判定する。 */
function isMeetingUrl_(url) {
  var host;
  try {
    host = url.match(/^https?:\/\/([^\/]+)/)[1].toLowerCase();
  } catch (e) {
    return false;
  }

  if (host === 'meet.google.com') {
    return true;
  }
  if (host === 'zoom.us' || host.endsWith('.zoom.us')) {
    return /\/(j|my)\//.test(url); // 参加リンク(/j/…)・個人ミーティングルーム(/my/…)のみ
  }
  return false;
}

function dedupeUrls_(urls) {
  var seen = {};
  var result = [];
  urls.forEach(function (url) {
    if (!seen[url]) {
      seen[url] = true;
      result.push(url);
    }
  });
  return result;
}

/** 進行中の案件欄向け: 締切日・残り日数・警告マークを付けて整形する。 */
function formatProjectLine_(event, todayDate) {
  var title = event.getTitle() || '(無題の案件)';
  var dueDate = getProjectDueDate_(event);
  var remaining = Math.round((dateOnly_(dueDate).getTime() - todayDate.getTime()) / (24 * 60 * 60 * 1000));

  var remainingText;
  if (remaining < 0) {
    remainingText = '期限超過';
  } else if (remaining === 0) {
    remainingText = '本日締切';
  } else {
    remainingText = 'あと' + remaining + '日';
  }

  var mark = remaining <= PROJECT_WARNING_WITHIN_DAYS ? '⚠️ ' : '';
  return mark + title + '（締切' + formatMd_(dueDate) + '・' + remainingText + '）';
}

function formatYmd_(d) {
  return Utilities.formatDate(d, TIME_ZONE, 'yyyy-MM-dd');
}
function formatMd_(d) {
  return Utilities.formatDate(d, TIME_ZONE, 'MM/dd');
}
function formatHm_(d) {
  return Utilities.formatDate(d, TIME_ZONE, 'HH:mm');
}
function formatMdHm_(d) {
  return Utilities.formatDate(d, TIME_ZONE, 'MM/dd HH:mm');
}
function formatDateWithWeekday_(d) {
  return Utilities.formatDate(d, TIME_ZONE, 'yyyy/MM/dd') + '(' + WEEKDAY_JA[dateOnly_(d).getDay()] + ')';
}

// ---- 文字数超過時の省略つき本文組み立て ----------------------------------

function buildSectionsText_(sections) {
  function render(secs) {
    return secs
      .map(function (s) {
        var body = s.lines.length ? s.lines.join('\n') : s.empty;
        if (s.omitted) {
          body += '\n…ほか' + s.omitted + '件は長くなるため省略しました。Googleカレンダーでご確認ください。';
        }
        return s.header + '\n' + body;
      })
      .join('\n\n');
  }

  var working = sections.map(function (s) {
    return { header: s.header, lines: s.lines.slice(), empty: s.empty, omitted: 0 };
  });

  var text = render(working);
  while (
    text.length > LINE_MAX_TEXT_LENGTH &&
    working.some(function (s) {
      return s.lines.length > 0;
    })
  ) {
    var target = working.reduce(function (best, s, i) {
      return s.lines.length > working[best].lines.length ? i : best;
    }, 0);
    if (working[target].lines.length === 0) break;
    working[target].lines.pop();
    working[target].omitted += 1;
    text = render(working);
  }
  return text;
}

// ---- LINE送信 -----------------------------------------------------------

function getCalendarId_() {
  var props = PropertiesService.getScriptProperties();
  return props.getProperty('GOOGLE_CALENDAR_ID') || 'primary';
}

function sendLineMessage_(text) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  var userId = props.getProperty('LINE_USER_ID');

  if (!token || !userId) {
    return {
      ok: false,
      error: 'LINE_CHANNEL_ACCESS_TOKEN / LINE_USER_ID が未設定です。スクリプトプロパティを確認してください。',
    };
  }

  var payload = {
    to: userId,
    messages: [{ type: 'text', text: text }],
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true,
  };

  try {
    var response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', options);
    var code = response.getResponseCode();
    if (code === 200) {
      return { ok: true };
    }
    return { ok: false, error: 'status=' + code + ' body=' + response.getContentText() };
  } catch (e) {
    return { ok: false, error: '通信エラー: ' + e };
  }
}

// ---- ログ(スプレッドシートの「ログ」シートに保存) --------------------------

function appendLog_(status, message) {
  try {
    var sheet = getLogSheet_();
    sheet.appendRow([new Date(), status, message]);
  } catch (e) {
    Logger.log('ログ書き込みに失敗しました: ' + e);
  }
}

function getLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(['実行日時', '状態', '内容']);
  }
  return sheet;
}
