var DEFAULT_CONFIG = {
  RSS_URL: 'https://j-net21.smrj.go.jp/snavi/support/support.xml',
  TARGET_REGIONS: '東京都,神奈川県,埼玉県,千葉県',
  DAILY_TRIGGER_HOUR: '19',
  SHEET_CONFIG: 'config',
  SHEET_STATE: 'state',
  SHEET_RUN_LOG: 'run_log',
  SHEET_ITEM_LOG: 'item_log',
  SHEET_PUBLIC_ITEM_LOG: 'public_item_log',
  TIMEZONE: 'Asia/Tokyo',
};

var CONFIG_KEYS = [
  'RSS_URL',
  'TARGET_REGIONS',
  'DAILY_TRIGGER_HOUR',
];

var SHEET_DEFINITIONS = {
  config: ['key', 'value', 'description'],
  state: [
    'external_id',
    'title',
    'link',
    'published_at',
    'matched_regions',
    'status',
    'first_seen_at',
    'last_seen_at',
    'last_notified_at',
  ],
  run_log: [
    'run_id',
    'executed_at',
    'rss_total_count',
    'region_matched_count',
    'new_item_count',
    'notification_status',
    'message_text',
    'error_message',
  ],
  item_log: [
    'run_id',
    'executed_at',
    'external_id',
    'title',
    'link',
    'published_at',
    'matched_regions',
    'is_new',
    'slack_sent',
    'slack_status',
    'raw_region_text',
  ],
  public_item_log: [
    'executed_at',
    'title',
    'link',
    'published_at',
    'matched_regions',
  ],
};

var PREFECTURES = [
  '北海道', '青森県', '岩手県', '宮城県', '秋田県', '山形県', '福島県',
  '茨城県', '栃木県', '群馬県', '埼玉県', '千葉県', '東京都', '神奈川県',
  '新潟県', '富山県', '石川県', '福井県', '山梨県', '長野県',
  '岐阜県', '静岡県', '愛知県', '三重県',
  '滋賀県', '京都府', '大阪府', '兵庫県', '奈良県', '和歌山県',
  '鳥取県', '島根県', '岡山県', '広島県', '山口県',
  '徳島県', '香川県', '愛媛県', '高知県',
  '福岡県', '佐賀県', '長崎県', '熊本県', '大分県', '宮崎県', '鹿児島県',
  '沖縄県',
];

var XML_NAMESPACES = {
  dc: XmlService.getNamespace('dc', 'http://purl.org/dc/elements/1.1/'),
  rdf: XmlService.getNamespace('rdf', 'http://www.w3.org/1999/02/22-rdf-syntax-ns#'),
};

function initializeProject() {
  var spreadsheet = getSpreadsheet_();
  ensureSheets_(spreadsheet);
  ensureConfigRows_(spreadsheet);
  ensurePublicSheet_();
  ensureDailyTrigger_();
}

function runDaily() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  var runId = Utilities.getUuid();
  var executedAt = new Date();
  var spreadsheet = null;
  var matchedItems = [];
  var newItems = [];
  var runSummary = {
    rssTotalCount: 0,
    regionMatchedCount: 0,
    newItemCount: 0,
    notificationStatus: 'not_started',
    messageText: '',
    errorMessage: '',
  };

  try {
    spreadsheet = getSpreadsheet_();
    ensureSheets_(spreadsheet);
    ensureConfigRows_(spreadsheet);

    var config = loadConfig_(spreadsheet);
    var items = fetchRssItems_(config.RSS_URL);
    matchedItems = filterItemsByRegions_(items, config.targetRegions);
    var stateMap = loadStateMap_(spreadsheet);
    newItems = findNewItems_(matchedItems, stateMap);
    var messageText = '';
    var slackStatus = 'skipped_no_items';

    if (newItems.length > 0) {
      messageText = buildSlackMessage_(newItems, config.targetRegions, executedAt);
      slackStatus = postToSlack_(messageText);
    }

    syncState_(spreadsheet, matchedItems, stateMap, executedAt, slackStatus);

    runSummary = {
      rssTotalCount: items.length,
      regionMatchedCount: matchedItems.length,
      newItemCount: newItems.length,
      notificationStatus: slackStatus,
      messageText: messageText,
      errorMessage: '',
    };

    appendItemLogs_(spreadsheet, runId, executedAt, matchedItems, newItems, slackStatus);
    syncPublicItemLog_(spreadsheet);
    appendRunLog_(spreadsheet, runId, executedAt, runSummary);
  } catch (error) {
    runSummary.notificationStatus = 'error';
    runSummary.errorMessage = error && error.message ? error.message : String(error);

    try {
      if (spreadsheet) {
        appendItemLogs_(spreadsheet, runId, executedAt, matchedItems, newItems, 'error');
        syncPublicItemLog_(spreadsheet);
        appendRunLog_(spreadsheet, runId, executedAt, runSummary);
      }
    } catch (logError) {
      console.error(logError);
    }

    throw error;
  } finally {
    lock.releaseLock();
  }
}

function ensureDailyTrigger() {
  ensureDailyTrigger_();
}

function setupSheets() {
  var spreadsheet = getSpreadsheet_();
  ensureSheets_(spreadsheet);
  ensureConfigRows_(spreadsheet);
  ensurePublicSheet_();
}

function fetchRssItems_(rssUrl) {
  var response = UrlFetchApp.fetch(rssUrl, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Google-Apps-Script',
    },
  });

  var statusCode = response.getResponseCode();
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('RSSの取得に失敗しました。status=' + statusCode);
  }

  var document = XmlService.parse(response.getContentText());
  var root = document.getRootElement();
  var channel = root.getChild('channel');
  if (!channel) {
    throw new Error('RSSのchannel要素が見つかりません。');
  }

  return channel.getChildren('item').map(function(item) {
    var title = getChildText_(item, 'title');
    var link = normalizeJnet21Link_(getChildText_(item, 'link'));
    var description = sanitizeDescription_(htmlToText_(getChildText_(item, 'description')));
    var guid = getChildText_(item, 'guid') || link || title;
    var publishedAt = getDcDateText_(item) || getChildText_(item, 'pubDate');
    var detectedRegions = getCoverageLabels_(item);

    return {
      externalId: buildExternalId_(guid, title, publishedAt),
      title: title,
      link: link,
      description: description,
      publishedAt: publishedAt,
      detectedRegions: detectedRegions,
      rawRegionText: detectedRegions.join(','),
    };
  });
}

function filterItemsByRegions_(items, targetRegions) {
  var normalizedTargets = targetRegions.map(normalizePrefectureName_);

  return items.filter(function(item) {
    if (item.detectedRegions.indexOf('全国') !== -1) {
      item.matchedRegions = ['全国'];
      return true;
    }

    var matched = item.detectedRegions.filter(function(region) {
      return normalizedTargets.indexOf(normalizePrefectureName_(region)) !== -1;
    });

    item.matchedRegions = matched;
    return matched.length > 0;
  });
}

function getCoverageLabels_(item) {
  var coverage = item.getChild('coverage', XML_NAMESPACES.dc);
  if (!coverage) {
    return [];
  }

  var label = coverage.getChild('label', XML_NAMESPACES.rdf);
  if (!label) {
    return [];
  }

  return unique_([label.getText().trim()]).filter(Boolean);
}

function buildSlackMessage_(newItems, targetRegions, executedAt) {
  var timestamp = Utilities.formatDate(
    executedAt,
    Session.getScriptTimeZone() || DEFAULT_CONFIG.TIMEZONE,
    'yyyy/MM/dd HH:mm:ss'
  );
  var targetDate = Utilities.formatDate(
    executedAt,
    Session.getScriptTimeZone() || DEFAULT_CONFIG.TIMEZONE,
    'yyyy/MM/dd'
  );
  var regionLabel = targetRegions.join('、');
  var publicSpreadsheetUrl = getPublicSpreadsheetUrl_();

  var lines = [
    '*J-Net21 当日分の新着補助金・助成金・融資情報*',
    '対象日: ' + targetDate,
    '対象地域: ' + regionLabel,
    '件数: ' + newItems.length,
    '',
  ];

  newItems.forEach(function(item, index) {
    lines.push((index + 1) + '. <' + item.link + '|' + escapeSlackText_(item.title) + '>');
    lines.push('');
  });

  lines.push('確認時刻: ' + timestamp);
  lines.push('過去の配信履歴は下記リンク先にまとめてあります。');
  lines.push(publicSpreadsheetUrl);
  return lines.join('\n');
}

function postToSlack_(messageText) {
  var webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!webhookUrl) {
    throw new Error('Script Properties に SLACK_WEBHOOK_URL を設定してください。');
  }

  var response = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      text: messageText,
    }),
    muteHttpExceptions: true,
  });

  var statusCode = response.getResponseCode();
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Slack送信に失敗しました。status=' + statusCode + ' body=' + response.getContentText());
  }

  return 'sent';
}

function loadConfig_(spreadsheet) {
  var configSheet = spreadsheet.getSheetByName(DEFAULT_CONFIG.SHEET_CONFIG);
  var values = configSheet.getDataRange().getValues();
  var configMap = {};

  values.slice(1).forEach(function(row) {
    var key = row[0];
    var value = row[1];
    if (key) {
      configMap[key] = String(value || '').trim();
    }
  });

  var targetRegions = (configMap.TARGET_REGIONS || DEFAULT_CONFIG.TARGET_REGIONS)
    .split(',')
    .map(function(region) {
      return region.trim();
    })
    .filter(Boolean);

  return {
    RSS_URL: configMap.RSS_URL || DEFAULT_CONFIG.RSS_URL,
    targetRegions: targetRegions,
    DAILY_TRIGGER_HOUR: configMap.DAILY_TRIGGER_HOUR || DEFAULT_CONFIG.DAILY_TRIGGER_HOUR,
  };
}

function loadStateMap_(spreadsheet) {
  var stateSheet = spreadsheet.getSheetByName(DEFAULT_CONFIG.SHEET_STATE);
  var values = stateSheet.getDataRange().getValues();
  var stateMap = {};

  values.slice(1).forEach(function(row, index) {
    var externalId = row[0];
    if (!externalId) {
      return;
    }

    stateMap[externalId] = {
      rowNumber: index + 2,
      status: row[5],
    };
  });

  return stateMap;
}

function findNewItems_(matchedItems, stateMap) {
  return matchedItems.filter(function(item) {
    var existing = stateMap[item.externalId];
    return !existing || existing.status !== 'sent';
  });
}

function syncState_(spreadsheet, matchedItems, stateMap, executedAt, slackStatus) {
  var stateSheet = spreadsheet.getSheetByName(DEFAULT_CONFIG.SHEET_STATE);
  var status = slackStatus === 'sent' ? 'sent' : 'pending';

  matchedItems.forEach(function(item) {
    var existing = stateMap[item.externalId];
    var rowValues = [
      item.externalId,
      item.title,
      item.link,
      item.publishedAt,
      (item.matchedRegions || []).join(','),
      existing && existing.status === 'sent' ? 'sent' : status,
      existing ? '' : executedAt,
      executedAt,
      slackStatus === 'sent'
        ? executedAt
        : existing && existing.status === 'sent'
          ? stateSheet.getRange(existing.rowNumber, 9).getValue()
          : '',
    ];

    if (!existing) {
      stateSheet.appendRow(rowValues);
      return;
    }

    stateSheet.getRange(existing.rowNumber, 1, 1, rowValues.length).setValues([[
      item.externalId,
      item.title,
      item.link,
      item.publishedAt,
      (item.matchedRegions || []).join(','),
      existing.status === 'sent' ? 'sent' : status,
      stateSheet.getRange(existing.rowNumber, 7).getValue() || executedAt,
      executedAt,
      slackStatus === 'sent'
        ? executedAt
        : stateSheet.getRange(existing.rowNumber, 9).getValue(),
    ]]);
  });
}

function appendRunLog_(spreadsheet, runId, executedAt, summary) {
  var sheet = spreadsheet.getSheetByName(DEFAULT_CONFIG.SHEET_RUN_LOG);
  sheet.appendRow([
    runId,
    executedAt,
    summary.rssTotalCount,
    summary.regionMatchedCount,
    summary.newItemCount,
    summary.notificationStatus,
    summary.messageText,
    summary.errorMessage,
  ]);
}

function appendItemLogs_(spreadsheet, runId, executedAt, matchedItems, newItems, slackStatus) {
  var sheet = spreadsheet.getSheetByName(DEFAULT_CONFIG.SHEET_ITEM_LOG);
  var newItemIds = newItems.map(function(item) {
    return item.externalId;
  });
  var rows = [];

  if (matchedItems.length === 0) {
    rows.push([
      runId,
      executedAt,
      '',
      '',
      '',
      '',
      '',
      false,
      slackStatus === 'sent',
      slackStatus,
      '',
    ]);
    prependRows_(sheet, rows);
    return;
  }

  matchedItems.forEach(function(item) {
    var isNew = newItemIds.indexOf(item.externalId) !== -1;
    rows.push([
      runId,
      executedAt,
      item.externalId,
      item.title,
      item.link,
      item.publishedAt,
      (item.matchedRegions || []).join(','),
      isNew,
      slackStatus === 'sent',
      slackStatus,
      item.rawRegionText,
    ]);
  });

  prependRows_(sheet, rows);
}

function syncPublicItemLog_(spreadsheet) {
  var sourceSheet = spreadsheet.getSheetByName(DEFAULT_CONFIG.SHEET_ITEM_LOG);
  var publicSpreadsheet = getPublicSpreadsheet_();
  var targetSheet = ensurePublicSheet_(publicSpreadsheet);
  var values = sourceSheet.getDataRange().getValues();
  var headers = SHEET_DEFINITIONS.public_item_log;
  var rows = values.slice(1)
    .map(function(row) {
      return [
        row[1],
        row[3],
        row[4],
        row[5],
        row[6],
      ];
    })
    .filter(function(row) {
      return row[1] || row[2] || row[3] || row[4];
    })
    .sort(function(a, b) {
      var timeA = new Date(a[0]).getTime();
      var timeB = new Date(b[0]).getTime();
      return timeB - timeA;
    });

  targetSheet.clearContents();
  targetSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  targetSheet.setFrozenRows(1);

  if (rows.length > 0) {
    targetSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function prependRows_(sheet, rows) {
  if (!rows || rows.length === 0) {
    return;
  }

  sheet.insertRowsBefore(2, rows.length);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function ensureSheets_(spreadsheet) {
  Object.keys(SHEET_DEFINITIONS).forEach(function(sheetName) {
    if (sheetName === DEFAULT_CONFIG.SHEET_PUBLIC_ITEM_LOG) {
      return;
    }

    var headers = SHEET_DEFINITIONS[sheetName];
    var sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  });
}

function ensureConfigRows_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(DEFAULT_CONFIG.SHEET_CONFIG);
  var values = sheet.getDataRange().getValues();
  var existingKeys = values.slice(1).map(function(row) {
    return row[0];
  });
  var rowsToAppend = [];

  CONFIG_KEYS.forEach(function(key) {
    if (existingKeys.indexOf(key) !== -1) {
      return;
    }

    rowsToAppend.push([
      key,
      DEFAULT_CONFIG[key],
      describeConfigKey_(key),
    ]);
  });

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);
  }
}

function ensurePublicSheet_(publicSpreadsheet) {
  var spreadsheet = publicSpreadsheet || getPublicSpreadsheet_();
  var headers = SHEET_DEFINITIONS.public_item_log;
  var sheet = spreadsheet.getSheetByName(DEFAULT_CONFIG.SHEET_PUBLIC_ITEM_LOG);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(DEFAULT_CONFIG.SHEET_PUBLIC_ITEM_LOG);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function ensureDailyTrigger_() {
  var spreadsheet = getSpreadsheet_();
  ensureSheets_(spreadsheet);
  ensureConfigRows_(spreadsheet);
  var config = loadConfig_(spreadsheet);
  var hour = Number(config.DAILY_TRIGGER_HOUR || DEFAULT_CONFIG.DAILY_TRIGGER_HOUR);

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runDaily') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  [
    ScriptApp.WeekDay.MONDAY,
    ScriptApp.WeekDay.TUESDAY,
    ScriptApp.WeekDay.WEDNESDAY,
    ScriptApp.WeekDay.THURSDAY,
    ScriptApp.WeekDay.FRIDAY,
  ].forEach(function(weekDay) {
    ScriptApp.newTrigger('runDaily')
      .timeBased()
      .onWeekDay(weekDay)
      .atHour(hour)
      .create();
  });
}

function describeConfigKey_(key) {
  var descriptions = {
    RSS_URL: 'J-Net21のRSS URL',
    TARGET_REGIONS: 'カンマ区切りの対象都道府県',
    DAILY_TRIGGER_HOUR: '日次実行の時刻（0-23）',
  };

  return descriptions[key] || '';
}

function getSpreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }

  var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }

  throw new Error('Active Spreadsheet が見つかりません。Script Properties に SPREADSHEET_ID を設定するか、スプレッドシート紐づきで実行してください。');
}

function getPublicSpreadsheet_() {
  var spreadsheetId = PropertiesService.getScriptProperties().getProperty('PUBLIC_SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('Script Properties に PUBLIC_SPREADSHEET_ID を設定してください。');
  }

  return SpreadsheetApp.openById(spreadsheetId);
}

function getPublicSpreadsheetUrl_() {
  var spreadsheetUrl = PropertiesService.getScriptProperties().getProperty('PUBLIC_SPREADSHEET_URL');
  if (!spreadsheetUrl) {
    throw new Error('Script Properties に PUBLIC_SPREADSHEET_URL を設定してください。');
  }

  return spreadsheetUrl;
}

function getChildText_(element, childName) {
  var child = element.getChild(childName);
  return child ? child.getText() : '';
}

function getDcDateText_(element) {
  var child = element.getChild('date', XML_NAMESPACES.dc);
  return child ? child.getText() : '';
}

function htmlToText_(html) {
  return String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+\n/g, '\n')
    .replace(/\n\s+/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .trim();
}

function sanitizeDescription_(text) {
  return String(text || '')
    .replace(/(?:申請|申込|応募)方法ほか詳細情報は、?「詳細情報を見る」からご確認(?:ください|下さい)。?/g, '')
    .replace(/\s+\n/g, '\n')
    .replace(/\n\s+/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .trim();
}

function buildExternalId_(guid, title, publishedAt) {
  var base = [guid || '', title || '', publishedAt || ''].join('||');
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, base);
  return digest.map(function(byte) {
    var value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('');
}

function normalizeJnet21Link_(url) {
  var value = String(url || '').trim();
  if (!value) {
    return value;
  }

  return value.replace(
    'https://j-net21.smrj.go.jp/snavi/articles/',
    'https://j-net21.smrj.go.jp/snavi2/articles/'
  );
}

function normalizeText_(text) {
  return String(text || '')
    .replace(/\s+/g, '')
    .replace(/[()（）【】［］\[\]「」『』]/g, '')
    .toLowerCase();
}

function prefectureVariants_(prefecture) {
  var variants = [prefecture];
  if (prefecture === '東京都') {
    variants.push('東京');
  } else if (prefecture.endsWith('県')) {
    variants.push(prefecture.replace(/県$/, ''));
  } else if (prefecture.endsWith('府')) {
    variants.push(prefecture.replace(/府$/, ''));
  } else if (prefecture === '北海道') {
    variants.push('道内');
  }
  return unique_(variants);
}

function normalizePrefectureName_(value) {
  var text = String(value || '').trim();
  var map = {
    東京: '東京都',
    神奈川: '神奈川県',
    埼玉: '埼玉県',
    千葉: '千葉県',
  };
  return map[text] || text;
}

function unique_(items) {
  var result = [];
  items.forEach(function(item) {
    if (result.indexOf(item) === -1) {
      result.push(item);
    }
  });
  return result;
}

function escapeSlackText_(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
