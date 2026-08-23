/**
 * Unified Merchant Funnel Product Control Extended V2
 *
 * Google Ads Script entry point. Edit only SPREADSHEET_URL before launch.
 *
 * V2 goals:
 * - keep the V1 user-facing sheets where possible: Settings, Products,
 *   ProductTypes, ProductDiagnostics, Dashboard, DashboardData,
 *   QuarantineRegistry, QuarantineLog;
 * - add service sheets RunState, RunLog, MerchantSnapshot_Work,
 *   AdsStatsSnapshot_Work, ProductsDraft_Work, ProductsPublish_Work;
 * - split the heavy flow across hourly launches using RunState;
 * - write STARTED before heavy work, then DONE / PARTIAL / FAILED after work;
 * - never partially rebuild the public Products feed. ProductsDraft_Work is
 *   built first, verified, then published into Products.
 */


var SPREADSHEET_URL = 'PASTE_SPREADSHEET_URL_HERE';


var SETTINGS_SHEET = 'Settings';
var PRODUCTS_SHEET = 'Products';
var PRODUCT_TYPES_SHEET = 'ProductTypes';
var PRODUCT_DIAGNOSTICS_SHEET = 'ProductDiagnostics';
var DASHBOARD_SHEET = 'Dashboard';
var DASHBOARD_DATA_SHEET = 'DashboardData';
var QUARANTINE_REGISTRY_SHEET = 'QuarantineRegistry';
var QUARANTINE_LOG_SHEET = 'QuarantineLog';
var RUN_STATE_SHEET = 'RunState';
var RUN_LOG_SHEET = 'RunLog';
var MERCHANT_SNAPSHOT_SHEET = 'MerchantSnapshot_Work';
var ADS_STATS_SNAPSHOT_SHEET = 'AdsStatsSnapshot_Work';
var PRODUCTS_DRAFT_SHEET = 'ProductsDraft_Work';
var PRODUCTS_PUBLISH_WORK_SHEET = 'ProductsPublish_Work';


var HEADER_BACKGROUND = '#c9daf8';
var SECTION_BACKGROUND = '#e6e6e6';
var MANUAL_BACKGROUND = '#fff2cc';
var REQUIRED_SETTING_BACKGROUND = '#f4cccc';


var STAGES = [
  'MERCHANT_SNAPSHOT',
  'ADS_STATS_SNAPSHOT',
  'PRODUCT_TYPES_BUILD',
  'QUARANTINE_UPDATE',
  'PRODUCTS_DRAFT_BUILD',
  'PUBLISH_PRODUCTS',
  'PRODUCT_DIAGNOSTICS_BUILD',
  'DASHBOARD_BUILD',
  'COMPLETE'
];


var START_TIME = new Date().getTime();


function main() {
  var ctx = null;
  var startup = null;
  try {
    if (!SPREADSHEET_URL || SPREADSHEET_URL === 'PASTE_SPREADSHEET_URL_HERE') {
      throw new Error('Set SPREADSHEET_URL at the top of the script.');
    }


    var ss = SpreadsheetApp.openByUrl(SPREADSHEET_URL);
    startup = startStartupRun_(ss);
    var sheets = ensureSheets_(ss);
    ensureProductsFirst_(ss, sheets.products);
    var settings = readSettings_(sheets.settings);
    ensureSpreadsheetLocale_(ss, settings);
    finishStartupRun_(startup, settings, 'Startup complete. Dispatching pipeline stage.');
    ctx = startRun_(ss, sheets, settings, startup.runId);
    while (ctx && !shouldStopSoon_(settings)) {
      var result = dispatchStage_(ctx);
      finishRun_(ctx, result.status, result.message);
      if (ctx.stage === 'LOCKED' || result.status !== 'DONE') break;
      var state;
      try {
        state = readState_(sheets.runState);
      } catch (stateErr) {
        Logger.log('Post-stage state read failed after ' + ctx.stage + ': ' + String(stateErr));
        break;
      }
      if (!state.stage) break;
      if (state.stage === 'COMPLETE') break;
      ctx = startNextStageInSameRun_(ctx, state.stage);
    }
  } catch (err) {
    if (ctx) failRun_(ctx, err);
    else failStartupRun_(startup, err);
    throw err;
  }
}


function ensureSheets_(ss) {
  var names = {};
  names.settings = SETTINGS_SHEET;
  names.products = PRODUCTS_SHEET;
  names.productTypes = PRODUCT_TYPES_SHEET;
  names.productDiagnostics = PRODUCT_DIAGNOSTICS_SHEET;
  names.dashboard = DASHBOARD_SHEET;
  names.dashboardData = DASHBOARD_DATA_SHEET;
  names.quarantineRegistry = QUARANTINE_REGISTRY_SHEET;
  names.quarantineLog = QUARANTINE_LOG_SHEET;
  names.runState = RUN_STATE_SHEET;
  names.runLog = RUN_LOG_SHEET;
  names.merchantSnapshot = MERCHANT_SNAPSHOT_SHEET;
  names.adsStatsSnapshot = ADS_STATS_SNAPSHOT_SHEET;
  names.productsDraft = PRODUCTS_DRAFT_SHEET;
  names.productsPublish = PRODUCTS_PUBLISH_WORK_SHEET;


  var out = {};
  for (var key in names) {
    out[key] = ss.getSheetByName(names[key]) || ss.insertSheet(names[key]);
  }
  ensureSettingsTemplate_(out.settings);
  ensureRunStateHeader_(out.runState);
  ensureRunLogHeader_(out.runLog);
  ensureMerchantSnapshotHeader_(out.merchantSnapshot);
  ensureAdsStatsHeader_(out.adsStatsSnapshot);
  ensureProductsDraftHeader_(out.products, readSettings_(out.settings));
  ensureProductsDraftHeader_(out.productsDraft, readSettings_(out.settings));
  ensureProductsDraftHeader_(out.productsPublish, readSettings_(out.settings));
  ensureQuarantineHeaders_(out.quarantineRegistry, out.quarantineLog);
  compactManagedSheetGrids_(out, readSettings_(out.settings));
  return out;
}


function ensureProductsFirst_(ss, productsSheet) {
  if (productsSheet.getIndex() !== 1) {
    ss.setActiveSheet(productsSheet);
    ss.moveActiveSheet(1);
  }
}


function ensureSettingsTemplate_(sheet) {
  var headers = ['setting', 'value', 'description'];
  var existing = readSettingsMap_(sheet);
  var rows = [
    ['[Required]', '', ''],
    ['merchant_id', '', 'ID облікового запису Merchant Center.'],
    ['developer_email', 'bazhinalex05@gmail.com', 'Email розробника для реєстрації Merchant API.'],
    ['auto_register_gcp_project', 'TRUE', 'TRUE автоматично реєструє GCP-проєкт Google Ads Scripts у Merchant API, якщо реєстрації ще немає.'],
    ['wait_after_gcp_registration_seconds', '0', 'Скільки секунд чекати після реєстрації Merchant API. 0 означає не чекати й повторити на наступному запуску.'],


    ['[Pipeline / run control]', '', ''],
    ['pipeline_test_mode', 'FALSE', 'TRUE вмикає тестові ліміти обробки. FALSE вмикає робочі ліміти.'],
    ['force_restart_pipeline', 'FALSE', 'TRUE перезапускає водоспад з етапу Merchant snapshot на наступному запуску.'],
    ['max_run_minutes', '25', 'Цільовий ліміт часу запуску в хвилинах для перевірок між безпечними порціями.'],
    ['run_lock_minutes', '40', 'Через скільки хвилин RUNNING-етап вважається завислим і доступним для відновлення.'],
    ['merchant_snapshot_test_row_limit', '0', 'Тестовий ліміт рядків Merchant snapshot. 0 означає без обмеження.'],
    ['max_merchant_pages_per_run', '2', 'Скільки сторінок Merchant API обробляти за запуск у тестовому режимі.'],
    ['max_merchant_pages_per_run_prod', '50', 'Скільки сторінок Merchant API обробляти за запуск у робочому режимі.'],
    ['ads_periods_per_run', '1', 'Скільки періодів статистики Google Ads забирати за один запуск.'],
    ['products_chunk_size', '5000', 'Скільки рядків ProductsDraft обробляти за запуск у тестовому режимі.'],
    ['products_chunk_size_prod', '250000', 'Скільки рядків ProductsDraft обробляти за запуск у робочому режимі.'],


    ['[Merchant API / filters]', '', ''],
    ['merchant_api_page_size', '1000', 'Скільки товарів запитувати з Merchant API за одну сторінку.'],
    ['merchant_api_retry_count', '5', 'Кількість повторних спроб після помилки Merchant API.'],
    ['merchant_api_retry_sleep_seconds', '10', 'Базова пауза в секундах між повторними спробами Merchant API.'],
    ['merchant_data_source_id_filter', '', 'ID джерел даних Merchant через кому, які потрібно включити. Порожньо означає включити всі.'],
    ['merchant_feed_label_filter', '', 'Feed labels Merchant через кому, які потрібно включити. Порожньо означає включити всі.'],
    ['merchant_content_language_filter', '', 'Мови контенту через кому, які потрібно включити. Порожньо означає включити всі.'],


    ['[Products output]', '', ''],
    ['funnel_stage_output_attribute', 'custom_label_0', 'Колонка custom label, у яку записується етап воронки.'],
    ['benchmark_output_attribute', 'custom_label_3', 'Додаткова колонка custom label у Products для benchmark / priority. Порожньо вимикає цю колонку.'],
    ['benchmark_label_field', 'custom_label_4', 'Вихідний custom label з Merchant API, з якого береться benchmark / priority.'],
    ['enable_products_write', 'TRUE', 'TRUE публікує Products після завершення ProductsDraft. FALSE залишає результат тільки у ProductsDraft_Work.'],


    ['[Modules]', '', ''],
    ['enable_product_type_filter', 'TRUE', 'TRUE застосовує правила ProductTypes.enabled для включення або виключення товарів.'],
    ['enable_funnel_builder', 'TRUE', 'TRUE записує етап воронки.'],
    ['enable_quarantine', 'TRUE', 'TRUE застосовує правила карантину та активні карантинні виключення.'],
    ['enable_product_diagnostics', 'TRUE', 'TRUE будує зведений лист ProductDiagnostics.'],
    ['enable_dashboard_data', 'TRUE', 'TRUE будує лист DashboardData.'],
    ['enable_dashboard', 'TRUE', 'TRUE будує лист Dashboard.'],
    ['enable_external_product_type_feed', 'FALSE', 'TRUE доповнює product_type із зовнішньої Google-таблиці.'],
    ['enable_previous_state_read', 'TRUE', 'TRUE читає збережені ручні значення перед перебудовою керованих листів.'],
    ['enable_managed_sheet_formatting', 'FALSE', 'TRUE повторно застосовує службове форматування. Для великих запусків тримати FALSE.'],


    ['[External product_type source]', '', ''],
    ['external_product_type_spreadsheet_url', '', 'URL Google-таблиці із зовнішніми значеннями product_type.'],
    ['external_product_type_sheet_name', '', 'Назва листа у зовнішньому джерелі product_type. Порожньо означає перший лист.'],
    ['external_product_type_id_column', 'id', 'Назва заголовка з ID товару у зовнішньому джерелі.'],
    ['external_product_type_value_column', 'product_type', 'Назва заголовка з product_type у зовнішньому джерелі.'],
    ['external_product_type_override', 'FALSE', 'TRUE дозволяє зовнішньому джерелу замінювати product_type з Merchant. FALSE заповнює тільки порожні значення.'],


    ['[Funnel and quarantine rules]', '', ''],
    ['funnel_period_days', '14', 'Період аналізу для етапів воронки в днях.'],
    ['funnel_exclude_last_days', '0', 'Скільки останніх днів виключити зі статистики воронки. 0 відповідає 14-денному вікну V1.'],
    ['no_sales_click_threshold', '60', 'Правило карантину NO_SALES: поріг кліків за відсутності конверсій.'],
    ['no_sales_period_days', '30', 'Період аналізу для правила NO_SALES у днях.'],
    ['spend_over_margin_share', '0.35', 'Правило SPEND_OVER_MARGIN: поріг витрат відносно ціни товару.'],
    ['spend_over_margin_period_days', '30', 'Період аналізу для правила SPEND_OVER_MARGIN у днях.'],
    ['expensive_click_cpc_threshold', '3', 'Правило EXPENSIVE_CLICK: поріг середньої ціни кліку.'],
    ['expensive_click_period_days', '30', 'Період аналізу для правила EXPENSIVE_CLICK у днях.'],
    ['quarantine_days', '14', 'Скільки днів товар залишається виключеним після нового спрацювання карантину.'],


    ['[Service]', '', ''],
    ['run_log_max_rows', '200', 'Максимальна кількість рядків у RunLog. Старі рядки видаляються.'],
    ['run_log_max_message_chars', '800', 'Максимальна кількість символів в одному повідомленні RunLog.'],
    ['spreadsheet_locale', 'en_US', 'Локаль таблиці для стабільних дат, чисел і формул.']
  ];
  var values = [headers];
  for (var i = 0; i < rows.length; i++) {
    var key = rows[i][0];
    if (key && !isSettingsSectionKey_(key) && existing.hasOwnProperty(key)) {
      rows[i][1] = existing[key];
    }
    values.push(rows[i]);
  }
  sheet.clear();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  sheet.getRange(1, 1, values.length, 3).setValues(values);
  formatSettingsSheet_(sheet, values);
  fillBlankSettingValue_(sheet, 'developer_email', 'bazhinalex05@gmail.com');
}


function formatSettingsSheet_(sheet, values) {
  var rowCount = values.length;
  if (!rowCount) return;
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 560);
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), rowCount), 3)
    .setFontWeight('normal')
    .setBackground(null)
    .setWrap(true)
    .setVerticalAlignment('middle');
  sheet.getRange(1, 1, 1, 3)
    .setFontWeight('bold')
    .setBackground(HEADER_BACKGROUND);
  if (rowCount > 1) {
    sheet.getRange(2, 2, rowCount - 1, 1).setBackground(MANUAL_BACKGROUND);
  }
  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || '').trim();
    var row = i + 1;
    if (isSettingsSectionKey_(key)) {
      sheet.getRange(row, 1, 1, 3)
        .setFontWeight('bold')
        .setBackground(SECTION_BACKGROUND);
    }
    if (key === 'merchant_id' || key === 'developer_email') {
      sheet.getRange(row, 1, 1, 3).setBackground(REQUIRED_SETTING_BACKGROUND);
    }
  }
  applySettingsCheckboxes_(sheet, values);
  applySettingsDropdowns_(sheet, values);
}


function applySettingsCheckboxes_(sheet, values) {
  if (values.length < 2) return;
  sheet.getRange(2, 2, values.length - 1, 1).clearDataValidations();
  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || '').trim();
    if (!isBooleanSettingKey_(key)) continue;
    var row = i + 1;
    sheet.getRange(row, 2).setValue(bool_(values[i][1], false));
    sheet.getRange(row, 2).insertCheckboxes();
  }
}


function isBooleanSettingKey_(key) {
  var bools = {
    auto_register_gcp_project: true,
    pipeline_test_mode: true,
    force_restart_pipeline: true,
    enable_benchmark_grouping: true,
    enable_products_write: true,
    enable_product_type_filter: true,
    enable_funnel_builder: true,
    enable_quarantine: true,
    enable_product_diagnostics: true,
    enable_dashboard_data: true,
    enable_dashboard: true,
    enable_external_product_type_feed: true,
    enable_previous_state_read: true,
    enable_managed_sheet_formatting: true,
    external_product_type_override: true
  };
  return bools.hasOwnProperty(String(key || '').trim());
}


function applySettingsDropdowns_(sheet, values) {
  var customLabels = ['custom_label_0', 'custom_label_1', 'custom_label_2', 'custom_label_3', 'custom_label_4'];
  var benchmarkSources = customLabels.concat(['product_type', 'product_type_l1', 'product_type_l2', 'product_type_l3', 'product_type_l4', 'product_type_l5', 'brand', 'title']);
  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || '').trim();
    var row = i + 1;
    if (key === 'funnel_stage_output_attribute') {
      sheet.getRange(row, 2).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(customLabels, true)
          .setAllowInvalid(false)
          .build()
      );
    } else if (key === 'benchmark_output_attribute') {
      sheet.getRange(row, 2).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(customLabels, true)
          .setAllowInvalid(true)
          .build()
      );
    } else if (key === 'benchmark_label_field') {
      sheet.getRange(row, 2).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(benchmarkSources, true)
          .setAllowInvalid(false)
          .build()
      );
    }
  }
}


function readSettings_(sheet) {
  var raw = readSettingsMap_(sheet);
  return {
    merchantId: String(raw.merchant_id || '').trim(),
    autoRegisterGcpProject: bool_(raw.auto_register_gcp_project, true),
    developerEmail: String(raw.developer_email || '').trim(),
    waitAfterGcpRegistrationSeconds: num_(raw.wait_after_gcp_registration_seconds, 0),
    testMode: bool_(raw.pipeline_test_mode, false),
    maxRunMinutes: num_(raw.max_run_minutes, 25),
    runLockMinutes: num_(raw.run_lock_minutes, 40),
    merchantPageSize: num_(raw.merchant_api_page_size, 1000),
    merchantApiRetryCount: num_(raw.merchant_api_retry_count, 5),
    merchantApiRetrySleepSeconds: num_(raw.merchant_api_retry_sleep_seconds, 10),
    merchantSnapshotTestRowLimit: num_(raw.merchant_snapshot_test_row_limit, 0),
    runLogMaxRows: num_(raw.run_log_max_rows, 200),
    runLogMaxMessageChars: num_(raw.run_log_max_message_chars, 800),
    spreadsheetLocale: String(raw.spreadsheet_locale || 'en_US').trim(),
    merchantDataSourceIdFilter: listSetting_(raw.merchant_data_source_id_filter),
    merchantFeedLabelFilter: listSetting_(raw.merchant_feed_label_filter),
    merchantContentLanguageFilter: listSetting_(raw.merchant_content_language_filter),
    maxMerchantPagesPerRun: bool_(raw.pipeline_test_mode, false) ? num_(raw.max_merchant_pages_per_run, 2) : num_(raw.max_merchant_pages_per_run_prod, 50),
    adsPeriodsPerRun: num_(raw.ads_periods_per_run, 1),
    productsChunkSize: bool_(raw.pipeline_test_mode, false) ? num_(raw.products_chunk_size, 5000) : num_(raw.products_chunk_size_prod, 250000),
    funnelStageOutputAttribute: String(raw.funnel_stage_output_attribute || 'custom_label_0').trim(),
    benchmarkOutputAttribute: String(raw.benchmark_output_attribute || 'custom_label_3').trim(),
    benchmarkLabelField: String(raw.benchmark_label_field || 'custom_label_4').trim(),
    enableExternalProductTypeFeed: bool_(raw.enable_external_product_type_feed, false),
    externalProductTypeSpreadsheetUrl: String(raw.external_product_type_spreadsheet_url || '').trim(),
    externalProductTypeSheetName: String(raw.external_product_type_sheet_name || '').trim(),
    externalProductTypeIdColumn: String(raw.external_product_type_id_column || 'id').trim(),
    externalProductTypeValueColumn: String(raw.external_product_type_value_column || 'product_type').trim(),
    externalProductTypeOverride: bool_(raw.external_product_type_override, false),
    funnelPeriodDays: num_(raw.funnel_period_days, 14),
    funnelExcludeLastDays: num_(raw.funnel_exclude_last_days, 0),
    noSalesClickThreshold: num_(raw.no_sales_click_threshold, 60),
    noSalesPeriodDays: num_(raw.no_sales_period_days, 30),
    spendOverMarginShare: num_(raw.spend_over_margin_share, 0.35),
    spendOverMarginPeriodDays: num_(raw.spend_over_margin_period_days, 30),
    expensiveClickCpcThreshold: num_(raw.expensive_click_cpc_threshold, 3),
    expensiveClickPeriodDays: num_(raw.expensive_click_period_days, 30),
    quarantineDays: num_(raw.quarantine_days, 14),
    enableProductTypeFilter: bool_(raw.enable_product_type_filter, true),
    enableFunnelBuilder: bool_(raw.enable_funnel_builder, true),
    enableQuarantine: bool_(raw.enable_quarantine, true),
    enableProductsWrite: bool_(raw.enable_products_write, true),
    enableProductDiagnostics: bool_(raw.enable_product_diagnostics, false),
    enableDashboardData: bool_(raw.enable_dashboard_data, false),
    enableDashboard: bool_(raw.enable_dashboard, false),
    enablePreviousStateRead: bool_(raw.enable_previous_state_read, true),
    enableManagedSheetFormatting: bool_(raw.enable_managed_sheet_formatting, false),
    forceRestartPipeline: bool_(raw.force_restart_pipeline, false)
  };
}


function readSettingsMap_(sheet) {
  var map = {};
  if (sheet.getLastRow() < 2) return map;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(2, sheet.getLastColumn())).getValues();
  for (var i = 0; i < values.length; i++) {
    var key = String(values[i][0] || '').trim();
    if (key && !isSettingsSectionKey_(key)) map[key] = values[i][1];
  }
  return map;
}


function isSettingsSectionKey_(key) {
  key = String(key || '').trim();
  return key.length > 2 && key.charAt(0) === '[' && key.charAt(key.length - 1) === ']';
}


function fillBlankSettingValue_(sheet, key, value) {
  if (sheet.getLastRow() < 2) return;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0] || '').trim() === key && String(rows[i][1] || '').trim() === '') {
      sheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
}


function startRun_(ss, sheets, settings, existingRunId) {
  var state = readState_(sheets.runState);
  var now = new Date();
  var runId = existingRunId || Utilities.formatDate(now, tz_(), 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 100000);


  if (settings.forceRestartPipeline || !state.stage || state.stage === 'COMPLETE') {
    state = resetPipelineState_(sheets, settings);
  }


  if (state.stage_status === 'RUNNING' && !isLockExpired_(state.lock_until)) {
    appendRunLog_(sheets.runLog, runId, state.stage, 'SKIPPED_LOCKED', 'Previous run still owns lock until ' + state.lock_until);
    return {
      ss: ss,
      sheets: sheets,
      settings: settings,
      runId: runId,
      stage: 'LOCKED',
      logRow: 0
    };
  }


  if (state.stage_status === 'RUNNING' && isLockExpired_(state.lock_until)) {
    appendRunLog_(sheets.runLog, runId, state.stage, 'RECOVERED_STALE_RUN', 'Previous run did not finish before lock_until=' + state.lock_until);
  }


  var stage = state.stage || 'MERCHANT_SNAPSHOT';
  setStateValues_(sheets.runState, {
    pipeline_status: 'RUNNING',
    stage: stage,
    stage_status: 'RUNNING',
    current_run_id: runId,
    run_started_at: iso_(now),
    heartbeat_at: iso_(now),
    lock_until: iso_(new Date(now.getTime() + settings.runLockMinutes * 60000)),
    last_error_at: '',
    last_error_message: ''
  });
  var row = appendRunLog_(sheets.runLog, runId, stage, 'STARTED', 'Started stage ' + stage);
  return {
    ss: ss,
    sheets: sheets,
    settings: settings,
    runId: runId,
    stage: stage,
    logRow: row
  };
}


function startStartupRun_(ss) {
  var runLog = ss.getSheetByName(RUN_LOG_SHEET) || ss.insertSheet(RUN_LOG_SHEET);
  var runState = ss.getSheetByName(RUN_STATE_SHEET) || ss.insertSheet(RUN_STATE_SHEET);
  ensureRunLogHeader_(runLog);
  ensureRunStateHeader_(runState);
  var now = new Date();
  var runId = Utilities.formatDate(now, tz_(), 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 100000);
  var row = appendRunLog_(runLog, runId, 'STARTUP', 'STARTED', 'Opened spreadsheet. Preparing sheets before pipeline dispatch.');
  return {
    ss: ss,
    sheets: { runLog: runLog, runState: runState },
    runId: runId,
    stage: 'STARTUP',
    logRow: row
  };
}


function finishStartupRun_(startup, settings, message) {
  if (!startup) return;
  updateRunLog_(startup.sheets.runLog, startup.logRow, 'DONE', message, settings);
  trimRunLog_(startup.sheets.runLog, settings);
}


function failStartupRun_(startup, err) {
  if (!startup) return;
  var msg = truncateLogMessage_(err && err.message ? err.message : String(err), null);
  updateRunLog_(startup.sheets.runLog, startup.logRow, 'FAILED', msg, null);
  setStateValues_(startup.sheets.runState, {
    pipeline_status: 'FAILED',
    startup_status: 'FAILED',
    heartbeat_at: iso_(new Date()),
    lock_until: '',
    last_error_at: iso_(new Date()),
    last_error_message: msg
  });
}


function dispatchStage_(ctx) {
  if (ctx.stage === 'LOCKED') return { status: 'SKIPPED', message: 'Previous run is still locked.' };
  if (ctx.stage === 'MERCHANT_SNAPSHOT') return stageMerchantSnapshot_(ctx);
  if (ctx.stage === 'ADS_STATS_SNAPSHOT') return stageAdsStatsSnapshot_(ctx);
  if (ctx.stage === 'PRODUCT_TYPES_BUILD') return stageProductTypesBuild_(ctx);
  if (ctx.stage === 'QUARANTINE_UPDATE') return stageQuarantineUpdate_(ctx);
  if (ctx.stage === 'PRODUCTS_DRAFT_BUILD') return stageProductsDraftBuild_(ctx);
  if (ctx.stage === 'PUBLISH_PRODUCTS') return stagePublishProducts_(ctx);
  if (ctx.stage === 'PRODUCT_DIAGNOSTICS_BUILD') return stageProductDiagnostics_(ctx);
  if (ctx.stage === 'DASHBOARD_BUILD') return stageDashboard_(ctx);
  return { status: 'DONE', message: 'Pipeline is complete.' };
}


function startNextStageInSameRun_(ctx, stage) {
  var now = new Date();
  setStateValues_(ctx.sheets.runState, {
    pipeline_status: 'RUNNING',
    stage: stage,
    stage_status: 'RUNNING',
    heartbeat_at: iso_(now),
    lock_until: iso_(new Date(now.getTime() + ctx.settings.runLockMinutes * 60000)),
    last_error_at: '',
    last_error_message: ''
  });
  ctx.stage = stage;
  ctx.logRow = appendRunLog_(ctx.sheets.runLog, ctx.runId, stage, 'STARTED', 'Started stage ' + stage);
  return ctx;
}


function stageMerchantSnapshot_(ctx) {
  requireMerchantId_(ctx.settings);
  ensureMerchantApiServicesOn_();
  if (!ensureMerchantApiDeveloperRegistration_(ctx.settings)) {
    return { status: 'PARTIAL', message: 'Merchant API GCP project registered. Waiting for propagation; next launch will retry this stage.' };
  }
  var state = readState_(ctx.sheets.runState);
  var token = state.merchant_next_page_token || '';
  var written = num_(state.merchant_rows_written, 0);
  if (!token && written === 0) {
    clearBelowHeader_(ctx.sheets.merchantSnapshot);
  }


  var pageCount = 0;
  while (pageCount < ctx.settings.maxMerchantPagesPerRun && !shouldStopSoon_(ctx.settings)) {
    if (ctx.settings.merchantSnapshotTestRowLimit > 0 && written >= ctx.settings.merchantSnapshotTestRowLimit) break;
    logProgress_(ctx, 'MERCHANT_SNAPSHOT', 'CHECKPOINT', 'Before Merchant API page. Written=' + written + ', pageInRun=' + (pageCount + 1) + '/' + ctx.settings.maxMerchantPagesPerRun + ', hasToken=' + Boolean(token));
    var page = fetchMerchantProductsPage_(ctx.settings.merchantId, token, ctx.settings.merchantPageSize, ctx.settings);
    logProgress_(ctx, 'MERCHANT_SNAPSHOT', 'CHECKPOINT', 'After Merchant API page. Items=' + page.items.length + ', nextToken=' + Boolean(page.nextPageToken));
    var rows = [];
    var filteredOut = 0;
    for (var i = 0; i < page.items.length; i++) {
      if (merchantProductMatchesFilters_(page.items[i], ctx.settings)) rows.push(merchantRow_(page.items[i], ctx.settings));
      else filteredOut++;
    }
    if (rows.length) {
      logProgress_(ctx, 'MERCHANT_SNAPSHOT', 'CHECKPOINT', 'Before Sheets append. AppendRows=' + rows.length + ', targetStartRow=' + (ctx.sheets.merchantSnapshot.getLastRow() + 1));
      appendRows_(ctx.sheets.merchantSnapshot, rows);
      written += rows.length;
      logProgress_(ctx, 'MERCHANT_SNAPSHOT', 'CHECKPOINT', 'After Sheets append. Written=' + written);
    } else if (filteredOut) {
      logProgress_(ctx, 'MERCHANT_SNAPSHOT', 'CHECKPOINT', 'Merchant API page filtered out. FilteredRows=' + filteredOut);
    }
    token = page.nextPageToken || '';
    pageCount++;
    heartbeat_(ctx, 'merchant_rows_written', written, 'merchant_next_page_token', token);
    logProgress_(ctx, 'MERCHANT_SNAPSHOT', 'CHECKPOINT', 'Heartbeat saved. Written=' + written + ', nextToken=' + Boolean(token));
    if (shouldStopSoon_(ctx.settings)) break;
    if (ctx.settings.merchantSnapshotTestRowLimit > 0 && written >= ctx.settings.merchantSnapshotTestRowLimit) break;
    if (!token) break;
  }


  var truncatedForTest = ctx.settings.merchantSnapshotTestRowLimit > 0 && written >= ctx.settings.merchantSnapshotTestRowLimit && token;
  if (truncatedForTest) {
    Logger.log('Merchant snapshot test row limit reached: ' + written + '/' + ctx.settings.merchantSnapshotTestRowLimit + '. Advancing pipeline with partial Merchant sample.');
    token = '';
  }
  setStateValues_(ctx.sheets.runState, {
    merchant_rows_written: written,
    merchant_next_page_token: token,
    merchant_snapshot_truncated_at: truncatedForTest ? written : ''
  });
  if (token) return { status: 'PARTIAL', message: 'Merchant snapshot partial. Rows=' + written + ', next page remains.' };
  advanceStage_(ctx, 'ADS_STATS_SNAPSHOT');
  if (truncatedForTest) return { status: 'DONE', message: 'Merchant snapshot test limit reached. Rows=' + written + '. Advancing with partial sample.' };
  return { status: 'DONE', message: 'Merchant snapshot complete. Rows=' + written };
}


function stageAdsStatsSnapshot_(ctx) {
  var state = readState_(ctx.sheets.runState);
  var periods = statsPeriods_(ctx.settings);
  var index = num_(state.ads_period_index, 0);
  if (index === 0) clearBelowHeader_(ctx.sheets.adsStatsSnapshot);


  var done = 0;
  while (index < periods.length && done < ctx.settings.adsPeriodsPerRun && !shouldStopSoon_(ctx.settings)) {
    var p = periods[index];
    var rows = fetchAdsStatsRows_(p.key, p.from, p.to);
    if (rows.length) appendRows_(ctx.sheets.adsStatsSnapshot, rows);
    index++;
    done++;
    heartbeat_(ctx, 'ads_period_index', index, 'ads_last_period', p.key);
  }


  setStateValues_(ctx.sheets.runState, {
    ads_period_index: index,
    ads_period_keys: periodKeys_(periods).join(',')
  });
  if (index < periods.length) return { status: 'PARTIAL', message: 'Ads stats partial. Periods=' + index + '/' + periods.length };
  advanceStage_(ctx, 'PRODUCT_TYPES_BUILD');
  return { status: 'DONE', message: 'Ads stats complete. Periods=' + periods.length };
}


function stageProductTypesBuild_(ctx) {
  if (ctx.settings.enableExternalProductTypeFeed) {
    applyExternalProductTypeFeed_(ctx);
  }
  var preserved = ctx.settings.enablePreviousStateRead ? readProductTypeState_(ctx.sheets.productTypes) : {};
  var hasEnabled = hasEnabledProductType_(preserved);
  var paths = collectProductTypes_(ctx.sheets.merchantSnapshot);
  var statsByPath = buildProductTypeStats_(ctx.sheets.merchantSnapshot, ctx.sheets.adsStatsSnapshot);
  var rows = [[
    'enabled', 'product_type_l1', 'product_type_l2', 'product_type_l3', 'product_type_l4', 'product_type_l5',
    'full_path', 'target_cpa', 'comment', 'products', 'impressions', 'clicks', 'cost',
    'conversions', 'conversion_value', 'avg_order_value', 'cpa', 'roas'
  ]];
  for (var i = 0; i < paths.length; i++) {
    var path = paths[i];
    var parts = path ? path.split(' > ') : [''];
    var saved = preserved[path] || {};
    var enabled = hasEnabled && saved.hasOwnProperty('enabled') ? saved.enabled === true : true;
    var stats = productTypeOutputStats_(statsByPath[path]);
    rows.push([
      enabled,
      parts[0] || '',
      parts[1] || '',
      parts[2] || '',
      parts[3] || '',
      parts[4] || '',
      path,
      saved.target_cpa || '',
      saved.comment || '',
      stats.products,
      stats.impressions,
      stats.clicks,
      stats.cost,
      stats.conversions,
      stats.value,
      stats.avgOrderValue,
      stats.cpa,
      stats.roas
    ]);
  }
  ctx.sheets.productTypes.clearContents();
  ctx.sheets.productTypes.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  applyProductTypesCheckboxes_(ctx.sheets.productTypes, rows.length);
  if (ctx.settings.enableManagedSheetFormatting) ctx.sheets.productTypes.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  advanceStage_(ctx, 'QUARANTINE_UPDATE');
  return { status: 'DONE', message: 'ProductTypes rebuilt. Paths=' + (rows.length - 1) };
}


function applyProductTypesCheckboxes_(sheet, rowCount) {
  var maxRows = sheet.getMaxRows();
  if (maxRows > 1) sheet.getRange(2, 1, maxRows - 1, 1).clearDataValidations();
  if (rowCount > 1) sheet.getRange(2, 1, rowCount - 1, 1).insertCheckboxes();
  if (maxRows > rowCount) sheet.getRange(rowCount + 1, 1, maxRows - rowCount, 1).clearContent();
}


function hasEnabledProductType_(preserved) {
  for (var path in preserved) {
    if (preserved[path] && preserved[path].enabled === true) return true;
  }
  return false;
}


function stageQuarantineUpdate_(ctx) {
  if (!ctx.settings.enableQuarantine) {
    advanceStage_(ctx, 'PRODUCTS_DRAFT_BUILD');
    return { status: 'DONE', message: 'Quarantine skipped by enable_quarantine=FALSE.' };
  }
  var registry = ctx.settings.enablePreviousStateRead ? readQuarantineRegistry_(ctx.sheets.quarantineRegistry) : {};
  var prices = readMerchantPriceMap_(ctx.sheets.merchantSnapshot);
  var stats = readStatsByPeriod_(ctx.sheets.adsStatsSnapshot);
  var today = dateOnly_(new Date());
  var release = dateOnly_(addDays_(new Date(), ctx.settings.quarantineDays));
  var changed = 0;
  var logRows = [];
  var periods = statsPeriods_(ctx.settings);


  applyNoSales_(registry, stats.no_sales, ctx.settings.noSalesClickThreshold, release, today, logRows);
  applySpend_(registry, stats.spend_over_margin, prices, ctx.settings.spendOverMarginShare, release, today, logRows);
  applyExpensiveClick_(registry, stats.expensive_click, ctx.settings.expensiveClickCpcThreshold, release, today, logRows);


  for (var id in registry) {
    var r = registry[id];
    r.problematic = isActiveDate_(r.active_until, today);
    changed++;
  }
  writeQuarantineRegistry_(ctx.sheets.quarantineRegistry, registry);
  if (logRows.length) appendRows_(ctx.sheets.quarantineLog, logRows);
  advanceStage_(ctx, 'PRODUCTS_DRAFT_BUILD');
  return { status: 'DONE', message: 'Quarantine updated. Registry=' + changed + ', new reasons=' + logRows.length + ', periods=' + periods.length };
}


function stageProductsDraftBuild_(ctx) {
  ensureProductsDraftHeader_(ctx.sheets.productsDraft, ctx.settings);
  var state = readState_(ctx.sheets.runState);
  var offset = num_(state.products_offset, 0);
  if (offset === 0) clearBelowHeader_(ctx.sheets.productsDraft);


  var merchantLastRow = ctx.sheets.merchantSnapshot.getLastRow();
  var total = Math.max(0, merchantLastRow - 1);
  if (total === 0) throw new Error('MerchantSnapshot_Work is empty. Cannot build ProductsDraft.');


  var rules = buildProductTypeRules_(ctx.sheets.productTypes);
  var quarantine = readActiveQuarantineMap_(ctx.sheets.quarantineRegistry);
  var funnel = buildFunnelStageMap_(ctx.sheets.merchantSnapshot, ctx.sheets.adsStatsSnapshot, ctx.settings);
  var count = Math.min(ctx.settings.productsChunkSize, total - offset);
  if (count > 0) {
    var data = ctx.sheets.merchantSnapshot.getRange(offset + 2, 1, count, 10).getValues();
    var rows = [];
    for (var i = 0; i < data.length; i++) {
      rows.push(productDraftRow_(data[i], rules, quarantine, funnel, ctx.settings));
    }
    appendRows_(ctx.sheets.productsDraft, rows);
    offset += rows.length;
  }


  setStateValues_(ctx.sheets.runState, {
    products_offset: offset,
    products_draft_rows: Math.max(0, ctx.sheets.productsDraft.getLastRow() - 1)
  });
  heartbeat_(ctx, 'products_offset', offset, 'products_total', total);
  if (offset < total) return { status: 'PARTIAL', message: 'Products draft partial. Rows=' + offset + '/' + total };
  advanceStage_(ctx, 'PUBLISH_PRODUCTS');
  return { status: 'DONE', message: 'Products draft complete. Rows=' + offset };
}


function stagePublishProducts_(ctx) {
  var draftRows = Math.max(0, ctx.sheets.productsDraft.getLastRow() - 1);
  var merchantRows = Math.max(0, ctx.sheets.merchantSnapshot.getLastRow() - 1);
  logProgress_(ctx, 'PUBLISH_PRODUCTS', 'CHECKPOINT', 'Publish stage opened. DraftRows=' + draftRows + ', merchantRows=' + merchantRows);
  if (draftRows <= 0) throw new Error('ProductsDraft_Work has no rows.');
  if (draftRows !== merchantRows) throw new Error('ProductsDraft row count ' + draftRows + ' does not match Merchant snapshot ' + merchantRows + '.');
  if (!ctx.settings.enableProductsWrite) {
    setStateValues_(ctx.sheets.runState, { publish_rows_written: 0, products_write_skipped_at: iso_(new Date()) });
    advanceStage_(ctx, 'PRODUCT_DIAGNOSTICS_BUILD');
    return { status: 'DONE', message: 'Products publish skipped by enable_products_write=FALSE. Draft rows=' + draftRows };
  }


  var state = readState_(ctx.sheets.runState);
  var written = num_(state.publish_rows_written, 0);
  var cols = ctx.sheets.productsDraft.getLastColumn();
  var publicRows = Math.max(0, ctx.sheets.products.getLastRow() - 1);
  var publishWorkRows = Math.max(0, ctx.sheets.productsPublish.getLastRow() - 1);
  logProgress_(ctx, 'PUBLISH_PRODUCTS', 'CHECKPOINT', 'Publish state read. Written=' + written + ', publicRows=' + publicRows + ', workRows=' + publishWorkRows);
  if (publicRows === draftRows && publishWorkRows === 0 && productsSheetLooksPublished_(ctx.sheets.products, ctx.sheets.productsDraft, draftRows, cols)) {
    ctx.sheets.products = ctx.ss.getSheetByName(PRODUCTS_SHEET);
    ensureProductsFirst_(ctx.ss, ctx.sheets.products);
    try {
      setStateValues_(ctx.sheets.runState, { last_successful_publish_at: iso_(new Date()) });
    } catch (stateErr) {
      Logger.log('Publish recovery state update skipped: ' + String(stateErr));
    }
    var recoveryAdvanced = true;
    try {
      advanceStage_(ctx, 'PRODUCT_DIAGNOSTICS_BUILD');
    } catch (advanceErr) {
      recoveryAdvanced = false;
      Logger.log('Publish recovery stage advance skipped: ' + String(advanceErr));
    }
    if (!recoveryAdvanced) return { status: 'PARTIAL', message: 'Products publish already completed. Stage advance deferred after Sheets timeout. Rows=' + draftRows };
    return { status: 'DONE', message: 'Products publish already completed before previous logging failure. Rows=' + draftRows };
  }
  if (written === 0) {
    logProgress_(ctx, 'PUBLISH_PRODUCTS', 'CHECKPOINT', 'Preparing publish work sheet. Columns=' + cols);
    ctx.sheets.productsPublish.clearContents();
    var header = ctx.sheets.productsDraft.getRange(1, 1, 1, cols).getValues();
    ctx.sheets.productsPublish.getRange(1, 1, 1, cols).setValues(header);
  }
  var chunk = 10000;
  while (written < draftRows && !shouldStopSoon_(ctx.settings)) {
    var count = Math.min(chunk, draftRows - written);
    var values = ctx.sheets.productsDraft.getRange(written + 2, 1, count, cols).getValues();
    ctx.sheets.productsPublish.getRange(written + 2, 1, count, cols).setValues(values);
    written += count;
    heartbeat_(ctx, 'publish_rows_written', written, 'publish_total', draftRows);
  }
  if (written < draftRows) return { status: 'PARTIAL', message: 'Publish work partial. Public Products is still untouched. Rows=' + written + '/' + draftRows };
  if (Math.max(0, ctx.sheets.productsPublish.getLastRow() - 1) !== draftRows) {
    throw new Error('ProductsPublish_Work row count does not match draft after copy.');
  }


  logProgress_(ctx, 'PUBLISH_PRODUCTS', 'CHECKPOINT', 'Publish work copied. Swapping Products sheet. Rows=' + written);
  swapPublishedProductsSheet_(ctx);
  Logger.log('[PUBLISH_PRODUCTS] Products sheet swapped. Updating final publish state.');
  try {
    ctx.sheets.products = ctx.ss.getSheetByName(PRODUCTS_SHEET);
    if (ctx.sheets.products) ensureProductsFirst_(ctx.ss, ctx.sheets.products);
  } catch (sheetErr) {
    Logger.log('Products sheet post-swap refresh skipped: ' + String(sheetErr));
  }
  try {
    setStateValues_(ctx.sheets.runState, { last_successful_publish_at: iso_(new Date()) });
  } catch (stateErr) {
    Logger.log('Publish final timestamp update skipped after successful sheet swap: ' + String(stateErr));
  }
  var finalAdvanceDone = true;
  try {
    advanceStage_(ctx, 'PRODUCT_DIAGNOSTICS_BUILD');
  } catch (advanceErr) {
    finalAdvanceDone = false;
    Logger.log('Publish final stage advance skipped after successful sheet swap: ' + String(advanceErr));
  }
  if (!finalAdvanceDone) return { status: 'PARTIAL', message: 'Products published. Stage advance deferred after Sheets timeout. Rows=' + draftRows };
  return { status: 'DONE', message: 'Products published. Rows=' + draftRows };
}


function productsSheetLooksPublished_(productsSheet, draftSheet, rowCount, colCount) {
  try {
    if (!productsSheet || !draftSheet) return false;
    if (rowCount < 0 || colCount < 1) return false;
    if (productsSheet.getLastColumn() < colCount || draftSheet.getLastColumn() < colCount) return false;
    if (!sameRowValues_(productsSheet.getRange(1, 1, 1, colCount).getValues()[0], draftSheet.getRange(1, 1, 1, colCount).getValues()[0])) return false;
    if (rowCount === 0) return true;
    var rowsToCheck = [2, rowCount + 1];
    if (rowCount > 2) rowsToCheck.push(Math.floor(rowCount / 2) + 1);
    var seen = {};
    for (var i = 0; i < rowsToCheck.length; i++) {
      var row = rowsToCheck[i];
      if (seen[row]) continue;
      seen[row] = true;
      var productRow = productsSheet.getRange(row, 1, 1, colCount).getValues()[0];
      var draftRow = draftSheet.getRange(row, 1, 1, colCount).getValues()[0];
      if (!sameRowValues_(productRow, draftRow)) return false;
    }
    return true;
  } catch (e) {
    Logger.log('Published Products verification skipped: ' + String(e));
    return false;
  }
}


function sameRowValues_(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  for (var i = 0; i < a.length; i++) {
    if (String(a[i]) !== String(b[i])) return false;
  }
  return true;
}


function swapPublishedProductsSheet_(ctx) {
  var ss = ctx.ss;
  var oldProducts = ss.getSheetByName(PRODUCTS_SHEET);
  var publish = ss.getSheetByName(PRODUCTS_PUBLISH_WORK_SHEET);
  if (!publish) throw new Error('ProductsPublish_Work is missing.');
  var backupName = PRODUCTS_SHEET + '_Previous_' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMddHHmmss');
  if (oldProducts) oldProducts.setName(backupName);
  publish.setName(PRODUCTS_SHEET);
  ctx.sheets.products = publish;
  ss.setActiveSheet(publish);
  ss.moveActiveSheet(1);
  if (oldProducts) ss.deleteSheet(oldProducts);
  try {
    var nextPublish = ss.getSheetByName(PRODUCTS_PUBLISH_WORK_SHEET) || ss.insertSheet(PRODUCTS_PUBLISH_WORK_SHEET);
    ensureProductsDraftHeader_(nextPublish, ctx.settings);
    ctx.sheets.productsPublish = nextPublish;
  } catch (e) {
    Logger.log('ProductsPublish_Work preparation skipped after publish swap: ' + String(e));
  }
  try {
    compactManagedSheetGrids_(ctx.sheets, ctx.settings);
  } catch (compactErr) {
    Logger.log('Sheet grid compaction skipped after publish swap: ' + String(compactErr));
  }
}


function stageProductDiagnostics_(ctx) {
  if (!ctx.settings.enableProductDiagnostics) {
    advanceStage_(ctx, 'DASHBOARD_BUILD');
    return { status: 'DONE', message: 'ProductDiagnostics skipped by enable_product_diagnostics=FALSE.' };
  }
  buildProductDiagnostics_(ctx);
  advanceStage_(ctx, 'DASHBOARD_BUILD');
  return { status: 'DONE', message: 'ProductDiagnostics built. Rows=' + Math.max(0, ctx.sheets.productDiagnostics.getLastRow() - 1) };
}


function stageDashboard_(ctx) {
  if (!ctx.settings.enableDashboardData && !ctx.settings.enableDashboard) {
    advanceStage_(ctx, 'COMPLETE');
    return { status: 'DONE', message: 'Dashboard skipped by settings.' };
  }
  setStateValues_(ctx.sheets.runState, {
    pipeline_status: 'COMPLETE',
    stage: 'COMPLETE',
    stage_status: 'DONE',
    last_completed_stage: ctx.stage,
    heartbeat_at: iso_(new Date()),
    lock_until: ''
  });
  buildDashboardData_(ctx);
  buildDashboard_(ctx);
  return { status: 'DONE', message: 'Dashboard built.' };
}


function buildProductDiagnostics_(ctx) {
  var sheet = ctx.sheets.productDiagnostics;
  clearBelowHeader_(sheet);
  var header = [
    'id', 'title', 'price', 'benchmark_group', 'product_type_full_path',
    'funnel_stage', 'impressions', 'clicks', 'cost', 'conversions',
    'conversion_value', 'quarantine_active_until', 'quarantine_problematic',
    'excluded_destination_1', 'excluded_destination_2'
  ];
  sheet.getRange(1, 1, 1, header.length).setValues([header]);


  var stats = readStatsByPeriod_(ctx.sheets.adsStatsSnapshot).funnel || {};
  var quarantine = readQuarantineRegistry_(ctx.sheets.quarantineRegistry);
  var rules = buildProductTypeRules_(ctx.sheets.productTypes);
  var funnel = buildFunnelStageMap_(ctx.sheets.merchantSnapshot, ctx.sheets.adsStatsSnapshot, ctx.settings);
  var today = dateOnly_(new Date());
  var merchantRows = ctx.sheets.merchantSnapshot.getLastRow() > 1
    ? ctx.sheets.merchantSnapshot.getRange(2, 1, ctx.sheets.merchantSnapshot.getLastRow() - 1, 10).getValues()
    : [];
  var rows = [];
  for (var i = 0; i < merchantRows.length; i++) {
    var m = merchantRows[i];
    var id = String(m[1] || m[0] || '');
    var s = stats[id] || { impressions: 0, clicks: 0, cost: 0, conversions: 0, value: 0 };
    var q = quarantine[id] || {};
    var qActive = isActiveDate_(q.active_until, today);
    var path = String(m[5] || '');
    var excluded = (!isAllowedProductType_(path, rules) || qActive) ? ['Shopping_Ads', 'Display_Ads'] : ['', ''];
    rows.push([
      m[0] || id,
      m[2] || '',
      m[3] || 0,
      m[4] || 'other',
      path,
      funnel[id] || '6 без стат',
      s.impressions || 0,
      s.clicks || 0,
      s.cost || 0,
      s.conversions || 0,
      s.value || 0,
      q.active_until || '',
      qActive,
      excluded[0],
      excluded[1]
    ]);
  }
  if (rows.length) appendRows_(sheet, rows);
}


function buildDashboardData_(ctx) {
  var sheet = ctx.sheets.dashboardData;
  sheet.clearContents();
  var state = readState_(ctx.sheets.runState);
  var diagRows = ctx.sheets.productDiagnostics.getLastRow() > 1
    ? ctx.sheets.productDiagnostics.getRange(2, 1, ctx.sheets.productDiagnostics.getLastRow() - 1, 15).getValues()
    : [];
  var stageCounts = {};
  var excludedCount = 0;
  var quarantineCount = 0;
  for (var i = 0; i < diagRows.length; i++) {
    var stage = String(diagRows[i][5] || 'unknown');
    stageCounts[stage] = (stageCounts[stage] || 0) + 1;
    if (diagRows[i][12] === true || String(diagRows[i][12]).toUpperCase() === 'TRUE') quarantineCount++;
    if (diagRows[i][13] || diagRows[i][14]) excludedCount++;
  }
  var rows = [
    ['section', 'metric', 'value'],
    ['pipeline', 'pipeline_status', state.pipeline_status || ''],
    ['pipeline', 'stage', state.stage || ''],
    ['pipeline', 'last_successful_publish_at', state.last_successful_publish_at || ''],
    ['merchant', 'merchant_rows_written', state.merchant_rows_written || 0],
    ['merchant', 'merchant_snapshot_truncated_at', state.merchant_snapshot_truncated_at || ''],
    ['ads', 'ads_period_index', state.ads_period_index || 0],
    ['products', 'products_draft_rows', state.products_draft_rows || 0],
    ['products', 'publish_rows_written', state.publish_rows_written || 0],
    ['diagnostics', 'diagnostic_rows', diagRows.length],
    ['diagnostics', 'excluded_products', excludedCount],
    ['diagnostics', 'quarantine_products', quarantineCount],
    ['product_types', 'product_type_paths', countProductTypePaths_(ctx.sheets.productTypes)],
    ['quarantine', 'registry_rows', Math.max(0, ctx.sheets.quarantineRegistry.getLastRow() - 1)],
    ['stats', 'ads_stats_rows', Math.max(0, ctx.sheets.adsStatsSnapshot.getLastRow() - 1)]
  ];
  for (var key in stageCounts) rows.push(['funnel_stage', key, stageCounts[key]]);
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
}


function buildDashboard_(ctx) {
  var sheet = ctx.sheets.dashboard;
  sheet.clearContents();
  var state = readState_(ctx.sheets.runState);
  var rows = [
    ['Unified Merchant Funnel Product Control Extended V2', ''],
    ['pipeline_status', state.pipeline_status || ''],
    ['stage', state.stage || ''],
    ['last_completed_stage', state.last_completed_stage || ''],
    ['last_successful_publish_at', state.last_successful_publish_at || ''],
    ['merchant_rows_written', state.merchant_rows_written || 0],
    ['products_draft_rows', state.products_draft_rows || 0],
    ['publish_rows_written', state.publish_rows_written || 0],
    ['product_diagnostics_rows', Math.max(0, ctx.sheets.productDiagnostics.getLastRow() - 1)],
    ['product_type_paths', countProductTypePaths_(ctx.sheets.productTypes)],
    ['quarantine_registry_rows', Math.max(0, ctx.sheets.quarantineRegistry.getLastRow() - 1)]
  ];
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
}


function finishRun_(ctx, status, message) {
  if (!ctx) return;
  var state = {
    stage_status: status,
    heartbeat_at: iso_(new Date()),
    lock_until: '',
    last_run_finished_at: iso_(new Date())
  };
  if (status === 'DONE' && readState_(ctx.sheets.runState).stage === 'COMPLETE') {
    state.pipeline_status = 'COMPLETE';
  } else if (status === 'PARTIAL') {
    state.pipeline_status = 'PARTIAL';
  } else if (status === 'SKIPPED') {
    state.pipeline_status = 'PARTIAL';
  }
  try {
    setStateValues_(ctx.sheets.runState, state);
  } catch (stateErr) {
    Logger.log('RunState finish update failed: ' + String(stateErr));
  }
  try {
    updateRunLog_(ctx.sheets.runLog, ctx.logRow, status, message, ctx.settings);
  } catch (logErr) {
    Logger.log('RunLog finish update failed: ' + String(logErr));
  }
  try {
    trimRunLog_(ctx.sheets.runLog, ctx.settings);
  } catch (trimErr) {
    Logger.log('RunLog finish trim failed: ' + String(trimErr));
  }
}


function failRun_(ctx, err) {
  if (!ctx) return;
  var msg = truncateLogMessage_(err && err.message ? err.message : String(err), ctx.settings);
  try {
    setStateValues_(ctx.sheets.runState, {
      pipeline_status: 'FAILED',
      stage_status: 'FAILED',
      heartbeat_at: iso_(new Date()),
      lock_until: '',
      last_error_at: iso_(new Date()),
      last_error_message: msg
    });
  } catch (stateErr) {
    Logger.log('RunState fail update failed: ' + String(stateErr));
  }
  try {
    updateRunLog_(ctx.sheets.runLog, ctx.logRow, 'FAILED', msg, ctx.settings);
  } catch (logErr) {
    Logger.log('RunLog fail update failed: ' + String(logErr));
  }
  try {
    trimRunLog_(ctx.sheets.runLog, ctx.settings);
  } catch (trimErr) {
    Logger.log('RunLog fail trim failed: ' + String(trimErr));
  }
}


function resetPipelineState_(sheets, settings) {
  clearBelowHeader_(sheets.merchantSnapshot);
  clearBelowHeader_(sheets.adsStatsSnapshot);
  clearBelowHeader_(sheets.productsDraft);
  clearBelowHeader_(sheets.productsPublish);
  var state = {
    pipeline_date: dateOnly_(new Date()),
    pipeline_status: 'READY',
    stage: 'MERCHANT_SNAPSHOT',
    stage_status: 'READY',
    merchant_next_page_token: '',
    merchant_rows_written: 0,
    ads_period_index: 0,
    ads_period_keys: '',
    products_offset: 0,
    products_draft_rows: 0,
    publish_rows_written: 0,
    force_restart_seen_at: settings.forceRestartPipeline ? iso_(new Date()) : ''
  };
  setStateValues_(sheets.runState, state);
  return state;
}


function advanceStage_(ctx, nextStage) {
  setStateValues_(ctx.sheets.runState, {
    stage: nextStage,
    stage_status: 'READY',
    last_completed_stage: ctx.stage,
    heartbeat_at: iso_(new Date()),
    lock_until: ''
  });
}


function heartbeat_(ctx, key1, value1, key2, value2) {
  var update = { heartbeat_at: iso_(new Date()), lock_until: iso_(new Date(new Date().getTime() + ctx.settings.runLockMinutes * 60000)) };
  if (key1) update[key1] = value1;
  if (key2) update[key2] = value2;
  setStateValues_(ctx.sheets.runState, update);
}


function logProgress_(ctx, stage, status, message) {
  var text = '[' + stage + '] ' + message;
  Logger.log(text);
  if (ctx && ctx.sheets && ctx.sheets.runLog) {
    try {
      appendRunLog_(ctx.sheets.runLog, ctx.runId, stage, status, message);
    } catch (logErr) {
      Logger.log('RunLog progress append failed: ' + String(logErr));
    }
    try {
      trimRunLog_(ctx.sheets.runLog, ctx.settings);
    } catch (trimErr) {
      Logger.log('RunLog progress trim failed: ' + String(trimErr));
    }
  }
}


function fetchMerchantProductsPage_(merchantId, pageToken, pageSize, settings) {
  var service = getMerchantProductsService_();
  var parent = 'accounts/' + String(merchantId);
  Logger.log('[MERCHANT_SNAPSHOT] Calling Merchant Products.list. pageSize=' + pageSize + ', hasToken=' + Boolean(pageToken));
  var response = listMerchantProductsPageWithRetry_(service, parent, pageToken, pageSize, settings);
  Logger.log('[MERCHANT_SNAPSHOT] Merchant Products.list returned. products=' + (response && response.products ? response.products.length : 0) + ', nextToken=' + Boolean(response && response.nextPageToken));
  return {
    items: response && response.products ? response.products : [],
    nextPageToken: response && response.nextPageToken ? response.nextPageToken : ''
  };
}


function ensureMerchantApiServicesOn_() {
  var accountsService = getMerchantAccountsService_();
  if (!accountsService ||
      !accountsService.Accounts ||
      !accountsService.Accounts.DeveloperRegistration ||
      typeof accountsService.Accounts.DeveloperRegistration.registerGcp !== 'function') {
    throw new Error('Enable Advanced APIs: Merchant API -> Accounts. This is required to register the Google Ads Scripts GCP project.');
  }
  var productsService = getMerchantProductsService_();
  if (!productsService ||
      !productsService.Accounts ||
      !productsService.Accounts.Products ||
      typeof productsService.Accounts.Products.list !== 'function') {
    throw new Error('Enable Advanced APIs: Merchant API -> Products. V1 uses MerchantApiProducts or MerchantProducts.');
  }
}


function ensureMerchantApiDeveloperRegistration_(settings) {
  if (!settings.autoRegisterGcpProject) return true;
  if (!settings.developerEmail) {
    throw new Error('Settings.developer_email is required when auto_register_gcp_project=TRUE.');
  }
  var accountsService = getMerchantAccountsService_();
  var merchantId = String(settings.merchantId);
  try {
    var existing = getMerchantApiDeveloperRegistration_(accountsService, merchantId);
    Logger.log('Merchant API developer registration already exists: ' + JSON.stringify(existing));
    return true;
  } catch (e) {
    Logger.log('Developer registration is not confirmed or not readable yet: ' + String(e));
  }
  registerMerchantApiDeveloper_(accountsService, merchantId, settings.developerEmail);
  Logger.log('Merchant API registration was created. Stop this launch and retry after propagation on the next scheduled run.');
  return false;
}


function getMerchantProductsService_() {
  if (typeof MerchantApiProducts !== 'undefined') return MerchantApiProducts;
  if (typeof MerchantProducts !== 'undefined') return MerchantProducts;
  return null;
}


function getMerchantAccountsService_() {
  if (typeof MerchantApiAccounts !== 'undefined') return MerchantApiAccounts;
  if (typeof MerchantAccounts !== 'undefined') return MerchantAccounts;
  return null;
}


function getMerchantApiDeveloperRegistration_(accountsService, merchantId) {
  return accountsService.Accounts.DeveloperRegistration.getDeveloperRegistration(
    'accounts/' + String(merchantId) + '/developerRegistration'
  );
}


function registerMerchantApiDeveloper_(accountsService, merchantId, developerEmail) {
  var name = 'accounts/' + String(merchantId) + '/developerRegistration';
  Logger.log('Registering GCP project for Merchant API. Merchant ID: ' + merchantId + '. Developer email: ' + developerEmail);
  var response = accountsService.Accounts.DeveloperRegistration.registerGcp(
    {
      developerEmail: developerEmail
    },
    name
  );
  Logger.log('GCP project registered for Merchant API: ' + JSON.stringify(response));
  return response;
}


function merchantProductMatchesFilters_(p, settings) {
  var a = productAttributes_(p);
  var dataSource = String(p.dataSource || p.source || a.dataSource || '');
  var dataSourceId = extractDataSourceId_(dataSource);
  var feedLabel = String(p.feedLabel || a.feedLabel || '');
  var contentLanguage = String(p.contentLanguage || a.contentLanguage || '');
  if (!matchesAnyFilter_(dataSourceId, settings.merchantDataSourceIdFilter) &&
      !matchesAnyFilter_(dataSource, settings.merchantDataSourceIdFilter)) return false;
  if (!matchesAnyFilter_(feedLabel, settings.merchantFeedLabelFilter)) return false;
  if (!matchesAnyFilter_(contentLanguage, settings.merchantContentLanguageFilter)) return false;
  return true;
}


function extractDataSourceId_(value) {
  var s = String(value || '').trim();
  var m = s.match(/dataSources\/([^\/]+)/);
  if (m && m[1]) return m[1];
  var parts = s.split('/');
  return parts.length ? parts[parts.length - 1] : s;
}


function matchesAnyFilter_(value, filters) {
  if (!filters || !filters.length) return true;
  var normalized = String(value || '').trim().toLowerCase();
  for (var i = 0; i < filters.length; i++) {
    if (normalized === String(filters[i]).trim().toLowerCase()) return true;
  }
  return false;
}


function listMerchantProductsPageWithRetry_(service, parent, pageToken, pageSize, settings) {
  var attempts = Math.max(1, Number(settings.merchantApiRetryCount || 0) + 1);
  var baseSleepMs = Math.max(1, Number(settings.merchantApiRetrySleepSeconds || 10)) * 1000;
  var lastMessage = '';
  for (var attempt = 1; attempt <= attempts; attempt++) {
    try {
      return service.Accounts.Products.list(parent, {
        pageToken: pageToken || undefined,
        pageSize: pageSize
      });
    } catch (e) {
      lastMessage = e && e.message ? e.message : String(e);
      if (lastMessage.indexOf('not registered with the merchant account') !== -1) {
        throw new Error('Merchant API is not registered for this Merchant ID. Use the original V1 registration flow or enable Merchant API -> Accounts and developer registration. Details: ' + lastMessage);
      }
      if (attempt >= attempts) break;
      Utilities.sleep(baseSleepMs * attempt);
    }
  }
  throw new Error('Merchant API Products.list failed after retries: ' + lastMessage);
}


function merchantRow_(p, settings) {
  var a = productAttributes_(p);
  var id = normOfferId_(p.offerId || p.id || '');
  var price = parsePrice_(a.price || p.price);
  var pt = firstProductType_(p);
  var label = merchantCustomLabelValue_(p, settings && settings.benchmarkLabelField);
  return [
    p.offerId || p.id || '',
    id,
    a.title || p.title || '',
    price,
    label || 'other',
    pt,
    allProductTypes_(p).join(' | '),
    p.dataSource || p.source || '',
    p.feedLabel || '',
    p.contentLanguage || ''
  ];
}


function merchantCustomLabelValue_(p, attributeName) {
  var a = productAttributes_(p);
  var key = String(attributeName || '').trim().toLowerCase().replace(/_/g, '');
  var map = {
    customlabel0: 'customLabel0',
    customlabel1: 'customLabel1',
    customlabel2: 'customLabel2',
    customlabel3: 'customLabel3',
    customlabel4: 'customLabel4'
  };
  var prop = map[key];
  if (!prop) return 'other';
  return a[prop] || p[prop] || 'other';
}


function applyExternalProductTypeFeed_(ctx) {
  if (!ctx.settings.externalProductTypeSpreadsheetUrl) {
    logProgress_(ctx, 'PRODUCT_TYPES_BUILD', 'CHECKPOINT', 'External product_type feed enabled but URL is empty. Skipping.');
    return;
  }
  var map = readExternalProductTypeMap_(ctx.settings);
  var last = ctx.sheets.merchantSnapshot.getLastRow();
  if (last < 2) return;
  var range = ctx.sheets.merchantSnapshot.getRange(2, 1, last - 1, 10);
  var data = range.getValues();
  var changed = 0;
  for (var i = 0; i < data.length; i++) {
    var id = String(data[i][1] || data[i][0] || '');
    var externalPath = map[id];
    if (!externalPath) continue;
    if (ctx.settings.externalProductTypeOverride || !String(data[i][5] || '').trim()) {
      data[i][5] = externalPath;
      changed++;
    }
  }
  if (changed) range.setValues(data);
  logProgress_(ctx, 'PRODUCT_TYPES_BUILD', 'CHECKPOINT', 'External product_type feed applied. Changed=' + changed + ', sourceRows=' + Object.keys(map).length);
}


function readExternalProductTypeMap_(settings) {
  var ss = SpreadsheetApp.openByUrl(settings.externalProductTypeSpreadsheetUrl);
  var sheet = settings.externalProductTypeSheetName
    ? ss.getSheetByName(settings.externalProductTypeSheetName)
    : ss.getSheets()[0];
  if (!sheet) throw new Error('External product_type sheet not found: ' + settings.externalProductTypeSheetName);
  if (sheet.getLastRow() < 2) return {};
  var lastCol = sheet.getLastColumn();
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idCol = headerIndex_(header, settings.externalProductTypeIdColumn, -1);
  var ptCol = headerIndex_(header, settings.externalProductTypeValueColumn, -1);
  if (idCol < 0 || ptCol < 0) {
    throw new Error('External product_type source must contain headers: ' + settings.externalProductTypeIdColumn + ', ' + settings.externalProductTypeValueColumn);
  }
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  var out = {};
  for (var i = 0; i < rows.length; i++) {
    var id = normOfferId_(rows[i][idCol]);
    var pt = String(rows[i][ptCol] || '').trim();
    if (id && pt) out[id] = pt;
  }
  return out;
}


function fetchAdsStatsRows_(periodKey, from, to) {
  var query = 'SELECT OfferId, Impressions, Clicks, Cost, Conversions, ConversionValue ' +
    'FROM SHOPPING_PERFORMANCE_REPORT DURING ' + from + ',' + to;
  var report = AdsApp.report(query);
  var rows = report.rows();
  var out = [];
  while (rows.hasNext()) {
    var r = rows.next();
    var offer = r.OfferId || r['OfferId'] || r['Offer ID'] || '';
    out.push([
      periodKey,
      normOfferId_(offer),
      offer,
      num_(r.Impressions || r['Impressions'], 0),
      num_(r.Clicks || r['Clicks'], 0),
      money_(r.Cost || r['Cost']),
      num_(r.Conversions || r['Conversions'], 0),
      money_(r.ConversionValue || r['ConversionValue'] || r['Conv. value'])
    ]);
  }
  return out;
}


function statsPeriods_(settings) {
  return [
    period_('funnel', settings.funnelPeriodDays, settings.funnelExcludeLastDays),
    period_('no_sales', settings.noSalesPeriodDays, 1),
    period_('spend_over_margin', settings.spendOverMarginPeriodDays, 1),
    period_('expensive_click', settings.expensiveClickPeriodDays, 1)
  ];
}


function period_(key, days, excludeLastDays) {
  var toDate = addDays_(new Date(), -excludeLastDays);
  var fromDate = addDays_(toDate, -days + 1);
  return { key: key, from: ymd_(fromDate), to: ymd_(toDate) };
}


function readStatsByPeriod_(sheet) {
  var out = {};
  if (sheet.getLastRow() < 2) return out;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    var period = String(data[i][0] || '');
    var id = String(data[i][1] || '');
    if (!period || !id) continue;
    if (!out[period]) out[period] = {};
    out[period][id] = {
      impressions: num_(data[i][3], 0),
      clicks: num_(data[i][4], 0),
      cost: num_(data[i][5], 0),
      conversions: num_(data[i][6], 0),
      value: num_(data[i][7], 0)
    };
  }
  return out;
}


function buildFunnelStageMap_(merchantSheet, statsSheet, settings) {
  if (!settings.enableFunnelBuilder) return {};
  var stats = readStatsByPeriod_(statsSheet).funnel || {};
  var groups = {};
  var rows = merchantSheet.getLastRow() > 1 ? merchantSheet.getRange(2, 1, merchantSheet.getLastRow() - 1, 10).getValues() : [];
  for (var i = 0; i < rows.length; i++) {
    var id = String(rows[i][1] || '');
    var group = String(rows[i][4] || 'other');
    var s = stats[id] || { impressions: 0, clicks: 0, conversions: 0 };
    if (!groups[group]) groups[group] = { clicks: [], impressions: [] };
    groups[group].clicks.push(s.clicks);
    groups[group].impressions.push(s.impressions);
  }
  var thresholds = {};
  for (var g in groups) {
    thresholds[g] = {
      clicks: percentile_(groups[g].clicks, 0.5),
      impressions: percentile_(groups[g].impressions, 0.5)
    };
  }
  var out = {};
  for (var j = 0; j < rows.length; j++) {
    var offer = String(rows[j][1] || '');
    var bg = String(rows[j][4] || 'other');
    out[offer] = funnelStage_(stats[offer], thresholds[bg] || thresholds.other || { clicks: 1, impressions: 1 });
  }
  return out;
}


function funnelStage_(s, t) {
  s = s || { impressions: 0, clicks: 0, conversions: 0 };
  if (s.conversions > 0) return '1 продажі';
  if (s.clicks >= Math.max(1, t.clicks) && s.impressions >= Math.max(1, t.impressions)) return '2 вк+вп';
  if (s.clicks >= Math.max(1, t.clicks) && s.impressions < Math.max(1, t.impressions)) return '3 вк+нп';
  if (s.clicks < Math.max(1, t.clicks) && s.impressions >= Math.max(1, t.impressions)) return '4 нк+вп';
  if (s.clicks > 0 || s.impressions > 0) return '5 нк+нп';
  return '6 без стат';
}


function productDraftRow_(m, rules, quarantine, funnel, settings) {
  var rawId = String(m[0] || '');
  var id = String(m[1] || rawId);
  var path = String(m[5] || '');
  var blockedByProductType = settings.enableProductTypeFilter && !isAllowedProductType_(path, rules);
  var blockedByQuarantine = settings.enableQuarantine && quarantine[id];
  var excluded = (blockedByProductType || blockedByQuarantine) ? ['Shopping_Ads', 'Display_Ads'] : ['', ''];
  var row = [rawId || id, excluded[0], excluded[1], settings.enableFunnelBuilder ? (funnel[id] || '6 без стат') : ''];
  if (settings.benchmarkOutputAttribute) row.push(String(m[4] || 'other'));
  return row;
}


function ensureProductsDraftHeader_(sheet, settings) {
  var header = ['id', 'excluded_destination', 'excluded_destination', settings.funnelStageOutputAttribute || 'custom_label_0'];
  if (settings.benchmarkOutputAttribute) header.push(settings.benchmarkOutputAttribute);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    return;
  }
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
}


function buildProductTypeRules_(sheet) {
  var rules = [];
  if (sheet.getLastRow() < 2) return rules;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(7, sheet.getLastColumn())).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === true || String(data[i][0]).toUpperCase() === 'TRUE') {
      rules.push(String(data[i][6] || '').trim());
    }
  }
  return rules;
}


function isAllowedProductType_(path, rules) {
  if (!rules.length) return true;
  for (var i = 0; i < rules.length; i++) {
    if (path === rules[i] || path.indexOf(rules[i] + ' > ') === 0) return true;
  }
  return false;
}


function readProductTypeState_(sheet) {
  var out = {};
  if (sheet.getLastRow() < 2) return out;
  var lastCol = Math.max(9, sheet.getLastColumn());
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var fullPathCol = headerIndex_(header, 'full_path', 6);
  var targetCpaCol = headerIndex_(header, 'target_cpa', 7);
  var commentCol = headerIndex_(header, 'comment', 8);
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  for (var i = 0; i < data.length; i++) {
    var path = String(data[i][fullPathCol] || '').trim();
    if (!path) continue;
    out[path] = {
      enabled: data[i][0] === true || String(data[i][0]).toUpperCase() === 'TRUE',
      target_cpa: data[i][targetCpaCol] || '',
      comment: data[i][commentCol] || ''
    };
  }
  return out;
}


function buildProductTypeStats_(merchantSheet, statsSheet) {
  var stats = readStatsByPeriod_(statsSheet).funnel || {};
  var out = {};
  if (merchantSheet.getLastRow() < 2) return out;
  var data = merchantSheet.getRange(2, 1, merchantSheet.getLastRow() - 1, 10).getValues();
  for (var i = 0; i < data.length; i++) {
    var id = String(data[i][1] || data[i][0] || '');
    var path = String(data[i][5] || '').trim() || '(empty)';
    var parts = path === '(empty)' ? ['(empty)'] : path.split(' > ');
    var s = stats[id] || { impressions: 0, clicks: 0, cost: 0, conversions: 0, value: 0 };
    var acc = [];
    for (var level = 0; level < Math.min(5, parts.length); level++) {
      acc.push(parts[level]);
      addProductTypeStats_(out, acc.join(' > '), s);
    }
  }
  return out;
}


function addProductTypeStats_(out, path, s) {
  if (!out[path]) out[path] = { products: 0, impressions: 0, clicks: 0, cost: 0, conversions: 0, value: 0 };
  out[path].products++;
  out[path].impressions += num_(s.impressions, 0);
  out[path].clicks += num_(s.clicks, 0);
  out[path].cost += num_(s.cost, 0);
  out[path].conversions += num_(s.conversions, 0);
  out[path].value += num_(s.value, 0);
}


function productTypeOutputStats_(s) {
  s = s || { products: 0, impressions: 0, clicks: 0, cost: 0, conversions: 0, value: 0 };
  var avgOrderValue = s.conversions > 0 ? s.value / s.conversions : 0;
  var cpa = s.conversions > 0 ? s.cost / s.conversions : 0;
  var roas = s.cost > 0 ? s.value / s.cost : 0;
  return {
    products: s.products || 0,
    impressions: s.impressions || 0,
    clicks: s.clicks || 0,
    cost: s.cost || 0,
    conversions: s.conversions || 0,
    value: s.value || 0,
    avgOrderValue: avgOrderValue,
    cpa: cpa,
    roas: roas
  };
}


function collectProductTypes_(sheet) {
  var set = {};
  if (sheet.getLastRow() > 1) {
    var data = sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      var path = String(data[i][0] || '').trim();
      if (!path) path = '(empty)';
      var parts = path === '(empty)' ? ['(empty)'] : path.split(' > ');
      var acc = [];
      for (var p = 0; p < Math.min(5, parts.length); p++) {
        acc.push(parts[p]);
        set[acc.join(' > ')] = true;
      }
    }
  }
  var out = [];
  for (var k in set) out.push(k);
  out.sort();
  return out;
}


function applyNoSales_(registry, stats, clickThreshold, release, today, logRows) {
  stats = stats || {};
  for (var id in stats) {
    var s = stats[id];
    if (s.clicks >= clickThreshold && s.conversions <= 0) addQuarantineReason_(registry, id, 'NO_SALES', 'no_sales_until', release, today, logRows);
  }
}


function applySpend_(registry, stats, prices, share, release, today, logRows) {
  stats = stats || {};
  for (var id in stats) {
    var price = prices[id] || 0;
    if (price > 0 && stats[id].cost >= price * share) addQuarantineReason_(registry, id, 'SPEND_OVER_MARGIN', 'spend_until', release, today, logRows);
  }
}


function applyExpensiveClick_(registry, stats, cpcThreshold, release, today, logRows) {
  stats = stats || {};
  for (var id in stats) {
    var s = stats[id];
    if (s.clicks > 0 && s.cost / s.clicks >= cpcThreshold) addQuarantineReason_(registry, id, 'EXPENSIVE_CLICK', 'expensive_click_until', release, today, logRows);
  }
}


function addQuarantineReason_(registry, id, reason, untilField, release, today, logRows) {
  var r = registry[id] || { id: id, quarantine_count: 0, active_until: '', no_sales_until: '', spend_until: '', expensive_click_until: '', last_added: '' };
  if (!isActiveDate_(r[untilField], today)) {
    r.quarantine_count = num_(r.quarantine_count, 0) + 1;
    r.last_added = today;
    logRows.push([today, id, reason, release]);
    r[untilField] = release;
  }
  r.active_until = maxDateString_(maxDateString_(r.no_sales_until, r.spend_until), r.expensive_click_until);
  r.problematic = true;
  registry[id] = r;
}


function readQuarantineRegistry_(sheet) {
  var out = {};
  if (sheet.getLastRow() < 2) return out;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    var id = String(data[i][0] || '');
    if (!id) continue;
    out[id] = {
      id: id,
      quarantine_count: num_(data[i][1], 0),
      active_until: dateString_(data[i][2]),
      no_sales_until: dateString_(data[i][3]),
      spend_until: dateString_(data[i][4]),
      expensive_click_until: dateString_(data[i][5]),
      problematic: data[i][6] === true || String(data[i][6]).toUpperCase() === 'TRUE',
      last_added: dateString_(data[i][7])
    };
  }
  return out;
}


function writeQuarantineRegistry_(sheet, registry) {
  clearBelowHeader_(sheet);
  var rows = [];
  for (var id in registry) {
    var r = registry[id];
    rows.push([id, r.quarantine_count || 0, r.active_until || '', r.no_sales_until || '', r.spend_until || '', r.expensive_click_until || '', r.problematic === true, r.last_added || '']);
  }
  rows.sort(function(a, b) { return String(a[0]).localeCompare(String(b[0])); });
  if (rows.length) appendRows_(sheet, rows);
}


function readActiveQuarantineMap_(sheet) {
  var reg = readQuarantineRegistry_(sheet);
  var today = dateOnly_(new Date());
  var out = {};
  for (var id in reg) if (isActiveDate_(reg[id].active_until, today)) out[id] = true;
  return out;
}


function readMerchantPriceMap_(sheet) {
  var out = {};
  if (sheet.getLastRow() < 2) return out;
  var data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < data.length; i++) out[String(data[i][0] || '')] = num_(data[i][2], 0);
  return out;
}


function ensureRunStateHeader_(sheet) {
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
}


function ensureRunLogHeader_(sheet) {
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, 8).setValues([['started_at', 'finished_at', 'run_id', 'stage', 'status', 'message', 'duration_sec', 'user']]);
}


function ensureMerchantSnapshotHeader_(sheet) {
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, 10).setValues([['id', 'norm_id', 'title', 'price', 'benchmark_group', 'product_type_full_path', 'product_type_all_paths', 'data_source_id', 'feed_label', 'content_language']]);
}


function ensureAdsStatsHeader_(sheet) {
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, 8).setValues([['period_key', 'norm_id', 'offer_id', 'impressions', 'clicks', 'cost', 'conversions', 'conversion_value']]);
}


function ensureQuarantineHeaders_(registry, log) {
  if (registry.getLastRow() === 0) registry.getRange(1, 1, 1, 8).setValues([['id', 'quarantine_count', 'active_until', 'no_sales_until', 'spend_until', 'expensive_click_until', 'problematic', 'last_added']]);
  if (log.getLastRow() === 0) log.getRange(1, 1, 1, 4).setValues([['date', 'id', 'reason', 'active_until']]);
}


function compactManagedSheetGrids_(sheets, settings) {
  var productOutputCols = settings.benchmarkOutputAttribute ? 5 : 4;
  ensureSheetColumns_(sheets.products, productOutputCols);
  ensureSheetColumns_(sheets.productsDraft, productOutputCols);
  ensureSheetColumns_(sheets.productsPublish, productOutputCols);
  ensureSheetColumns_(sheets.settings, 3);
  ensureSheetColumns_(sheets.productTypes, 18);
  ensureSheetColumns_(sheets.productDiagnostics, 15);
  ensureSheetColumns_(sheets.dashboard, 3);
  ensureSheetColumns_(sheets.dashboardData, 3);
  ensureSheetColumns_(sheets.quarantineRegistry, 8);
  ensureSheetColumns_(sheets.quarantineLog, 4);
  ensureSheetColumns_(sheets.runState, 2);
  ensureSheetColumns_(sheets.runLog, 8);
  ensureSheetColumns_(sheets.merchantSnapshot, 10);
  ensureSheetColumns_(sheets.adsStatsSnapshot, 8);
}


function ensureSheetColumns_(sheet, requiredCols) {
  if (!sheet || !requiredCols) return;
  var currentCols = sheet.getMaxColumns();
  if (currentCols > requiredCols) {
    sheet.deleteColumns(requiredCols + 1, currentCols - requiredCols);
  } else if (currentCols < requiredCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }
}


function readState_(sheet) {
  var out = {};
  if (sheet.getLastRow() < 2) return out;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) if (data[i][0]) out[String(data[i][0])] = data[i][1];
  return out;
}


function countProductTypePaths_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var values = sheet.getRange(2, 7, last - 1, 1).getValues();
  var count = 0;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim()) count++;
  }
  return count;
}


function setStateValues_(sheet, values) {
  var lastRow = Math.max(1, sheet.getLastRow());
  var data = sheet.getRange(1, 1, lastRow, 2).getValues();
  if (!data.length || String(data[0][0] || '') !== 'key') data.unshift(['key', 'value']);
  else data[0] = ['key', 'value'];
  var keys = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) keys[String(data[i][0])] = i;
  }
  for (var key in values) {
    if (keys.hasOwnProperty(key)) data[keys[key]][1] = values[key];
    else {
      keys[key] = data.length;
      data.push([key, values[key]]);
    }
  }
  sheet.getRange(1, 1, data.length, 2).setValues(data);
}


function appendRunLog_(sheet, runId, stage, status, message) {
  var msg = truncateLogMessage_(message, null);
  Logger.log('RunLog ' + stage + ' ' + status + ': ' + msg);
  try {
    var row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, 8).setValues([[iso_(new Date()), '', runId, stage, status, msg, '', 'google_ads_script']]);
    return row;
  } catch (e) {
    Logger.log('RunLog append failed: ' + String(e));
    return null;
  }
}


function updateRunLog_(sheet, row, status, message, settings) {
  if (!row) return;
  try {
    var started = sheet.getRange(row, 1).getValue();
    var duration = started ? Math.round((new Date().getTime() - new Date(started).getTime()) / 1000) : '';
    var stage = sheet.getRange(row, 4).getValue();
    var msg = truncateLogMessage_(message, settings);
    Logger.log('RunLog ' + stage + ' ' + status + ': ' + msg);
    sheet.getRange(row, 2, 1, 5).setValues([[iso_(new Date()), sheet.getRange(row, 3).getValue(), stage, status, msg]]);
    sheet.getRange(row, 7).setValue(duration);
  } catch (e) {
    Logger.log('RunLog update failed: ' + String(e));
  }
}


function trimRunLog_(sheet, settings) {
  if (!sheet) return;
  try {
    var maxRows = Math.max(1, num_(settings && settings.runLogMaxRows, 200));
    var dataRows = Math.max(0, sheet.getLastRow() - 1);
    var extraRows = dataRows - maxRows;
    if (extraRows > 0) sheet.deleteRows(2, extraRows);
  } catch (e) {
    Logger.log('RunLog trim failed: ' + String(e));
  }
}


function truncateLogMessage_(message, settings) {
  var text = String(message || '');
  var maxChars = Math.max(100, num_(settings && settings.runLogMaxMessageChars, 800));
  if (text.length <= maxChars) return text;
  return text.substring(0, maxChars) + '... [truncated]';
}


function appendRows_(sheet, rows) {
  if (!rows || !rows.length) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}


function clearBelowHeader_(sheet) {
  var last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, Math.max(1, sheet.getLastColumn())).clearContent();
}


function requireMerchantId_(settings) {
  if (!settings.merchantId) throw new Error('Settings.merchant_id is required.');
}


function ensureSpreadsheetLocale_(ss, settings) {
  var locale = settings && settings.spreadsheetLocale ? String(settings.spreadsheetLocale).trim() : '';
  if (!locale) return;
  try {
    if (ss.getSpreadsheetLocale && ss.getSpreadsheetLocale() === locale) return;
    if (ss.setSpreadsheetLocale) ss.setSpreadsheetLocale(locale);
  } catch (e) {
    Logger.log('Spreadsheet locale setup skipped: ' + String(e));
  }
}


function shouldStopSoon_(settings) {
  var maxMinutes = Math.max(1, num_(settings && settings.maxRunMinutes, 25));
  var workMinutes = Math.max(1, maxMinutes - 5);
  return new Date().getTime() - START_TIME > workMinutes * 60000;
}


function isLockExpired_(lockUntil) {
  if (!lockUntil) return true;
  return new Date(lockUntil).getTime() < new Date().getTime();
}


function firstProductType_(p) {
  var all = allProductTypes_(p);
  return all.length ? all[0] : '';
}


function allProductTypes_(p) {
  var out = [];
  var a = productAttributes_(p);
  if (a.productTypes && a.productTypes.length) out = a.productTypes;
  else if (p.productTypes && p.productTypes.length) out = p.productTypes;
  else if (a.productType) out = [a.productType];
  else if (p.productType) out = [p.productType];
  else {
    var customProductType = customAttributeValue_(p, 'product_type') || customAttributeValue_(p, 'product type');
    if (customProductType) out = [customProductType];
  }
  return out;
}


function productAttributes_(p) {
  return (p && (p.productAttributes || p.attributes)) || {};
}


function customAttributeValue_(p, name) {
  var attrs = p && p.customAttributes ? p.customAttributes : [];
  var target = String(name || '').toLowerCase();
  for (var i = 0; i < attrs.length; i++) {
    var attrName = String(attrs[i].name || '').toLowerCase();
    if (attrName === target) return attrs[i].value || attrs[i].textValue || '';
  }
  return '';
}


function parsePrice_(price) {
  if (!price) return 0;
  if (typeof price === 'string') return num_(price.replace(/[^\d.,-]/g, '').replace(',', '.'), 0);
  return num_(price.value || price.amountMicros / 1000000, 0);
}


function normOfferId_(id) {
  return String(id || '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ')
    .replace(/[\u2010-\u2015\u2212]/g, '-')
    .replace(/^online:[^:]+:[^:]+:/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}


function money_(value) {
  if (typeof value === 'number') return value;
  return num_(String(value || '').replace(/,/g, ''), 0);
}


function num_(value, fallback) {
  var n = parseFloat(value);
  return isNaN(n) ? fallback : n;
}


function bool_(value, fallback) {
  if (value === true || value === false) return value;
  var s = String(value || '').toUpperCase();
  if (s === 'TRUE' || s === 'YES' || s === '1') return true;
  if (s === 'FALSE' || s === 'NO' || s === '0') return false;
  return fallback;
}


function listSetting_(value) {
  var text = String(value || '').trim();
  if (!text) return [];
  var parts = text.split(',');
  var out = [];
  for (var i = 0; i < parts.length; i++) {
    var v = String(parts[i] || '').trim();
    if (v) out.push(v);
  }
  return out;
}


function headerIndex_(header, name, fallback) {
  var target = String(name || '').trim().toLowerCase();
  for (var i = 0; i < header.length; i++) {
    if (String(header[i] || '').trim().toLowerCase() === target) return i;
  }
  return fallback;
}


function percentile_(values, p) {
  var nums = [];
  for (var i = 0; i < values.length; i++) if (!isNaN(parseFloat(values[i]))) nums.push(parseFloat(values[i]));
  if (!nums.length) return 1;
  nums.sort(function(a, b) { return a - b; });
  return nums[Math.min(nums.length - 1, Math.max(0, Math.floor((nums.length - 1) * p)))];
}


function periodKeys_(periods) {
  var out = [];
  for (var i = 0; i < periods.length; i++) out.push(periods[i].key);
  return out;
}


function iso_(date) {
  return Utilities.formatDate(date, tz_(), "yyyy-MM-dd'T'HH:mm:ss");
}


function ymd_(date) {
  return Utilities.formatDate(date, tz_(), 'yyyyMMdd');
}


function dateOnly_(date) {
  return Utilities.formatDate(date, tz_(), 'yyyy-MM-dd');
}


function tz_() {
  try {
    return AdsApp.currentAccount().getTimeZone();
  } catch (e) {
    return 'Etc/UTC';
  }
}


function dateString_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return dateOnly_(v);
  return String(v).substring(0, 10);
}


function addDays_(date, days) {
  return new Date(date.getTime() + days * 86400000);
}


function isActiveDate_(dateString, today) {
  if (!dateString) return false;
  return String(dateString).substring(0, 10) >= today;
}


function maxDateString_(a, b) {
  a = dateString_(a);
  b = dateString_(b);
  if (!a) return b;
  if (!b) return a;
  return a > b ? a : b;
}
