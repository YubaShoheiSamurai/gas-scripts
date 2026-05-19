/**
 * summary_dashboard.gs
 * 売上データを月ごとに集計し、サマリーシートへ書き込んで棒グラフを作成する
 */

// シート名の定数
var SALES_SHEET_NAME = '売上データ';
var SUMMARY_SHEET_NAME = '月次サマリー';

// デバッグログの出力を制御するフラグ（本番環境では false に設定する）
var DEBUG_LOG = false;

// 金額の有効範囲（0円以上・10億円以下）
var AMOUNT_MIN = 0;
var AMOUNT_MAX = 1000000000;

/**
 * メイン処理：売上データを集計してサマリーシートに書き込み、グラフを更新する
 * 毎朝9時のトリガーから自動実行される
 */
function updateSummaryDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var salesSheet = ss.getSheetByName(SALES_SHEET_NAME);
  var summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);

  // サマリーシートが存在しない場合は新規作成する
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SUMMARY_SHEET_NAME);
    // 誤編集を防ぐために編集保護を設定する（オーナーのみ編集可能）
    summarySheet.protect().setDescription('月次サマリー - 自動集計シートのため編集禁止');
  }

  // 売上データを読み込む（1行目はヘッダーなのでスキップ）
  var salesData = salesSheet.getDataRange().getValues();
  var monthlyMap = aggregateSalesByMonth(salesData);

  // サマリーシートを書き直す
  writeSummary(summarySheet, monthlyMap);

  // 棒グラフを更新する
  updateBarChart(ss, summarySheet);
}

/**
 * 売上データを月ごとに集計してMapで返す
 * @param {Array} data - 売上データシートの全行（2次元配列）
 * @returns {Map} key: "YYYY年M月", value: { total: 合計金額, count: 件数 }
 */
function aggregateSalesByMonth(data) {
  var monthlyMap = new Map();

  // 1行目はヘッダーなのでスキップ
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateCell = row[0];
    var amount = row[3];

    // 日付と金額が空の行はスキップする
    if (!dateCell || amount === '') continue;

    // GASはDate型・文字列・数値（シリアル値）を返すことがあるため型に応じて変換する
    var date;
    if (dateCell instanceof Date) {
      date = dateCell;
    } else if (typeof dateCell === 'number') {
      // Excelシリアル値（1900-01-01起算）をDateに変換する
      date = new Date((dateCell - 25567) * 86400 * 1000);
    } else {
      date = new Date(dateCell);
    }
    if (isNaN(date.getTime())) continue;

    var monthKey = date.getFullYear() + '年' + (date.getMonth() + 1) + '月';
    var numAmount = Number(amount);
    // 数値かつ有効な金額範囲（0円〜10億円）であることを確認する
    if (isNaN(numAmount) || numAmount < AMOUNT_MIN || numAmount > AMOUNT_MAX) continue;

    if (monthlyMap.has(monthKey)) {
      var entry = monthlyMap.get(monthKey);
      entry.total += numAmount;
      entry.count += 1;
    } else {
      monthlyMap.set(monthKey, { total: numAmount, count: 1 });
    }
  }

  return monthlyMap;
}

/**
 * 月次サマリーシートをクリアして集計結果を書き込む
 * @param {Sheet} sheet - 書き込み先シート
 * @param {Map} monthlyMap - 月別集計データ
 */
function writeSummary(sheet, monthlyMap) {
  // シートの内容を全クリアする
  sheet.clearContents();

  // ヘッダー行を書き込む
  sheet.getRange(1, 1, 1, 3).setValues([['月', '合計売上', '件数']]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');

  if (monthlyMap.size === 0) return;

  // 月キーを年月順にソートする
  var sortedKeys = Array.from(monthlyMap.keys()).sort(function(a, b) {
    return parseMonthKey(a) - parseMonthKey(b);
  });

  // データ行を書き込む
  var rows = sortedKeys.map(function(key) {
    var entry = monthlyMap.get(key);
    return [key, entry.total, entry.count];
  });

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);

  // 合計売上列に通貨フォーマットを適用する
  sheet.getRange(2, 2, rows.length, 1).setNumberFormat('¥#,##0');
}

/**
 * "YYYY年M月" 形式の文字列を比較用の数値（YYYYMM）に変換する
 * @param {string} key - 月キー（例: "2026年1月"）
 * @returns {number} 比較用の数値（例: 202601）
 */
function parseMonthKey(key) {
  var match = key.match(/(\d{4})年(\d{1,2})月/);
  if (!match) return 0;
  var month = parseInt(match[2]);
  // 月が1〜12の範囲外の場合は不正なキーとして除外する
  if (month < 1 || month > 12) return 0;
  return parseInt(match[1]) * 100 + month;
}

/**
 * サマリーシートのデータをもとに棒グラフを作成・更新する
 * 既存のグラフがあれば削除してから新しく作成する
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {Sheet} summarySheet - サマリーシート
 */
function updateBarChart(ss, summarySheet) {
  // 既存のグラフをすべて削除する
  var existingCharts = summarySheet.getCharts();
  existingCharts.forEach(function(chart) {
    summarySheet.removeChart(chart);
  });

  var lastRow = summarySheet.getLastRow();
  // データが1行もない（ヘッダーのみ）場合はグラフを作成しない
  if (lastRow < 2) return;

  // グラフのデータ範囲（月と合計売上の2列）を指定する
  var dataRange = summarySheet.getRange(1, 1, lastRow, 2);

  var chart = summarySheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    .setPosition(2, 5, 0, 0)   // サマリーデータの右側に配置する
    .setNumHeaders(1)           // 1行目をヘッダーとして扱う
    .setOption('title', '月次売上推移')
    .setOption('hAxis', { title: '月' })
    .setOption('vAxis', { title: '合計売上（円）', format: '¥#,##0' })
    .setOption('legend', { position: 'none' })
    .setOption('width', 600)
    .setOption('height', 400)
    .build();

  summarySheet.insertChart(chart);
}

/**
 * 毎朝9時に updateSummaryDashboard を自動実行するトリガーを登録する
 * GASエディタでこの関数を一度だけ手動実行してください
 * 重複登録を防ぐため、既存の同名トリガーは削除してから新規登録します
 */
function setupDailyTrigger() {
  var functionName = 'updateSummaryDashboard';

  // 既存の同名トリガーを削除する（重複防止）
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎朝9時のトリガーを新規登録する
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();

  // DEBUG_LOG が true のときのみログを出力する（本番環境では false に設定すること）
  if (DEBUG_LOG) {
    Logger.log('トリガーを登録しました：毎朝9時に ' + functionName + ' が実行されます');
  }
}
