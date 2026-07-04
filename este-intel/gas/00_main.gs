/**
 * ============================================================
 * エステティック通信 Intel Platform - メイン/共通設定
 * ============================================================
 *
 * 構成:
 *   00_main.gs            … 共通設定・メニュー・トリガー
 *   01_collectNews.gs     … Googleニュース収集 → シート1
 *   02_analyzeNews.gs     … Gemini分析 → シート2・3
 *   03_exportToGitHub.gs  … data.json を GitHub へ公開
 *
 * 事前設定（プロジェクトの設定 > スクリプトプロパティ）:
 *   GEMINI_API_KEY : Google AI Studio で発行した無料APIキー
 *   GITHUB_TOKEN   : GitHub Personal Access Token (repoスコープ)
 *   GITHUB_OWNER   : GitHubユーザー名
 *   GITHUB_REPO    : リポジトリ名 (例: este-intel)
 *   GITHUB_BRANCH  : main
 * ============================================================
 */

const ESTE_CONFIG = {
  SHEETS: {
    NEWS: '1_収集ニュース一覧',
    SCORE: '2_定量評価スコア',
    ANALYSIS: '3_経営分析ダッシュボード',
    TARGETS: '4_監視対象'
  },
  // ニュース収集:過去何日分を対象にするか
  COLLECT_DAYS: 3,
  // 1回の実行で分析する最大件数(Gemini無料枠・GAS実行時間対策)
  ANALYZE_BATCH: 12,
  // Gemini呼び出し間隔(ミリ秒)。無料枠のレート制限対策
  GEMINI_SLEEP_MS: 7000,
  // 使用モデル(無料枠対応)。新モデルが出たらここを変更
  GEMINI_MODEL: 'gemini-2.5-flash',
  TARGET_PATH: 'docs/este/data.json',  // vaz-intelリポジトリ内のエステ版パス
  // 総合スコア(30点満点)の足切り。これ未満はシート3に載せない
  MIN_SCORE_FOR_ANALYSIS: 8
};

/** スプレッドシートを開いたときのメニュー */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📡 Intel運用')
    .addItem('1️⃣ ニュース収集を実行', 'collectNews')
    .addItem('2️⃣ AI分析を実行(未処理分)', 'analyzeNews')
    .addItem('3️⃣ GitHubへ公開', 'exportToGitHub')
    .addSeparator()
    .addItem('▶ 全工程を連続実行', 'runAll')
    .addItem('⏰ 自動実行トリガーを設定', 'setupTriggers')
    .addItem('🔍 JSONプレビュー(ログ)', 'previewJson')
    .addToUi();
}

/** 手動で全工程を回す(収集→分析→公開) */
function runAll() {
  collectNews();
  analyzeNews();
  exportToGitHub();
}

/**
 * 自動実行トリガーを一括設定
 *  - 毎朝5時台: ニュース収集
 *  - 毎時: AI分析(未処理があれば少しずつ処理)
 *  - 毎朝8時台・毎夕17時台: GitHub公開
 * 何度実行しても重複しないよう、既存の同名トリガーは一旦削除
 */
function setupTriggers() {
  const mine = ['collectNews', 'analyzeNewsTriggered', 'exportToGitHubSilent'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (mine.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('collectNews').timeBased().atHour(5).everyDays(1).create();
  ScriptApp.newTrigger('analyzeNewsTriggered').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('exportToGitHubSilent').timeBased().atHour(8).everyDays(1).create();
  ScriptApp.newTrigger('exportToGitHubSilent').timeBased().atHour(17).everyDays(1).create();

  SpreadsheetApp.getUi().alert('✅ トリガーを設定しました\n収集: 毎朝5時台 / 分析: 毎時 / 公開: 8時台・17時台');
}

/* ========== 共通ユーティリティ ========== */

function esteGetSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function esteFormatDate(d) {
  if (!d) return '';
  if (d instanceof Date) return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(d).substring(0, 10);
}

function esteToNumber(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}
