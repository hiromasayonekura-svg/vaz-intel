/**
 * ============================================================
 * 03_exportToGitHub.gs - data.json を GitHub Pages へ公開
 * ============================================================
 * スプレッドシートの最新データを data.json に変換し、
 * GitHubリポジトリの docs/data.json へ commit する。
 * (VAZ Intel Platform の exportToGitHub.gs をエステ版に改修)
 * ============================================================
 */

/** メニュー用 */
function exportToGitHub() {
  try {
    const json = buildDashboardJson();
    pushToGitHub(JSON.stringify(json, null, 2));
    SpreadsheetApp.getUi().alert(
      '✅ 公開完了',
      `data.json を GitHub へ送信しました。\n件数: ニュース ${json.news.length} / 分析 ${json.analyses.length}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log('❌ Export Error: ' + e);
    SpreadsheetApp.getUi().alert('❌ エラー', String(e), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/** トリガー用(UIなし) */
function exportToGitHubSilent() {
  try {
    const json = buildDashboardJson();
    pushToGitHub(JSON.stringify(json, null, 2));
    Logger.log('✅ 自動公開完了');
  } catch (e) {
    Logger.log('❌ 自動公開エラー: ' + e);
  }
}

/** JSONプレビュー(ログ確認用) */
function previewJson() {
  const json = buildDashboardJson();
  Logger.log(JSON.stringify(json, null, 2).substring(0, 5000));
  SpreadsheetApp.getUi().alert('JSONをログに出力しました。「実行数」から確認してください。');
}

/* ---------- JSON構築 ---------- */

function buildDashboardJson() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const news = readSheetAsObjects(ss, ESTE_CONFIG.SHEETS.NEWS);
  const analyses = readSheetAsObjects(ss, ESTE_CONFIG.SHEETS.ANALYSIS);
  const targets = readSheetAsObjects(ss, ESTE_CONFIG.SHEETS.TARGETS);

  const formattedAnalyses = analyses.map(formatAnalysis);

  return {
    meta: {
      generatedAt: new Date().toISOString(),
      generatedAtJst: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
      totalNews: news.length,
      totalAnalyses: analyses.length
    },
    news: news.map(formatNewsRow),
    analyses: formattedAnalyses,
    competitors: targets
      .filter(t => t['企業名'] && !String(t['企業名']).startsWith('(キーワード)'))
      .map(t => ({ name: String(t['企業名']), category: String(t['カテゴリ'] || '業界全般') }))
  };
}

function readSheetAsObjects(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  return values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function formatNewsRow(row) {
  return {
    date: esteFormatDate(row['日付']),
    category: row['カテゴリ'] || '業界全般',
    company: row['対象企業'] || '',
    title: row['ニュースタイトル'] || '',
    url: row['ソースURL'] || '',
    status: row['ステータス'] || ''
  };
}

let _analysisIdCounter = 0;
function formatAnalysis(row) {
  return {
    id: _analysisIdCounter++,
    date: esteFormatDate(row['日付']),
    category: row['カテゴリ'] || '業界全般',
    company: row['対象企業'] || '',
    title: row['ニュースタイトル'] || '',
    summary: row['要約'] || '',
    url: row['ソースURL'] || '',
    totalScore: esteToNumber(row['総合スコア']),
    salesScore: esteToNumber(row['営業スコア']),
    editScore: esteToNumber(row['編集スコア']),
    bizScore: esteToNumber(row['事業スコア']),
    tags: parseTags(row['タグ']),
    fact: row['客観的ファクト'] || '',
    intent: row['背景・意図'] || '',
    salesAngle: row['営業への示唆'] || '',
    editorialAngle: row['編集への示唆'] || '',
    bizAngle: row['新規事業への示唆'] || '',
    agenda: row['リーダーへの問い'] || ''
  };
}

function parseTags(s) {
  if (!s) return [];
  return String(s).split(/[ 　,、]/).map(t => t.trim()).filter(t => t.startsWith('#'));
}

/* ---------- GitHub API ---------- */

function pushToGitHub(content) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('GITHUB_TOKEN');
  const owner = props.getProperty('GITHUB_OWNER');
  const repo = props.getProperty('GITHUB_REPO');
  const branch = props.getProperty('GITHUB_BRANCH') || 'main';

  if (!token || !owner || !repo) {
    throw new Error('スクリプトプロパティに GITHUB_TOKEN / GITHUB_OWNER / GITHUB_REPO を設定してください。');
  }

  const apiUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${ESTE_CONFIG.TARGET_PATH}`;

  let sha = null;
  const getRes = UrlFetchApp.fetch(apiUrl + `?ref=${branch}`, {
    method: 'get',
    headers: { Authorization: 'token ' + token },
    muteHttpExceptions: true
  });
  if (getRes.getResponseCode() === 200) {
    sha = JSON.parse(getRes.getContentText()).sha;
  }

  const payload = {
    message: `chore: update data.json (${new Date().toISOString()})`,
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch: branch
  };
  if (sha) payload.sha = sha;

  const putRes = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    headers: { Authorization: 'token ' + token, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = putRes.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error(`GitHub API エラー (${code}): ${putRes.getContentText()}`);
  }
}
