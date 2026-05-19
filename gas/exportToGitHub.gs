/**
 * ============================================================
 * VAZ Intel Platform - GitHub Pages 自動更新スクリプト
 * ============================================================
 *
 * 役割：スプレッドシートの最新データを data.json に変換し、
 *       GitHub リポジトリへ commit する。
 *
 * 設定手順：
 *   1. GitHub で Personal Access Token (classic) を発行
 *      ・必要スコープ: repo
 *   2. GAS の「プロジェクトの設定 > スクリプトプロパティ」に以下を登録
 *      ・GITHUB_TOKEN  : 発行したトークン
 *      ・GITHUB_OWNER  : GitHubユーザー名（例: hiromasa-yonekura）
 *      ・GITHUB_REPO   : リポジトリ名（例: vaz-intel）
 *      ・GITHUB_BRANCH : ブランチ（通常は main）
 *
 * 実行方法：
 *   ・onOpen のメニューから「📤 GitHubへ公開」をクリック
 *   ・または時間ベースのトリガーで自動実行（推奨：1時間ごと）
 * ============================================================
 */

const CONFIG = {
  SHEETS: {
    NEWS: '1_収集ニュース一覧',
    SCORE: '2_定量評価スコア',
    ANALYSIS: '3_経営分析ダッシュボード',
    COMPETITOR: '4_競合'
  },
  TARGET_PATH: 'docs/data.json'  // GitHub Pages を docs/ から配信する想定
};

/**
 * onOpen メニュー登録（既存メニューに追加）
 * ※ 既存の onOpen と統合する場合は、addItem を既存コードに追記してください
 */
function addExportMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🌐 ダッシュボード公開')
    .addItem('📤 GitHubへ公開', 'exportToGitHub')
    .addItem('🔍 JSONプレビュー（ログ）', 'previewJson')
    .addToUi();
}

/**
 * メイン関数：スプレッドシート全体を JSON 化して GitHub に push
 */
function exportToGitHub() {
  try {
    const json = buildDashboardJson();
    const jsonString = JSON.stringify(json, null, 2);

    pushToGitHub(jsonString);

    SpreadsheetApp.getUi().alert(
      '✅ 公開完了',
      `data.json を GitHub へ送信しました。\n` +
      `件数: ニュース ${json.news.length} / 分析 ${json.analyses.length} / 競合 ${json.competitors.length}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    Logger.log('❌ Export Error: ' + e.toString());
    SpreadsheetApp.getUi().alert('❌ エラー', e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * JSON プレビュー（ログ確認用、GitHub には送らない）
 */
function previewJson() {
  const json = buildDashboardJson();
  Logger.log(JSON.stringify(json, null, 2).substring(0, 5000));
  SpreadsheetApp.getUi().alert('JSON を Logger に出力しました。「表示 > ログ」で確認してください。');
}

/**
 * 各シートを読み取って統合JSONを構築
 */
function buildDashboardJson() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const news        = readSheet(ss, CONFIG.SHEETS.NEWS);
  const scores      = readSheet(ss, CONFIG.SHEETS.SCORE);
  const analyses    = readSheet(ss, CONFIG.SHEETS.ANALYSIS);
  const competitors = readSheet(ss, CONFIG.SHEETS.COMPETITOR);

  // データ整形
  const normalized = {
    meta: {
      generatedAt: new Date().toISOString(),
      generatedAtJst: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
      totalNews: news.length,
      totalAnalyses: analyses.length,
      totalCompetitors: competitors.length
    },
    news: news.map(formatNews),
    scores: scores.map(formatScore),
    analyses: analyses.map(formatAnalysis),
    competitors: competitors.map(formatCompetitor),
    // ホーム画面用の集計データ（事前計算）
    home: buildHomeAggregates(news, scores, analyses)
  };

  return normalized;
}

/**
 * シートを1行=1オブジェクトとして読み込み
 */
function readSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`⚠️ シートが見つかりません: ${sheetName}`);
    return [];
  }
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];
  return values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

/** ニュース1件のフォーマット */
function formatNews(row) {
  return {
    id: row['🔐 Softr Record ID'] || '',
    date: formatDate(row['日付']),
    company: row['対象企業'] || '',
    title: row['ニュースタイトル'] || '',
    url: row['ソースURL'] || '',
    status: row['ステータス'] || ''
  };
}

/** スコア1件のフォーマット */
function formatScore(row) {
  return {
    id: row['🔐 Softr Record ID'] || '',
    date: formatDate(row['日付']),
    company: row['対象企業'] || '',
    title: row['ニュースタイトル'] || '',
    url: row['ソースURL'] || '',
    totalScore: toNumber(row['総合スコア']),
    breakdown: {
      finance: toNumber(row['財務・規模']),
      platform: toNumber(row['PF適応']),
      humanCapital: toNumber(row['人材・IP流動']),
      directRivalry: toNumber(row['直接競合度']),
      coreBusiness: toNumber(row['注力事業合致']),
      benchmark: toNumber(row['ベンチマーク度'])
    },
    status: row['ステータス'] || ''
  };
}

/** 分析1件のフォーマット */
function formatAnalysis(row) {
  return {
    id: row['🔐 Softr Record ID'] || '',
    date: formatDate(row['日付']),
    company: row['対象企業'] || '',
    title: row['ニュースタイトル'] || '',
    url: row['ソースURL'] || '',
    totalScore: toNumber(row['総合スコア']),
    tags: parseTags(row['タグ']),
    fact: row['客観的ファクト(レイヤー1)'] || '',
    intent: row['背景・競合の意図'] || '',
    impact: row['自社・業界への影響(レイヤー2)'] || '',
    agenda: row['経営陣への問いかけ(レイヤー3)'] || ''
  };
}

/** 競合1件のフォーマット */
function formatCompetitor(row) {
  return {
    id: row['🔐 Softr Record ID'] || '',
    name: row['競合一覧'] || ''
  };
}

/**
 * ホーム画面用の集計データを事前計算
 * フロント側で重い計算をさせないため、ここで作っておく
 */
function buildHomeAggregates(news, scores, analyses) {
  // 今月の月初日
  const now = new Date();
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);

  // 今月のニュースに絞る
  const thisMonthNews = news.filter(n => {
    const d = n['日付'];
    return d && new Date(d) >= monthStart;
  });

  const thisMonthAnalyses = analyses.filter(a => {
    const d = a['日付'];
    return d && new Date(d) >= monthStart;
  });

  // 企業別累積スコア
  const companyScores = {};
  analyses.forEach(a => {
    const c = a['対象企業'];
    if (!c) return;
    if (!companyScores[c]) companyScores[c] = { company: c, totalScore: 0, count: 0 };
    companyScores[c].totalScore += toNumber(a['総合スコア']);
    companyScores[c].count += 1;
  });
  const ranking = Object.values(companyScores)
    .sort((a, b) => b.totalScore - a.totalScore);

  // タグ別件数
  const tagCounts = {};
  analyses.forEach(a => {
    parseTags(a['タグ']).forEach(t => {
      tagCounts[t] = (tagCounts[t] || 0) + 1;
    });
  });
  const topTags = Object.entries(tagCounts)
    .map(([tag, count]) => ({ tag, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);

  // ハイライト（今月の高スコア分析 TOP3）
  const highlights = thisMonthAnalyses
    .sort((a, b) => toNumber(b['総合スコア']) - toNumber(a['総合スコア']))
    .slice(0, 3)
    .map(formatAnalysis);

  return {
    monthLabel: Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年M月'),
    thisMonthNewsCount: thisMonthNews.length,
    thisMonthAnalysisCount: thisMonthAnalyses.length,
    highlights: highlights,
    ranking: ranking,
    topTags: topTags
  };
}

/* ========== ユーティリティ ========== */

function formatDate(d) {
  if (!d) return '';
  if (d instanceof Date) {
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return String(d).substring(0, 10);
}

function toNumber(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function parseTags(s) {
  if (!s) return [];
  return String(s).split(/[ 　,、]/).map(t => t.trim()).filter(t => t.startsWith('#'));
}

/* ========== GitHub API 連携 ========== */

/**
 * data.json を GitHub リポジトリへ push
 * Contents API を使用：https://docs.github.com/en/rest/repos/contents
 */
function pushToGitHub(content) {
  const props = PropertiesService.getScriptProperties();
  const token  = props.getProperty('GITHUB_TOKEN');
  const owner  = props.getProperty('GITHUB_OWNER');
  const repo   = props.getProperty('GITHUB_REPO');
  const branch = props.getProperty('GITHUB_BRANCH') || 'main';

  if (!token || !owner || !repo) {
    throw new Error('スクリプトプロパティに GITHUB_TOKEN / GITHUB_OWNER / GITHUB_REPO を設定してください。');
  }

  const path = CONFIG.TARGET_PATH;
  const apiUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;

  // 既存ファイルの SHA を取得（更新時は必須）
  let sha = null;
  try {
    const getRes = UrlFetchApp.fetch(apiUrl + `?ref=${branch}`, {
      method: 'get',
      headers: { Authorization: 'token ' + token },
      muteHttpExceptions: true
    });
    if (getRes.getResponseCode() === 200) {
      sha = JSON.parse(getRes.getContentText()).sha;
    }
  } catch (e) {
    Logger.log('既存ファイルなし（新規作成）');
  }

  // base64 エンコード（日本語対応のため Utilities.base64Encode を使用）
  const encoded = Utilities.base64Encode(content, Utilities.Charset.UTF_8);

  const payload = {
    message: `chore: update data.json (${new Date().toISOString()})`,
    content: encoded,
    branch: branch
  };
  if (sha) payload.sha = sha;

  const putRes = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    headers: {
      Authorization: 'token ' + token,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = putRes.getResponseCode();
  if (code !== 200 && code !== 201) {
    throw new Error(`GitHub API エラー (${code}): ${putRes.getContentText()}`);
  }

  Logger.log(`✅ GitHub 更新成功: ${apiUrl}`);
}
