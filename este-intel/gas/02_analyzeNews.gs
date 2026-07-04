/**
 * ============================================================
 * 02_analyzeNews.gs - Gemini APIによるスコアリング&3視点分析
 * ============================================================
 * 「1_収集ニュース一覧」の未処理行をGeminiに渡し、
 *   ・6項目スコア(営業2+編集2+事業2、各0〜5点、計30点満点)
 *   ・3視点の示唆(営業/編集/新規事業)+リーダーへの問い
 * を生成して「2_定量評価スコア」「3_経営分析ダッシュボード」へ書き込む。
 *
 * スコア設計:
 *   [営業] 出稿ポテンシャル / 成長・投資活発度
 *   [編集] 読者関心度(サロン経営への実用性) / 話題性・新規性
 *   [事業] 市場インパクト / 自社転用・連携可能性
 * ============================================================
 */

const SCORE_HEADERS = ['日付', 'カテゴリ', '対象企業', 'ニュースタイトル', 'ソースURL',
  '総合スコア', '出稿ポテンシャル', '成長・投資', '読者関心', '話題性', '市場インパクト', '転用可能性', 'ステータス'];

const ANALYSIS_HEADERS = ['日付', 'カテゴリ', '対象企業', 'ニュースタイトル', '要約', 'ソースURL',
  '総合スコア', '営業スコア', '編集スコア', '事業スコア', 'タグ',
  '客観的ファクト', '背景・意図', '営業への示唆', '編集への示唆', '新規事業への示唆', 'リーダーへの問い'];

/** トリガー用(UIアラートを出さない) */
function analyzeNewsTriggered() {
  analyzeNewsCore(false);
}

/** メニュー用 */
function analyzeNews() {
  analyzeNewsCore(true);
}

function analyzeNewsCore(showUi) {
  const newsSheet = esteGetSheet(ESTE_CONFIG.SHEETS.NEWS);
  const scoreSheet = esteGetSheet(ESTE_CONFIG.SHEETS.SCORE);
  const analysisSheet = esteGetSheet(ESTE_CONFIG.SHEETS.ANALYSIS);

  if (scoreSheet.getLastRow() === 0) scoreSheet.appendRow(SCORE_HEADERS);
  if (analysisSheet.getLastRow() === 0) analysisSheet.appendRow(ANALYSIS_HEADERS);

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('スクリプトプロパティに GEMINI_API_KEY を設定してください。');

  if (newsSheet.getLastRow() < 2) return;
  const values = newsSheet.getRange(2, 1, newsSheet.getLastRow() - 1, 6).getValues();

  let processed = 0;
  for (let i = 0; i < values.length; i++) {
    if (processed >= ESTE_CONFIG.ANALYZE_BATCH) break;
    const row = values[i];
    if (String(row[5]) !== '未処理') continue;

    const news = {
      date: esteFormatDate(row[0]),
      category: String(row[1] || '業界全般'),
      company: String(row[2] || ''),
      title: String(row[3] || ''),
      url: String(row[4] || '')
    };

    let status;
    try {
      const result = callGemini(apiKey, news);
      if (!result || result.relevant === false) {
        status = '対象外';
      } else {
        writeScoreRow(scoreSheet, news, result);
        if (result.totalScore >= ESTE_CONFIG.MIN_SCORE_FOR_ANALYSIS) {
          writeAnalysisRow(analysisSheet, news, result);
        }
        status = '処理済';
      }
    } catch (e) {
      Logger.log(`⚠️ 分析失敗 [${news.title}]: ${e}`);
      status = 'エラー';
    }

    newsSheet.getRange(i + 2, 6).setValue(status);
    processed++;
    Utilities.sleep(ESTE_CONFIG.GEMINI_SLEEP_MS);
  }

  Logger.log(`✅ 分析完了: ${processed} 件処理`);
  if (showUi && processed === 0) {
    SpreadsheetApp.getUi().alert('未処理のニュースはありません。');
  }
}

/* ---------- Gemini呼び出し ---------- */

function callGemini(apiKey, news) {
  const prompt = buildPrompt(news);
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${ESTE_CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.3,
      responseMimeType: 'application/json'
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    throw new Error(`Gemini API エラー (${res.getResponseCode()}): ${res.getContentText().substring(0, 300)}`);
  }

  const body = JSON.parse(res.getContentText());
  const text = body.candidates && body.candidates[0]
    && body.candidates[0].content && body.candidates[0].content.parts
    && body.candidates[0].content.parts[0].text;
  if (!text) throw new Error('Geminiの応答が空です');

  const result = JSON.parse(text);
  // 総合スコアを計算(6項目合計、30点満点)
  const s = result.scores || {};
  result.totalScore = round1(
    n5(s.adPotential) + n5(s.growth) + n5(s.readerInterest)
    + n5(s.novelty) + n5(s.marketImpact) + n5(s.reusability)
  );
  result.salesScore = round1(n5(s.adPotential) + n5(s.growth));
  result.editScore = round1(n5(s.readerInterest) + n5(s.novelty));
  result.bizScore = round1(n5(s.marketImpact) + n5(s.reusability));
  return result;
}

function buildPrompt(news) {
  return `あなたはエステティック業界の専門誌「エステティック通信(エス通)」の経営戦略アナリストです。
エス通は、エステサロン経営者・エステティシャン向けの業界専門メディアで、
社内には (1)広告営業部隊 (2)編集部 (3)新規事業を検討するリーダー陣 がいます。

# エス通のビジネス文脈(分析の前提として必ず考慮すること)
- 収益商品: 月刊誌の純広告・記事広告・同梱チラシ(エリア別配布)、WEB記事(esthe.news)+SNS/LINE広告、
  アワード事業(ベストアイテム/エステセレクション/日本美容企業大賞)、イベント・セミナー・ビジネス交流会
- 広告主は業務用機器・化粧品メーカー、ディーラー、OEM会社、サロン支援会社(集客/制作/コンサル)が中心
- 営業の好機シグナル: 新製品発表、展示会出展(BWJ/BEYOND BEAUTY TOKYO/COSME Week等)、
  OEM・卸強化、上場・資金調達・M&A、新業態参入、TVショッピング・EC強化(受賞実績PRの需要が生まれる)
- 編集部が重視するテーマ: フェムケアへの業態転換、成分トレンド(エクソソーム/幹細胞/NAD+等)、
  薬機法・広告表現規制(工業会通達含む)、脱毛サロン問題と都度払いシフト、サロン集客(MEO/SNS/ライブコマース)、
  AI・DX活用、韓国美容、メンズ市場、経営者のセルフブランディング
- リーダー陣の関心: アワード/イベント/コミュニティ/データ販売など業界メディアの収益多角化、
  競合メディア(美容経済新聞/ザ・ビューレック/ビュートピア/美容の窓口/健康美容EXPO等)の動き

以下のニュースを分析し、JSONのみで回答してください。

# ニュース
- 日付: ${news.date}
- カテゴリ: ${news.category}
- 関連企業: ${news.company}
- タイトル: ${news.title}
- URL: ${news.url}

# 判定基準
タイトルから判断して、エステ業界・プロ美容業界・業界メディア経営のいずれにも無関係
(例:無関係な芸能ニュース、同名の別企業)なら relevant を false にしてください。

# 出力形式(JSON)
{
  "relevant": true,
  "summary": "30字程度の日本語要約",
  "scores": {
    "adPotential": 0-5,   // 出稿ポテンシャル: この企業/領域が広告・タイアップ出稿する見込み
    "growth": 0-5,        // 成長・投資活発度: 出店、M&A、資金調達、新製品などの勢い
    "readerInterest": 0-5,// 読者関心度: エステサロン経営者・技術者にとっての実用性
    "novelty": 0-5,       // 話題性・新規性: 記事にしたとき読者の惹きになるか
    "marketImpact": 0-5,  // 市場インパクト: エステ・プロ美容市場の構造への影響度
    "reusability": 0-5    // 転用可能性: エステティック通信の新規事業のヒントになるか
  },
  "tags": ["#タグ1", "#タグ2", "#タグ3"],
  "fact": "客観的ファクト。何が起きたかを2文以内で",
  "intent": "背景・当事者の意図。なぜこの動きをしたのかの推察を2-3文で",
  "salesAngle": "営業への示唆。攻めるべき企業・提案の切り口を具体的に1-2文で。該当なしなら「—」",
  "editorialAngle": "編集への示唆。記事化するならどんな企画・切り口が読者に刺さるかを1-2文で",
  "bizAngle": "新規事業への示唆。エステティック通信が模倣・連携・参入できる可能性を1-2文で。該当なしなら「—」",
  "agenda": "リーダー陣への問いかけを1文で"
}`;
}

/* ---------- シート書き込み ---------- */

function writeScoreRow(sheet, news, r) {
  const s = r.scores || {};
  sheet.appendRow([
    news.date, news.category, news.company, news.title, news.url,
    r.totalScore, n5(s.adPotential), n5(s.growth), n5(s.readerInterest),
    n5(s.novelty), n5(s.marketImpact), n5(s.reusability), '処理済'
  ]);
}

function writeAnalysisRow(sheet, news, r) {
  sheet.appendRow([
    news.date, news.category, news.company, news.title, r.summary || '', news.url,
    r.totalScore, r.salesScore, r.editScore, r.bizScore,
    (r.tags || []).join(' '),
    r.fact || '', r.intent || '', r.salesAngle || '', r.editorialAngle || '',
    r.bizAngle || '', r.agenda || ''
  ]);
}

/* ---------- ユーティリティ ---------- */

function n5(v) {
  const n = esteToNumber(v);
  return Math.max(0, Math.min(5, n));
}

function round1(n) {
  return Math.round(n * 10) / 10;
}
