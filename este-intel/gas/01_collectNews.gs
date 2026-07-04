/**
 * ============================================================
 * 01_collectNews.gs - Googleニュース自動収集
 * ============================================================
 * 「4_監視対象」シートの各行についてGoogleニュースRSSを検索し、
 * 新着記事を「1_収集ニュース一覧」に追記する。
 *
 * 4_監視対象 の列構成:
 *   A: カテゴリ (エステ業界/クライアント動向/隣接プロ美容/出版・ビジネス)
 *   B: 企業名   (キーワード監視の場合は「(キーワード)〇〇」)
 *   C: 検索キーワード (空ならB列で検索)
 *   D: 有効 (TRUE/FALSE)
 *
 * 1_収集ニュース一覧 の列構成:
 *   日付 / カテゴリ / 対象企業 / ニュースタイトル / ソースURL / ステータス
 * ============================================================
 */

const NEWS_HEADERS = ['日付', 'カテゴリ', '対象企業', 'ニュースタイトル', 'ソースURL', 'ステータス'];

function collectNews() {
  const targetSheet = esteGetSheet(ESTE_CONFIG.SHEETS.TARGETS);
  const newsSheet = esteGetSheet(ESTE_CONFIG.SHEETS.NEWS);

  // ヘッダーがなければ作成
  if (newsSheet.getLastRow() === 0) {
    newsSheet.appendRow(NEWS_HEADERS);
  }

  // 監視対象を読み込み
  const targets = targetSheet.getDataRange().getValues().slice(1)
    .filter(r => r[1] && String(r[3]).toUpperCase() !== 'FALSE')
    .map(r => ({
      category: String(r[0] || '業界全般'),
      company: String(r[1]),
      query: String(r[2] || r[1])
    }));

  if (targets.length === 0) {
    Logger.log('⚠️ 監視対象がありません。「4_監視対象」シートを確認してください。');
    return;
  }

  // 既存URL・タイトルで重複判定
  const existing = newsSheet.getLastRow() > 1
    ? newsSheet.getRange(2, 1, newsSheet.getLastRow() - 1, NEWS_HEADERS.length).getValues()
    : [];
  const seenUrls = new Set(existing.map(r => String(r[4])));
  const seenTitles = new Set(existing.map(r => String(r[3])));

  const since = new Date();
  since.setDate(since.getDate() - ESTE_CONFIG.COLLECT_DAYS);

  const newRows = [];
  targets.forEach(t => {
    try {
      const items = fetchGoogleNewsRss(t.query);
      items.forEach(item => {
        if (item.pubDate < since) return;
        if (seenUrls.has(item.link) || seenTitles.has(item.title)) return;
        seenUrls.add(item.link);
        seenTitles.add(item.title);
        newRows.push([
          esteFormatDate(item.pubDate),
          t.category,
          t.company,
          item.title,
          item.link,
          '未処理'
        ]);
      });
      Utilities.sleep(500); // RSS取得の連続アクセスを緩和
    } catch (e) {
      Logger.log(`⚠️ 収集失敗 [${t.company}]: ${e}`);
    }
  });

  if (newRows.length > 0) {
    newsSheet.getRange(newsSheet.getLastRow() + 1, 1, newRows.length, NEWS_HEADERS.length)
      .setValues(newRows);
  }
  Logger.log(`✅ 収集完了: 新規 ${newRows.length} 件`);
}

/**
 * GoogleニュースRSSを取得してパース
 * 1クエリあたり最大10件に制限(ノイズ・処理量対策)
 */
function fetchGoogleNewsRss(query) {
  const url = 'https://news.google.com/rss/search?q=' + encodeURIComponent(query)
    + '&hl=ja&gl=JP&ceid=JP:ja';
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return [];

  const doc = XmlService.parse(res.getContentText());
  const channel = doc.getRootElement().getChild('channel');
  if (!channel) return [];

  return channel.getChildren('item').slice(0, 10).map(item => {
    const title = item.getChildText('title') || '';
    const link = item.getChildText('link') || '';
    const pub = item.getChildText('pubDate');
    return {
      title: title,
      link: link,
      pubDate: pub ? new Date(pub) : new Date()
    };
  });
}
