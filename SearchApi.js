/**
 * Cast-only Search API (desc_people専用)
 * /exec?token=...&people=...&limit=10
 *
 * Script Properties:
 * - API_TOKEN
 * - SEARCH_SHEET (default: videos)
 */

function doGet(e) {
  try {
    const p = (e && e.parameter) ? e.parameter : {};

    // クエリパラメータなし → Web UI を配信
    if (!p.people && !p.q) {
      return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('動画検索')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // JSON API モード
    let result;
    if (p.q) {
      result = searchByText_(p);
    } else {
      result = searchByPeople_(p);
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    const out = { ok: false, error: String(err && err.message ? err.message : err) };
    return ContentService
      .createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** google.script.run から呼ばれるサーバー側関数 */
function searchVideos(params) {
  if (!params) return { ok: false, error: 'params is null' };
  if (params.q) return searchByText_(params);
  if (params.people) return searchByPeople_(params);
  return { ok: false, error: 'q または people が必要' };
}

/** チャンネルタイトルを返す（HTML初期表示用） */
function getChannelTitle() {
  return PropertiesService.getScriptProperties().getProperty('CHANNEL_TITLE') || 'YouTube動画検索';
}

/** キーワード検索（title + desc_core + desc_topics + desc_people の横断検索） */
function searchByText_(p) {
  const queryRaw = p.q || '';
  const query = normalize_(queryRaw);
  if (!query) throw new Error("q が空。例: ?q=安保法制");

  let limit = parseInt(p.limit || '20', 10);
  if (!Number.isFinite(limit) || limit <= 0) limit = 20;
  limit = Math.min(limit, 100);

  let offset = parseInt(p.offset || '0', 10);
  if (!Number.isFinite(offset) || offset < 0) offset = 0;

  const sheetName = PropertiesService.getScriptProperties().getProperty('SEARCH_SHEET') || 'videos';
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) throw new Error(`シートが見つからない: ${sheetName}`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) {
    return { ok: true, query: { q: queryRaw, limit, offset }, totalMatches: 0, returned: 0, items: [], nextOffset: offset };
  }

  const data = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const header = data[0].map(String);
  const idx = (name) => header.indexOf(name);

  const iId       = idx('videoId');
  const iUrl      = idx('url');
  const iTitle    = idx('title');
  const iPublished = idx('publishedAt');
  const iCore     = idx('desc_core');
  const iTopics   = idx('desc_topics');
  const iPeople   = idx('desc_people');

  if (iId === -1 || iUrl === -1 || iTitle === -1) throw new Error('videosシートに videoId/url/title がない');

  const tokens = tokenize_(query);
  const items = [];
  let matchedTotal = 0;
  let taken = 0;

  for (let r = lastRow - 1; r >= 1; r--) {
    const row = data[r];
    const hay = normalize_([
      row[iTitle]   || '',
      iCore    !== -1 ? row[iCore]    || '' : '',
      iTopics  !== -1 ? row[iTopics]  || '' : '',
      iPeople  !== -1 ? row[iPeople]  || '' : '',
    ].join(' '));

    if (!allTokensIn_(tokens, hay)) continue;

    matchedTotal++;
    if (matchedTotal <= offset) continue;

    if (taken < limit) {
      items.push({
        videoId:     String(row[iId]  || ''),
        url:         String(row[iUrl] || ''),
        publishedAt: String(row[iPublished] || ''),
        title:       shrink_(String(row[iTitle] || ''), 180),
        people:      iPeople !== -1 ? shrink_(String(row[iPeople] || ''), 120) : '',
      });
      taken++;
    }
  }

  return {
    ok: true,
    query: { q: queryRaw, limit, offset },
    totalMatches: matchedTotal,
    returned: items.length,
    items,
    nextOffset: offset + items.length,
  };
}

function searchByPeople_(p) {
  const peopleRaw = p.people || "";
  const people = normalize_(peopleRaw);
  if (!people) throw new Error("people が空。例: ?people=青木理");

  let limit = parseInt(p.limit || "20", 10);
  if (!Number.isFinite(limit) || limit <= 0) limit = 20;
  limit = Math.min(limit, 100); // ★上限100（30で丸めない）

  let offset = parseInt(p.offset || "0", 10);
  if (!Number.isFinite(offset) || offset < 0) offset = 0;

  const sheetName = PropertiesService.getScriptProperties().getProperty("SEARCH_SHEET") || "videos";
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) throw new Error(`シートが見つからない: ${sheetName}`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) {
    return { ok: true, query: { people: peopleRaw, limit, offset }, totalMatches: 0, returned: 0, items: [], nextOffset: offset };
  }

  const data = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const header = data[0].map(String);
  const idx = (name) => header.indexOf(name);

  const iId = idx("videoId");
  const iUrl = idx("url");
  const iTitle = idx("title");
  const iPublished = idx("publishedAt");
  const iCast = idx("desc_people");

  if (iId === -1 || iUrl === -1 || iTitle === -1) throw new Error("videosシートに videoId/url/title がない");
  if (iCast === -1) throw new Error("desc_people 列がない（buildDescFieldsBatch を回して作って）");

  const tokens = tokenize_(people);

  const items = [];
  let matchedTotal = 0; // ★全ヒット数
  let taken = 0;

  // 最新行から走査（新しい順）
  for (let r = lastRow - 1; r >= 1; r--) {
    const row = data[r];
    const cast = normalize_(row[iCast] || "");
    if (!cast) continue;
    if (!allTokensIn_(tokens, cast)) continue;

    matchedTotal++;

    // offsetまでは読み飛ばす
    if (matchedTotal <= offset) continue;

    // limit件だけ採用（ただし全件数カウントは続ける）
    if (taken < limit) {
      items.push({
        videoId: String(row[iId] || ""),
        url: String(row[iUrl] || ""),
        publishedAt: String(row[iPublished] || ""),
        title: shrink_(String(row[iTitle] || ""), 180)
      });
      taken++;
    }
  }

  const nextOffset = offset + items.length;

  return {
    ok: true,
    query: { people: peopleRaw, limit, offset },
    totalMatches: matchedTotal,
    returned: items.length,
    items,
    nextOffset
  };
}

function normalize_(s) {
  return String(s || "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function tokenize_(q) {
  const t = normalize_(q);
  if (!t) return [];
  if (!t.includes(" ")) return [t];
  return t.split(" ").map(x => x.trim()).filter(Boolean);
}

function allTokensIn_(tokens, hay) {
  for (const tok of tokens) {
    if (!hay.includes(tok)) return false;
  }
  return true;
}

function shrink_(s, n) {
  const t = String(s || "").replace(/\s+/g, " ").trim();
  if (t.length <= n) return t;
  return t.slice(0, n) + "…";
}
