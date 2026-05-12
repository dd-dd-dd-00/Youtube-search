/**
 * Script Properties に入れるもの:
 * - YT_API_KEY: YouTube Data API v3 の APIキー
 *
 * 初回は runOnceInit() → syncNewPublic() の順で実行
 */

function runOnceInit() {
  const apiKey = getProp_("YT_API_KEY");
  const handle = "@PolitasTV"; // ここだけ変えれば他チャンネルにも使える

  const chUrl = "https://www.googleapis.com/youtube/v3/channels"
    + "?part=contentDetails,snippet"
    + "&forHandle=" + encodeURIComponent(handle)
    + "&key=" + encodeURIComponent(apiKey);

  const ch = JSON.parse(UrlFetchApp.fetch(chUrl).getContentText());
  if (!ch.items || ch.items.length === 0) throw new Error("チャンネル取得失敗: handle=" + handle);

  const uploads = ch.items[0].contentDetails.relatedPlaylists.uploads;
  const channelTitle = ch.items[0].snippet?.title || "";

  setProp_("UPLOADS_PLAYLIST_ID", uploads);
  setProp_("CHANNEL_TITLE", channelTitle);

  ensureSheet_();
}

/**
 * 新着（= uploadsプレイリスト先頭側）だけ拾って videos シートに追記する
 * 6時間ごとトリガー推奨
 */
function syncNewPublic() {
  const apiKey = getProp_("YT_API_KEY");
  const uploads = getProp_("UPLOADS_PLAYLIST_ID"); // runOnceInitで保存済み前提
  const channelTitle = getProp_("CHANNEL_TITLE") || "";

  if (!uploads) throw new Error("UPLOADS_PLAYLIST_ID が未設定。先に runOnceInit() を実行して。");

  const sh = ensureSheet_();

  // 既存IDをSet化（重複防止）
  const lastRow = sh.getLastRow();
  const existing = new Set();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, 1).getValues().flat()
      .forEach(id => { if (id) existing.add(String(id)); });
  }

  const nowIso = new Date().toISOString();
  let pageToken = "";
  let pages = 0;
  const MAX_PAGES = 4; // 50×4=200本分だけ先頭から見る。十分。
  let addedTotal = 0;

  // 「既存ばかり」になったら打ち切るためのカウンタ
  let noAddPageStreak = 0;

  while (pages < MAX_PAGES) {
    const plUrl = "https://www.googleapis.com/youtube/v3/playlistItems"
      + "?part=snippet,contentDetails"
      + "&maxResults=50"
      + "&playlistId=" + encodeURIComponent(uploads)
      + (pageToken ? "&pageToken=" + encodeURIComponent(pageToken) : "")
      + "&key=" + encodeURIComponent(apiKey);

    const pl = JSON.parse(UrlFetchApp.fetch(plUrl).getContentText());
    if (!pl.items) break;

    const rows = [];
    for (const it of pl.items) {
      const vid = it.contentDetails?.videoId;
      const sn = it.snippet || {};
      if (!vid) continue;

      // 既に拾ってるならスキップ
      if (existing.has(vid)) continue;

      rows.push([
        vid,
        "https://www.youtube.com/watch?v=" + vid,
        sn.title || "",
        sn.description || "",
        sn.publishedAt || "",
        sn.channelTitle || channelTitle,
        nowIso // firstSeenAt（公開中に拾えた証拠）
      ]);
      existing.add(vid);
    }

    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      addedTotal += rows.length;
      noAddPageStreak = 0;
    } else {
      noAddPageStreak += 1;
      // 先頭側ページが連続で追加ゼロなら、もう新着はないと判断して止める
      if (noAddPageStreak >= 2) break;
    }

    pageToken = pl.nextPageToken;
    if (!pageToken) break;
    pages += 1;
  }

  Logger.log("Added: " + addedTotal);
}

/**
 * 公開中の「今見える分」を uploads プレイリストの最後まで走査して拾う（分割実行・再開対応）
 *
 * 使い方：
 * 1) runOnceInit() を1回実行（UPLOADS_PLAYLIST_IDが入る）
 * 2) backfillAllPublicResume() を何回か実行（トリガーで回すと楽）
 * 3) ログに DONE が出たら完了
 *
 * 注意：メン限化済みでAPIから見えない過去動画はここでは取れない（今公開されてる分＝今見える分だけ）
 */
function backfillAllPublicResume() {
  const apiKey = getProp_("YT_API_KEY");
  const uploads = getProp_("UPLOADS_PLAYLIST_ID");
  if (!uploads) throw new Error("UPLOADS_PLAYLIST_ID が未設定。先に runOnceInit() を実行して。");

  const sh = ensureSheet_();

  // 既存IDをSet化（重複防止）
  const lastRow = sh.getLastRow();
  const existing = new Set();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, 1).getValues().flat()
      .forEach(id => { if (id) existing.add(String(id)); });
  }

  // 前回の続き（pageToken）を読み込む
  let pageToken = PropertiesService.getScriptProperties().getProperty("BACKFILL_PAGE_TOKEN") || "";
  const nowIso = new Date().toISOString();

  // 1回の実行で回すページ数（多すぎるとタイムアウトする）
  const PAGES_PER_RUN = 8;  // 50×8=400件ぶん
  let pages = 0;
  let addedTotal = 0;

  while (pages < PAGES_PER_RUN) {
    const plUrl = "https://www.googleapis.com/youtube/v3/playlistItems"
      + "?part=snippet,contentDetails"
      + "&maxResults=50"
      + "&playlistId=" + encodeURIComponent(uploads)
      + (pageToken ? "&pageToken=" + encodeURIComponent(pageToken) : "")
      + "&key=" + encodeURIComponent(apiKey);

    const pl = JSON.parse(UrlFetchApp.fetch(plUrl).getContentText());
    if (!pl.items) break;

    const rows = [];
    for (const it of pl.items) {
      const vid = it.contentDetails?.videoId;
      const sn = it.snippet || {};
      if (!vid) continue;
      if (existing.has(vid)) continue;

      rows.push([
        vid,
        "https://www.youtube.com/watch?v=" + vid,
        sn.title || "",
        sn.description || "",
        sn.publishedAt || "",
        sn.channelTitle || "",
        nowIso
      ]);
      existing.add(vid);
    }

    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      addedTotal += rows.length;
    }

    pageToken = pl.nextPageToken || "";
    pages += 1;

    // 最後まで到達
    if (!pageToken) break;
  }

  // 続きトークンを保存（空なら完了）
  PropertiesService.getScriptProperties().setProperty("BACKFILL_PAGE_TOKEN", pageToken);

  if (!pageToken) {
    Logger.log("BACKFILL DONE. added=" + addedTotal);
  } else {
    Logger.log("BACKFILL CONTINUE. added=" + addedTotal + " nextPageToken exists");
  }
}

function syncMembersPlaylist() {
  const playlistId = PropertiesService.getScriptProperties().getProperty("MEMBERS_PLAYLIST_ID");
  if (!playlistId) throw new Error("Script Properties に MEMBERS_PLAYLIST_ID を入れて");

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("members_videos") || ss.insertSheet("members_videos");

  // videos と同じヘッダ
  if (sh.getLastRow() === 0) {
    sh.appendRow(["videoId","url","title","description","publishedAt","channelTitle","firstSeenAt"]);
  }

  // 重複防止（videoId）
  const existing = new Set();
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, 1).getValues().flat()
      .forEach(id => { if (id) existing.add(String(id)); });
  }

  const nowIso = new Date().toISOString();
  let pageToken = "";
  let added = 0;

  while (true) {
    // Apps Script 高度なサービス YouTube（OAuth）
    const res = YouTube.PlaylistItems.list(
      "snippet,contentDetails",
      { playlistId, maxResults: 50, pageToken: pageToken || undefined }
    );

    const items = res.items || [];
    const rows = [];

    for (const it of items) {
      const vid = it.contentDetails?.videoId;
      if (!vid || existing.has(vid)) continue;

      const sn = it.snippet || {};
      const cd = it.contentDetails || {};

      const title = sn.title || "";
      const desc = sn.description || "";
      const videoPublishedAt = cd.videoPublishedAt || ""; // 動画の公開日（取れるならこれが最適）
      const ownerTitle = sn.videoOwnerChannelTitle || ""; // 元動画の投稿者チャンネル名

      rows.push([
        vid,
        "https://www.youtube.com/watch?v=" + vid,
        title,
        desc,
        videoPublishedAt,
        ownerTitle,
        nowIso
      ]);

      existing.add(vid);
    }

    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      added += rows.length;
    }

    pageToken = res.nextPageToken || "";
    if (!pageToken) break;
  }

  Logger.log("members_videos added=" + added);
}

/***************
 *  準備：列追加
 ***************/
function initSearchColumns_() {
  const sh = SpreadsheetApp.getActive().getSheetByName("videos");
  if (!sh) throw new Error("videos シートが見つからない");

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const need = ["desc_head","desc_core","desc_topics","desc_people","desc_refs"];

  for (const name of need) {
    if (header.indexOf(name) === -1) {
      sh.getRange(1, header.length + 1).setValue(name);
      header.push(name);
    }
  }
}

/*******************
 *  検索用：本文抽出
 *******************/
function makeDescHead_(desc, n) {
  const t = normalizeText_(desc);
  return t.length > n ? t.slice(0, n) : t;
}

function makeDescCore_(desc) {
  let t = normalizeText_(desc);

  // A) 明確なテンプレ開始で切る（desc_core はカッターが有効）
  t = cutAfterMarker_(t, [
    "ポリタスTVの番組は一週間後の19時まで見逃し配信",
    "ポリタスTVの過去の番組アーカイブは下記の有料プランにご加入の上ご視聴ください",
    "【ポリタスTV】毎日（日本時間）午後7時より配信中！",
    "【ポリタスTV】（日本時間）午後7時より配信中！",
  ]);

  // B) 保険：テンプレっぽい開始ワードでも切る
  t = cutAfterRegex_(t, [
    /^ポリタスTVの番組は一週間後/m,
    /^ポリタスTVの過去の番組アーカイブは/m,
    /^【ポリタスTV】（日本時間）/m,
    /^【ポリタスTV】毎日/m,
    /^毎週\s+/m,
    /^ジャーナリストの津田大介/m,
    /^※\s*/m,
  ]);

  // C) 行単位でノイズ除去
  const lines = t.split("\n").map(s => s.trim()).filter(Boolean);

  let inBooks = false;
  const kept = [];

  for (const line of lines) {
    if (/^【今週の未読本】/.test(line)) { inBooks = true; continue; }
    if (inBooks) {
      if (/^【/.test(line)) inBooks = false;
      else continue;
    }

    if (/https?:\/\/|www\./i.test(line)) continue;
    if (/^\d{1,2}:\d{2}(:\d{2})?\s+/.test(line)) continue;
    if ((line.match(/#/g) || []).length >= 2) continue;

    if (/会員限定|二部|メンバー|メンバーシップ|加入|チャンネル登録|見逃し配信|有料アーカイブ|ご視聴|お問い合わせ|配信中/.test(line)) continue;
    if (/^[■▼▶👉]+/.test(line) && line.length < 30) continue;

    kept.push(line);
  }

  let core = kept.join("\n").trim();
  if (core.length < 60) core = makeDescHead_(desc, 300);
  return core;
}

function extractTopics_(desc) {
  const t0 = normalizeText_(desc);

  // topics はフッター文言だけで軽く切る（罫線・メンバー加入は切らない）
  const t = cutAfterMarker_(t0, [
    "ポリタスTVの番組は一週間後の19時まで見逃し配信",
    "ポリタスTVの過去の番組アーカイブは下記の有料プランにご加入の上ご視聴ください",
    "【ポリタスTV】毎日（日本時間）午後7時より配信中！",
    "【ポリタスTV】（日本時間）午後7時より配信中！",
  ]);

  const lines = t.split("\n").map(s => s.trim()).filter(Boolean);

  const keep = [];
  for (const line of lines) {
    if (/^\d{1,2}:\d{2}(:\d{2})?\s+/.test(line)) { keep.push(line); continue; }
    if (/^[0-9]+️⃣/.test(line) && !/https?:\/\//i.test(line)) { keep.push(line); continue; }
    if (/^トピックス|^Topics|^テクノロジー/.test(line)) { keep.push(line); continue; }
  }
  return keep.join("\n").trim();
}

/***************
 * 出演者抽出（フッターで切らない版）
 ***************/
/**
 * 見出し判定（新仕様）
 * 単語: 出演 / 出演者 / 登壇 / ゲスト / パネリスト / 司会 / MC
 *
 * 条件1: あらゆるカッコで単語が囲まれている（前後3文字まで許容）
 * 条件2: 単語のあとにコロン（: / ：）（前後3文字まで許容）
 * 条件3: 単語のみの単独行
 *
 * 戻り値:
 * { hit: boolean, kw: string|null, remainder: string }
 * remainder はコロン見出しの「コロン後ろ文字列」
 */
function detectCastHeader_(line, opts) {
  const o = Object.assign({ maxLineLen: 80 }, opts || {});
  const raw = String(line || "");
  let s = raw.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();

  if (!s) return { hit: false, kw: null, remainder: "" };
  if (s.length > o.maxLineLen) return { hit: false, kw: null, remainder: "" };
  if (/https?:\/\/|www\./i.test(s)) return { hit: false, kw: null, remainder: "" };

  const KEYWORDS = [
    "出演",
    "出演者",
    "登壇",
    "ゲスト",
    "パネリスト",
    "司会",
    "MC"
  ];

  // 条件3: 単語のみ単独行
  for (const kw of KEYWORDS) {
    if (s === kw) return { hit: true, kw, remainder: "" };
  }

  // あらゆるカッコ（開き/閉じ）
  const ob = "\\(（\\[［\\{｛<＜【「『〈《";
  const cb = "\\)）\\]］\\}｝>＞】」』〉》";

  // 条件1: カッコで囲まれている（内側で前後3文字まで許容）
  for (const kw of KEYWORDS) {
    const re = new RegExp(`^[${ob}]\\s*.{0,3}${kw}.{0,3}\\s*[${cb}]\\s*$`);
    if (re.test(s)) return { hit: true, kw, remainder: "" };
  }

  // 条件2: コロン形式（行頭から見て、単語の前後3文字まで許容）
  for (const kw of KEYWORDS) {
    const re = new RegExp(`^.{0,3}${kw}.{0,3}\\s*[:：]\\s*(.*)$`);
    const m = s.match(re);
    if (m) {
      const remainder = String(m[1] || "").trim();
      return { hit: true, kw, remainder };
    }
  }

  return { hit: false, kw: null, remainder: "" };
}

/**
 * 互換：既存コードが isCastHeaderTitle_ を呼ぶ前提なのでラッパーで残す
 */
function isCastHeaderTitle_(s) {
  return detectCastHeader_(s).hit;
}

function looksLikePersonLine_(l) {
  let s = String(l || "").trim();
  if (!s) return false;

  // ★URLが混ざっても人名行を殺さない：URLだけ除去して判定する
  s = s.replace(/https?:\/\/\S+/gi, "").replace(/www\.\S+/gi, "").trim();
  if (!s) return false;

  // 名前（肩書）※同一行で括弧が閉じている
  if (/（[^）]+）/.test(s)) return true;

  // 列挙（区切り揺れを許容）
  const hasSep = /(\s\/\s|\/|／|、|・|,|\||｜)/.test(s);
  if (hasSep && s.length <= 140 && (/[一-龠々]/.test(s) || /[ぁ-んァ-ヶー]/.test(s))) return true;

  // 名前だけ（漢字）
  if (s.length <= 30 && /[一-龠々]/.test(s)) return true;

  // ★ひらがな/カタカナだけの短い名前
  if (s.length <= 20 && /^[ぁ-んァ-ヶー]+$/.test(s)) return true;

  // 英字名だけ（控えめ）
  if (s.length <= 40 && /^[A-Za-z][A-Za-z .'\-]+$/.test(s)) return true;

  return false;
}

// 括弧が改行で割れている行を結合する（例： "名前（肩書\n）" を1行にする）
function joinBrokenParenLine_(lines, j, maxJoinLines) {
  const limit = Math.min(lines.length, j + (maxJoinLines || 4));
  let s = String(lines[j] || "").trim();
  if (!s) return { text: s, nextIndex: j };

  // 「（」はあるが「）」が無い → 次行を結合して「）」が出るまで拾う
  if (s.includes("（") && !s.includes("）")) {
    let k = j;
    while (k + 1 < limit && !s.includes("）")) {
      k++;
      const add = String(lines[k] || "").trim();
      if (!add) continue;
      s = s + add; // 改行を潰して連結
    }
    return { text: s, nextIndex: k };
  }

  return { text: s, nextIndex: j };
}

function extractPeople_(desc) {
  const t0 = normalizeText_(desc);
  const lines = t0.split("\n").map(s => String(s || "").trim());

  // 見出し候補を収集（上から順に）
  const heads = [];
  for (let i = 0; i < lines.length; i++) {
    const h = detectCastHeader_(lines[i]);
    if (h.hit) heads.push({ idx: i, kw: h.kw, remainder: h.remainder });
  }

  // 見出しが無い場合：保険（現状維持だが、ゼロなら空を返す）
  if (heads.length === 0) {
    const keep = [];
    for (const line of lines) {
      if (/（[^）]+）/.test(line) && looksLikePersonLine_(line) && String(line).trim().length <= 160) {
        keep.push(String(line).trim());
      }
    }
    return keep.length ? keep.join("\n").trim() : "";
  }

  // 最大3候補を評価して一番マシを選ぶ
  let bestPeopleLines = []; // ★人名ブロック（見出し以外）だけ保持
  let bestScore = -1e9;

  for (let k = 0; k < Math.min(heads.length, 3); k++) {
    const start = heads[k].idx;

    // ★候補の「人名行」だけを集める（見出し行は入れない）
    const people = [];

    // コロン見出しで「同一行に名前」がある場合は人名候補に入れる
    if (heads[k].remainder) {
      people.push(heads[k].remainder);
    }

    let blankStreak = 0;

    for (let j = start + 1; j < Math.min(start + 120, lines.length); j++) {
      let l = lines[j];

      if (!l) {
        blankStreak++;
        if (blankStreak >= 2) break;
        continue;
      }
      blankStreak = 0;

      // 次の見出しが来たら終了
      if (detectCastHeader_(l).hit) break;

      // 既存の終端条件は維持
      if (/^【/.test(l)) break;
      if (/^［/.test(l) && !detectCastHeader_(l).hit) break;
      if (/^\[/.test(l) && !detectCastHeader_(l).hit) break;
      if (/^●DAY\./.test(l) || /^DAY\.\d/.test(l)) break;
      if (/^BONUS TRACK/i.test(l)) break;
      if (/^--------------------------------------------------------/.test(l)) break;
      if (/^[-－ー]{5,}$/.test(l) || /^={5,}$/.test(l) || /^—{5,}$/.test(l)) break;

      // ★追加：括弧が改行で割れてる行を結合
      const joined = joinBrokenParenLine_(lines, j, 4);
      l = joined.text;
      j = joined.nextIndex;

      // 人名行っぽくなければ終了
      if (!looksLikePersonLine_(l)) break;

      people.push(l);
    }

    // スコアリング（現状維持）
    let score = 0;
    score += (k === 0 ? 20 : 0);
    score += Math.min(people.length, 6) * 5;
    if (people.length === 0) score -= 50;
    if (people.length >= 12) score -= 10;

    const joined = people.join(" ");
    if (people.length >= 3 && (joined.match(/MC|司会/g) || []).length >= people.length) score -= 8;

    if (score > bestScore) {
      bestScore = score;
      bestPeopleLines = people;
    }
  }

  // ★抽出ゼロなら「完全に空白」
  if (!bestPeopleLines || bestPeopleLines.length === 0) return "";

  // ★見出しは返さない（人名行だけ返す）
  return bestPeopleLines.join("\n").trim();
}

function extractRefs_(desc) {
  const t0 = normalizeText_(desc);

  // refs は切らなくてOK（URLだけ拾うのでノイズを受けない）
  const lines = t0.split("\n").map(s => s.trim()).filter(Boolean);
  const urls = [];

  for (const line of lines) {
    const m = line.match(/https?:\/\/\S+/g);
    if (m) urls.push(...m);
  }
  return Array.from(new Set(urls)).join("\n");
}

/*******************
 *  共通ユーティリティ
 *******************/
function ensureSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("videos") || ss.insertSheet("videos");
  if (sh.getLastRow() === 0) {
    sh.appendRow(["videoId","url","title","description","publishedAt","channelTitle","firstSeenAt"]);
  }
  return sh;
}

function getProp_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  if (!v && key === "YT_API_KEY") throw new Error("Script Properties に YT_API_KEY を入れて");
  return v;
}

function setProp_(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, String(value));
}

function normalizeText_(s) {
  return String(s || "")
    .replace(/[\u200B-\u200D\uFEFF]/g, "") // ゼロ幅文字除去
    .replace(/\r\n?/g, "\n")
    .replace(/^"+|"+$/g, "")
    .replace(/[ \t]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function cutAfterMarker_(t, markers) {
  let cutPos = -1;
  for (const m of markers) {
    const p = t.indexOf(m);
    if (p !== -1) cutPos = (cutPos === -1) ? p : Math.min(cutPos, p);
  }
  return (cutPos === -1) ? t : t.slice(0, cutPos).trim();
}

function cutAfterRegex_(t, regexes) {
  let cutPos = -1;
  for (const re of regexes) {
    const m = t.match(re);
    if (m && typeof m.index === "number") {
      cutPos = (cutPos === -1) ? m.index : Math.min(cutPos, m.index);
    }
  }
  return (cutPos === -1) ? t : t.slice(0, cutPos).trim();
}

/**
 * 概要欄：検索用フィールド生成（デフォルト：空欄だけ埋める）
 * - “最新（末尾）から上へ” 進む
 * - 1回の実行で時間いっぱいまで進める（回数を減らす）
 */
function buildDescFieldsBatch() {
  buildDescFieldsBatchBottomUp_(true); // ★空欄のみ
}

/**
 * 概要欄：検索用フィールド生成（全件再生成）
 * - ロジック変更後に過去行も作り直したい時だけ使う
 * - これも “最新（末尾）から上へ”
 */
function buildDescFieldsBatchRebuildAll() {
  // カーソルをリセットして末尾からやり直す
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("DESC_BUILD_CURSOR_BOTTOM");
  buildDescFieldsBatchBottomUp_(false); // ★上書きOK
}

/**
 * 内部：実体（列単位の空欄埋め + desc_people LOCK対応）
 * @param {boolean} onlyBlank true=空欄のみ（列単位） / false=全件上書き
 */
function buildDescFieldsBatchBottomUp_(onlyBlank) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("videos");
  if (!sh) throw new Error("videos シートが見つからない");

  initSearchColumns_();

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = (name) => header.indexOf(name);

  const iDesc   = idx("description");
  const iHead   = idx("desc_head");
  const iCore   = idx("desc_core");
  const iTopics = idx("desc_topics");
  const iPeople = idx("desc_people");
  const iRefs   = idx("desc_refs");

  if (iDesc === -1) throw new Error("description 列が見つからない");
  if (iHead === -1 || iCore === -1 || iTopics === -1 || iPeople === -1 || iRefs === -1) {
    throw new Error("desc_* 列が不足。initSearchColumns_ を確認して。");
  }

  const props = PropertiesService.getScriptProperties();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    Logger.log("DESC BUILD: no data");
    return;
  }

  // ★末尾から上へ：次に処理する「終端行」
  let cursorRow = parseInt(props.getProperty("DESC_BUILD_CURSOR_BOTTOM") || String(lastRow), 10);
  if (!Number.isFinite(cursorRow) || cursorRow < 2) cursorRow = lastRow;
  if (cursorRow > lastRow) cursorRow = lastRow;

  // 1回の実行で使う時間（GASの実行制限に当てないため）
  const started = Date.now();
  const TIME_LIMIT_MS = 5 * 60 * 1000; // 5分
  const CHUNK = 300; // 1回の読み書き単位

  let processedTotal = 0;
  let skippedTotal = 0;
  let windows = 0;

  while (cursorRow >= 2 && (Date.now() - started) < TIME_LIMIT_MS) {
    const startRow = Math.max(2, cursorRow - CHUNK + 1);
    const endRow = cursorRow;
    const numRows = endRow - startRow + 1;

    const rng = sh.getRange(startRow, 1, numRows, header.length);
    const rows = rng.getValues();

    // ★desc_people 列のメモ（LOCK判定用）
    const peopleNotes = sh.getRange(startRow, iPeople + 1, numRows, 1).getNotes();

    let processed = 0;
    let skipped = 0;

    for (let r = 0; r < rows.length; r++) {
      const desc = String(rows[r][iDesc] || "");

      // onlyBlank=true は「列単位で空欄だけ埋める」
      // onlyBlank=false は「全件上書き」ただし desc_people の LOCK は常に尊重

      // desc_head
      if (!onlyBlank || !rows[r][iHead]) {
        rows[r][iHead] = makeDescHead_(desc, 600);
        processed++;
      } else {
        skipped++;
      }

      // desc_core
      if (!onlyBlank || !rows[r][iCore]) {
        rows[r][iCore] = makeDescCore_(desc);
        processed++;
      } else {
        skipped++;
      }

      // desc_topics
      if (!onlyBlank || !rows[r][iTopics]) {
        rows[r][iTopics] = extractTopics_(desc);
        processed++;
      } else {
        skipped++;
      }

      // desc_people（LOCK対応）
      const lockedPeople = String(peopleNotes[r][0] || "").toUpperCase().includes("LOCK");
      if (!lockedPeople) {
        if (!onlyBlank || !rows[r][iPeople]) {
          rows[r][iPeople] = extractPeople_(desc);
          processed++;
        } else {
          skipped++;
        }
      } else {
        skipped++;
      }

      // desc_refs
      if (!onlyBlank || !rows[r][iRefs]) {
        rows[r][iRefs] = extractRefs_(desc);
        processed++;
      } else {
        skipped++;
      }
    }

    // 列単位更新なので常に書き戻す
    rng.setValues(rows);

    processedTotal += processed;
    skippedTotal += skipped;
    windows++;

    // 次はこのウィンドウのさらに上
    cursorRow = startRow - 1;
    props.setProperty("DESC_BUILD_CURSOR_BOTTOM", String(cursorRow));
  }

  if (cursorRow < 2) {
    props.deleteProperty("DESC_BUILD_CURSOR_BOTTOM");
    Logger.log(`DESC BUILD DONE. processed=${processedTotal} skipped=${skippedTotal} windows=${windows} onlyBlank=${onlyBlank}`);
  } else {
    Logger.log(`DESC BUILD CONTINUE. processed=${processedTotal} skipped=${skippedTotal} windows=${windows} nextCursor=${cursorRow} onlyBlank=${onlyBlank}`);
  }
}

/**
 * 直近N行だけ「全部上書き」したい時用（例：100）
 * 使い方：rebuildDescFieldsRecent100()
 */
function rebuildDescFieldsRecent100() {
  rebuildDescFieldsRecent_(3000);
}

function rebuildDescFieldsRecent_(n) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("videos");
  if (!sh) throw new Error("videos シートが見つからない");

  initSearchColumns_();

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const idx = (name) => header.indexOf(name);

  const iDesc   = idx("description");
  const iHead   = idx("desc_head");
  const iCore   = idx("desc_core");
  const iTopics = idx("desc_topics");
  const iPeople = idx("desc_people");
  const iRefs   = idx("desc_refs");

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const nn = Math.max(1, parseInt(n, 10) || 100);
  const startRow = Math.max(2, lastRow - nn + 1);
  const numRows = lastRow - startRow + 1;

  const rng = sh.getRange(startRow, 1, numRows, header.length);
  const rows = rng.getValues();

  // ★desc_people 列のメモ（LOCK判定用）
  const peopleNotes = sh.getRange(startRow, iPeople + 1, numRows, 1).getNotes();

  for (let r = 0; r < rows.length; r++) {
    const desc = String(rows[r][iDesc] || "");

    rows[r][iHead]   = makeDescHead_(desc, 600);
    rows[r][iCore]   = makeDescCore_(desc);
    rows[r][iTopics] = extractTopics_(desc);

    // ★LOCKされている desc_people は上書き禁止
    const lockedPeople = String(peopleNotes[r][0] || "").toUpperCase().includes("LOCK");
    if (!lockedPeople) {
      rows[r][iPeople] = extractPeople_(desc);
    }

    rows[r][iRefs]   = extractRefs_(desc);
  }

  rng.setValues(rows);
  Logger.log(`rebuild recent: ${startRow}-${lastRow} (n=${nn})`);
}
