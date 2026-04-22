import { google } from "googleapis";
import { chromium } from "playwright";

const SERVER_ID = 2002;
const RACE = 2;

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME || "Sheet1";
const GOOGLE_SERVICE_ACCOUNT_JSON = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
const MAX_ROWS = Number(process.env.MAX_ROWS || 300);

if (!SPREADSHEET_ID) throw new Error("SPREADSHEET_ID is missing");
if (!GOOGLE_SERVICE_ACCOUNT_JSON) throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is missing");

function nowKSTString() {
  const now = new Date();
  const kst = new Date(now.getTime() + 9 * 60 * 60 * 1000);
  return kst.toISOString().slice(0, 19).replace("T", " ");
}

function stripHtml(text) {
  return String(text || "").replace(/<[^>]*>/g, "").trim();
}

function buildCharacterPageUrl(serverId, characterId) {
  return `https://aion2.plaync.com/ko-kr/characters/${serverId}/${characterId}`;
}

async function getSheetsClient() {
  const credentials = JSON.parse(GOOGLE_SERVICE_ACCOUNT_JSON);

  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });

  return google.sheets({ version: "v4", auth });
}

async function readNicknames(sheets) {
  const range = `${SHEET_NAME}!B3:B${MAX_ROWS + 2}`;
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range
  });

  const values = res.data.values || [];
  const rows = [];

  for (let i = 0; i < MAX_ROWS; i++) {
    const cell = values[i]?.[0] ?? "";
    rows.push(String(cell).trim());
  }

  return rows;
}

async function fetchCharacterData(page, nickname) {
  if (!nickname) {
    return ["", "", "", "", ""];
  }

  const searchUrl =
    "https://aion2.plaync.com/ko-kr/api/search/aion2/search/v2/character" +
    `?keyword=${encodeURIComponent(nickname)}` +
    `&race=${RACE}` +
    `&serverId=${SERVER_ID}` +
    "&page=1&size=30";

  const searchJson = await page.evaluate(async (url) => {
    const res = await fetch(url, {
      method: "GET",
      headers: { "Accept": "application/json" },
      credentials: "include"
    });
    return await res.json();
  }, searchUrl);

  const list = Array.isArray(searchJson?.list) ? searchJson.list : [];
  if (!list.length) {
    return ["검색결과없음", "", "", nowKSTString(), ""];
  }

  const picked =
    list.find((x) => stripHtml(x.name).toLowerCase() === nickname.toLowerCase()) ||
    list[0];

  const characterId = String(picked?.characterId || "").trim();
  const serverId = String(picked?.serverId || SERVER_ID).trim();

  if (!characterId) {
    return ["ID없음", "", "", nowKSTString(), ""];
  }

  const characterPageUrl = buildCharacterPageUrl(serverId, characterId);

  await page.goto(characterPageUrl, {
    waitUntil: "domcontentloaded",
    timeout: 60000
  });

  await page.waitForTimeout(1500);

  const detailUrl =
    "https://aion2.plaync.com/api/character/info?lang=ko" +
    `&characterId=${characterId}` +
    `&serverId=${serverId}`;

  const detailJson = await page.evaluate(async (url) => {
    const res = await fetch(url, {
      method: "GET",
      headers: { "Accept": "application/json" },
      credentials: "include"
    });
    return await res.json();
  }, detailUrl);

  const className = detailJson?.profile?.className || "";
  const combatPower = detailJson?.profile?.combatPower || "";

  const statList = Array.isArray(detailJson?.stat?.statList)
    ? detailJson.stat.statList
    : [];

  const itemLevel =
    statList.find((x) => String(x?.name || "").trim() === "아이템레벨")?.value || "";

  if (!className && !combatPower && !itemLevel) {
    return ["상세빈값", "", "", nowKSTString(), `=HYPERLINK("${characterPageUrl}","정보보기")`];
  }

  return [
    className,
    itemLevel,
    combatPower,
    nowKSTString(),
    `=HYPERLINK("${characterPageUrl}","정보보기")`
  ];
}

async function writeResults(sheets, rows) {
  const startRow = 3;
  const endRow = startRow + rows.length - 1;

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!C${startRow}:G${endRow}`,
    valueInputOption: "USER_ENTERED",
    requestBody: {
      values: rows
    }
  });
}

async function main() {
  const sheets = await getSheetsClient();
  const nicknames = await readNicknames(sheets);

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    locale: "ko-KR",
    timezoneId: "Asia/Seoul",
    userAgent:
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36"
  });

  const page = await context.newPage();

  await page.goto("https://aion2.plaync.com/ko-kr/characters/index", {
    waitUntil: "domcontentloaded",
    timeout: 60000
  });

  const results = [];
  for (const nickname of nicknames) {
    try {
      const row = await fetchCharacterData(page, nickname);
      results.push(row);
      console.log(nickname || "(blank)", row);
      await page.waitForTimeout(800);
    } catch (err) {
      console.error("row failed:", nickname, err);
      results.push([`오류: ${err.message}`, "", "", nowKSTString(), ""]);
    }
  }

  await writeResults(sheets, results);
  await browser.close();
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});