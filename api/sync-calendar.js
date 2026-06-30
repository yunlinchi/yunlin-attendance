// =========================================================================
// Vercel Serverless Function：主管核准後，自動把記錄寫入「數辦公用 Google 行事曆」
// 路徑：/api/sync-calendar   （Vercel 會自動把 /api 目錄變成後端函式）
//
// 安全設計：前端只傳「類型 + 文件ID + appId」，後端用服務帳戶自己去 Firestore
// 讀該筆記錄，確認狀態確實為「核准 / 已核准」後，才寫入行事曆。
// 因此這支 API 無法被用來偽造任意行事曆事件。
//
// 需要的環境變數（全部設定在 Vercel → Settings → Environment Variables）：
//   GOOGLE_PROJECT_ID    服務帳戶所在的 GCP/Firebase 專案 ID（例：yunlin-digital-travel）
//   GOOGLE_CLIENT_EMAIL  服務帳戶 email（xxx@xxx.iam.gserviceaccount.com）
//   GOOGLE_PRIVATE_KEY   服務帳戶私鑰（PEM，整段貼上，含 -----BEGIN...，換行用 \n）
//   OFFICE_CALENDAR_ID   數辦公用行事曆的「行事曆 ID」（在行事曆設定頁可找到）
// =========================================================================

import admin from 'firebase-admin';
import { google } from 'googleapis';

// 把環境變數裡以 \n 表示的換行還原成真正換行
const PRIVATE_KEY = (process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');
const CLIENT_EMAIL = process.env.GOOGLE_CLIENT_EMAIL;
const PROJECT_ID = process.env.GOOGLE_PROJECT_ID;
const CALENDAR_ID = process.env.OFFICE_CALENDAR_ID;
const TIME_ZONE = 'Asia/Taipei';

// --- Firebase Admin 初始化（serverless 會重用，所以要避免重複初始化）---
function getDb() {
  if (!admin.apps.length) {
    admin.initializeApp({
      credential: admin.credential.cert({
        projectId: PROJECT_ID,
        clientEmail: CLIENT_EMAIL,
        privateKey: PRIVATE_KEY,
      }),
    });
  }
  return admin.firestore();
}

// --- 用同一組服務帳戶取得「寫行事曆」用的 Google Calendar client ---
// 若設定了 GOOGLE_IMPERSONATE_SUBJECT，服務帳戶會「假扮」該校內帳號（網域委派 DWD），
// 用於 Workspace 禁止對外部共用行事曆的情況；未設定則走一般共用模式。
function getCalendar() {
  const subject = process.env.GOOGLE_IMPERSONATE_SUBJECT || undefined;
  const auth = new google.auth.JWT({
    email: CLIENT_EMAIL,
    key: PRIVATE_KEY,
    scopes: ['https://www.googleapis.com/auth/calendar.events'],
    subject, // 網域委派：假扮成此校內帳號（需在 Workspace 後台授權）
  });
  return google.calendar({ version: 'v3', auth });
}

// 民國年 → 西元年（與前端 convertToWestern 相同邏輯）
function toWesternDate(dateStr) {
  if (!dateStr) return '';
  const parts = String(dateStr).split('-');
  if (parts.length === 3) {
    let year = parseInt(parts[0], 10);
    if (year < 1000) year += 1911;
    return `${year}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
  }
  return dateStr;
}

// 全天事件的結束日是「不含當天」，需 +1 天（用 UTC 計算避免時區位移）
function addOneDay(dateStr) {
  const d = new Date(`${dateStr}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + 1);
  return d.toISOString().slice(0, 10);
}

function normalizeTime(t, fallback) {
  if (!t) return fallback;
  // 接受 "8:00" / "08:00" / "0800"
  const clean = String(t).replace(/[^0-9:]/g, '');
  if (clean.includes(':')) {
    const [h, m] = clean.split(':');
    return `${h.padStart(2, '0')}:${(m || '00').padStart(2, '0')}`;
  }
  if (clean.length === 4) return `${clean.slice(0, 2)}:${clean.slice(2)}`;
  return fallback;
}

// 依記錄內容組出 Google 行事曆事件
function buildEvent(type, data) {
  if (type === 'leave') {
    const startDate = toWesternDate(data.startDate || data.date);
    const endDate = toWesternDate(data.endDate || data.startDate || data.date);
    const startTime = normalizeTime(data.startTime, '08:00');
    const endTime = normalizeTime(data.endTime, '17:00');
    const title = `【請假-${data.leaveType || '出差'}】${data.applicant || ''}（代理人：${data.agent || '無'}）`;
    let description =
      `🕒 時間：${startDate} ${startTime} ~ ${endDate} ${endTime}\n` +
      `請假同仁：${data.applicant || ''}\n` +
      `請假類別：${data.leaveType || '出差公出'}\n` +
      `時數：${data.hours || 0} 小時\n` +
      `事由：${data.reason || '未填寫'}\n` +
      `核准主管：${data.approver || ''}`;
    if (data.leaveType === '公假') {
      description += `\n起訖地點：${data.location || ''}\n出差旅費：${data.isBusinessTrip ? '可請領差旅費' : '不具出差性質'}`;
    }
    return {
      summary: title,
      description,
      location: data.location || data.destination || '',
      // 以「全天事件」呈現（避免時數非整天時在行事曆上變成一個小點）；實際時間寫在備註
      start: { date: startDate },
      end: { date: addOneDay(endDate) },
    };
  }

  // type === 'overtime'
  const workDate = toWesternDate(data.workDate || data.date);
  const startTime = normalizeTime(data.startTime, '08:00');
  const endTime = normalizeTime(data.endTime, '17:00');
  const people = Array.isArray(data.participants) ? data.participants.join('、') : (data.applicant || '');
  return {
    summary: `【加班-補休】${data.activityName || '假日加班'}（${people}）`,
    description:
      `🕒 時間：${workDate} ${startTime} ~ ${endTime}\n` +
      `活動／加班事由：${data.activityName || ''}\n` +
      `加班同仁：${people}\n` +
      `時數：${data.hours || 0} 小時\n` +
      `核准主管：${data.approver || ''}`,
    location: data.location || '',
    // 以「全天事件」呈現；實際時間寫在備註
    start: { date: workDate },
    end: { date: addOneDay(workDate) },
  };
}

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ ok: false, error: '只接受 POST' });
  }

  // 啟動前先檢查環境變數，給出清楚錯誤
  if (!PRIVATE_KEY || !CLIENT_EMAIL || !PROJECT_ID || !CALENDAR_ID) {
    return res.status(500).json({
      ok: false,
      error: '伺服器尚未設定完整環境變數（GOOGLE_PROJECT_ID / GOOGLE_CLIENT_EMAIL / GOOGLE_PRIVATE_KEY / OFFICE_CALENDAR_ID）',
    });
  }

  try {
    const body = typeof req.body === 'string' ? JSON.parse(req.body) : (req.body || {});
    const { type, id, appId } = body;

    if (!['leave', 'overtime'].includes(type) || !id || !appId) {
      return res.status(400).json({ ok: false, error: '參數不足：需要 type(leave|overtime)、id、appId' });
    }

    const collectionName = type === 'leave' ? 'leaves' : 'overtimes';
    const approvedStatus = type === 'leave' ? '核准' : '已核准';

    const db = getDb();
    const docRef = db
      .collection('artifacts').doc(appId)
      .collection('public').doc('data')
      .collection(collectionName).doc(id);

    const snap = await docRef.get();
    if (!snap.exists) {
      return res.status(404).json({ ok: false, error: '找不到該筆記錄' });
    }
    const data = snap.data();

    // 關鍵安全檢查：只有「已核准」的記錄才允許寫入行事曆
    if (data.status !== approvedStatus) {
      return res.status(403).json({ ok: false, error: `此記錄尚未核准（目前狀態：${data.status}），不予同步` });
    }

    const calendar = getCalendar();
    const event = buildEvent(type, data);

    let result;
    if (data.calendarEventId) {
      // 已同步過 → 更新事件，避免重複
      result = await calendar.events.patch({
        calendarId: CALENDAR_ID,
        eventId: data.calendarEventId,
        requestBody: event,
      });
    } else {
      result = await calendar.events.insert({
        calendarId: CALENDAR_ID,
        requestBody: event,
      });
      // 回寫事件 ID 到 Firestore，作為「已同步」標記與防重複依據
      await docRef.set(
        { calendarEventId: result.data.id, calendarSyncedAt: new Date().toISOString() },
        { merge: true }
      );
    }

    return res.status(200).json({ ok: true, eventId: result.data.id, htmlLink: result.data.htmlLink });
  } catch (e) {
    console.error('行事曆同步失敗：', e);
    return res.status(500).json({ ok: false, error: e.message || '未知錯誤' });
  }
}
