// =========================================================================
// 🌐 雲端人事系統 - 企業個別環境設定檔 (env.js)
// 💡 提示：不同公司部署時，僅需修改此檔案的參數即可，核心主程式免變更。
// =========================================================================

const SYSTEM_ENV = {
  // 1. 各公司專屬的 Google Apps Script 網路應用程式 URL
  GAS_WEB_APP_URL: "https://script.google.com/macros/s/xxxx/exec",

  // 2. 各公司專屬的 LINE LIFF 設定 (用於前端身分綁定或員工打卡端跳轉)
  LIFF_ID: "1234567890-XXXXXXXX",                      // 填入該公司的 LIFF ID
  LIFF_URL: "https://liff.line.me/1234567890-XXXXXXXX",  // 填入該公司的 LIFF 專屬網址

  // 3. 當前測試或預設登入的 LINE UID (正式對接 LIFF 後會改由 LINE 登入自動動態取得)
  CURRENT_LINE_UID: "2010383109-EolilvJ0",

  // 4. 企業客製化設定 (方便未來擴充給不同公司)
  COMPANY_NAME: "鼎新科技管理部",
  THEME_COLOR: "indigo", // 可用來動態切換主題色彩
  VERSION: "v2.1-CH"
};