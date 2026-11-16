/* 設定檔：把 Apps Script 部署後的 /exec 網址貼在 sheetApi 的引號裡。
 * 留空＝本機模式（只能看介面，不能批改）。改完請把 index.html 裡 config.js?v= 的數字加 1。
 */
window.PZ_CONFIG = {
  sheetApi: "",
  // 疑似亂答：某一話用時少於幾秒（只是提醒）
  suspectSeconds: 20
};
