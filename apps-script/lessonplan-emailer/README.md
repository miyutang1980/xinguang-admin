# Lesson Plan 老師入口 Gmail 寄送自動化

## 用途

太平新光分校 Lesson Plan 老師入口 Gmail 寄送自動化。當老師在 Google Sheet「老師入口」選好班級與週次後，系統自動從 `LessonPlan_Index` 找到對應備課包，透過 Gmail 寄給該班老師，並寫入「寄送紀錄」避免重複寄送。

## Google Sheets 需要的分頁

本 Apps Script 需搭配一份 Google Sheet，並具備以下分頁：

| 分頁名稱 | 用途 |
| --- | --- |
| `老師入口` | 老師選擇班級（B4）、週次（B5）、顯示課程（E4）、單元（E5）、Lesson Plan 連結（B7） |
| `LessonPlan_Index` | 備課包索引，欄位需包含 `lookup_key`、`class_id`、`week_no`、`course`、`unit_range`、`file_name`、`google_doc_link` |
| `表單設定` | key/value 設定表，key 包含 `lesson_log_form_link`、`test_score_form_link`、`observation_form_link`、`enable_auto_send` |
| `教師Email` | 每班對應的寄送對象，欄位需包含 `class_id`、`teacher_name`、`email` |
| `寄送紀錄` | 系統自動建立；紀錄每次寄送的 `send_key`（格式：`班級|週次`），避免同班同週重複寄送 |

## 安裝與部署步驟

1. 將工作簿以 Google Sheet 形式開啟（若為 `.xlsx`，請先用 Google Drive 另存為 Google 試算表格式）。
2. 開啟該 Google Sheet → 「擴充功能」→「Apps Script」。
3. 將 `Code.gs` 內容貼進 Apps Script 編輯器（或若專案已設定 [`clasp`](https://github.com/google/clasp)，用 `clasp push` 同步）。
4. 儲存後，在編輯器中執行 `setupLessonPlanAutomation()`。
5. 依指示授權 Gmail 及 Spreadsheet 權限（第一次需要 Google 帳號確認）。
6. 回到 Google Sheet，試算表選單會出現「老師備課包」。
7. 測試方式：
   - 在「教師Email」分頁填入一筆測試班級的 email。
   - 在「老師入口」選擇該班級與週次。
   - 使用選單「老師備課包 → 寄出目前班級週次」確認可正常送出。
   - 檢查「寄送紀錄」分頁是否新增一筆。

## 測試模式（預設啟用）

目前程式預設啟用「測試模式」，用於驗收階段避免誤寄給真實老師。設定位置在 `Code.gs` 頂部的 `TEST_MODE_CONFIG`：

```js
const TEST_MODE_CONFIG = {
  test_mode_enabled: true,
  test_recipient_email: 'miyutang1980@gmail.com',
};
```

測試模式啟用時：

- 不論「教師Email」分頁中老師的 email 是什麼，**所有** Gmail 寄送（自動寄送與手動寄出皆同一條路徑）都會改寄到 `test_recipient_email`。
- 信件主旨會加上 `[測試模式｜原收件人：xxx@xxx]` 標記；信件內容最上方會顯示黃色提醒橫幅，標示原訂收件人姓名與 email。
- 「寄送紀錄」分頁會同時記錄：
  - `email`、`intended_email`：原本應寄送的老師 email（稽核用）
  - `actual_email`：實際寄出的 Gmail 收件人（測試模式下會是 `test_recipient_email`）
  - `test_mode`：`TRUE` / `FALSE`
- 自動寄送的重複寄送防護（`send_key = class_id|week_no`）仍有效；若要重寄同班同週，請使用「重設目前選擇寄送紀錄」選單。

### 切換到正式寄送（Production）

完成測試後請依序操作：

1. 到「教師Email」分頁，**逐筆確認每個 `class_id` 的 `teacher_name` 與 `email` 正確無誤**（這是正式寄送對象，錯一個就會寄錯人）。
2. 編輯 `Code.gs`，將 `TEST_MODE_CONFIG.test_mode_enabled` 由 `true` 改為 `false`：
   ```js
   const TEST_MODE_CONFIG = {
     test_mode_enabled: false,
     test_recipient_email: 'miyutang1980@gmail.com',
   };
   ```
3. 儲存並同步到 Apps Script 編輯器（或使用 `clasp push`）。
4. 建議先用「寄出目前班級週次」對一個已確認 email 的班級做最終驗證，確認實際收件人已變回該班老師 email，且「寄送紀錄」中 `actual_email` 等於老師 email、`test_mode = FALSE`。

> 安全提醒：請勿同時保留測試模式並對外發送重要通知；正式切換前務必完成第 1 步的「教師Email」資料複查。

## 選單功能

試算表開啟後，上方會出現「老師備課包」自訂選單，包含以下三個項目：

- **寄出目前班級週次**（對應 `sendCurrentSelectionManual`）：不受 auto-send 開關與重複寄送限制，直接寄出目前「老師入口」選好的班級與週次。
- **設定自動寄送觸發器**（對應 `setupLessonPlanAutomation`）：重建 onEdit trigger，讓老師在「老師入口」更改班級或週次時自動寄出。
- **重設目前選擇寄送紀錄**（對應 `resetCurrentSelectionSendLog`）：刪除「寄送紀錄」中對應的 `send_key`，允許重新自動寄送。

## 選單疑難排解

若試算表上方看不到「老師備課包」選單（常見於剛貼上新版 `Code.gs`、重新授權後、或 `onOpen` 尚未執行完成的情況），請依序嘗試下列步驟：

1. 先回到 Google Sheet 整個頁籤重新整理一次（`Ctrl/Cmd + R`），讓 Apps Script 重新觸發 `onOpen(e)` 建立選單。
2. 若重新整理後仍沒有「老師備課包」選單，回到 Apps Script 編輯器，從函式下拉選單選擇 **`forceCreateMenu`** 並按執行；執行完成後回到 Google Sheet 再重新整理一次，選單應該就會出現。
3. 若 `forceCreateMenu()` 之後重新整理仍看不到選單（例如帳號授權尚未完成、或 UI 暫時無法載入自訂選單），可以暫時略過選單，直接在 Apps Script 編輯器中選擇 **`sendCurrentSelectionManual`** 函式並按執行，即可寄出目前「老師入口」選好的班級週次備課包；事後再排查選單為何未載入即可。

## 安全性注意事項

- **老師 email 存在 Google Sheet 的「教師Email」分頁**，不要把實際 email 寫進本 repo。
- 不要 commit 任何憑證、token、service account key、學生個資或老師個資。
- 本 repo 僅保存程式邏輯；所有個資、連結與設定值皆應留在 Google Sheet 中。
- 重複寄送防護：自動模式以「寄送紀錄」分頁的 `send_key = class_id|week_no` 檢查；要重寄時使用「重設本班本週寄送紀錄」或切換到手動寄出。
- `表單設定` 中的 `enable_auto_send = FALSE` 可全域關閉自動寄送（手動寄出不受影響）。

## 相關 Google Drive 檔案

- Apps Script 對應的試算表（寄送工具所在 workbook）：
  <https://docs.google.com/spreadsheets/d/1NQXkoIsZSSAtSmNQg5uDZRpA7aF2ntNN/edit?usp=drivesdk&ouid=118041506256640130565&rtpof=true&sd=true>
- `Code.gs` 在 Google Drive 上的備份：
  <https://drive.google.com/file/d/1ab_AraLKmN6DsCTPsK7IEGc4dAxCnzKK/view?usp=drivesdk>

## 維護建議

- 有任何邏輯調整，請以此 repo 為單一事實來源，改完再同步到 Apps Script 編輯器。
- 若未來導入 `clasp`，建議在本資料夾加上 `.clasp.json`（記得列入 `.gitignore` 以避免 commit `scriptId`）與 `appsscript.json` 設定檔。
