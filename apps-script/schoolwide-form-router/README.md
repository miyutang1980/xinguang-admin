# 太平新光分校全校版 Google Forms 回應自動分流系統

## 用途

將 Google Forms 的問卷回應自動分流寫入全校版 Google 試算表對應的資料分頁：

- `Lesson_Log`：老師上課紀錄
- `Quiz_Result`：小考 / 測驗成績
- `Parent_Feedback_Biweekly`：雙週家長回饋（聯絡簿）

系統同時會讀取 `Class_Master` / `Teacher_Master` / `Student_Master` 做主檔校驗與 ID 補齊，並保留原始回應、分流紀錄與錯誤紀錄，方便日後稽核。

## Google Sheets 需要的分頁

本 Apps Script 需搭配一份全校版 Google Sheet，並具備以下分頁：

### 主檔分頁（需先建立並維護）

| 分頁名稱 | 用途 |
| --- | --- |
| `Class_Master` | 班級主檔，提供 `class_id`、班級名稱、導師等資訊 |
| `Teacher_Master` | 老師主檔，提供 `teacher_id`、姓名、Email |
| `Student_Master` | 學生主檔，提供 `student_id`、姓名、所屬 `class_id` |

### 目標資料分頁（分流寫入）

| 分頁名稱 | 用途 |
| --- | --- |
| `Lesson_Log` | 老師每週上課紀錄的主要資料表 |
| `Quiz_Result` | 小考、測驗成績記錄 |
| `Parent_Feedback_Biweekly` | 雙週家長回饋 / 聯絡簿 |

### 系統自動建立的 Log 分頁

首次執行 `setupFormRouter()` 時，以下分頁若不存在會自動建立：

| 分頁名稱 | 用途 |
| --- | --- |
| `Form_Response_Raw` | 保留所有 Google Forms 原始回應（含時間戳、來源表單、原始 JSON） |
| `Form_Routing_Log` | 每筆回應成功分流到哪個目標分頁、對應的 row id |
| `Form_Error_Log` | 分流失敗、欄位缺漏或主檔找不到對應 ID 的錯誤紀錄 |

## 安裝與部署步驟

1. 將工作簿以 Google Sheet 形式開啟（若為 `.xlsx`，請先用 Google Drive 另存為 Google 試算表格式）。
2. 開啟該 Google Sheet → 「擴充功能」→「Apps Script」。
3. 將 `Code.gs` 內容完整貼進 Apps Script 編輯器。
4. 儲存後，在編輯器中執行 `setupFormRouter()`。
5. 依指示授權 Spreadsheet 權限。
6. 將 Google Forms 的回應表連結到同一份 Google Sheet（Forms → 回應 → 連結到試算表 → 選擇現有試算表）。
7. 設定 Forms 提交觸發條件：回到 Apps Script → 觸發條件 → 新增 `onFormSubmitRouter`，事件來源選「試算表」、事件類型選「提交表單時」。

## 表單設計要求

為了讓分流判斷正確，建議在每份 Google Form 新增一個隱藏或預填的欄位 `form_type`，值可使用以下任一：

- `lesson_log` / `lesson log` / `上課紀錄` / `課堂紀錄` / `lesson`
- `quiz_result` / `quiz result` / `小考成績` / `測驗成績` / `quiz`
- `parent_feedback` / `parent feedback` / `家長回饋` / `聯絡簿` / `contactbook`

若表單沒有 `form_type` 欄位，系統會改用欄位名稱特徵（例如是否有 `week_no`、`listening_score`、`parent_reply` 等）自動判斷分流目標；請盡量使用可辨識的欄位名稱，或直接加上 `form_type` 以確保準確。

## 測試方式

在 Apps Script 編輯器內依序執行以下測試函式，驗證三種表單分流都能正常寫入：

- `testRouteLessonLog()`：模擬一筆 Lesson_Log 回應，應寫入 `Lesson_Log` 並在 `Form_Routing_Log` 留下紀錄。
- `testRouteQuizResult()`：模擬一筆 Quiz_Result 回應，應寫入 `Quiz_Result`。
- `testRouteParentFeedback()`：模擬一筆 Parent_Feedback 回應，應寫入 `Parent_Feedback_Biweekly`。

若測試失敗，請檢查 `Form_Error_Log` 了解錯誤原因（常見為主檔尚未建立 `class_id` / `teacher_id` / `student_id` 或欄位名稱未對齊）。

## 安全性說明

- 本 Apps Script **不會寄送任何 Email**，僅在試算表內附加資料列與寫入錯誤紀錄。
- 所有失敗情境（主檔查不到、欄位缺漏、型別錯誤）都會保留到 `Form_Error_Log`，不會中斷後續表單提交。
- 原始回應一律完整保留在 `Form_Response_Raw`，即使分流失敗也可以日後重跑。
- 建議把 Google Sheet 的分享權限控制在「全校行政管理者」層級，並定期檢視 `Form_Error_Log`。
