# 太平新光分校：系統連結與操作說明書

更新日期：2026-09-23。本文件與後台文件中心、操作說明書及連結中心共用資料來源；「待部署」不代表已上線。

## 版本與上線狀態

本手冊更新至 2026-09-23，涵蓋本後台、公開網站、Gateway、家長請假與接送流程；Eagle TMS／WriteC 等獨立教學平台不在本後台的維護範圍，不憑舊資料補入未知入口。

四項獨立勾選版已由管理者回報部署；小巴下午班與學生名單已由管理者確認恢復。家長中心 v2 的手機改版、驗證綁定、即時分流已完成隔離測試，尚待正式 Gateway 與網站成對切版。

版本依實際 API 回應與正式頁面驗收，不因手冊更新而視為已上線。最近一次 Gateway 唯讀查核回報 leave-services-v1。家長中心新功能及 5 分鐘重試觸發器仍列待驗收；GitHub 審核單 #2 保留部署閘門。

38 項新版／既有隔離檢查通過，不代表真實家長登入或群組收件已通過。既有每日彙整函式保留，觸發器是否啟用、實際收件須在 Apps Script 與 LINE 分別核對。

## 群組名稱與通知分流

顯示名稱統一為「美語課請假」與「課輔/托育請假」。此為原 LINE 群組改名，不建立新群、不更換群組 ID、不搬移既有通知紀錄。

家長仍分別勾選取消接送、美語、安親與托育；課輔/托育是收件群名稱，不把安親與托育合併成單一請假項目。

新版上線後：新增通知家長與所選服務群；修改通知異動前後涉及的群；取消通知原相關群。安親、托育同時勾選，課輔/托育請假群只收到一次。

TEACHING_ANNOUNCE_GROUP_ID 為檢測／體驗課備課通知，和 TEACHING_GROUP_ID 美語課請假不同，不能因「教學」二字相同就互換。

LINE 已受理只表示訊息 API 接受，不代表老師已讀；封鎖、未加好友或群內無機器人等情形須看實際收件端。群名改動本身不需重建每日排程。

| 項目 | 統一群名 | 原設定鍵 | 規則 |
| --- | --- | --- | --- |
| 取消接送 | 接送群 | TRANSPORT_GROUP_ID | 保留原群組與路線流程 |
| 美語請假 | 美語課請假 | TEACHING_GROUP_ID | 原教學請假群；不是招生預約教學公告群 |
| 安親請假、托育請假 | 課輔/托育請假 | DAYCARE_GROUP_ID | 兩項同時勾選只通知同一群一次 |

## 家長請假與取消接送

家長從官方 LINE 的請假入口進入，選孩子、日期與要調整的服務。只請美語假不會自動取消接送；未勾選的項目照常。

新版 parent-center-v2 待部署後：已綁定 LINE 身分直接顯示有效在校孩子；未綁定者走校方一次性綁定碼，不以知道電話為身分證明。

新版每次最多 5 位孩子、14 天，總計不超過 20 筆；可選今天起 180 天內日期。區間逐日建立，包含週末，送出前請家長核對。補課聯繫只適用美語項目。

送出前看確認內容，成功後看紀錄與通知狀態。等待重試不等於紀錄沒寫入；網路中斷請重試同一請求，不另建一筆。正式版只在目前分頁暫存未確認內容，不儲存 LINE token。

本人建立、尚未彙整鎖定的紀錄可依日期規則修改或取消；當日 11:30 後請直接聯絡校方。查不到孩子或紀錄不能猜測他人學生編號。

取消接送不等於整天不上課；課程請假也不等於接送取消。臨時接送或學生安全問題請直接電話聯繫 04-2396-0585，不只等系統訊息。

## 首次綁定與校方核驗

新版部署後，行政／管理員使用家長綁定碼核發頁，依既有校方聯絡資料確認家長身分，輸入學生編號並選正確家長欄位。

綁定碼有效 30 分鐘，只交給核驗過的家長。重新核發使相同孩子與欄位的未使用舊碼失效；已有他人綁定的欄位不會被覆寫。

新學生模型啟用時寫入學生主檔的媽媽／爸爸 LINE 欄位；未啟用時沿用既有家長一／家長二欄位。不得擅自開啟模型、改歷史學期或整批清空 LINE 資料。

新主檔更新不強制重建舊投影；其他舊模組仍依賴在校投影時，另依既有同步流程確認。LINE好友對照不是新版唯一找孩子來源。

## 學生主檔與學期班級指派

先在學生管理維護唯一學生編號及基本資料，再在學期班級指派選正確學期、班代號、姓名、教師、生效日與狀態。不要因姓名相同就合併學生。

115-1 為本輪維護學期；114-2 與歷史學期維持唯讀。轉班、停學、離校保留歷程，不刪除學生或搬動歷史名冊。

先唯讀預檢與備份，再依差異審核決定資料模型切換。前端顯示功能不代表 Gateway 或 STUDENT_MODEL_ACTIVE 已啟用。

中英文姓名以主檔／指派來源為準，顯示年級班級不可回寫成姓名。接送明細使用穩定學生編號；同名或舊資料無法唯一辨識時停止自動配對。

## 接送排程與當日回報

依學期、日期、路線與中午／下午班核對名單，儲存後重新讀取確認人數。小巴下午班已由管理者確認正常，但每次改排程仍須核對當日回報。

整週複製先看來源週、目標週與模式；覆蓋會改動目標排程，須先備份及確認，不直接對有正式回報的歷史週操作。

接送群通知中的完成、轉接送、未接到學生連結帶路線上下文。老師從該通知進入，不直接打開空白頁、不手工改 row 或 session。

未接到學生先確認孩子安全並聯繫校方／家長，再依流程回報。畫面未顯示學生時先重新讀取及核對學期，不用測試重設函式覆蓋正式名單。

## 招生、檢測與體驗流程

問課諮詢 → 預約參訪／檢測 → 檢測與分班建議 → 體驗課或正式報名。依實際需求分流，不把舊手冊的每一步都當不可跳過的強制門檻。

預約與改期讀同一份日曆規則，核對可預約日期、封鎖時段與班級。純參訪不等於需出考卷；檢測、參訪加檢測與體驗課由既有教學公告群收備課通知。

取消保留原紀錄與原因，釋放名額後重查衝突；體驗預約不會自動建立正式學期入班。

公開頁面採多頁網址。舊首頁 #booking、#inquiry、#register、#camp 不再作為本手冊推薦入口；冬夏季課程使用對應頁面。

## 課程設定與公告

開班資訊由後台班級設定與公開頁面讀取同一份資料；更新後核對家長端。不要執行 TEST_classesSave、TEST_doPostFlow 或 FULL_RESET 來測正式資料。

公告由提報、主管審核、PDF 生成至通知提報人；依目前程式，不自動把所有公告推到全員工群。不要以程式中的舊註解推定實際收件者。

公告撤回保留軌跡並產生作廢版本，但不能收回他人已下載的 PDF。核對 PDF 是否生成成功、連結分享範圍及實際通知結果。

審核與預約詳情含敏感資料，使用系統產生的個別通知連結，不把帶電話或紀錄參數的網址放入一般連結中心。

## 部署與排程維護

公開站是 Cloudflare Worker + Static Assets。後台有 GitHub Pages 發布紀錄，但自訂網域是否更新仍要讀正式頁；GitHub push、CI 成功、Cloudflare 發布與 Apps Script 新版本是不同核對點。

Gateway 沿用已確認 Apps Script 專案，先備份，再管理部署 → 原 Web App → 新版本，保留網址。完整檔含校方設定，不能提交公開程式庫。

家長中心 v2 必須前後端成對切換。PARENT_LINE_CHANNEL_ID 須與 LINE Developers 實際 Login Channel 一致；LIFF endpoint 為 /leave/，profile scope 與 Provider 一致性須管理者確認。

新版正式部署且目的群確認後，手動執行一次 INSTALL_parentCenterNotificationRetry，建立每 5 分鐘通知重試；重跑不重複建同名排程。此手冊不會替你安裝觸發器。

既有每日摘要：cronLeave12 中午接送、cronLeavePM15 下午接送、cronLeave1230 美語與課輔/托育彙整。時間觸發器與 Asia/Taipei 時區由管理者在 Console 核對；不要點公開網址或執行測試函式試推正式群。

新通知以固定 retry key 防重複；首次嘗試超過 23 小時停止自動重試，needs_review 要人工核對。不要刪事件表、改新請求編號或盲目重送。

## 速度、錯誤處理與資料安全

版本查詢、孩子名冊、請假寫入及 LINE 通知分開量測。既有一次版本查詢約 1.32 秒，不能代表所有功能達到 3 秒；正式家長中心整段速度尚待新版部署後驗收。

白畫面先核對前端檔案與部署；找不到孩子先核對LINE身分、在學狀態與綁定；送出超時重試同一請求；LINE等待重試先查事件狀態、Bot權限與群組，不重建請假。

尚在 planned 的事件只由同一原請求復原，背景通知排程只處理 committed 事件。紀錄需人工核對時保留現場，不直接刪表清空。

個人帳號、最小權限，試算表不開任何人編輯；密碼、LINE token 與完整 Gateway 不放公開手冊。歷史公開過的密碼應撤銷／更換，刪掉文字不等於撤銷憑證。

本輪只更新後台文件與群組名稱，並強化家長請假入口；預約、公告與其他舊API尚未全面安全重構，不宣稱全系統已完成安全稽核。

## 正式驗收與待辦

待辦一：部署完整 parent-center-v2 Gateway 與家長網站，確認新版 capabilities 回應、頁面版本及回復方案。

待辦二：授權測試家長從官方 LINE 登入，只看到自己的有效在校孩子；驗證未綁定、成功綁定與錯誤碼。

待辦三：事前確認測試學生、日期、內容與收件群後，驗證只美語、只取消接送、安親加托育、新增／修改／取消實際收件，不把 API 受理當已讀。

待辦四：核對每日彙整與新版 5 分鐘通知重試排程，量測冷啟動、名冊與送單速度。

待辦五：既有公告審核通知指向 /announce-approve/，2026-09-23 裸路徑查核回傳 404；須恢復頁面並再驗證帶參數的審核流程。本次未發出或核准公告。

回復時前後端一起回復並保留事件與綁定紀錄；已推播訊息不會因回復程式撤回。此次文件更新不寫正式請假、名冊，不發 LINE。

## 系統連結索引

以下網址為入口與維護索引，並非全部都已做登入後驗收。僅將「家長」入口分享給家長；校務及維護入口仍需權限。

| 對象 | 名稱 | 網址 | 使用說明 |
| --- | --- | --- | --- |
| 家長 | 官方網站 | [官方網站](https://taipingxinguang.org/) | 公開入口 |
| 家長 | 課程介紹 | [課程介紹](https://taipingxinguang.org/curriculum.html) | 多頁網站；不再使用首頁 ?course 連結 |
| 家長 | 開班資訊 | [開班資訊](https://taipingxinguang.org/classes.html) | 與後台班級設定同步 |
| 家長 | 學費方案 | [學費方案](https://taipingxinguang.org/pricing.html) | 公開方案，以頁面當期資訊為準 |
| 家長 | 冬夏季課程 | [冬夏季課程](https://taipingxinguang.org/summer-winter.html) | 取代舊首頁 #camp |
| 家長 | 問課諮詢 | [問課諮詢](https://taipingxinguang.org/inquiry.html) | 取代舊首頁 #inquiry |
| 家長 | 預約參訪／檢測 | [預約參訪／檢測](https://taipingxinguang.org/booking.html) | 取代舊首頁 #booking |
| 家長 | 正式註冊 | [正式註冊](https://taipingxinguang.org/register.html) | 取代舊首頁 #register |
| 家長 | 國際認證 | [國際認證](https://taipingxinguang.org/certifications.html) | 公開資訊 |
| 家長 | 接送路線試算 | [接送路線試算](https://taipingxinguang.org/route-planner.html) | 試算不等於已安排接送 |
| 家長 | 分校介紹 | [分校介紹](https://taipingxinguang.org/about.html) | 公開資訊 |
| 家長 | 招生說明 | [招生說明](https://taipingxinguang.org/sales/) | 公開導客頁 |
| 家長 | LINE 官方帳號 | [LINE 官方帳號](https://lin.ee/tdpTDdu) | 加好友與校方聯絡 |
| 家長 | LINE 歡迎選單 | [LINE 歡迎選單](https://liff.line.me/2009757754-ZtkX6Igq) | 既有選單；不等於已綁定在校學生 |
| 家長 | 請假與取消接送 | [請假與取消接送](https://liff.line.me/2009757754-paHJ5QJO) | 官方 LINE 固定入口；新版 parent-center-v2 待正式部署驗收 |
| 校務 | 後台管理 | [後台管理](https://admin.taipingxinguang.org/) | 個人帳號登入；依角色授權 |
| 校務 | 家長綁定碼核發 | [家長綁定碼核發](https://taipingxinguang.org/parent-binding/) | 待新版前後端部署後使用；限行政／管理員（待部署） |
| 校務 | 學生資料操作手冊 | [學生資料操作手冊](https://admin.taipingxinguang.org/STUDENT_MANUAL.md) | 學生主檔與學期班級指派 |
| 校務 | Google Sheets 資料 | [Google Sheets 資料](https://docs.google.com/spreadsheets/d/1Q3lZwp8BiA6dcz5Bu_WLOitn-a_4O1Z0jDQoG6VMC7w/edit) | 僅授權校務人員；不得公開分享 |
| 校務 | Apps Script Gateway 專案 | [Apps Script Gateway 專案](https://script.google.com/home/projects/1a58uIi0Zbtxr6esICbVuE8CM9i7gtdozANhnu1SLUfhTrwNHAtqB_5lN/edit) | 沿用既有專案與部署網址 |
| 校務 | LINE OA 管理 | [LINE OA 管理](https://manager.line.biz/) | 圖文選單、家長訊息 |
| 校務 | LINE Developers | [LINE Developers](https://developers.line.biz/console/) | LIFF、Provider、Channel 設定 |
| 校務 | Cloudflare | [Cloudflare](https://dash.cloudflare.com/) | 公開站 Worker + Static Assets；發布與 GitHub 提交分開核對 |
| 維護 | 後台程式庫 | [後台程式庫](https://github.com/miyutang1980/xinguang-admin) | 後台版本與發布紀錄 |
| 維護 | 公開網站程式庫 | [公開網站程式庫](https://github.com/miyutang1980/xinguang-website) | 網站與 LIFF 頁面原始碼 |
| 維護 | 家長中心待部署更新 | [家長中心待部署更新](https://github.com/miyutang1980/xinguang-website/pull/2) | 獨立分支；不可在舊 Gateway 上單獨切新版前端 |
| 維護 | BotBonnie | [BotBonnie](https://app.botbonnie.com/) | 保留既有聊天分流；不擅改 Webhook |

## 通知情境頁

不得自行猜測參數或轉傳帶有家長個資的實際網址。以下 URL 是技術索引，不是一般家長入口。

| 用途 | 頁面 | 必要上下文 | 規則 |
| --- | --- | --- | --- |
| 預約改期 | [預約改期](https://taipingxinguang.org/booking-edit/) | row | 由該筆預約的通知進入，不手工猜列號 |
| 預約取消 | [預約取消](https://taipingxinguang.org/booking-cancel/) | row | 保留取消軌跡，重新核對名額 |
| 預約詳情 | [預約詳情](https://taipingxinguang.org/booking-detail/) | phone、date、slot | 含家長資料；不要轉貼帶參數的連結到不相關群 |
| 接送完成回報 | [接送完成回報](https://taipingxinguang.org/return-pickup/) | row、session、route | 接送老師回報，不是家長請假／找人代接 |
| 轉接送 | [轉接送](https://taipingxinguang.org/transfer-pickup/) | row、session、route、driver | 依路線通知與校方接送流程操作 |
| 未接到學生 | [未接到學生](https://taipingxinguang.org/missing-pickup/) | row、session、route | 接送老師回報；先確認學生安全 |
| 公告審核 | [公告審核](https://taipingxinguang.org/announce-approve/) | row；拒絕時 action=reject | 2026-09-23 裸路徑查核回傳 404，需恢復頁面並另驗證帶參數入口；不轉傳審核連結 |

## 版本核對依據

本輪以[正式後台](https://admin.taipingxinguang.org/)、[後台程式庫](https://github.com/miyutang1980/xinguang-admin)、[網站程式庫](https://github.com/miyutang1980/xinguang-website)及[家長中心待部署更新](https://github.com/miyutang1980/xinguang-website/pull/2)交叉核對。使用者提供的 Gateway 完整檔、截圖及本機隔離測試作為補充，不公開其中的憑證或學生資料。
