# A車加指定班次編輯修正：部署與操作

本次只開放 2026/10/01「A車加」下午班的學生、多校與人員編輯。這筆原設定是 A6、15:55、上限 6 人；它是獨立保留的加班趟，不是「半巴」，也不是週五 14:45 加班範本。

## 原因與修正

原程式將所有保留的特殊安排一律設為唯讀，因此即使沒有完成紀錄，學生與人員仍無法編輯。本版只為上述日期及路線加入編輯權限，不批次解除其他特殊車次的保護。

- 可編輯：下午班學生、多校接送點、接送順序、人員及備註；停開／恢復須符合原班次規則。
- 仍固定：日期、路線、車型 A6、15:55、容量 6 人；總人數由明細計算。
- 不開放：中午班、已接回班次、超過六人、其他未核准特殊車次。
- 保留：五分鐘版本化明細、儲存衝突檢查、批次儲存與完成通知防重複。

部署程式不會自動改學生或司機，不會搬移名單，不會推 LINE 測試通知。

## 更新 Gateway

1. 開啟[Gateway Apps Script](https://script.google.com/home/projects/1a58uIi0Zbtxr6esICbVuE8CM9i7gtdozANhnu1SLUfhTrwNHAtqB_5lN/edit)。
2. 備份目前「程式碼.gs」，用 `Gateway-A車加指定班次編輯修正-20260928.txt` 全文取代，儲存。不要動測試／救援檔。
3. 「部署」→「管理部署」→既有正式 Web app →鉛筆→版本選「新版本」。
4. 說明填：`A車加指定班次編輯修正｜保留固定時間與容量｜2026-09-28`。
5. 按部署，沿用原網址，不新增觸發器。

部署後能力回應應新增 `retainedEditVersion: "retained-trip-edit-v1"`，並保留 `detailsReadVersion: "pickup-revision-v1"`。[正式能力檢查](https://script.google.com/macros/s/AKfycbw-7_a_OfUVlgegcLxkux_9dr9UlYSVKhi3uQjV-0sr2X2TpRRmCXtSM7jbIqMHK4hNww/exec?_action=routes_policy_capabilities)

## 後台操作

在[正式後台](https://admin.taipingxinguang.org/)強制重新整理，進入「接送排程」，找 2026/10/01 的 A車加下午班，按「安排多校」或「選」。前端版號為 `20260928.1`；若顯示「加班趟編輯待 Gateway 更新」，表示尚未讀到新後端，按鈕仍保護性停用。

本版通過 67 項後端回歸與 9 項介面檢查，含舊 Gateway 不解鎖、完成後不可改、超載禁止、多校儲存、其他特殊趟次維持保護及桌面／手機畫面。均為合成資料測試，不代表已代校方儲存這筆正式名單。
