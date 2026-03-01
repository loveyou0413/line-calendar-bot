# LINE 行程登記 Bot（行程小幫手）

在 LINE 群組中輸入行程資訊，Bot 自動解析並登記到 Google 日曆。

## 功能

### 一般行程登記
- 自動解析訊息中的行程名稱、日期、時間、地點、主持人、出席人員
- 支援線上會議資訊（會議連結、會議號、會議密碼）
- 自動偵測重複行程，可選擇替換或新增
- 備註自動記錄登記時間

### 同仁請假登記
- 自動辨識請假類型訊息
- 同一天多次傳送會自動整合到同一個日曆事件
- 同一人重複登記會以新的覆蓋舊的

### 每日報表
- 每天晚上 6 點自動產生當日行程更新摘要
- 透過 LINE 私訊傳送給管理員

---

## 系統架構

```
同仁在 LINE 群組傳「登記 ...」
    ↓
LINE Bot 收到 Webhook
    ↓
Claude AI 解析行程資訊
    ↓
檢查是否有重複行程
    ↓（有重複）→ 詢問確認 →「確認」替換 /「新增」新增
    ↓（無重複）
建立 Google Calendar 事件
    ↓
Bot 回覆確認訊息
```

---

## 使用方式

### 登記一般行程
在 LINE 群組中傳送以「登記」開頭的訊息：

```
登記 3月15日下午兩點到四點，在A棟301會議室，產品規劃會議
主持人：王經理
出席人員：朱理事長(O) 幕僚 貝珊
會議連結：https://reurl.cc/xKng2Z
會議號/密碼：2510 771 7063/2026
```

Bot 會回覆確認訊息，並在 Google 日曆上建立事件。

### 出席人員格式說明
- `朱理事長(O)` → 會出席
- `朱理事長(X)` → 不出席
- `幕僚(貝珊)` → 會議幕僚是貝珊
- 括號使用半形 `()`，人員之間用頓號 `、` 分隔

### 登記請假
```
登記 0315同仁請假
麗如1.5小時(09-1030)
倫羽1小時(09-1000)
```

### 重複行程處理
如果同一天已有相同名稱或地點的行程，Bot 會詢問：
- 回覆「確認」→ 替換舊行程
- 回覆「新增」→ 另外新增一筆

---

## 部署需求

- **Synology NAS**（Plus 系列，需支援 Docker / Container Manager）
- **MikroTik 路由器**（或其他支援 Port Forwarding 的路由器）
- 以下帳號和金鑰：
  - LINE Developers 帳號（免費）
  - Google Cloud Platform 帳號（免費）
  - Anthropic Claude API 帳號（需儲值，每月約 $1 美元）

---

## 從零開始部署（完整步驟）

### 第一步：建立 LINE Bot

1. 到 [LINE Developers Console](https://developers.line.biz/) 登入
2. 建立 Provider → 建立 Messaging API Channel
3. Channel name 填「行程小幫手」
4. 記下 **Channel Secret**（Basic settings 分頁）
5. 產生並記下 **Channel Access Token**（Messaging API 分頁最下方，點 Issue）
6. 記下 **Your user ID**（Basic settings 分頁最下方）
7. 關閉自動回覆：到 [LINE Official Account Manager](https://manager.line.biz/) → 設定 → 回應設定 → 自動回應訊息設為停用

### 第二步：設定 Google Calendar API

1. 到 [Google Cloud Console](https://console.cloud.google.com/) 建立專案
2. 搜尋並啟用「Google Calendar API」
3. 到 APIs & Services → Credentials → CREATE CREDENTIALS → Service account
4. 名稱填「calendar-bot」，建立後到 Keys 分頁 → Add Key → Create new key → JSON
5. 下載 JSON 金鑰檔案，**重新命名為 `credentials.json`**
6. 打開 JSON 檔，複製 `client_email` 的值
7. 到 [Google 日曆](https://calendar.google.com/) → 你的日曆 → 設定和共用 → 與特定使用者共用 → 加入該 email，權限選「變更活動」
8. 在同一個頁面的「整合日曆」區塊，記下 **日曆 ID**（主日曆通常是你的 Gmail 地址）

### 第三步：取得 Claude API Key

1. 到 [Anthropic Console](https://console.anthropic.com/) 註冊
2. Settings → Billing 儲值至少 $5 美元
3. API Keys → Create Key，記下 API Key

### 第四步：準備檔案

1. 複製 `docker-compose.example.yml` 為 `docker-compose.yml`
2. 填入你的金鑰：

```yaml
environment:
  - LINE_CHANNEL_SECRET=你的值
  - LINE_CHANNEL_ACCESS_TOKEN=你的值
  - ANTHROPIC_API_KEY=你的值
  - GOOGLE_CALENDAR_ID=你的值
  - ADMIN_USER_ID=你的值（LINE User ID，用於接收每日報表）
  - REPORT_HOUR=18
  - PORT=8000
```

3. 將以下檔案放到 NAS 的同一個資料夾（例如 `/docker/line-calendar-bot/`）：
   - `app.py`
   - `requirements.txt`
   - `Dockerfile`
   - `docker-compose.yml`（已填好金鑰）
   - `credentials.json`（Google 服務帳戶金鑰）

### 第五步：在 NAS 上啟動

#### 方法 A：SSH

```bash
ssh 管理員帳號@NAS的IP
cd /volume1/docker/line-calendar-bot
sudo docker-compose up -d --build
```

#### 方法 B：Container Manager

1. 開啟 Container Manager → 專案 → 建立
2. 專案名稱填 `line-calendar-bot`
3. 路徑選 `/docker/line-calendar-bot`
4. 建立

### 第六步：設定 Synology 反向代理

1. 控制台 → 登入入口 → 進階 → 反向代理 → 新增
2. 來源：HTTPS、你的 DDNS 網址、連接埠 `58443`
3. 目的地：HTTP、localhost、連接埠 `8000`
4. 到 控制台 → 安全性 → 憑證 → 設定，確認反向代理使用 Let's Encrypt 憑證

### 第七步：設定路由器 Port Forwarding

在 MikroTik 上：

```
/ip firewall nat add \
    chain=dstnat \
    action=dst-nat \
    to-addresses=NAS的內網IP \
    to-ports=58443 \
    protocol=tcp \
    dst-address=你的外網IP \
    dst-port=58443 \
    in-interface=pppoe-out1 \
    comment="LINE Bot Webhook"
```

確認防火牆 forward chain 有允許 port 58443。

### 第八步：設定 LINE Webhook

1. LINE Developers Console → 你的 Channel → Messaging API
2. Webhook URL 填：`https://你的DDNS網址:58443/callback`
3. 開啟 Use webhook
4. 把 Bot 加入 LINE 群組

### 第九步：測試

在群組中傳送：

```
登記 明天下午兩點到四點，會議室A，測試會議，出席人員：張三(O) 幕僚 李四
```

---

## Synology 防火牆建議設定

控制台 → 安全性 → 防火牆 → 編輯規則，由上到下：

| 順序 | 連接埠 | 來源 IP | 動作 |
|------|--------|---------|------|
| 1 | 全部 | 你的內網網段（如 192.168.x.0/24） | 允許 |
| 2 | 58443 | 全部 | 允許 |
| 3 | 其他你需要的 port | 特定 IP | 允許 |
| 4 | 全部 | 全部 | 拒絕 |

注意：如果 NAS 防火牆有限制國家/地區，LINE 的伺服器在日本，需確保允許日本 IP 連入 port 58443。

---

## 維護注意事項

- **Let's Encrypt 憑證**：每 90 天到期，Synology 通常自動續約
- **Claude API 餘額**：到 console.anthropic.com 查看，每天 5 筆約每月 $1 美元
- **查看日誌**：`sudo docker logs --tail 50 line-calendar-bot`
- **重啟 Bot**：`cd /volume1/docker/line-calendar-bot && sudo docker-compose restart`
- **更新程式**：替換 `app.py` 後執行 `sudo docker-compose down && sudo docker-compose up -d --build`

---

## 檔案結構

```
line-calendar-bot/
├── app.py                      # 主程式
├── requirements.txt            # Python 套件
├── Dockerfile                  # Docker 映像設定
├── docker-compose.example.yml  # docker-compose 範本（不含金鑰）
├── docker-compose.yml          # 實際使用的設定（含金鑰，不上傳 Git）
├── credentials.json            # Google 服務帳戶金鑰（不上傳 Git）
├── .gitignore                  # Git 忽略規則
├── README.md                   # 本文件
└── 部署教學.md                  # 繁體中文部署教學
```

---

## 安全注意事項

- **絕對不要**將 `docker-compose.yml` 和 `credentials.json` 上傳到 GitHub
- `docker-compose.example.yml` 是範本，不含真實金鑰，可以安全上傳
- 所有金鑰都透過環境變數注入，不會寫死在程式碼中
- LINE Webhook 有簽章驗證，防止偽造請求
