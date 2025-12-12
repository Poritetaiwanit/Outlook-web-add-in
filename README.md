# Porite Smart Alerts Outlook Add-in

此專案為 Outlook Web Add-in，可在按下寄送時彈出視窗顯示：
- 收件人完整資訊（名稱、Email、domain）
- 郵件主旨
- 附件列表

只有按下「寄送」才允許寄出郵件。

## Deployment (Azure Static Web Apps)

1. 將本 repo push 到 GitHub。
2. Azure Portal → Create Static Web App。
3. Source 選 GitHub。
4. Branch 選 main。
5. App location：public
6. Output location：public
