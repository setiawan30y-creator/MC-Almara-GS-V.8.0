# MC-Almara GS V8.0

Foundation Google Apps Script + Google Sheets based on the uploaded MC-ALMARA MASTER BLUEPRINT v1.0.

## Included
- Idempotent installer and Google Sheets schema.
- Server-side session/role/permission checks.
- Core transaction engine with BUY/SELL, cash/transfer/split payment, cash/bank/FX-stock mutations and audit trail.
- Dashboard, customer, transaction/history, calculator and public-rate UI.
- Responsive desktop/mobile UI.

## Deploy
Run setupSystem() once, then createInitialAdmin(), then deploy as Apps Script Web App. Secrets belong in PropertiesService. Regulator/OCR/WhatsApp/external-rate integrations remain configurable integration points.

Local preview is in preview/index.html in the delivered project archive.