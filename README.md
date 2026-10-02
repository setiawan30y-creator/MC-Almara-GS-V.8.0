# MC-Almara GS V8.0
Structured Google Apps Script + Google Sheets ERP baseline derived from MC-ALMARA MASTER BLUEPRINT v1.0.

Architecture: 00_Config, 01_App, 02_Auth, 03_CoreEngine, 04_Database, 05_Utils, 06_Router, isolated module files, and shared assets.

Financial rule: modules do not calculate independent balances. Financial mutations are routed through the Core Engine.

This is the clean baseline. Full production completion still requires the blueprint's detailed OCR, DTOTT enforcement, threshold rules, bank/cash/stock/closing workflows, regulator exports, WhatsApp provider integration, translation/voice, gallery, old-money, coin, public display, approvals, and hardening.
