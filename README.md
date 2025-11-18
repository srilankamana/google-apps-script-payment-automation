# 支払い通知書 自動作成・送付システム (Google Apps Script)

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google-apps-script&logoColor=white)

## 📖 概要 (Overview)

Googleスプレッドシートで管理されている経理データを元に、取引先ごとの支払い通知書（PDF）を自動作成し、社内承認フローを経てメールで自動送付するシステムです。

**「誤送信を絶対に防ぐ」**という業務要件を満たすため、システムを「作成」と「送付」の2つのプロジェクトに物理的に分離し、人の目による承認プロセスをシステム的に強制する設計を採用しています。

---

## 🔄 業務フローと安全設計 (Workflow & Safety Design)

安全性と効率を両立するため、以下の**3段階フロー**で運用します。

```mermaid
graph TD
    subgraph STEP1 [1. 作成フェーズ]
        A[👤 担当者] -->|GAS実行| B(PDF自動生成 & ドライブ保存)
        B --> C[ステータス更新: 承認待ち]
    end

    subgraph STEP2 [2. 確認・承認フェーズ]
        C --> D{👤 上長}
        D -->|内容確認| E[📂 作成されたPDF]
        E --> D
        D -->|NG| F[修正・再作成]
        D -->|OK| G[ステータス変更: 承認済み送付OK]
    end

    subgraph STEP3 [3. 送付フェーズ]
        G --> H[👤 担当者]
        H -->|GAS実行| I{最終チェック}
        I -->|承認済を確認| J(メール自動送信)
        I -->|未承認なら| K[送信中断・スキップ]
        J --> L[ステータス更新: 支払通知書送付済]
    end
