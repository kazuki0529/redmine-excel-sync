# RedmineExcelSync

Redmine上のチケット情報をExcelに同期するVBAツールです。  
プロジェクトマネージャや非エンジニア向けに、Redmineの情報を見やすく共有することを目的としています。

## ✨ 主な機能

- RedmineのクエリIDを指定してチケット情報を取得
- Excel上でチケットIDの突合と更新／新規追加を自動処理
- チケットの説明欄から `h2`（Markdown/Textile）レベルの見出しを解析し、対応する列に自動反映
- チケット本文をセルのコメントとして記録
- 将来的に、ExcelからRedmineへのチケット更新にも対応予定

## 📥 セットアップ手順

### 1. 必要な外部ライブラリの導入

以下の外部ライブラリを使用しています。  
`./lib/` フォルダに配置してください。

| ライブラリ名 | URL | ライセンス |
|--------------|-----|------------|
| **VBA-JSON** | [github.com/VBA-tools/VBA-JSON](https://github.com/VBA-tools/VBA-JSON) | MIT |
| **ariawase** | [github.com/vbaidiot/ariawase](https://github.com/vbaidiot/ariawase) | MIT |

```bash
# ariawase
mkdir -p ./lib/ariawase
curl -L https://github.com/vbaidiot/ariawase/archive/refs/tags/v0.9.0.tar.gz | tar -xz --strip-components=1 -C lib/ariawase

# vba-json
mkdir -p ./lib/vba-json
curl -L https://github.com/VBA-tools/VBA-JSON/archive/refs/tags/v2.3.1.tar.gz | tar -xz --strip-components=1 -C lib/vba-json
```

## 📜 ライセンス

このプロジェクトは [MITライセンス](LICENSE) の下で公開されています。  
ただし、使用している外部ライブラリ（VBA-JSON、ariawase）もMITライセンスです。各ライブラリのライセンスに準じてください。

## 📌 備考・今後の展望

- Excel → Redmineへのチケット更新
- `h2` 見出しの抽出ルールの柔軟化
- GitHub Actions等を使ったテンプレートExcel自動生成（将来的に）
