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

### 必要な外部ライブラリの導入

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

## ⚠ Excelのセキュリティ設定について

このツールは VBA を通じてモジュールをインポート／エクスポートするため、  
以下の設定を有効にする必要があります（無効のままだとスクリプト実行時にエラーが発生します）。

### ✅ 必要な設定：VBAプロジェクトへのアクセスを許可

1. Excel を起動
2. [ファイル] → [オプション] を開く
3. [セキュリティセンター] → [セキュリティセンターの設定] をクリック
4. [マクロの設定] を開く
5. 以下のオプションにチェックを入れる：
   - 「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」
6. Excel を再起動

> この設定を有効にしない場合、以下のようなエラーが発生します：
> Microsoft Excel: プログラミングによる Visual Basic プロジェクトへのアクセスは信頼性に欠けます
> 参考: [Microsoft公式ドキュメント](https://learn.microsoft.com/ja-jp/office/vba/library-reference/concepts/trust-access-to-the-vba-project-object-model)

## 🛠 ビルド手順

1. `build.bat` を実行する
2. `dist/` フォルダに `redmine-excel-sync.xlsm` が生成されます。

```cmd
# Windowsコマンドプロンプトでのビルド
>build.bat
0 個のファイルをコピーしました
lib\ariawase\src\Ariawase.xlsm\Assert.bas
lib\ariawase\src\Ariawase.xlsm\Core.bas
lib\ariawase\src\Ariawase.xlsm\Ext.bas
lib\ariawase\src\Ariawase.xlsm\IO.bas
4 個のファイルをコピーしました
lib\ariawase\src\Ariawase.xlsm\ArrayEx.cls
lib\ariawase\src\Ariawase.xlsm\Func.cls
lib\ariawase\src\Ariawase.xlsm\MonkeyTest.cls
lib\ariawase\src\Ariawase.xlsm\Tuple.cls
4 個のファイルをコピーしました
lib\vba-json\JsonConverter.bas
1 個のファイルをコピーしました
lib\ariawase\vbac.wsf
1 個のファイルをコピーしました
begin combine

> Target: redmine-excel-sync.xlsm
- Import: ArrayEx.cls
- Import: Assert.bas
- Import: Core.bas
- Import: Ext.bas
- Import: Func.cls
- Import: IO.bas
- Import: JsonConverter.bas
- Import: MonkeyTest.cls
- Import: Tuple.cls

end
> dir .\dist
 ドライブ D のボリューム ラベルは ボリューム です
 ボリューム シリアル番号は B218-EAB5 です

 D:\Users\kazuki\workspace\private\redmine-excel-eync\dist のディレクトリ

2025/07/05  21:48    <DIR>          .
2025/07/05  21:48    <DIR>          ..
2025/07/05  21:48           127,448 redmine-excel-sync.xlsm
               1 個のファイル             127,448 バイト
               2 個のディレクトリ  611,374,161,920 バイトの空き領域
> 
```

## 📜 ライセンス

このプロジェクトは [MITライセンス](LICENSE) の下で公開されています。  
ただし、使用している外部ライブラリ（VBA-JSON、ariawase）もMITライセンスです。各ライブラリのライセンスに準じてください。
