# PowerPoint アドイン「ログ」タブ表示の解消要因

## 背景

PowerPoint VSTO アドイン（PowerPointAddIn1）では、当初リボンにカスタムタブが表示されない問題がありました。以下の対応により「**ログ**」タブが表示されるようになりました。本ドキュメントは、解消に至った要因と手順をまとめたものです。

---

## 解消の要因（なぜタブが表示されるようになったか）

### 1. Ribbon クラスを COM に公開した（最重要）

- **原因**: アセンブリで `[assembly: ComVisible(false)]` が指定されており、すべての型が既定で COM に非公開でした。Office は **COM 経由**でアドインの `IRibbonExtensibility`（GetCustomUI 等）を呼び出すため、Ribbon クラスが COM に公開されていないとリボンが読み込まれませんでした。
- **対応**: Ribbon クラスに **`[ComVisible(true)]`** を付与し、`using System.Runtime.InteropServices` を追加しました。Word アドイン（New_MOSWordVSTOAddIn）の Ribbon と同様の指定です。
- **対象ファイル**: `PowerPointAddIn1/Ribbon.cs`

### 2. タブの挿入位置を明示した

- **原因**: カスタムタブのみを定義し、標準リボン上の「どこに出すか」を指定していませんでした。
- **対応**: タブ要素に **`insertAfterMso="Help"`** を追加し、標準の「ヘルプ」タブの右隣に表示されるようにしました。Word アドインの「MOSデバッグ」タブでも同様の指定をしています。
- **対象ファイル**: `PowerPointAddIn1/Ribbon.xml`、`Ribbon.cs` 内のフォールバック用 XML（GetRibbonXmlContent）

### 3. タブの表示名を「ログ」に統一した

- **対応**: タブの `label` を「ログ」に設定しました（当初は「デバック」等だった場合の整理として）。

### 4. デバッグ用ログ出力を追加した

- **対応**: 以下を追加し、リボンが正しく読み込まれているか確認しやすくしました。
  - **ThisAddIn.cs**: `CreateRibbonExtensibilityObject` 内で `[ThisAddIn] CreateRibbonExtensibilityObject called` を出力
  - **Ribbon.cs**: `GetCustomUI` で `ribbonID` と返却 XML の文字数を出力
  - **Ribbon.cs**: `GetResourceText` で XML の取得元（埋め込みリソース / ファイル / フォールバック）を出力
  - **Ribbon.cs**: `Ribbon_Load` で `[Ribbon] Ribbon_Load completed - ログタブ有効` を出力

これにより、タブが出ない場合に「CreateRibbonExtensibilityObject が呼ばれていない」「GetCustomUI が呼ばれていない」等を切り分けられます。

### 5. 手順の徹底（スタートアップ・キャッシュ）

- **スタートアッププロジェクト**: ソリューションのスタートアップを **PowerPointAddIn1** にし、F5 で PowerPoint が起動するようにする。別プロジェクト（例: 採点アプリ）がスタートアップだと、アドインがデバッグビルドで読み込まれず、タブが更新されないように見えることがあります。
- **古いアドインのキャッシュ**: ビルドし直してもタブが出ない場合は、アドインをアンインストールし、最新の .vsto で再インストールしてから PowerPoint を起動し直す。

---

## まとめ（解消に効いた主な変更）

| 要因 | 対応内容 |
|------|----------|
| COM 非公開 | Ribbon クラスに `[ComVisible(true)]` を付与 |
| タブの位置 | `insertAfterMso="Help"` を指定 |
| タブ名 | `label="ログ"` に統一 |
| 切り分け | CreateRibbonExtensibilityObject / GetCustomUI / Ribbon_Load 等にデバッグ出力を追加 |

---

## ログタブを表示する手順（参照用）

1. Visual Studio でソリューションを開く
2. ソリューションエクスプローラーで **PowerPointAddIn1** を右クリック → **スタートアッププロジェクトに設定**
3. F5 でデバッグ開始 → PowerPoint が起動する
4. リボンに「**ログ**」タブ（ヘルプの右隣）が表示されることを確認
5. 出力ウィンドウで次が出ていれば成功の目安：
   - `[ThisAddIn] CreateRibbonExtensibilityObject called`
   - `[Ribbon] GetCustomUI called, ribbonID=...`
   - `[Ribbon] Ribbon_Load completed - ログタブ有効`

タブが出ない場合は、PowerPoint を終了 → アドインをアンインストール → PowerPointAddIn1 をリビルド → 生成された .vsto で再インストール → PowerPoint を起動し直す。

---

## 参照

- [Wordアドイン_MOSデバッグタブ表示のプロンプト.md](Wordアドイン_MOSデバッグタブ表示のプロンプト.md) — Word 側の同種の事象と手順（別ソリューション）
- PowerPointAddIn1: `Ribbon.cs`, `Ribbon.xml`, `ThisAddIn.cs`
