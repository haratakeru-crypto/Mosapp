# PowerPoint 類題採点対応表

修正版問題文（新タスク）と旧採点ロジックの対応。Legacy 退避は P1 改修前 baseline を参照。

## Baseline

| 項目 | 値 |
|------|-----|
| Legacy 作成日 | 2026-06-15 |
| P1 改修前 commit（Checker1_1 旧版） | `d03e149` |
| P1 改修 commit（参考） | `81ecce9` |
| JSON 旧版 source（P2以降） | commit `9df05ff`（旧タスク構成・旧文言。新CSVのタスク数とは一致しない） |
| Legacy 配置（Checker） | `Libraries/Group1/Legacy/*.Legacy.cs` |
| Legacy 配置（破壊的操作） | `Libraries/Legacy/PPTaskValidationConfig.Legacy.cs`（**凍結**・P3 完了時点・旧 projectId 4〜11） |
| 破壊的操作一覧（主） | 本ファイル「破壊的操作免除一覧」 |
| 元 CSV | `Mos PowerPoint Mogi App/PP修正版問題文一覧_類似付き.csv` |

> **git 確認（依頼1実施時）**: `PowerPointChecker1_1.cs` は HEAD（`81ecce9`）に P1 改修済み。作業ツリーに**未コミットの追加差分**あり。Legacy の 1_1 は commit `d03e149` から復元。

> **JSON について**: P1 のみ新問題文（8タスク）。P2〜P10 は commit `9df05ff` の旧文言・旧タスク構成に戻した。各プロジェクトの採点実装完了時に、新CSVどおりの問題文へ更新する。

## 凡例

- **旧 X-Y**: 旧プロジェクト X のタスク Y
- **Legacy ファイル**: `PowerPointChecker1_{旧プロジェクト}.Legacy.cs`
- **Legacy メソッド**: `CheckTask_1_{旧P}_{旧T}`（ゼロ埋め2桁）。**P2 以降の現行**は `CheckTask_1_{P}_{新T}`（P2-X → `CheckTask_1_2_0X`）に揃える
- **類似なし**: CSV の類似問題列が空欄 → 新規採点実装

## 全タスク対応表

| 新 | 旧 | 新問題文（先頭80字） | Legacy ファイル | Legacy メソッド | 上書きタイミング | 差分メモ | 実装 |
|----|-----|----------------------|-----------------|-----------------|------------------|----------|------|
| P1-1 | 旧1-1 | あなたは英語教育プログラムのご提案資料を作成しています。 スライド４に、レイアウト「テーブルスライド」のスライドを挿入します。 | PowerPointChecker1_1.Legacy.cs | CheckTask_1_1_01 | P1 実装時に PowerPointChecker1_1.cs 上書き |  |  |
| P1-2 | 旧1-3 | スライド４を非表示にします。 | PowerPointChecker1_1.Legacy.cs | CheckTask_1_1_03 | P1 実装時に PowerPointChecker1_1.cs 上書き |  |  |
| P1-3 | 旧1-5 | スライド５のレイアウトを「テキスト１スライド」に変更します。上側の［マスターテキストのスタイルを編集する］のプレースホルダーに「"英語教育を始めたばかりのケース | PowerPointChecker1_1.Legacy.cs | CheckTask_1_1_05 | P1 実装時に PowerPointChecker1_1.cs 上書き |  |  |
| P1-4 | 旧1-6 | スライド６の箇条書きを2段組みに変更します。 | PowerPointChecker1_1.Legacy.cs | CheckTask_1_1_06 | P1 実装時に PowerPointChecker1_1.cs 上書き |  |  |
| P1-5 | 旧10-3 | スライド８の箇条書きのプレースホルダーの文字の間隔を広げます。幅を「"4"pt」にします。 | PowerPointChecker1_10.Legacy.cs | CheckTask_1_10_03 | P10 実装時に PowerPointChecker1_10.cs 上書き |  |  |
| P1-6 | 旧7-1 | スライド８にセクションを追加します。セクション名は「"まとめ"」にします。 | PowerPointChecker1_7.Legacy.cs | CheckTask_1_7_01 | P7 実装時に PowerPointChecker1_7.cs 上書き |  |  |
| P1-7 | 旧11-2 | スライド１のセクション名を「"はじめに"」とします。 | PowerPointChecker1_11.Legacy.cs | CheckTask_1_11_02 | P11 実装時に PowerPointChecker1_11.cs 上書き |  |  |
| P1-8 | 旧3-4 | スライド１の後ろにサマリーズームスライドを挿入し［1.教育理念］、［4.募集要項］の各スライドへのリンクを作成します。スライド１と８へのリンクは含めません。タイ | PowerPointChecker1_3.Legacy.cs | CheckTask_1_3_04 | P3 実装時に PowerPointChecker1_3.cs 上書き | 旧3-4 は P1-8, P3-6 でも参照 |  |
| P2-1 | 旧2-1 | あなたはMOS合格対策講座のコースガイダンス資料を作成しています。 すべてのスライドに、画面切り替え「スプリット」を設定します。効果のオプションを「ワイプアウト | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_01 | P2 実装時に PowerPointChecker1_2.cs 上書き |  | 済 |
| P2-2 | 旧2-2 | すべての画面切り替えの継続時間を3秒に設定します。 | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_02 | P2 実装時に PowerPointChecker1_2.cs 上書き |  | 済 |
| P2-3 | 旧2-3 | スライド3、4、5に「切り替え」の画面切り替え効果を設定します。 | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_03 | P2 実装時に PowerPointChecker1_2.cs 上書き |  | 済 |
| P2-4 | — | すべてのスライドが、"５"秒後に自動で次のスライドへ進むように画面切り替えのタイミングを設定します。 | — | — | — | [画面切り替え]タブ
↓
[自動]チェックボックスを
クリックしオン
↓
[自動… | 済 |
| P2-5 | 旧2-4 | スライド６の3Ｄモデル「黒板」に、アニメーション「ジャンプしてターン」を設定します。 | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_04 | P2 実装時に PowerPointChecker1_2.cs 上書き |  | 済 |
| P2-6 | 旧2-5 | スライド2の男の子と？マークの２つの画像が、スライドの左上隅から登場するようにします。アニメーションの継続時間は"0.55"秒にします。 | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_05 | P2 実装時に PowerPointChecker1_2.cs 上書き |  | 済 |
| P2-7 | 旧2-6 | スライド「今年度募集について」の箇条書きに設定されたアニメーションの効果を「プラス」に変更します。また、クリックするとすべて同時に動くようにします。 | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_06 | P2 実装時に PowerPointChecker1_2.cs 上書き |  | 済 |
| P2-8 | 旧2-7 | スライド４の星の図形にアニメーションの軌跡「ニュートロン」を設定します。円の図形にはアニメーションを設定しません。 | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_07 | P2 実装時に PowerPointChecker1_2.cs 上書き |  | 済 |
| P3-1 | 旧3-1 | あなたはPower Point 新機能についての発表用資料を作成しています。 スライド７にSmartArtグラフィック「タイムライン」の手順を追加し、文字列「" | PowerPointChecker1_3.Legacy.cs | CheckTask_1_3_01 | P3 実装時に PowerPointChecker1_3.cs 上書き | 基本タイムラインは layout Id=hProcess11 | 済 |
| P3-2 | 旧3-2 | スライド７のSmartArtグラフィックに、色「グラデーション循環-アクセント６」を設定します。 | PowerPointChecker1_3.Legacy.cs | CheckTask_1_3_02 | P3 実装時に PowerPointChecker1_3.cs 上書き |  | 済 |
| P3-3 | 旧3-3 | スライド６の箇条書きを「ターゲットリスト」のSmartArtに変更します。 | PowerPointChecker1_3.Legacy.cs | CheckTask_1_3_03 | P3 実装時に PowerPointChecker1_3.cs 上書き | ターゲットリストは layout Id=target3 | 済 |
| P3-4 | 旧6-3 | スライド１に3Dモデル「虫眼鏡」を挿入します。幅を「"2.5"」に変更して中央の楕円の図形の中に配置します。正確な位置は問いません。 | PowerPointChecker1_6.Legacy.cs | CheckTask_1_6_03 | P3 実装時に PowerPointChecker1_3.cs 上書き | 旧6-3 を P3-4 に移植 | 済 |
| P3-5 | 旧6-4 | スライド10の3Dモデルのビューを上前面にし、高さを「"6.5"」に変更します。 | PowerPointChecker1_6.Legacy.cs | CheckTask_1_6_04 | P3 実装時に PowerPointChecker1_3.cs 上書き | 旧6-4 を P3-5 に移植 | 済 |
| P3-6 | 旧3-4 | 「機能の概要」のスライドにスライドズームを挿入して「画面録画で説明」「ズーム機能で訴求力アップ」「デザインアイデアで魅力的に！」へリンクを作成します。タイトルの | PowerPointChecker1_3.Legacy.cs | CheckTask_1_3_04 | P3 実装時に PowerPointChecker1_3.cs 上書き | 旧3-4 は P1-8, P3-6 でも参照。3件ズーム。タイトルPH完全一致。配置は0pt隙間許容 | 済 |
| P3-7 | — | スライド２にセクションズームのリンクを挿入します。セクション「1.機能の概要」と「2.伝わるスライドの要素」にリンクを作成し、それぞれ文字の下に配置します。 | — | — | P3 実装時に PowerPointChecker1_3.cs 上書き | 新規。ラベル文字とセクション名の1対1ペアリング検証（入れ替えは×） | 済 |
| P4-1 | 旧1-7 | あなたは英語教育プログラムの提案用資料を作成しています。 スライド１の青い図形に、「"教育者必見"」と入力します。 | PowerPointChecker1_1.Legacy.cs | CheckTask_1_1_07 | P1 実装時に PowerPointChecker1_1.cs 上書き |  |  |
| P4-2 | 旧5-2 | スライド４の文字列「いつでも体験可能です!」に、文字の塗りつぶし「青、アクセント1」を設定します。 | PowerPointChecker1_5.Legacy.cs | CheckTask_1_5_02 | P5 実装時に PowerPointChecker1_5.cs 上書き |  |  |
| P4-3 | 旧4-2 | スライド１枚目の子供の画像に、図の効果「光彩:18pt;緑､アクセントカラー6」を設定します。 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_02 | P4 実装時に PowerPointChecker1_4.cs 上書き |  |  |
| P4-4 | 旧4-1 | スライド１の子供の画像に、スタイル「楕円 ぼかし」を設定し、「テクスチャライザー」のアート効果を設定します。 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_01 | P4 実装時に PowerPointChecker1_4.cs 上書き |  |  |
| P4-5 | 旧4-5 | スライド５の右の画像を左の画像の上端に合わせます。水平方向には動かさないようにします。 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_05 | P4 実装時に PowerPointChecker1_4.cs 上書き | 旧4-5 は P4-5, P5-1 でも参照 |  |
| P4-6 | 旧4-4 | スライド５の右側の画像の右端を、スライドの右端に揃えてトリミングします。画像の右端以外は変更しないでください。トリミングした領域をプレゼンテーションから完全に削 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_04 | P4 実装時に PowerPointChecker1_4.cs 上書き |  |  |
| P4-7 | 旧11-3 | 2枚目のスライドのテキストボックスに、塗りつぶし「青、アクセント1、白+基本色60％」、枠線「濃い青」、太さ「0.75pt」を設定します。 | PowerPointChecker1_11.Legacy.cs | CheckTask_1_11_03 | P11 実装時に PowerPointChecker1_11.cs 上書き |  |  |
| P4-8 | 旧11-6 | スライド３のコンテンツ領域にあるテキストボックスを、スライドの垂直方向の中央に配置します。 | PowerPointChecker1_11.Legacy.cs | CheckTask_1_11_06 | P11 実装時に PowerPointChecker1_11.cs 上書き |  |  |
| P5-1 | 旧4-5 | あなたはMOS合格対策講座のコースガイダンス資料を作成しています。 スライド３の丸の図形４個の右端を揃えます。 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_05 | P5 実装時に PowerPointChecker1_5.cs 上書き | スライド3・円ちょうど4個・右端2pt | 済 |
| P5-2 | 旧5-4 | スライド５の小さい四角の大きさを他の四角と同じ幅にします。 | PowerPointChecker1_5.Legacy.cs | CheckTask_1_5_04 | P5 実装時に PowerPointChecker1_5.cs 上書き | 角丸四角含む・max-min幅0.5pt | 済 |
| P5-3 | 旧5-3 | スライド４の星の図形をスマイルに変更します。 | PowerPointChecker1_5.Legacy.cs | CheckTask_1_5_03 | P5 実装時に PowerPointChecker1_5.cs 上書き |  | 済 |
| P5-4 | 旧4-6 | スライド６の図形を手前から「対策講座」「PC教室」「通信講座」になるように重なりの順番を変更します。 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_06 | P5 実装時に PowerPointChecker1_5.cs 上書き | スライド6・Z順テキスト指定 | 済 |
| P5-5 | 旧5-5 | スライド６の3つの論理積ゲートの図形をグループ化します。 | PowerPointChecker1_5.Legacy.cs | CheckTask_1_5_05 | P5 実装時に PowerPointChecker1_5.cs 上書き | スライド6・3メンバー同一種類 | 済 |
| P5-6 | 旧4-3 | スライド「MOSって何？」の男の子の画像の代替テキストを装飾化し、スクリーンリーダーに表示させないようにします。 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_03 | P5 実装時に PowerPointChecker1_5.cs 上書き | タイトル指定・最大画像・Decorative | 済 |
| P5-7 | 旧11-5 | スライド２の[?]のアイコンに「濃い赤」の塗りつぶしを設定します。 | PowerPointChecker1_11.Legacy.cs | CheckTask_1_11_05 | P5 実装時に PowerPointChecker1_5.cs 上書き | スライド2・Graphic/Icon・濃い赤RGB | 済 |
| P6-1 | 旧10-1 | あなたは、「伝わる」プレゼンテーションのテクニックを紹介する資料を作成しています。 プレゼンテーションからドキュメントのプロパティと個人情報を削除します。他の情 | PowerPointChecker1_10.Legacy.cs | CheckTask_1_10_01 | P6 実装時に PowerPointChecker1_6.cs 上書き |  | 済 |
| P6-2 | 旧8-5 | プレゼンテーションを常に読み取り専用にします。 | PowerPointChecker1_8.Legacy.cs | CheckTask_1_8_05 | P6 実装時に PowerPointChecker1_6.cs 上書き |  | 済 |
| P6-3 | 旧7-4 | スライドショーを自動プレゼンテーションとして設定します。 | PowerPointChecker1_7.Legacy.cs | CheckTask_1_7_04 | P6 実装時に PowerPointChecker1_6.cs 上書き | VSTO キオスク証跡要 | 済 |
| P6-4 | 旧10-2 | スライド４，５，６を選択し、「"書式のポイント"」という名前の目的別スライドショーを作成します。スライドショーは実行しません。 | PowerPointChecker1_10.Legacy.cs | CheckTask_1_10_02 | P6 実装時に PowerPointChecker1_6.cs 上書き | Legacy 名称「教育」→「書式のポイント」 | 済 |
| P6-5 | 旧5-1 | すべてのスライドをアウトラインで、部単位に"６"部印刷するように設定します。 | PowerPointChecker1_5.Legacy.cs | CheckTask_1_5_01 | P6 実装時に PowerPointChecker1_6.cs 上書き | 旧5-1 は P6-5, P6-7 でも参照。Outline/6部 | 済 |
| P6-6 | 旧11-7 | ノートですべてのスライドを"3"部印刷しなさい。ただし、1ページ目を全て印刷したあとに2ページ目を印刷するようにします。 | PowerPointChecker1_11.Legacy.cs | CheckTask_1_11_07 | P6 実装時に PowerPointChecker1_6.cs 上書き | VSTO 印刷証跡要 | 済 |
| P6-7 | 旧5-1 | 印刷オプションで、グレースケールの配布資料を、１ページに3スライドのレイアウトで"４"部印刷するように設定します。印刷は実行しないでください。 | PowerPointChecker1_5.Legacy.cs | CheckTask_1_5_01 | P6 実装時に PowerPointChecker1_6.cs 上書き | Grayscale/3スライド/4部/PrintColorType | 済 |
| P7-1 | 旧6-1 | あなたは、情報モラルの重要性について発表する資料を作成しています。 スライド１にコメント「"情報発信の責任を考える"」を挿入します。 | PowerPointChecker1_6.Legacy.cs | CheckTask_1_6_01 | P7 実装時に PowerPointChecker1_7.cs 上書き | コメント文言は問題文差分あり | 済 |
| P7-2 | 旧9-6 | スライド１枚目の文字列「情報学習支援」に、Webページ「"https://rabbitway.jp/"」を表示するハイパーリンクを設定します。 | PowerPointChecker1_9.Legacy.cs | CheckTask_1_9_06 | P7 実装時に PowerPointChecker1_7.cs 上書き | 旧9-6 とは対象文字列・URL が異なる | 済 |
| P7-3 | 旧7-3 | スライド６の後ろに、文書「まとめ」のアウトラインを使用してスライドを挿入します。 | PowerPointChecker1_7.Legacy.cs | CheckTask_1_7_03 | P7 実装時に PowerPointChecker1_7.cs 上書き | 旧7-3 はスライド5後・「弊社の他の講座一覧」 | 済 |
| P7-4 | 旧9-4 | スライドのフッターに、スライド番号と「"rabbitway.jp"」をタイトルスライド以外に追加します。 | PowerPointChecker1_9.Legacy.cs | CheckTask_1_9_04 | P7 実装時に PowerPointChecker1_7.cs 上書き | 旧9-4 は www.MOS.jp | 済 |
| P7-5 | 旧9-5 | ５枚目の「事例１」と６枚目の「事例2」のフッターに「"参考事例"」と挿入します。他のスライドには表示しません。 | PowerPointChecker1_9.Legacy.cs | CheckTask_1_9_05 | P7 実装時に PowerPointChecker1_7.cs 上書き | 旧9-5 はスライド5のみ「集中的に」 | 済 |
| P8-1 | 旧9-2 | あなたは、「見せる！」　プレゼンテーションの極意を紹介する資料作成しています。 スライド「見せる！スライドの基本ルール」の表を、スタイル「中間スタイル４-アクセ | PowerPointChecker1_9.Legacy.cs | CheckTask_1_9_02 | P9 実装時に PowerPointChecker1_9.cs 上書き |  |  |
| P8-2 | 旧6-2 | スライド１の背景を「濃い緑、テキスト２、白+基本色80％」に変更します。 | PowerPointChecker1_6.Legacy.cs | CheckTask_1_6_02 | P6 実装時に PowerPointChecker1_6.cs 上書き |  |  |
| P8-3 | 旧11-1 | スライドのサイズを「画面に合わせる16：9」にします | PowerPointChecker1_11.Legacy.cs | CheckTask_1_11_01 | P11 実装時に PowerPointChecker1_11.cs 上書き |  |  |
| P8-4 | 旧9-7 | スライドの大きさを、高さ「"17.46"cm」、幅「"27.54"cm」に変更します。スライドは画面に合わせます。 | PowerPointChecker1_9.Legacy.cs | CheckTask_1_9_07 | P9 実装時に PowerPointChecker1_9.cs 上書き |  |  |
| P8-5 | 旧10-4 | スライドをグレースケールで表示し、スライド２のイラストを「明るいグレースケール」にします。 | PowerPointChecker1_10.Legacy.cs | CheckTask_1_10_04 | P10 実装時に PowerPointChecker1_10.cs 上書き |  |  |
| P9-1 | 旧8-1 | あなたは、英語教育プログラムのご提案資料を作成しています。 スライド「スクールの様子」に、動画「受講の様子.mp4」を挿入します。挿入にはアイコンを使います。 | PowerPointChecker1_8.Legacy.cs | CheckTask_1_8_01 | P8 実装時に PowerPointChecker1_8.cs 上書き |  |  |
| P9-2 | 旧8-3 | スライド５のビデオを、開始を「"4"秒」、終了を「"9"秒」に設定します。 | PowerPointChecker1_8.Legacy.cs | CheckTask_1_8_03 | P8 実装時に PowerPointChecker1_8.cs 上書き |  |  |
| P9-3 | 旧8-4 | スライド１のオーディオを、スライドを切り替えても1回だけ再生するように設定します。再生は"3"秒かけてフェードアウトするようにします。 | PowerPointChecker1_8.Legacy.cs | CheckTask_1_8_04 | P8 実装時に PowerPointChecker1_8.cs 上書き |  |  |
| P9-4 | 旧9-1 | スライド２のプレースホルダーに表の内容を表す「集合縦棒」グラフを作成します。「結果」の列を項目、「人数」の列をデータ系列として使用します。表のデータはグラフシー | PowerPointChecker1_9.Legacy.cs | CheckTask_1_9_01 | P9 実装時に PowerPointChecker1_9.cs 上書き |  |  |
| P9-5 | 旧9-3 | スライド２のグラフに［凡例マーカーなし］のデータテーブルを表示し、タイトルと凡例を削除します。 | PowerPointChecker1_9.Legacy.cs | CheckTask_1_9_03 | P9 実装時に PowerPointChecker1_9.cs 上書き |  |  |
| P10-1 | 旧10-5 | あなたは、日本の「四季」について紹介する資料を作成しています。 スライドマスターにテーマ「木版活字」を設定します。 | PowerPointChecker1_10.Legacy.cs | CheckTask_1_10_05 | P10 実装時に PowerPointChecker1_10.cs 上書き |  |  |
| P10-2 | — | スライドマスターにスライド番号を挿入します。 | — | — | — | [表示]タブ
↓
[スライドマスター]をクリッ
ク
↓
[スライドマスター]タブ… | 新規 |
| P10-3 | — | スライドマスターの［タイトルスライド］レイアウトのスライド番号を非表示にします。 | — | — | — | [表示]タブ
↓
[スライドマスター]をクリッ
ク
↓
[スライドマスター]タブ… | 新規 |
| P10-4 | 旧10-6 | スライドマスターの［２つのコンテンツ］レイアウトのスライドの背景のデザインを非表示にします。 | PowerPointChecker1_10.Legacy.cs | CheckTask_1_10_06 | P10 実装時に PowerPointChecker1_10.cs 上書き |  |  |
| P10-5 | — | スライドマスターで［フッター］のプレースホルダーを削除します。 | — | — | — | [表示]タブ
↓
[スライドマスター]をクリッ
ク
↓
[スライドマスター]タブ… | 新規 |
| P10-6 | 旧10-7 | スライドマスターの［タイトルのみ］レイアウトをもとに「"タイトル付きの図と表"」の名前でレイアウトを作成します。図のプレースホルダーをスライドの左側、表のプレー | PowerPointChecker1_10.Legacy.cs | CheckTask_1_10_07 | P10 実装時に PowerPointChecker1_10.cs 上書き |  |  |
| P10-7 | 旧11-4 | 配布資料マスターの日付を削除します。 | PowerPointChecker1_11.Legacy.cs | CheckTask_1_11_04 | P11 実装時に PowerPointChecker1_11.cs 上書き | 旧11-4 は P10-7, P10-8 でも参照 |  |
| P10-8 | 旧11-4 | 配布資料マスターのフッターに「"四季を楽しむ"」と表示します。 | PowerPointChecker1_11.Legacy.cs | CheckTask_1_11_04 | P11 実装時に PowerPointChecker1_11.cs 上書き | 旧11-4 は P10-7, P10-8 でも参照 |  |

## P2 現行メソッド・破壊的操作 taskId 対応

`PowerPointGrader` の `projectId==2` では **taskId = P2-X**。**現行** `PowerPointChecker1_2.cs` のメソッド名は **P2-X = `CheckTask_1_2_0X`** に揃えた。Legacy は旧2-x番号のまま `Legacy/` に退避。破壊的操作検知（`PPTaskValidationConfig` / スナップショット）は **taskId（P2-X）** で紐づく。

| P2（taskId） | 現行メソッド | Legacy メソッド（旧2-x） |
|--------------|-------------|-------------------------|
| P2-1 (1) | `CheckTask_1_2_01` | 同左 |
| P2-2 (2) | `CheckTask_1_2_02` | 同左 |
| P2-3 (3) | `CheckTask_1_2_03` | 同左 |
| P2-4 (4) | `CheckTask_1_2_04` | —（新規） |
| P2-5 (5) | `CheckTask_1_2_05` | `CheckTask_1_2_04` |
| P2-6 (6) | `CheckTask_1_2_06` | `CheckTask_1_2_05` |
| P2-7 (7) | `CheckTask_1_2_07` | `CheckTask_1_2_06` |
| P2-8 (8) | `CheckTask_1_2_08` | `CheckTask_1_2_07` |

> **既知課題（P2-4）**: P2-1〜8 を一連で行うと、P2-4 の「すべてのスライドに適用」により P2-1 または P2-3 が最終採点で × になる。詳細・対策案は [`PPtasks_md/P2-4_累積採点と画面切り替え_課題と対策.md`](../PPtasks_md/P2-4_累積採点と画面切り替え_課題と対策.md) を参照。

## P3 現行メソッド・破壊的操作 taskId 対応

`PowerPointGrader` の `projectId==3` では **taskId = P3-X**。**現行** `PowerPointChecker1_3.cs` のメソッド名は **P3-X = `CheckTask_1_3_0X`** に揃えた。P3-4/5 は旧6-3/6-4 ロジックを移植。P3-6 は旧3-4 を3件ズームに拡張。P3-7 は新規（`PptxSlideZoomLinkReader.TryValidateSectionZoomPlacedUnderLabels` でラベル下のセクションズームを1対1検証）。

| P3（taskId） | 現行メソッド | Legacy メソッド（参照元） | 検証メモ |
|--------------|-------------|-------------------------|----------|
| P3-1 (1) | `CheckTask_1_3_01` | 旧3-1 | SmartArt layout Id=hProcess11 |
| P3-2 (2) | `CheckTask_1_3_02` | 旧3-2 | 色のみ |
| P3-3 (3) | `CheckTask_1_3_03` | 旧3-3 | SmartArt layout Id=target3 |
| P3-4 (4) | `CheckTask_1_3_04` | 旧6-3 (`CheckTask_1_6_03`) | 3Dモデル「虫眼鏡」幅2.5 |
| P3-5 (5) | `CheckTask_1_3_05` | 旧6-4 (`CheckTask_1_6_04`) | 上前面ビュー・高さ6.5 |
| P3-6 (6) | `CheckTask_1_3_06` | 旧3-4 (`CheckTask_1_3_04`) | 3件スライドズーム。タイトルPH「機能の概要」完全一致 |
| P3-7 (7) | `CheckTask_1_3_07` | —（新規） | セクションズーム2件。ラベルとリンク先セクションの対応必須 |

## P4 現行メソッド・破壊的操作 taskId 対応

`PowerPointGrader` の `projectId==4` では **taskId = P4-X**。**現行** `PowerPointChecker1_4.cs` のメソッド名は **P4-X = `CheckTask_1_4_0X`** に揃えた。

| P4（taskId） | 現行メソッド | Legacy メソッド（参照元） | 検証メモ |
|--------------|-------------|-------------------------|----------|
| P4-1 (1) | `CheckTask_1_4_01` | 旧1-7 (`CheckTask_1_1_07`) | 「教育者必見」テキスト |
| P4-2 (2) | `CheckTask_1_4_02` | 旧5-2 (`CheckTask_1_5_02`) | スライド4・アクセント1塗り |
| P4-3 (3) | `CheckTask_1_4_03` | 旧4-2 相当（新規） | 光彩18pt・アクセント6 |
| P4-4 (4) | `CheckTask_1_4_04` | 旧4-1 (`CheckTask_1_4_01`) | 楕円ぼかし+テクスチャライザー |
| P4-5 (5) | `CheckTask_1_4_05` | 旧4-5 (`CheckTask_1_4_05`) | 画像上端揃え |
| P4-6 (6) | `CheckTask_1_4_06` | 旧4-4 (`CheckTask_1_4_04`) | スライド5右画像トリミング |
| P4-7 (7) | `CheckTask_1_4_07` | 旧11-3 (`CheckTask_1_11_03`) | スライド2・アクセント1塗り60％・濃い青枠線0.75pt |
| P4-8 (8) | `CheckTask_1_4_08` | 旧11-6 (`CheckTask_1_11_06`) | スライド3・垂直中央 |

## P5 現行メソッド・破壊的操作 taskId 対応

`PowerPointGrader` の `projectId==5` では **taskId = P5-X**。**現行** `PowerPointChecker1_5.cs` のメソッド名は **P5-X = `CheckTask_1_5_0X`** に揃えた（Phase A: 全スタブ。Phase B で Legacy 参照しつつ順次実装）。

| P5（taskId） | 現行メソッド | Legacy メソッド（参照元） | 検証メモ |
|--------------|-------------|-------------------------|----------|
| P5-1 (1) | `CheckTask_1_5_01` | 旧4-5 (`CheckTask_1_4_05`) | スライド3・円ちょうど4個・右端2pt |
| P5-2 (2) | `CheckTask_1_5_02` | 旧5-4 (`CheckTask_1_5_04`) | スライド5・角丸四角3つ以上・max-min幅0.5pt未満 |
| P5-3 (3) | `CheckTask_1_5_03` | 旧5-3 (`CheckTask_1_5_03`) | スライド4スマイル1・星0・全体スマイル1 |
| P5-4 (4) | `CheckTask_1_5_04` | 旧4-6 (`CheckTask_1_4_06`) | スライド6・Z: 対策講座 > PC教室 > 通信講座 |
| P5-5 (5) | `CheckTask_1_5_05` | 旧5-5 (`CheckTask_1_5_05`) | スライド6・3メンバーグループ・同一AutoShape・寸法一致 |
| P5-6 (6) | `CheckTask_1_5_06` | 旧4-3 (`CheckTask_1_4_03`) | 「MOSって何？」・最大画像・Decorative |
| P5-7 (7) | `CheckTask_1_5_07` | 旧11-5 (`CheckTask_1_11_05`) | スライド2・Graphic/Icon・濃い赤#C00000付近 |

## P6 現行メソッド・破壊的操作 taskId 対応

`PowerPointGrader` の `projectId==6` では **taskId = P6-X**。**現行** `PowerPointChecker1_6.cs` のメソッド名は **P6-X = `CheckTask_1_6_0X`** に揃える（Phase A: 全スタブ `return false`。Phase B で Legacy 参照しつつタスク単位に実装）。

| P6（taskId） | 現行メソッド | Legacy メソッド（参照元） | 検証メモ | Phase B |
|--------------|-------------|-------------------------|----------|---------|
| P6-1 (1) | `CheckTask_1_6_01` | 旧10-1 (`CheckTask_1_10_01`) | 全スライドコメント0 + プロパティ空 | [x] |
| P6-2 (2) | `CheckTask_1_6_02` | 旧8-5 (`CheckTask_1_8_05`) | OpenXML readOnlyRecommended | [x] |
| P6-3 (3) | `CheckTask_1_6_03` | 旧7-4 (`CheckTask_1_7_04`) | ppShowTypeKiosk + VSTO `[Task6-3] Kiosk` | [x] |
| P6-4 (4) | `CheckTask_1_6_04` | 旧10-2 (`CheckTask_1_10_02`) | ショー名「書式のポイント」・スライド4-6 | [x] |
| P6-5 (5) | `CheckTask_1_6_05` | 旧5-1 (`CheckTask_1_5_01`) | Outline/6部/Collate + VSTO `[Task6-5] Print` | [x] |
| P6-6 (6) | `CheckTask_1_6_06` | 旧11-7 (`CheckTask_1_11_07`) | Notes/3部/ページ単位（Collate OFF）+ VSTO `[Task6-6] Print` | [x] |
| P6-7 (7) | `CheckTask_1_6_07` | 旧5-1 (`CheckTask_1_5_01`) | Grayscale/3スライド/4部/PrintColorType + VSTO `[Task6-7] Print` | [x] |

## P7 現行メソッド・破壊的操作 taskId 対応

`PowerPointGrader` の `projectId==7` では **taskId = P7-X**。**現行** `PowerPointChecker1_7.cs` のメソッド名は **P7-X = `CheckTask_1_7_0X`** に揃える（Phase A: 全スタブ `return false`。Phase B で Legacy 参照しつつタスク単位に実装）。

| P7（taskId） | 現行メソッド | Legacy メソッド（参照元） | 検証メモ | Phase B |
|--------------|-------------|-------------------------|----------|---------|
| P7-1 (1) | `CheckTask_1_7_01` | 旧6-1 (`CheckTask_1_6_01`) | スライド1コメント「情報発信の責任を考える」 | [x] |
| P7-2 (2) | `CheckTask_1_7_02` | 旧9-6 (`CheckTask_1_9_06`) | ハイパーリンク・情報学習支援→rabbitway.jp | [x] |
| P7-3 (3) | `CheckTask_1_7_03` | 旧7-3 (`CheckTask_1_7_03`) | スライド7に「まとめ」 | [x] |
| P7-4 (4) | `CheckTask_1_7_04` | 旧9-4 (`CheckTask_1_9_04`) | スライド2〜4に番号+rabbitway.jp・タイトル非表示（5以降は不問） | [x] |
| P7-5 (5) | `CheckTask_1_7_05` | 旧9-5 (`CheckTask_1_9_05`) | スライド5・6のみ「参考事例」 | [x] |

## 破壊的操作免除一覧（`PPTaskValidationConfig`）

**日常の参照は本セクション（MD）を主とする。** 旧 `(projectId, taskId)` のコード確認は `Libraries/Legacy/PPTaskValidationConfig.Legacy.cs`（**凍結・編集禁止**）。それでも不明なときのみ `git log -p -- Libraries/PPTaskValidationConfig.cs`。

### 破壊的操作 config の更新ルール

| ファイル | Px 実装時 | 役割 |
|----------|-----------|------|
| `Libraries/PPTaskValidationConfig.cs` | **更新する** | 実行時の正（Grader / スナップショット比較） |
| `PowerPointAddIn1/ThisAddIn.cs` | **更新する** | VSTO リアルタイム監視（デルタ判定は config と同期） |
| `tasks/PP_類題採点対応表.md` | **更新する** | 人間向け一覧（本セクション） |
| `Libraries/Legacy/PPTaskValidationConfig.Legacy.cs` | **更新しない** | P3 完了時点の旧番号スナップショット（Checker Legacy と同様） |

> Legacy を誤って編集しても**採点結果は変わらない**（ビルド対象外）。壊れるのは旧番号の移植参照のみ。詳細は `Libraries/Legacy/README.md`。

- **現行 config の正**: `Libraries/PPTaskValidationConfig.cs`（実行時はここだけ有効）
- **VSTO 重複**: `PowerPointAddIn1/ThisAddIn.cs` にデルタ判定のコピーあり → config 変更時は両方更新
- **taskId**: P1〜P6 は **新番号（Px-X = taskId）**。projectId 7〜11 は **旧番号のまま**（各 Px 実装時に付け替え）
- **AnimationRemoved**: `SlidesCount` または `ShapesCount` 免除時に自動付与（P1 の 1-1/1-3/1-8 を除く）

### 列の凡例

| 列 | 意味 |
|----|------|
| 免除フラグ | `ShapesCount` / `TextLength` / `SlidesCount` / `ShapePosition`（+ 自動 `AnimationRemoved`） |
| 図形数デルタ | スライド別の許容増減。`無制限` = `int.MaxValue`。`0 or +N` = `IsAllowedShapesCountDelta` で緩和 |
| 文字数デルタ | 同上。例: 9-6 スライド1は `0 or -57` |
| 既存図形位置 | `新規のみ` = `IsShapePositionExemptForNewShapesOnly`。`上限N` = 既存図形の移動・サイズ変更を N 件まで |

### P1（projectId=1, taskId=P1-X）

| taskId | 旧 | 免除フラグ | 図形数デルタ | 文字数デルタ | 既存図形位置 | 備考 |
|--------|-----|-----------|-------------|-------------|-------------|------|
| 1 (P1-1) | 旧1-1 | Slides, Shapes, Text, Position | 挿入スライドのみ無制限、他0 | 挿入スライドのみ無制限、他0 | 挿入スライドのみ免除（スライド番号マッピング） | 1-8 実行後は論理4→物理5 |
| 2 (P1-2) | 旧1-3 | なし | — | — | — | 非表示のみ |
| 3 (P1-3) | 旧1-5 | Shapes, Text, Position | 論理5のみ無制限、他0 | 論理5のみ無制限、他0 | 論理5のみ免除 | |
| 4 (P1-4) | 旧1-6 | Position | — | — | 全スライド免除 | 2段組みでレイアウト変化 |
| 5–7 | 旧10-3等 | なし | — | — | — | P1 実装済み・config は旧1-x系 |
| 8 (P1-8) | 旧3-4 | Shapes, Text, Position | スライド2のみ無制限、他0 | スライド2のみ無制限、他0 | スライド2のみ免除 | サマリーズーム挿入・スライド番号マッピング |

### P2（projectId=2, taskId=P2-X）

| taskId | 免除フラグ | 備考 |
|--------|-----------|------|
| 1–8 | なし | 画面切り替え・アニメーション中心 |

### P3（projectId=3, taskId=P3-X）

| taskId | 旧 | 免除フラグ | 図形数デルタ | 既存図形位置 | 備考 |
|--------|-----|-----------|-------------|-------------|------|
| 1 (P3-1) | 旧3-1 | Shapes, Text, Position | スライド7: 0 | 新規のみ | SmartArt はプレースホルダー内 |
| 2 (P3-2) | 旧3-2 | なし | — | — | 色変更のみ |
| 3 (P3-3) | 旧3-3 | Shapes, Text, Position | スライド6: 0 | 新規のみ | |
| 4 (P3-4) | 旧6-3 | Shapes, Position | スライド1: 0 or +1 | 新規のみ | 旧 config は projectId=6 taskId=3 |
| 5 (P3-5) | 旧6-4 | Position | — | 上限1 | 旧 config は projectId=6 taskId=4 |
| 6 (P3-6) | 旧3-4 | Shapes, Position | 全スライド: 0 or +3 | 新規のみ | タイトル指定のためスライド番号非固定。配置はタイトルPH下端に0pt隙間許容 |
| 7 (P3-7) | — | Shapes, **Text**, Position | スライド2: 0 or +2、スライド1: 0 or +1 | 無制限（副作用） | COM: ラベル下に対応セクションへのズーム必須（入れ替え×）。破壊的: セクション操作の副作用を許容 |

### P4（projectId=4, taskId=P4-X）

| taskId | 旧 | 免除フラグ | 図形数デルタ | 既存図形位置 | 備考 |
|--------|-----|-----------|-------------|-------------|------|
| 1 (P4-1) | 旧1-7 | Text | スライド1: 無制限、他0 | — | テキスト入力のみ |
| 2 (P4-2) | 旧5-2 | なし | — | — | 文字色変更のみ |
| 3 (P4-3) | 旧4-2 | なし | — | — | 図の効果（光彩） |
| 4 (P4-4) | 旧4-1 | なし | — | — | スタイル+アート効果 |
| 5 (P4-5) | 旧4-5 | Position | — | 上限1 | 画像上端揃え |
| 6 (P4-6) | 旧4-4 | Position | — | 上限1 | トリミング |
| 7 (P4-7) | 旧11-3 | なし | — | — | アクセント1塗り60％・濃い青枠線0.75pt |
| 8 (P4-8) | 旧11-6 | Position | — | 上限1 | 垂直中央配置 |

### P5（projectId=5, taskId=P5-X）

| taskId | 旧 | 免除フラグ | 図形数デルタ | 既存図形位置 | 備考 |
|--------|-----|-----------|-------------|-------------|------|
| 1 (P5-1) | 旧4-5 | Position | — | 上限4 | 丸右端揃え（円ちょうど4個・右端2pt） |
| 2 (P5-2) | 旧5-4 | Shapes, Position | — | 上限1 | 四角幅揃え |
| 3 (P5-3) | 旧5-3 | Position | 全スライド: 0 | 新規のみ | 図形変更（図形数不変）。checker: スライド4スマイル1・星0・全体1 |
| 4 (P5-4) | 旧4-6 | Position | — | 新規のみ | z-order（対策講座 > PC教室 > 通信講座） |
| 5 (P5-5) | 旧5-5 | Shapes, **Text**, Position | スライド6: 0 or -2 | 新規のみ | グループ化。Text免除は子図形テキスト集計差対策 |
| 6 (P5-6) | 旧4-3 | なし | — | — | 代替テキスト装飾化（最大画像・Decorative） |
| 7 (P5-7) | 旧11-5 | なし | — | — | 塗りつぶし色のみ（#C00000付近・ただの赤除外） |

### P6（projectId=6, taskId=P6-X）— 完了

| taskId | 旧 | 免除フラグ | 図形数デルタ | 文字数デルタ | 既存図形位置 | 備考 |
|--------|-----|-----------|-------------|-------------|-------------|------|
| 1 (P6-1) | 旧10-1 | なし | — | — | — | ドキュメント検査 |
| 2 (P6-2) | 旧8-5 | なし | — | — | — | 読み取り専用推奨 |
| 3 (P6-3) | 旧7-4 | なし | — | — | — | キオスク設定 |
| 4 (P6-4) | 旧10-2 | なし | — | — | — | 目的別スライドショー「書式のポイント」 |
| 5 (P6-5) | 旧5-1 | なし | — | — | — | アウトライン6部・部単位 |
| 6 (P6-6) | 旧11-7 | なし | — | — | — | ノート3部・**ページ単位**（Collate OFF） |
| 7 (P6-7) | 旧5-1 | なし | — | — | — | グレースケール配布資料3スライド/頁・4部 |

> 旧 `projectId=6 taskId=3/4`（6-3/6-4 3Dモデル）は **P3-4/5 に移植済み**。Phase A で config から削除済み。

### P7（projectId=7, taskId=P7-X）— Phase A 完了

| taskId | 旧 | 免除フラグ | 図形数デルタ | 文字数デルタ | 既存図形位置 | 備考 |
|--------|-----|-----------|-------------|-------------|-------------|------|
| 1 (P7-1) | 旧6-1 | なし | — | — | — | コメント挿入 |
| 2 (P7-2) | 旧9-6 | なし | — | — | — | ハイパーリンク（既存文字列にリンク設定のみ） |
| 3 (P7-3) | 旧7-3 | Slides, Shapes, Text, Position | 無制限 | 無制限 | 無制限 | アウトライン挿入 |
| 4 (P7-4) | 旧9-4 | Shapes, Text, Position | 無制限 | 無制限 | 無制限 | フッター |
| 5 (P7-5) | 旧9-5 | Shapes, Text, Position | 無制限 | 無制限 | 無制限 | フッター（特定スライド） |

> 旧 `projectId=7 taskId=1/2/4`（7-1 セクション、7-2 スライド再利用、7-4 キオスク）は新P7に含まれない。Phase A で config から削除済み。

### 旧 projectId 7〜11（P8 以降・config は旧番号のまま）

| projectId | taskId | 旧タスク | 免除フラグ | 図形数デルタ | 文字数デルタ | 既存図形位置 | 将来の新タスク（参考） |
|-----------|--------|----------|-----------|-------------|-------------|-------------|----------------------|
| 7 | 2 | 旧7-2 | Slides, Shapes, Text, Position | 無制限 | 無制限 | 無制限 | P7 系 |
| 7 | 3 | 旧7-3 | 同上 | 無制限 | 無制限 | 無制限 | P7-3 |
| 8 | 1 | 旧8-1 | Shapes | 無制限 | — | — | P9-1 等 |
| 8 | 2 | 旧8-2 | Shapes | 無制限 | — | — | |
| 9 | 1 | 旧9-1 | Shapes, Position | スライド2: 0 | — | 新規のみ | P9-4 |
| 9 | 4,5 | 旧9-4/5 | Shapes, Text, Position | 無制限 | 無制限 | 無制限 | P7-4, P7-5 |
| 9 | 6 | 旧9-6 | Text, Position | — | スライド1: 0 or -57 | 上限1 | P7-2 |
| 9 | 7 | 旧9-7 | Position | — | — | 無制限 | P8-4 |
| 10 | 5 | 旧10-5 | Shapes, Position | 無制限 | — | 無制限 | P10-1 |
| 10 | 7 | 旧10-7 | Shapes, Position | 無制限 | — | 新規のみ | P10-6 |
| 11 | 1 | 旧11-1 | Shapes, Position | 無制限 | — | 無制限 | P8-3 |
| 11 | 6 | 旧11-6 | Position | — | — | 上限1 | P4-8 |

### 破壊的操作の旧→新移植メモ（taskId 付け替え時に参照）

| 旧 config（projectId, taskId） | 移植先（新） | Legacy 参照 |
|-------------------------------|-------------|-------------|
| (3, 4) 旧3-4 スライドズーム | P3-6 (3,6)、P1-8 (1,8) は別設定 | `PPTaskValidationConfig.Legacy.cs` の case 6 / 1-8 |
| (6, 3) 旧6-3 3D挿入 | P3-4 (3, 4) | Legacy の `projectId == 6 && taskId == 3` |
| (6, 4) 旧6-4 3Dサイズ | P3-5 (3, 5) | Legacy の `projectId == 6 && taskId == 4` |
| (4, 5) 旧4-5 画像配置 | P4-5 (4,5)、P5-1 (5,1) | Legacy の `projectId==4 && taskId==5`（**旧番号**。現行 P4-5 は config の projectId=4 taskId=5） |
| (5, 1) 旧5-1 印刷 | P6-5 (6,5), P6-7 (6,7) | Phase A: P6 config 枠のみ。Phase B で checker + VSTO |

> P5 着手時: 上記「旧 5〜11」表と **Legacy（凍結）** の旧番号行を参照し、現行 `PPTaskValidationConfig.cs` と本 MD の P5 セクションに書き換える。**Legacy ファイル自体は更新しない。**

## 要注意（旧タスクの再利用）

同じ旧タスクを複数の新タスクが参照。Legacy をコピーする際は差分を必ず確認。

| 旧 | 参照する新タスク |
|----|------------------|
| 旧3-4 | P1-8, P3-6 |
| 旧4-5 | P4-5, P5-1 |
| 旧5-1 | P6-5, P6-7 |
| 旧11-4 | P10-7, P10-8 |

## 上書きリスク（他プロジェクト完了後に Legacy 参照が必要）

| 新 | 旧 | 参照 Legacy | 備考 |
|----|-----|-------------|------|
| P4-1 | 旧1-7 | PowerPointChecker1_1.Legacy.cs | CheckTask_1_1_07（教育者必見）。P1 で 1_1.cs 上書き済み |
| P1-5 | 旧10-3 | PowerPointChecker1_10.Legacy.cs | 文字間隔。旧3pt→新4pt、スライド番号も異なる |
| P1-8 | 旧3-4 | PowerPointChecker1_3.Legacy.cs | P1=サマリーズーム。P3-6 も同旧3-4（スライドズーム） |
| P3-6 | 旧3-4 | PowerPointChecker1_3.Legacy.cs | P1-8 実装後も Legacy 参照 |

## プロジェクト実装チェックリスト

- [ ] **P1**（8 タスク）— 完了確認中
- [x] **P2**（8 タスク）— 完了（P2-4 累積採点の既知課題あり）
- [x] **P3**（7 タスク）— 完了（採点・破壊的操作とも検証済み）
- [ ] **P4**（8 タスク）— 実装済み・検証待ち
- [ ] **P5**（7 タスク）— 完了・検証済み
- [x] **P6 Phase A**（7 タスク）— スキャフォールド完了（Checker スタブ・Grader・config）
- [x] **P6 Phase B** — P6-1 → … → P6-7 をタスク単位に実装・検証
- [x] **P7 Phase A**（5 タスク）— スキャフォールド完了（Checker スタブ・Grader・config・問題文 JSON）
- [x] **P7 Phase B** — P7-1 → … → P7-5 をタスク単位に実装・検証
- [ ] **P8**（5 タスク）— 未着手
- [ ] **P9**（5 タスク）— 未着手
- [ ] **P10**（8 タスク）— 未着手

## 推奨実装順

1. **P1** — 検証・コミット
2. ~~**P2**~~ — 完了
3. ~~**P3**~~ — 完了
4. **P4** — 実装済み・手動検証待ち
5. **P5 以降** — 上書きリスク表を確認してから着手
