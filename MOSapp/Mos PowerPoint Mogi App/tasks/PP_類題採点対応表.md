# PowerPoint 類題採点対応表

修正版問題文（新タスク）と旧採点ロジックの対応。Legacy 退避は P1 改修前 baseline を参照。

## Baseline

| 項目 | 値 |
|------|-----|
| Legacy 作成日 | 2026-06-15 |
| P1 改修前 commit（Checker1_1 旧版） | `d03e149` |
| P1 改修 commit（参考） | `81ecce9` |
| JSON 旧版 source（P2以降） | commit `9df05ff`（旧タスク構成・旧文言。新CSVのタスク数とは一致しない） |
| Legacy 配置 | `Libraries/Group1/Legacy/*.Legacy.cs` |
| 元 CSV | `Mos PowerPoint Mogi App/PP修正版問題文一覧_類似付き.csv` |

> **git 確認（依頼1実施時）**: `PowerPointChecker1_1.cs` は HEAD（`81ecce9`）に P1 改修済み。作業ツリーに**未コミットの追加差分**あり。Legacy の 1_1 は commit `d03e149` から復元。

> **JSON について**: P1 のみ新問題文（8タスク）。P2〜P10 は commit `9df05ff` の旧文言・旧タスク構成に戻した。各プロジェクトの採点実装完了時に、新CSVどおりの問題文へ更新する。

## 凡例

- **旧 X-Y**: 旧プロジェクト X のタスク Y
- **Legacy ファイル**: `PowerPointChecker1_{旧プロジェクト}.Legacy.cs`
- **Legacy メソッド**: `CheckTask_1_{旧P}_{旧T}`（ゼロ埋め2桁）
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
| P2-1 | 旧2-1 | あなたはMOS合格対策講座のコースガイダンス資料を作成しています。 すべてのスライドに、画面切り替え「スプリット」を設定します。効果のオプションを「ワイプアウト | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_01 | P2 実装時に PowerPointChecker1_2.cs 上書き |  |  |
| P2-2 | 旧2-2 | すべての画面切り替えの継続時間を3秒に設定します。 | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_02 | P2 実装時に PowerPointChecker1_2.cs 上書き |  |  |
| P2-3 | 旧2-3 | スライド3、4、5に「切り替え」の画面切り替え効果を設定します。 | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_03 | P2 実装時に PowerPointChecker1_2.cs 上書き |  |  |
| P2-4 | — | すべてのスライドが、"５"秒後に自動で次のスライドへ進むように画面切り替えのタイミングを設定します。 | — | — | — | [画面切り替え]タブ
↓
[自動]チェックボックスを
クリックしオン
↓
[自動… | 新規 |
| P2-5 | 旧2-4 | スライド６の3Ｄモデル「黒板」に、アニメーション「ジャンプしてターン」を設定します。 | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_04 | P2 実装時に PowerPointChecker1_2.cs 上書き |  |  |
| P2-6 | 旧2-5 | スライド2の男の子と？マークの２つの画像が、スライドの左上隅から登場するようにします。アニメーションの継続時間は"0.55"秒にします。 | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_05 | P2 実装時に PowerPointChecker1_2.cs 上書き |  |  |
| P2-7 | 旧2-6 | スライド「今年度募集について」の箇条書きに設定されたアニメーションの効果を「プラス」に変更します。また、クリックするとすべて同時に動くようにします。 | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_06 | P2 実装時に PowerPointChecker1_2.cs 上書き |  |  |
| P2-8 | 旧2-7 | スライド４の星の図形にアニメーションの軌跡「ニュートロン」を設定します。円の図形にはアニメーションを設定しません。 | PowerPointChecker1_2.Legacy.cs | CheckTask_1_2_07 | P2 実装時に PowerPointChecker1_2.cs 上書き |  |  |
| P3-1 | 旧3-1 | あなたはPower Point 新機能についての発表用資料を作成しています。 スライド７にSmartArtグラフィック「タイムライン」の手順を追加し、文字列「" | PowerPointChecker1_3.Legacy.cs | CheckTask_1_3_01 | P3 実装時に PowerPointChecker1_3.cs 上書き |  |  |
| P3-2 | 旧3-2 | スライド７のSmartArtグラフィックに、色「グラデーション循環-アクセント６」を設定します。 | PowerPointChecker1_3.Legacy.cs | CheckTask_1_3_02 | P3 実装時に PowerPointChecker1_3.cs 上書き |  |  |
| P3-3 | 旧3-3 | スライド６の箇条書きを「ターゲットリスト」のSmartArtに変更します。 | PowerPointChecker1_3.Legacy.cs | CheckTask_1_3_03 | P3 実装時に PowerPointChecker1_3.cs 上書き |  |  |
| P3-4 | 旧6-3 | スライド１に3Dモデル「虫眼鏡」を挿入します。幅を「"2.5"」に変更して中央の楕円の図形の中に配置します。正確な位置は問いません。 | PowerPointChecker1_6.Legacy.cs | CheckTask_1_6_03 | P6 実装時に PowerPointChecker1_6.cs 上書き |  |  |
| P3-5 | 旧6-4 | スライド10の3Dモデルのビューを上前面にし、高さを「"6.5"」に変更します。 | PowerPointChecker1_6.Legacy.cs | CheckTask_1_6_04 | P6 実装時に PowerPointChecker1_6.cs 上書き |  |  |
| P3-6 | 旧3-4 | 「機能の概要」のスライドにスライドズームを挿入して「画面録画で説明」「ズーム機能で訴求力アップ」「デザインアイデアで魅力的に！」へリンクを作成します。タイトルの | PowerPointChecker1_3.Legacy.cs | CheckTask_1_3_04 | P3 実装時に PowerPointChecker1_3.cs 上書き | 旧3-4 は P1-8, P3-6 でも参照 |  |
| P3-7 | — | スライド２にセクションズームのリンクを挿入します。セクション「1.機能の概要」と「2.伝わるスライドの要素」にリンクを作成し、それぞれ文字の下に配置します。 | — | — | — | スライド2を選択
↓
[挿入]タブ
↓
[ズーム]をクリック
↓
[セクションズ… | 新規 |
| P4-1 | 旧1-7 | あなたは英語教育プログラムの提案用資料を作成しています。 スライド１の青い図形に、「"教育者必見"」と入力します。 | PowerPointChecker1_1.Legacy.cs | CheckTask_1_1_07 | P1 実装時に PowerPointChecker1_1.cs 上書き |  |  |
| P4-2 | 旧5-2 | スライド４の文字列「いつでも体験可能です!」に、文字の塗りつぶし「青、アクセント1」を設定します。 | PowerPointChecker1_5.Legacy.cs | CheckTask_1_5_02 | P5 実装時に PowerPointChecker1_5.cs 上書き |  |  |
| P4-3 | 旧4-2 | スライド１枚目の子供の画像に、図の効果「光彩:18pt;緑､アクセントカラー6」を設定します。 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_02 | P4 実装時に PowerPointChecker1_4.cs 上書き |  |  |
| P4-4 | 旧4-1 | スライド１の子供の画像に、スタイル「楕円 ぼかし」を設定し、「テクスチャライザー」のアート効果を設定します。 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_01 | P4 実装時に PowerPointChecker1_4.cs 上書き |  |  |
| P4-5 | 旧4-5 | スライド５の右の画像を左の画像の上端に合わせます。水平方向には動かさないようにします。 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_05 | P4 実装時に PowerPointChecker1_4.cs 上書き | 旧4-5 は P4-5, P5-1 でも参照 |  |
| P4-6 | 旧4-4 | スライド５の右側の画像の右端を、スライドの右端に揃えてトリミングします。画像の右端以外は変更しないでください。トリミングした領域をプレゼンテーションから完全に削 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_04 | P4 実装時に PowerPointChecker1_4.cs 上書き |  |  |
| P4-7 | 旧11-3 | 2枚目のスライドのテキストボックスに、塗りつぶし「青、アクセント1、白+基本色60％」、枠線「濃い青」、太さ「0.75pt」を設定します。 | PowerPointChecker1_11.Legacy.cs | CheckTask_1_11_03 | P11 実装時に PowerPointChecker1_11.cs 上書き |  |  |
| P4-8 | 旧11-6 | スライド３のコンテンツ領域にあるテキストボックスを、スライドの垂直方向の中央に配置します。 | PowerPointChecker1_11.Legacy.cs | CheckTask_1_11_06 | P11 実装時に PowerPointChecker1_11.cs 上書き |  |  |
| P5-1 | 旧4-5 | あなたはMOS合格対策講座のコースガイダンス資料を作成しています。 スライド３の丸の図形４個の右端を揃えます。 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_05 | P4 実装時に PowerPointChecker1_4.cs 上書き | 旧4-5 は P4-5, P5-1 でも参照 |  |
| P5-2 | 旧5-4 | スライド５の小さい四角の大きさを他の四角と同じ幅にします。 | PowerPointChecker1_5.Legacy.cs | CheckTask_1_5_04 | P5 実装時に PowerPointChecker1_5.cs 上書き |  |  |
| P5-3 | 旧5-3 | スライド４の星の図形をスマイルに変更します。 | PowerPointChecker1_5.Legacy.cs | CheckTask_1_5_03 | P5 実装時に PowerPointChecker1_5.cs 上書き |  |  |
| P5-4 | 旧4-6 | スライド６の図形を手前から「対策講座」「PC教室」「通信講座」になるように重なりの順番を変更します。 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_06 | P4 実装時に PowerPointChecker1_4.cs 上書き |  |  |
| P5-5 | 旧5-5 | スライド６の3つの論理積ゲートの図形をグループ化します。 | PowerPointChecker1_5.Legacy.cs | CheckTask_1_5_05 | P5 実装時に PowerPointChecker1_5.cs 上書き |  |  |
| P5-6 | 旧4-3 | スライド「MOSって何？」の男の子の画像の代替テキストを装飾化し、スクリーンリーダーに表示させないようにします。 | PowerPointChecker1_4.Legacy.cs | CheckTask_1_4_03 | P4 実装時に PowerPointChecker1_4.cs 上書き |  |  |
| P5-7 | 旧11-5 | スライド２の[?]のアイコンに「濃い赤」の塗りつぶしを設定します。 | PowerPointChecker1_11.Legacy.cs | CheckTask_1_11_05 | P11 実装時に PowerPointChecker1_11.cs 上書き |  |  |
| P6-1 | 旧10-1 | あなたは、「伝わる」プレゼンテーションのテクニックを紹介する資料を作成しています。 プレゼンテーションからドキュメントのプロパティと個人情報を削除します。他の情 | PowerPointChecker1_10.Legacy.cs | CheckTask_1_10_01 | P10 実装時に PowerPointChecker1_10.cs 上書き |  |  |
| P6-2 | 旧8-5 | プレゼンテーションを常に読み取り専用にします。 | PowerPointChecker1_8.Legacy.cs | CheckTask_1_8_05 | P8 実装時に PowerPointChecker1_8.cs 上書き |  |  |
| P6-3 | 旧7-4 | スライドショーを自動プレゼンテーションとして設定します。 | PowerPointChecker1_7.Legacy.cs | CheckTask_1_7_04 | P7 実装時に PowerPointChecker1_7.cs 上書き |  |  |
| P6-4 | 旧10-2 | スライド４，５，６を選択し、「"書式のポイント"」という名前の目的別スライドショーを作成します。スライドショーは実行しません。 | PowerPointChecker1_10.Legacy.cs | CheckTask_1_10_02 | P10 実装時に PowerPointChecker1_10.cs 上書き |  |  |
| P6-5 | 旧5-1 | すべてのスライドをアウトラインで、部単位に"６"部印刷するように設定します。 | PowerPointChecker1_5.Legacy.cs | CheckTask_1_5_01 | P5 実装時に PowerPointChecker1_5.cs 上書き | 旧5-1 は P6-5, P6-7 でも参照 |  |
| P6-6 | 旧11-7 | ノートですべてのスライドを"3"部印刷しなさい。ただし、1ページ目を全て印刷したあとに2ページ目を印刷するようにします。 | PowerPointChecker1_11.Legacy.cs | CheckTask_1_11_07 | P11 実装時に PowerPointChecker1_11.cs 上書き |  |  |
| P6-7 | 旧5-1 | 印刷オプションで、グレースケールの配布資料を、１ページに3スライドのレイアウトで"４"部印刷するように設定します。印刷は実行しないでください。 | PowerPointChecker1_5.Legacy.cs | CheckTask_1_5_01 | P5 実装時に PowerPointChecker1_5.cs 上書き | 旧5-1 は P6-5, P6-7 でも参照 |  |
| P7-1 | 旧6-1 | あなたは、情報モラルの重要性について発表する資料を作成しています。 スライド１にコメント「"情報発信の責任を考える"」を挿入します。 | PowerPointChecker1_6.Legacy.cs | CheckTask_1_6_01 | P6 実装時に PowerPointChecker1_6.cs 上書き |  |  |
| P7-2 | 旧9-6 | スライド１枚目の文字列「情報学習支援」に、Webページ「"https://rabbitway.jp/"」を表示するハイパーリンクを設定します。 | PowerPointChecker1_9.Legacy.cs | CheckTask_1_9_06 | P9 実装時に PowerPointChecker1_9.cs 上書き |  |  |
| P7-3 | 旧7-3 | スライド６の後ろに、文書「まとめ」のアウトラインを使用してスライドを挿入します。 | PowerPointChecker1_7.Legacy.cs | CheckTask_1_7_03 | P7 実装時に PowerPointChecker1_7.cs 上書き |  |  |
| P7-4 | 旧9-4 | スライドのフッターに、スライド番号と「"rabbitway.jp"」をタイトルスライド以外に追加します。 | PowerPointChecker1_9.Legacy.cs | CheckTask_1_9_04 | P9 実装時に PowerPointChecker1_9.cs 上書き |  |  |
| P7-5 | 旧9-5 | ５枚目の「事例１」と６枚目の「事例2」のフッターに「"参考事例"」と挿入します。他のスライドには表示しません。 | PowerPointChecker1_9.Legacy.cs | CheckTask_1_9_05 | P9 実装時に PowerPointChecker1_9.cs 上書き |  |  |
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
- [ ] **P2**（8 タスク）— 未着手
- [ ] **P3**（7 タスク）— 未着手
- [ ] **P4**（8 タスク）— 未着手
- [ ] **P5**（7 タスク）— 未着手
- [ ] **P6**（7 タスク）— 未着手
- [ ] **P7**（5 タスク）— 未着手
- [ ] **P8**（5 タスク）— 未着手
- [ ] **P9**（5 タスク）— 未着手
- [ ] **P10**（8 タスク）— 未着手

## 推奨実装順

1. **P1** — 検証・コミット
2. **P2** — 旧2-x がすべて `PowerPointChecker1_2.Legacy.cs` に揃っている
3. **P3 以降** — 上書きリスク表を確認してから着手
