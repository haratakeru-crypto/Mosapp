# MOS採点QR公開パスの作成手順

オフラインや座席表送信に失敗したとき、結果画面のQRはスマホで開く。  
そのURLは **3科目共通の公開ID** を使い、管理サイトのログイン画面を出さない。

PCからの直接POST（`/api/mos-practice/submit`）は従来どおり。この手順は **QRのパスだけ** を扱う。

## 結論

- 公開パス: `{baseUrl}/p/{qrPathId}?university=…&name=…&wrong=…&sig=…`
- `qrPathId` は Excel / Word / PowerPoint と kouzakanri で **一字一句同じ**
- `/p/` 配下はログイン不要。IDが違うとログインではなく「このQRは無効です」
- 旧パス `/mos-practice/submit` もログイン不要のまま残す（古いQR用）

## 共通ID（現行）

| 場所 | 値 |
|---|---|
| MOS `practiceSubmit.qrPathId` | `5cabe93ea53c64713796ac05fab9239f` |
| kouzakanri 既定値 `DEFAULT_MOS_PRACTICE_QR_PATH_ID` | 同じ |
| kouzakanri 環境変数 `MOS_PRACTICE_QR_PATH_ID` | 未設定なら上記既定。本番で固定したいときだけ Infisical に同じ値を入れる |

完成URLの例（クエリは省略）:

`https://kouzakanri.netlify.app/p/5cabe93ea53c64713796ac05fab9239f`

`ingestKey`（HMAC）とは別物。パスIDはURLに出る。鍵は出さない。

## 新しいIDを作るとき

1. PowerShell で32桁hexを1つ作る。

```powershell
-join ((1..16) | ForEach-Object { '{0:x2}' -f (Get-Random -Maximum 256) })
```

2. 次を **全部同じ値** に更新する。どれか1つでも違うとQRは「このQRは無効です」。

| 対象 | ファイル / 設定 |
|---|---|
| MOS 共通既定 | `MOSapp/MosPracticeClient/MosPracticePublicQrPath.cs` の `DefaultPathId` |
| Excel | `MOSapp/mos_xaml_app/Assets/config.json` の `practiceSubmit.qrPathId` |
| Word | `MOSapp/MOS Word app/Assets/config.json` の `practiceSubmit.qrPathId` |
| PowerPoint | `MOSapp/Mos PowerPoint Mogi App/Assets/config.json` の `practiceSubmit.qrPathId` |
| kouzakanri 既定 | `src/lib/mosPractice/publicQrPath.ts` の `DEFAULT_MOS_PRACTICE_QR_PATH_ID` |
| kouzakanri 本番（任意） | Infisical の `MOS_PRACTICE_QR_PATH_ID` |

3. kouzakanri を本番デプロイしてから、3科目をビルドして教室PCへ配る。  
   サイト側が古いと、新しいIDのQRは無効表示になる（ログインは出ない）。

## MOSアプリに書く場所

`baseUrl` / `ingestKey` は座席表送信の本番設定と同じ。そこに `qrPathId` を足す。

```json
"practiceSubmit": {
  "baseUrl": "https://（kouzakanriの公開URL）",
  "ingestKey": "（Infisical の MOS_PRACTICE_INGEST_KEY と同じ）",
  "qrPathId": "5cabe93ea53c64713796ac05fab9239f"
}
```

`qrPathId` を省略してもアプリは既定IDを使う。運用では3科目の JSON に明示しておく。

ランチャー（`MOSapp/MOSapp/Assets/config.json`）には不要。

## サーバー側（kouzakanri）で揃えるもの

リポジトリ: `C:\Rabbit Website\kouzakanri`

1. ルート `src/app/p/[publicId]/page.tsx` がある（ログインレイアウトの外の `/p/`）
2. `isPublicAppPath` が `/p` と `/p/` を許可している（`src/lib/navigation/publicAppPaths.ts`）
3. `getMosPracticeQrPathId()` が MOS の `qrPathId` と一致する
4. Infisical に `MOS_PRACTICE_QR_PATH_ID` を入れる場合は、Netlify / 本番プロセスに届いていること

`.env.example` にキー名のコメントがある。値はチャットや Git に貼らない運用でも、現行の共通IDはアプリ JSON に載る。

## 入れたあとの確認

1. 管理サイトをデプロイする
2. ブラウザで（ログインせず）次を開く  
   `https://（公開URL）/p/5cabe93ea53c64713796ac05fab9239f`  
   → ログイン画面に飛ばない。クエリが無いので「このQRは無効です」でよい
3. わざと違うID  
   `https://（公開URL）/p/wrong-id`  
   → 同じくログインなしで無効表示
4. 3科目をビルドする
5. オフラインまたは送信失敗でQRを出す。URLが `/p/5cabe93ea53c64713796ac05fab9239f?` で始まる
6. スマホで読み、同意して「座席表に送る」

未設定のまま（`baseUrl` / `ingestKey` が空）だとQR自体が出ない。

## 関連

- 送信先URLと HMAC 鍵: `task/MOS採点座席表送信_本番設定.md`
- 受講生画面の確認: `task/MOS採点QR_確認チェックリスト.md`（正本は `kouzakanri/docs/MOS採点QR_確認チェックリスト.md`）
- 2026-09-18 の確認記録: `task/MOS採点QR_確認結果.md`
