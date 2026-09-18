# MOS採点結果を座席表へ送る本番設定

空のままだと結果画面に「送信先が設定されていません」と出る。  
入れる値は **2つだけ**。Excel / Word / PowerPoint の3科目で同じにする。

ランチャー（`MOSapp/MOSapp/Assets/config.json`）には `practiceSubmit` は不要。科目アプリだけ。

## 必要な値

| MOS側（config.json） | kouzakanri側 | 内容 |
|---|---|---|
| `practiceSubmit.baseUrl` | 管理サイトの公開URL | `https://` から始まるサイト原点。末尾スラッシュなし |
| `practiceSubmit.ingestKey` | 環境変数 `MOS_PRACTICE_INGEST_KEY` | PCからの POST ヘッダと、失敗時QRの HMAC に使う共通鍵 |

アプリが読むパスは、各 exe と同じフォルダの `Assets\config.json`。

## 値の取り方（kouzakanri）

リポジトリ: `C:\Rabbit Website\kouzakanri`

### 1. baseUrl

本番の管理サイトURL。ブラウザでログイン画面が開くアドレスのオリジンだけ。

例:

- OK: `https://kouzakanri.example.com`
- NG: `https://kouzakanri.example.com/`（末尾スラッシュはアプリ側で削るが、書かない）
- NG: `https://kouzakanri.example.com/mos-practice/submit`（パスは付けない）

確認先の例:

- ブラウザのアドレスバー
- GitHub Repository Variables の `APP_BASE_URL`（座席表クリーンアップと同じ）
- Netlify の本番サイトURL

確認用（認証なしで JSON が返ればURLは合っている）:

`https://（公開URL）/api/mos-practice/lookups`

### 2. ingestKey

Infisical（本番 / prod）の `MOS_PRACTICE_INGEST_KEY` と **一字一句同じ**。

- 未設定なら、十分に長いランダム文字列を1つ作り、**先に Infisical へ入れる**。そのあとアプリへコピーする。
- サーバーに無い鍵をアプリだけに入れると、POST は 401、QRは「このQRは無効です」になる。
- 鍵を変えたら、配布済み教室PCの `config.json` も同じ値に更新する。

PowerShell で候補を作る例（この値をチャットや Git に貼らない）:

```powershell
-join ((1..32) | ForEach-Object { '{0:x2}' -f (Get-Random -Maximum 256) })
```

## MOSアプリに書く場所

3ファイルとも、既存の空欄を埋める。他のキー（`excelDataPath` など）はいじらない。

```json
"practiceSubmit": {
  "baseUrl": "https://（kouzakanriの公開URL）",
  "ingestKey": "（Infisical の MOS_PRACTICE_INGEST_KEY と同じ）",
  "qrPathId": "（3科目共通。task/MOS採点QR公開パス_作成手順.md を参照）"
}
```

| 科目 | ファイル |
|---|---|
| Excel | `MOSapp/mos_xaml_app/Assets/config.json` |
| Word | `MOSapp/MOS Word app/Assets/config.json` |
| PowerPoint | `MOSapp/Mos PowerPoint Mogi App/Assets/config.json` |

ビルドすると `bin\Debug` / `bin\Release` の `Assets\config.json` にコピーされる。  
発行物は `Merge-AppOutputs-ToPublish.ps1` で `publish\App` へ載るので、教室配布は **科目をビルドし直してからマージ**する。

exe だけ差し替えて `Assets\config.json` が古いと、また空のままになる。

## サーバー側（kouzakanri）で先に揃えるもの

アプリを配る前に、本番で次が動いていること。

1. Infisical 本番に `MOS_PRACTICE_INGEST_KEY` がある（Netlify / 本番プロセスに届いている）
2. `GET /api/mos-practice/lookups` が公開で JSON を返す
3. `POST /api/mos-practice/submit` が同じ鍵で受け付ける
4. `/p/{qrPathId}` がログインなしで開ける（オフラインQR用。手順は `task/MOS採点QR公開パス_作成手順.md`。受講生画面の確認は `task/MOS採点QR_確認チェックリスト.md`）
5. 座席表スプレッドシート書き込み用のサービスアカウント設定が本番で有効

鍵が無いと submit API は `503`（`MOS_PRACTICE_INGEST_KEY is not configured`）。

## 入れたあとの確認

1. 3科目をビルドする
2. 大学名タブで大学名・氏名を登録（教室がある大学は自動／プルダウン）
3. 初回採点まで進めて結果画面を出す
4. ネットがあれば「座席表に送信しました」（QRなし）
5. オフラインや座席表不一致なら QR が出る。スマホで読み、同意して「座席表に送る」
6. 座席表の該当氏名セルに、時刻と ×の数が入る

未設定のままだと「送信先が設定されていません」（QRなし）。

## 注意

- `ingestKey` は教室PCの exe 隣の JSON に入る。Git に載せるかは運用判断。載せるならリポジトリ権限を限定する。
- 大学名・氏名は座席表の漢字表記と一致させる。教室が複数ある大学は、候補の教室名（座席表の `templateSheetName`）を選ぶ。
- 大学名の予測は `baseUrl` が空でも同梱 `universities.json` で動く。オンライン更新と座席表送信だけ `baseUrl` / `ingestKey` が必要。
