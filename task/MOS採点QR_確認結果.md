# MOS採点QR — 確認結果（2026-09-18）

`task/MOS採点QR_確認チェックリスト.md`（正本: `kouzakanri/docs/MOS採点QR_確認チェックリスト.md`）に沿って確認した記録。

**結論: MOSアプリ側の設定と画面コードは揃っている。本番サイトにはまだ載っていない。**

公開URL（3科目の `practiceSubmit.baseUrl`）:

`https://kouzakanri.netlify.app`

公開パスID:

`5cabe93ea53c64713796ac05fab9239f`

---

## 最短手順（チェックリスト §7）

| 手順 | 結果 |
|------|------|
| 1. 未ログインで `/api/mos-practice/lookups` が JSON | **NG**。本番は 404 |
| 2. 未ログインで `/p/{id}` が「このQRは無効です」 | **NG**。本番は 404。HTMLは「読み込み中…」のあと認証レイアウトに入り、ログインへ流れる |
| 3〜5. 科目アプリで採点 → 座席表 | 本番ルートが無いので未実施 |

---

## 2. 事前確認（本番・未ログイン）

対象オリジンは末尾スラッシュなし。パスは付けない。

| 項目 | URL | HTTP | 判定 |
|------|-----|------|------|
| 2-1 lookups | `/api/mos-practice/lookups` | 404（Next の HTML） | **NG**。JSONではない。公開URL違いというより、**本番に API 自体が無い** |
| 2-2 QRパス | `/p/5cabe93ea53c64713796ac05fab9239f` | 404 | **NG**。「このQRは無効です」ではない |
| 2-2 違うID | `/p/wrong-id` | 404 | **NG**。ログインなしの無効表示にもなっていない |
| 2-3 旧パス | `/mos-practice/submit` | 404 | **NG** |
| 対照 | `/login` | 200 | ログイン画面は存在する |

`/p/` の 404 HTML には「このQRは無効です」も「自己採点報告フォーム」も無い。スピナー（読み込み中）と `AuthProvider` があり、未知パスとして保護レイアウトに入る。

### 原因

ローカルの kouzakanri（ブランチ `mosapp`）には次があるが、**未コミット／未デプロイ**。

- `src/app/p/`
- `src/app/mos-practice/`
- `src/app/api/mos-practice/`
- `src/lib/mosPractice/`
- `src/lib/navigation/publicAppPaths.ts`

本番に載るまで、受講生QRはチェックリストどおり動かない。

---

## 3. 画面

クエリ付きの本番画面は、ルート未デプロイのため開けない。  
ローカル実装はチェックリスト §1 / §3-2 と一致している。

| 項目 | コード |
|------|--------|
| 見出し | 「自己採点報告フォーム」 |
| ×の数 / 大学名（教室があれば `大学名 / 教室名`）/ お名前 | 表示のみ。手入力なし |
| 個人情報の同意 | 初期オフ。同意するまで「座席表に送る」は押せない |
| 成功時 | 「座席表に送信しました」 |
| ログイン欄・共通アカウント | 無し |

アプリが出すQRのURL形:

```
https://kouzakanri.netlify.app/p/5cabe93ea53c64713796ac05fab9239f?university=…&name=…&wrong=…&at=…&sig=…
```

`/login` 始まりや、`baseUrl` に `/mos-practice/submit` まで付ける形にはなっていない。  
教室が複数ある大学は `classroom`、科目は `subject` が付く。

スマホ実読・同意送信・座席表セル反映（§3-2 後半 / §4-1 / §4-2）は本番デプロイ後に行う。

---

## 4. MOSアプリ設定

3科目とも同じ。ランチャー（`MOSapp/MOSapp/Assets/config.json`）に `practiceSubmit` は無い。

| 科目 | ファイル | `baseUrl` | `qrPathId` | `ingestKey` |
|------|----------|-----------|------------|-------------|
| Excel | `MOSapp/mos_xaml_app/Assets/config.json` | `https://kouzakanri.netlify.app` | `5cabe93ea53c64713796ac05fab9239f` | 3科目同一（空ではない） |
| Word | `MOSapp/MOS Word app/Assets/config.json` | 同じ | 同じ | 同じ |
| PowerPoint | `MOSapp/Mos PowerPoint Mogi App/Assets/config.json` | 同じ | 同じ | 同じ |

アプリ既定 `MosPracticePublicQrPath.DefaultPathId` も同じ ID。  
失敗時QRは `QrUrlBuilder` が `/p/{qrPathId}` を付ける。オンライン成功時は「座席表に送信しました」でQRは出さない。

発行物は exe 隣の `Assets\config.json` が使われる。exe だけ差し替えると古い JSON のままになる。

---

## 5. サーバー（kouzakanri）

ローカルコードは揃っている。

- 公開ページ `src/app/p/[publicId]/page.tsx`
- `isPublicAppPath` が `/p` と `/p/` を許可
- `GET /api/mos-practice/lookups`（認証不要）
- `POST /api/mos-practice/submit`（`MOS_PRACTICE_INGEST_KEY` または QR の `sig`）

本番 Infisical の鍵そのものはこの確認では見ていない。  
lookups / submit / `/p/` が本番 404 なので、**今の本番では座席表送信 API は動いていない**。

鍵が届いたあとの切り分けはチェックリスト §6 のとおり。

- 鍵なし → submit は `503`（`MOS_PRACTICE_INGEST_KEY is not configured`）
- 鍵違い → POST は `401`、QRは「このQRは無効です」

---

## 追記（9/18）デプロイ試行の結果

`mosapp` を `develop` に push 済み（`575c54c`）。CI のビルドは成功し、`/p/[publicId]` と
`/mos-practice/*` も生成された。**Netlify への配信だけが認証エラーで止まっている。**

```
Error: Unauthorized: could not retrieve project
```

- GitHub Actions の `NETLIFY_AUTH_TOKEN`（最終更新 2026-08-17）が失効した模様。9/14 の実行は成功していた。
- ローカルの netlify CLI は `haratakeru@rabbitway.jp` でログイン済みだが、`kouzakanri` プロジェクトが見えない（別チーム所有）。
- そのため本番 `kouzakanri.netlify.app` は依然 9/14 時点のまま。lookups と `/p/` は 404 のまま。

あわせて判明した点:

- CI の deploy は `netlify deploy --dir=.next --context develop` で、`--prod` が無い。
  9/14 の成功時も出力は Draft URL（`https://<deployId>--kouzakanri.netlify.app`）だった。
  トークンを直すだけでは本番 URL は更新されない。`--prod` か Netlify 側での publish が要る。
- 途中でビルドを止めていた型エラー（テスト3ファイル）を修正済み。
  9/16 の `develop` の CI はこれで失敗していた。

### 先に必要な対応

1. Netlify のパーソナルアクセストークンを再発行し、GitHub の `NETLIFY_AUTH_TOKEN` を更新する
2. `kouzakanri` プロジェクトを本番更新する経路を決める（CI に `--prod` を足す / Netlify UI で publish）

MOSアプリ側は変更不要。`baseUrl` / `qrPathId` / `ingestKey` は現状のままでよい。

---

## 次にやること

1. ~~kouzakanri の `mosapp` をコミットし、本番デプロイする~~ → 上記「追記（9/18）」のとおり認証で停止中
2. デプロイ後、チェックリスト §7 の 1 と 2 をやり直す  
   - lookups が JSON  
   - `/p/{id}` がログインなしで「このQRは無効です」
3. 3科目をビルドし、exe 隣の `Assets\config.json` が上表どおりか確認する
4. オンライン採点（QRなし）とオフライン／失敗時QR（スマホ同意送信）を実機で通す
5. 座席表の氏名セルに時刻と ×の数が入ることを確認する

ここまで通れば、旧 Formzu 相当の受講生フローは確認完了。

---

## 関連

- [MOS採点QR_確認チェックリスト.md](./MOS採点QR_確認チェックリスト.md)
- [MOS採点QR公開パス_作成手順.md](./MOS採点QR公開パス_作成手順.md)
- [MOS採点座席表送信_本番設定.md](./MOS採点座席表送信_本番設定.md)
