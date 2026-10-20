解答手順ガイド資産（Word / Group1）

配置:
  Group1/Project{N}/Task{T}/manifest.json
  Group1/Project{N}/Task{T}/step01.png
  Group1/Project{N}/Task{T}/step02.png
  ...

コミット対象:
  manifest.json と stepNN.png のみ。
  _source/ と *_original.png は gitignore。

再分割（開発用・gitignore）:
  _source/Word_Project{N}.pdf
  _source/render_project.py
  _source/hdr_ys.py
  _source/guide_common.py
  _source/split_project{N}.py     （Phase B でレイアウト確定後）
  _source/project{N}_pages/

動線:
  1. PDF を _source/Word_Project{N}.pdf に置く
  2. python render_project.py {N}     → project{N}_pages/pageNN.png
  3. python hdr_ys.py {N}             → 操作/完成見出しの Y
  4. ページ数 ≟ 問題文 JSON のタスク数（P1=5 / P2=5）
  5. split_project{N}.py で操作単位クロップ → Task{T}/

Phase B（Project1–2 実測、2x 1191×1684）:
  - 画像のみ PDF（text_len=0）。表紙・目次なし。1ページ＝1タスク。
  - 全ページ単カラム（左は操作手順テキスト、右が紙芝居用の操作画面）。
    hdr_ys の R=[] が二段組なしの確認。Project3 以降は再測。
  - X0 は Excel と同じで確定:
      奇数 285（☝ 271–283 の右。「操作」は残る）
      偶数 334（☝ 316–329 の右。破線・☝を入れない）
      右端 1175（本文インクは概ね 930–974 で終わる。余白込み）
  - 見出し Y 目安（操作…／最後が完成）:
      1-1: 201 / 551 / 903 / 1253
      1-2: 201 / 609 / 1026
      1-3: 201 / 583 / 918 / 1257
      1-4: 201 / 590 / 935 / 1279
      1-5: 201 / 586 / 958
      2-1: 201 / 557 / 907 / 1269
      2-2: 201 / 573 / 1005
      2-3: 201 / 586 / 983
      2-4: 201 / 543 / 913 / 1256
      2-5: 201 / 566 / 910 / 1261

Phase C:
  - Project1 Task1–5 切り取り済み。
  - 完成下端は窓枠直後（1-1≈1618 / 1-2≈1427 / 1-3≈1622 / 1-4≈1624 / 1-5≈1346 /
    2-1≈1622 / 2-2≈1400 / 2-3≈1382 / 2-4≈1622 / 2-5≈1622）。
    操作が少ないページは完成が上に来るので、ページ余白の線を下端にしない。
  - 偶数は X0=334、pad_left=0。Project1–2 Task1–5 切り取り済み。次は Project3。

仕様:
  mos_xaml_app/tasks/Excel_解答手順ガイド_仕様.md
  （本線は切り取り＋紙芝居 UI。優先: Excel → Word → PPT）
