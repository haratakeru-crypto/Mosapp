using System;
using System.Collections.Generic;

namespace Libraries
{
    /// <summary>
    /// Excel 破壊的操作検知で使用する操作カテゴリ。
    /// VSTO ログの operationType と 1:1 で合わせて運用する想定。
    /// 
    /// [発火タイミングの補足]
    /// ・[即時発火]: 操作した瞬間に記録。
    /// ・[タスク切替時フラッシュ]: タスク切替時に「全ブックの全シート」を走査して記録。
    /// ・[SheetDeactivate ＋ タスク切替時]: シート切替時、またはタスク切替時に「選択中のシートのみ」を走査して記録。
    /// ・[ハイブリッド]: 即時発火とタスク切替時（全シート走査）の両方で検知。
    /// </summary>
    public enum ExcelOperationType
    {
        EditCellValue,          // [即時発火] セル値の入力・上書き（手入力、貼り付けなど）
        EditCellFormula,        // [即時発火] 数式の入力・変更
        EditCellFormat,         // [SheetDeactivate ＋ タスク切替時] セル書式変更（表示形式、フォント、配置、罫線、塗りつぶし等）
        InsertRows,             // [タスク切替時フラッシュ] 行の挿入
        DeleteRows,             // [タスク切替時フラッシュ] 行の削除
        InsertColumns,          // [タスク切替時フラッシュ] 列の挿入
        DeleteColumns,          // [タスク切替時フラッシュ] 列の削除
        InsertShapeOrImage,     // [タスク切替時フラッシュ] 図形/画像/アイコン/テキストボックス等の追加
        DeleteShapeOrImage,     // [タスク切替時フラッシュ] 図形/画像等の削除
        MoveOrResizeShape,      // [タスク切替時フラッシュ] 図形/画像等の移動・サイズ変更・整列
        InsertHyperlink,        // [ハイブリッド] ハイパーリンクの挿入・変更 (SheetChange後＋タスク切替時)
        SortOrFilter,           // [タスク切替時フラッシュ] 並べ替え・フィルター実行
        SetPrintArea,           // [SheetDeactivate ＋ タスク切替時] 印刷範囲の設定
        SetPrintTitle,          // [SheetDeactivate ＋ タスク切替時] タイトル行/タイトル列の設定
        SetHeaderFooter,        // [SheetDeactivate ＋ タスク切替時] ヘッダー/フッターの設定
        SetFreezePanes,         // [SheetDeactivate ＋ タスク切替時] ウィンドウ枠の固定
        SetWorkbookProperty,    // [即時発火] Backstage 経由の文書プロパティ変更（ログ詳細は変更キーのカンマ区切り）
        ManageNamedRange,       // [SheetDeactivate ＋ タスク切替時] 名前定義の追加・削除・参照先変更
        ImportExternalData,     // [SheetDeactivate ＋ タスク切替時] 外部データ取り込み（テキスト/CSV等）
        SetPageBreak,           // [SheetDeactivate ＋ タスク切替時] 改ページ位置の設定
        /// <summary>[SheetDeactivate ＋ タスク切替時] 印刷の向き（PageSetup.Orientation）。</summary>
        SetPageOrientation,
        /// <summary>[SheetDeactivate ＋ タスク切替時] 余白（PageSetup の各 Margin）。</summary>
        SetPageMargins,
        /// <summary>[SheetDeactivate ＋ タスク切替時] 用紙サイズ・拡大/縮小・ページに合わせる・Zoom 等（PageSetup のスケーリング系）。</summary>
        SetPageScaling,
        SetTableStyle,          // [タスク切替時フラッシュ] テーブルスタイル/テーブルデザイン変更
        ResizeTable,            // [SheetDeactivate ＋ タスク切替時] テーブル範囲の拡張・縮小
        AddConditionalFormat,   // [SheetDeactivate ＋ タスク切替時] 条件付き書式の追加・変更
        AddChart                // [即時発火] グラフの新規作成
    }

    [Flags]
    public enum ExcelValidationExemptFlags
    {
        None = 0,
        RangeEdit = 1,            // セル値/数式/書式編集
        SheetStructure = 2,       // 行列挿入削除・シート構造変更
        ShapeOrImage = 4,         // 図形/画像追加削除・移動リサイズ
        PrintAndPage = 8,         // 印刷範囲、タイトル、改ページ、ヘッダーフッター
        WorkbookProperty = 16,    // ドキュメントプロパティ、名前定義など
        ExternalImport = 32,      // 外部データ取り込み
        /// <summary>セル書式のみ（値・数式編集とは別に免除するタスク用）。</summary>
        CellFormatOnly = 64,
        All = 127
    }

    /// <summary>
    /// Excel タスクごとの破壊的操作ルール（免除・禁止・許可集合・セル範囲）。
    /// PowerPoint の PPTaskValidationConfig と同様の責務。
    /// </summary>
    /// <remarks>
    /// <para>コード内は <c>#region</c> で免除 / 許可 / 禁止 / 許可範囲に分割。</para>
    /// <para>実際の採点フローは <c>ReviewPageWindow.ApplyDestructiveValidation</c> と
    /// <see cref="ExcelLogReader"/>（<c>mos_excel_log.txt</c> の <c>[Op]</c> 解析）。</para>
    /// <para>PowerPoint にある図形数・文字数などのデルタ厳密判定は、Excel は操作ログベースのため未実装。</para>
    /// </remarks>
    public static class ExcelTaskValidationConfig
    {
        #region 免除（カテゴリ対応・タスク別フラグ）

        /// <summary>
        /// 操作タイプが免除フラグのどのカテゴリに属するか（複合しない単一フラグ）。
        /// </summary>
        public static ExcelValidationExemptFlags GetExemptCategoryForOperation(ExcelOperationType op)
        {
            switch (op)
            {
                case ExcelOperationType.EditCellFormat:
                    return ExcelValidationExemptFlags.CellFormatOnly;
                case ExcelOperationType.EditCellValue:
                case ExcelOperationType.EditCellFormula:
                case ExcelOperationType.InsertHyperlink:
                case ExcelOperationType.AddConditionalFormat:
                    return ExcelValidationExemptFlags.RangeEdit;
                case ExcelOperationType.InsertRows:
                case ExcelOperationType.DeleteRows:
                case ExcelOperationType.InsertColumns:
                case ExcelOperationType.DeleteColumns:
                case ExcelOperationType.SortOrFilter:
                case ExcelOperationType.SetTableStyle:
                case ExcelOperationType.ResizeTable:
                    return ExcelValidationExemptFlags.SheetStructure;
                case ExcelOperationType.InsertShapeOrImage:
                case ExcelOperationType.DeleteShapeOrImage:
                case ExcelOperationType.MoveOrResizeShape:
                case ExcelOperationType.AddChart:
                    return ExcelValidationExemptFlags.ShapeOrImage;
                case ExcelOperationType.SetPrintArea:
                case ExcelOperationType.SetPrintTitle:
                case ExcelOperationType.SetHeaderFooter:
                case ExcelOperationType.SetFreezePanes:
                case ExcelOperationType.SetPageBreak:
                case ExcelOperationType.SetPageOrientation:
                case ExcelOperationType.SetPageMargins:
                case ExcelOperationType.SetPageScaling:
                    return ExcelValidationExemptFlags.PrintAndPage;
                case ExcelOperationType.SetWorkbookProperty:
                case ExcelOperationType.ManageNamedRange:
                    return ExcelValidationExemptFlags.WorkbookProperty;
                case ExcelOperationType.ImportExternalData:
                    return ExcelValidationExemptFlags.ExternalImport;
                default:
                    return ExcelValidationExemptFlags.None;
            }
        }

        /// <summary>
        /// 免除フラグにより当該操作の allow/forbid チェックをスキップするか。
        /// </summary>
        public static bool IsOperationExempt(ExcelOperationType op, ExcelValidationExemptFlags flags)
        {
            if (flags == ExcelValidationExemptFlags.None) return false;
            ExcelValidationExemptFlags cat = GetExemptCategoryForOperation(op);
            if (cat == ExcelValidationExemptFlags.None) return false;
            return flags.HasFlag(cat);
        }

        /// <summary>
        /// 指定タスクでの破壊的操作チェック免除フラグを返す。
        /// プロジェクト内で <see cref="taskId"/> ごとに上書き可能（既定は従来のプロジェクト単位値）。
        /// </summary>
        public static ExcelValidationExemptFlags GetExemptFlags(int projectId, int taskId)
        {
            switch (projectId)
            {
                case 1:
                    switch (taskId)
                    {
                        case 1: // 売上一覧・印刷の向きを横向き
                        case 2: // 売上一覧・印刷範囲
                        case 3: // 売上一覧・印刷タイトル（行）
                        case 4: // 販売実績・余白
                        case 5: // 販売実績・改ページ
                            return ExcelValidationExemptFlags.PrintAndPage;
                        case 6: // スキルアップ検定結果・セル内折り返し（書式のみ → CellFormatOnly）。許可範囲は A4:K4（TryGetFirstNonExemptViolation の範囲ゲート）
                            return ExcelValidationExemptFlags.CellFormatOnly;
                        case 7: // 売上一覧・G4 メモ（範囲制限あり。RangeEdit 免除は付けない）
                            return ExcelValidationExemptFlags.None;
                        default:
                            return ExcelValidationExemptFlags.PrintAndPage;
                    }
                case 2:
                    switch (taskId)
                    {
                        case 5: // イベント売上: テーブルサイズ変更時に Excel 内部更新で EditCellValue が発生し得る
                            return ExcelValidationExemptFlags.SheetStructure | ExcelValidationExemptFlags.RangeEdit;
                        case 1: // 試験結果テーブル: 縞模様(行/列)切替
                        case 2: // 試験結果テーブル: 最後の列
                        case 3: // 試験結果テーブル: スタイル変更
                        case 4: // 担当者リスト: フィルター抽出
                        default:
                            // プロジェクト2はテーブル操作/フィルター操作が中心。
                            // 方式Aに切り替えたため、RangeEdit を免除すると不要なセル編集まで許容してしまう。
                            return ExcelValidationExemptFlags.SheetStructure;
                    }
                case 3:
                    switch (taskId)
                    {
                        case 1: // 下半期売上・A2スタイル
                        case 2: // 社員リスト・B5:B44インデント
                        case 3: // 社員リスト・A2:F2配置
                        case 5: // 業務予定・C5:C11取り消し線
                        case 6: // 参加者一覧・B2結合解除
                            return ExcelValidationExemptFlags.None;
                        case 4: // 担当者別売上・H5:K19コピーA5貼付
                            return ExcelValidationExemptFlags.RangeEdit;
                        case 7: // 参加者一覧・8-9行削除
                            return ExcelValidationExemptFlags.SheetStructure | ExcelValidationExemptFlags.RangeEdit;
                        default:
                            return ExcelValidationExemptFlags.RangeEdit | ExcelValidationExemptFlags.SheetStructure;
                    }
                case 4:
                    switch (taskId)
                    {
                        case 1: // 上半期売上・スパークライン
                        case 2: // ５年間売上・積み上げ縦棒
                        case 3: // 下半期売上・3-D円
                        case 4: // 商品別売上・代替テキスト
                            return ExcelValidationExemptFlags.ShapeOrImage;
                        default:
                            return ExcelValidationExemptFlags.ShapeOrImage;
                    }
                case 5:
                    switch (taskId)
                    {
                        case 1: // 売上実績・レイアウト
                        case 2: // 商品別売上・スタイル/配色
                        case 3: // 商品別売上・データ追加
                        case 4: // 商品別売上・凡例
                        case 5: // 月別売上・データラベル
                        case 6: // 月別売上・横軸ラベル
                            return ExcelValidationExemptFlags.ShapeOrImage;
                        default:
                            return ExcelValidationExemptFlags.ShapeOrImage;
                    }
                case 6:
                    switch (taskId)
                    {
                        case 1: // 売上一覧・ウィンドウ枠固定
                            return ExcelValidationExemptFlags.PrintAndPage;
                        case 2: // 売上一覧・ハイパーリンク
                            return ExcelValidationExemptFlags.RangeEdit;
                        case 3: // 販売実績・数値の書式（通貨）
                            return ExcelValidationExemptFlags.None;
                        case 4: // プロパティ・タグ
                            return ExcelValidationExemptFlags.WorkbookProperty;
                        default:
                            return ExcelValidationExemptFlags.WorkbookProperty | ExcelValidationExemptFlags.RangeEdit;
                    }
                case 7:
                    switch (taskId)
                    {
                        case 1: // イベント売上・オートフィル
                        case 2: // イベント売上・MAX
                        case 3: // 試験結果・COUNT
                        case 4: // 試験結果・COUNTBLANK
                        case 5: // 試験結果・RANDBETWEEN/オートフィル
                        case 6: // 試験結果・LEFT/オートフィル
                        case 7: // 申込一覧・UNIQUE
                            return ExcelValidationExemptFlags.RangeEdit;
                        default:
                            return ExcelValidationExemptFlags.RangeEdit;
                    }
                case 8:
                    switch (taskId)
                    {
                        case 1: // 学生名簿・名前定義
                            return ExcelValidationExemptFlags.WorkbookProperty;
                        case 2: // 名前移動・日付変更
                            return ExcelValidationExemptFlags.WorkbookProperty | ExcelValidationExemptFlags.RangeEdit;
                        case 3: // 売上報告・SUM
                        case 4: // 学生名簿・CONCAT/オートフィル
                        case 5: // 担当者リスト・CONCAT/オートフィル
                        case 6: // 申込一覧・CONCAT/オートフィル
                            return ExcelValidationExemptFlags.RangeEdit;
                        default:
                            return ExcelValidationExemptFlags.WorkbookProperty | ExcelValidationExemptFlags.RangeEdit;
                    }
                case 9:
                    switch (taskId)
                    {
                        case 1: // 売上報告・数式表示
                            return ExcelValidationExemptFlags.WorkbookProperty;
                        case 2: // 受注明細・並べ替え
                            return ExcelValidationExemptFlags.RangeEdit | ExcelValidationExemptFlags.SheetStructure;
                        case 3: // 下半期売上・アイコンセット
                        case 4: // 下半期売上・条件付き書式
                        case 7: // 下半期売上・書式変更（アクセシビリティ）
                            return ExcelValidationExemptFlags.None;
                        case 5: // 受注明細・ヘッダー
                        case 6: // 受注明細・フッター
                            return ExcelValidationExemptFlags.PrintAndPage;
                        default:
                            return ExcelValidationExemptFlags.PrintAndPage | ExcelValidationExemptFlags.RangeEdit;
                    }
                case 10:
                    switch (taskId)
                    {
                        case 1: // 担当者リスト・IF
                        case 2: // 出張精算・IF
                        case 3: // 売上一覧・IF
                        case 4: // 担当者リスト・SEQUENCE
                        case 5: // 業務予定・SEQUENCE
                        case 6: // 売上集計・SORT
                        case 7: // 売上一覧・絶対参照
                            return ExcelValidationExemptFlags.RangeEdit;
                        case 8: // 在庫管理・インポート
                            // インポート操作は内部的に値入力、行列挿入、名前定義、ページ設定等、多岐にわたる自動処理を引き起こすため、
                            // 図形操作(ShapeOrImage)以外を広範に免除します。
                            return ExcelValidationExemptFlags.ExternalImport | 
                                   ExcelValidationExemptFlags.RangeEdit | 
                                   ExcelValidationExemptFlags.CellFormatOnly |
                                   ExcelValidationExemptFlags.SheetStructure | 
                                   ExcelValidationExemptFlags.PrintAndPage | 
                                   ExcelValidationExemptFlags.WorkbookProperty;
                        default:
                            return ExcelValidationExemptFlags.RangeEdit | ExcelValidationExemptFlags.ExternalImport;
                    }
                default:
                    return ExcelValidationExemptFlags.None;
            }
        }

        #endregion

        #region 許可範囲（セル）

        /// <summary>
        /// 指定タスクで編集を許可するセル範囲（A1形式）を返す。
        /// 空配列の場合は「範囲制限なし」ではなく「未設定」として扱う運用を推奨。
        /// </summary>
        public static List<string> GetAllowedRanges(int projectId, int taskId)
        {
            switch (projectId)
            {
                case 3:
                    switch (taskId)
                    {
                        case 1:
                            return new List<string> { "下半期売上!A2" };
                        case 2:
                            return new List<string> { "社員リスト!B5:B44" };
                        case 3:
                            return new List<string> { "社員リスト!A2:F2" };
                        case 5:
                            return new List<string> { "業務予定!C5:C11" };
                        case 6:
                            return new List<string> { "参加者一覧!B2:G2" };
                        default:
                            return new List<string>();
                    }
                case 4:
                    return new List<string>();
                case 5:
                    return new List<string>();
                case 6:
                    switch (taskId)
                    {
                        case 2:
                            return new List<string> { "売上一覧!G3" };
                        case 3:
                            return new List<string> { "販売実績!B5:G11" };
                        default:
                            return new List<string>();
                    }
                case 7:
                    switch (taskId)
                    {
                        case 1:
                            return new List<string> { "イベント売上!H5:H16" };
                        case 2:
                            return new List<string> { "イベント売上!J3" };
                        case 3:
                            return new List<string> { "試験結果!I4" };
                        case 4:
                            return new List<string> { "試験結果!J4" };
                        case 5:
                            return new List<string> { "試験結果!B7:B56" };
                        case 6:
                            return new List<string> { "試験結果!G7:G56" };
                        case 7:
                            return new List<string> { "申込一覧!I5" };
                        default:
                            return new List<string>();
                    }
                case 8:
                    switch (taskId)
                    {
                        case 2:
                            return new List<string> { "申込一覧!G3" };
                        case 3:
                            return new List<string> { "売上報告!J5" };
                        case 4:
                            return new List<string> { "学生名簿!G5:G24" };
                        case 5:
                            return new List<string> { "担当者リスト!H5:H19" };
                        case 6:
                            return new List<string> { "申込一覧!G5:G174" };
                        default:
                            return new List<string>();
                    }
                case 9:
                    switch (taskId)
                    {
                        case 3:
                        case 4:
                            return new List<string> { "下半期売上!D5:I12" };
                        case 7:
                            return new List<string> { "下半期売上!H7" };
                        default:
                            return new List<string>();
                    }
                case 10:
                    switch (taskId)
                    {
                        case 1:
                            return new List<string> { "担当者リスト!G5:G26" };
                        case 2:
                            return new List<string> { "出張精算!G5:G9" };
                        case 3:
                            return new List<string> { "売上一覧!G4:G99" };
                        case 4:
                            return new List<string> { "担当者リスト!A5" };
                        case 5:
                            return new List<string> { "業務予定!C4" };
                        case 6:
                            return new List<string> { "売上集計!D6" };
                        case 7:
                            return new List<string> { "売上一覧!I4:I99" };
                        case 8:
                            return new List<string> { "在庫管理!B4" };
                        default:
                            return new List<string>();
                    }
                case 1:
                    switch (taskId)
                    {
                        case 6:
                            return new List<string> { "スキルアップ検定結果!A4:K4" };
                        case 7:
                            return new List<string> { "売上一覧!G4" };
                        default:
                            return new List<string>();
                    }
                default:
                    return new List<string>();
            }
        }

        /// <summary>
        /// 許可範囲外編集を不正とするか。
        /// RangeEdit 免除時は範囲外チェックを行わない。
        /// </summary>
        public static bool ShouldDenyOutsideAllowedRanges(int projectId, int taskId)
        {
            if (GetExemptFlags(projectId, taskId).HasFlag(ExcelValidationExemptFlags.RangeEdit))
                return false;
            return GetAllowedRanges(projectId, taskId).Count > 0;
        }

        #endregion
    }
}
