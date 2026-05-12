using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;

namespace MOS_Word_app
{
    public class TabChecker
    {
        /// <summary>
        /// 現在選択されているタブが指定されたタブと一致するかチェック
        /// 注意: Word Interopではリボンタブの状態を直接取得できないため、
        /// 簡易的な実装として、ユーザーがタブをクリックしたことを前提とします
        /// </summary>
        public bool CheckTab(string expectedTabName)
        {
            Microsoft.Office.Interop.Word.Application wordApp = null;
            try
            {
                wordApp = (Microsoft.Office.Interop.Word.Application)Marshal.GetActiveObject("Word.Application");
                
                // タブ名を正規化
                string normalizedExpected = NormalizeTabName(expectedTabName);
                
                // Word 2013以降では、リボンUIの状態を直接取得するのは困難
                // 代替方法: UI Automationを使用するか、Wordのイベントを監視
                // ここでは簡易的な実装として、CommandBarsを使用
                
                CommandBars commandBars = wordApp.CommandBars;
                
                // 各タブの特徴的なコマンドIDを確認
                // 注意: この方法は完全ではないため、実際の使用ではUI Automationを推奨
                bool result = CheckTabByCommandBars(wordApp, commandBars, normalizedExpected);
                
                Marshal.ReleaseComObject(commandBars);
                return result;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (wordApp != null)
                    Marshal.ReleaseComObject(wordApp);
            }
        }

        private bool CheckTabByCommandBars(Microsoft.Office.Interop.Word.Application wordApp, CommandBars commandBars, string expectedTab)
        {
            try
            {
                // 各タブの特徴的なコマンドを確認
                // 注意: この実装は簡易版です。完全な実装にはUI Automationが必要です
                
                // ホームタブの確認（フォント関連のコマンドが有効）
                if (expectedTab == "ホーム")
                {
                    try
                    {
                        CommandBarControl control = commandBars.FindControl(Type: MsoControlType.msoControlButton, Id: 16); // FontSize
                        if (control != null)
                        {
                            Marshal.ReleaseComObject(control);
                            return true; // 簡易的な実装
                        }
                    }
                    catch { }
                }
                
                // 挿入タブの確認
                if (expectedTab == "挿入")
                {
                    // 挿入タブ特有のコマンドを確認
                    return true; // 簡易的な実装
                }
                
                // その他のタブも同様に実装
                // 実際の実装では、UI Automationを使用することを推奨
                
                return false;
            }
            catch
            {
                return false;
            }
        }

        private string NormalizeTabName(string tabName)
        {
            // タブ名を正規化（「タブ」を削除、空白を削除）
            string normalized = tabName.Replace("タブ", "").Replace(" ", "").Trim();
            
            // マッピング
            if (normalized.Contains("挿入")) return "挿入";
            if (normalized.Contains("デザイン")) return "デザイン";
            if (normalized.Contains("レイアウト")) return "レイアウト";
            if (normalized.Contains("参考資料")) return "参考資料";
            if (normalized.Contains("校閲")) return "校閲";
            if (normalized.Contains("ホーム")) return "ホーム";
            
            return normalized;
        }

    }
}

