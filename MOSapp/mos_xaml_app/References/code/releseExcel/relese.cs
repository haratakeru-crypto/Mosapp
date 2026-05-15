/ Excelの保存・終了処理は非同期で実行（UIをブロックしない）
            Task.Run(() =>
            {
                if (excelProcess != null)
                {
                    try
                    {
                        // ExcelのCOMオブジェクトを取得
                        Microsoft.Office.Interop.Excel.Application excelApp = null;
                        try
                        {
                            excelApp = (Microsoft.Office.Interop.Excel.Application)Marshal.GetActiveObject("Excel.Application");
                            
                            // すべてのブックを保存
                            var workbooks = excelApp.Workbooks;
                            foreach (Microsoft.Office.Interop.Excel.Workbook workbook in workbooks)
                            {
                                try
                                {
                                    if (workbook.Path != "")
                                    {
                                        workbook.Save();
                                    }
                                    else
                                    {
                                        // 未保存のブックは保存ダイアログを表示せずに閉じる
                                        workbook.Saved = true;
                                    }
                                    Marshal.ReleaseComObject(workbook);
                                }
                                catch { }
                            }
                            Marshal.ReleaseComObject(workbooks);
                            
                            // Excelを終了
                            excelApp.Quit();
                            Marshal.ReleaseComObject(excelApp);
                            excelApp = null;
                            
                            // プロセスが終了するまで待機（最大5秒）
                            if (!excelProcess.HasExited)
                            {
                                excelProcess.WaitForExit(5000);
                            }
                        }
                        catch (COMException)
                        {
                            // COMオブジェクトが取得できない場合は、プロセスを強制終了
                        }
                        catch (Exception)
                        {
                            // エラーが発生した場合は、プロセスを強制終了
                        }
                        finally
                        {
                            // プロセスがまだ実行中の場合は強制終了
                            try
                            {
                                if (excelProcess != null && !excelProcess.HasExited)
                                {
                                    excelProcess.Kill();
                                    excelProcess.WaitForExit(2000);
                                }
                            }
                            catch { }
                            
                            // プロセスを解放
                            if (excelProcess != null)
                            {
                                try
                                {
                                    excelProcess.Dispose();
                                }
                                catch { }
                                excelProcess = null;
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // エラーが発生した場合もプロセスをクリーンアップ
                        try
                        {
                            if (excelProcess != null && !excelProcess.HasExited)
                            {
                                excelProcess.Kill();
                            }
                            excelProcess?.Dispose();
                        }
                        catch { }
                        excelProcess = null;
                    }
                }
            });
        }