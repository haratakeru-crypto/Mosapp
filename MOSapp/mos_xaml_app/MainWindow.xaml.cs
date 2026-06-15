using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using Ui.ViewModels;
using Core.Adapters;
using Infrastructure;
using System.IO;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using Libraries;

namespace MOSExcelMogiApp
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        private MainViewModel _viewModel;
        private AppBarWindow _appBarWindow;
        public static bool IsTimerDisabled { get; private set; } = true; // タイマー無効化フラグ（静的プロパティ）。デフォルトは一時停止。
        
        public MainWindow()
        {
            // アプリ起動時にアドインの診断ログをクリアして肥大化を防ぐ
            ExcelLogReader.ClearDiagnosticLog();

            InitializeComponent();
            
            // Dependency Injection setup
            var repository = new ExcelCheckerRepository();
            var service = new ExcelCheckerService(repository);
            _viewModel = new MainViewModel(service);
            
            DataContext = _viewModel;
            
            // ViewModelのイベントを購読
            _viewModel.ShowAppBarRequested += OnShowAppBarRequested;
            _viewModel.HideMainWindowRequested += OnHideMainWindowRequested;
            _viewModel.ShowMainWindowRequested += OnShowMainWindowRequested;
            _viewModel.ExamEnded += OnExamEnded;
            _viewModel.UiTestRequested += OnUiTestRequested;

            Closing += MainWindow_Closing;
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            var result = MessageBox.Show("アプリ自体を終了します。本当にいいですか？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
            {
                e.Cancel = true;
                return;
            }

            try
            {
                // 終了処理中にカーソルを待機状態にする
                this.Cursor = Cursors.Wait;
                
                // Excel を確実に閉じる（同期実行して完了を待つことでゾンビプロセスを防止）
                _viewModel?.CloseExcelApplication();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Error during closing cleanup: {ex.Message}");
            }
        }
        
        private void OnShowAppBarRequested(object sender, EventArgs e)
        {
            if (_appBarWindow == null || !_appBarWindow.IsLoaded)
            {
                _appBarWindow = new AppBarWindow(_viewModel);
                // Excel の前面化を優先するため、表示時にフォーカスを奪わない。
                _appBarWindow.ShowActivated = false;
            }

            if (!_appBarWindow.IsVisible)
            {
                _appBarWindow.Show();
            }
        }
        
        private void OnHideMainWindowRequested(object sender, EventArgs e)
        {
            this.Hide();
        }
        
        private void OnShowMainWindowRequested(object sender, EventArgs e)
        {
            this.Show();
            this.Activate();
        }
        
        private void OnExamEnded(object sender, EventArgs e)
        {
            // アプリバーの参照をクリアして、次回新しいインスタンスを作成できるようにする
            _appBarWindow = null;
        }
        
        private void OnUiTestRequested(object sender, EventArgs e)
        {
            // UIテスト機能の実装（ダイアログなし）
        }
        
        private void TimerCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            IsTimerDisabled = false;
            System.Diagnostics.Debug.WriteLine($"MainWindow: TimerCheckBox checked, IsTimerDisabled = {IsTimerDisabled}");
        }
        
        private void TimerCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            IsTimerDisabled = true;
            System.Diagnostics.Debug.WriteLine($"MainWindow: TimerCheckBox unchecked, IsTimerDisabled = {IsTimerDisabled}");
        }
        
        private void ProjectResetButton_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[ResetProject] ProjectResetButton_Click called");
            try
            {
                // 現在選択されているタブを取得（0: 演習=Group1, 1: 応用編=Group3。模試①は非表示）
                int selectedTabIndex = _viewModel?.SelectedTabIndex ?? 0;
                int groupId = selectedTabIndex == 0 ? 1 : 3; // タブ0=演習(Group1), タブ1=応用編(Group3)
                
                string tabName = selectedTabIndex switch
                {
                    0 => "演習",
                    1 => "応用編",
                    _ => "選択中のタブ"
                };
                
                // 選択されているタブのプロジェクトをリセットするか確認
                var result = MessageBox.Show(
                    $"{tabName}のすべてのプロジェクトをリセットしますか？\n（すべての編集内容は失われます）", 
                    "確認", 
                    MessageBoxButton.YesNo, 
                    MessageBoxImage.Warning);
                
                System.Diagnostics.Debug.WriteLine($"[ResetProject] User response: {result}, SelectedTabIndex: {selectedTabIndex}, GroupId: {groupId}");
                
                if (result == MessageBoxResult.Yes)
                {
                    // 待機メッセージを表示するウィンドウを作成
                    var waitWindow = new Window
                    {
                        Title = "リセット中",
                        Width = 300,
                        Height = 150,
                        WindowStyle = WindowStyle.ToolWindow,
                        WindowStartupLocation = WindowStartupLocation.CenterOwner,
                        Owner = this,
                        ShowInTaskbar = false,
                        ResizeMode = ResizeMode.NoResize
                    };
                    
                    var stackPanel = new StackPanel
                    {
                        Margin = new Thickness(20),
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    
                    var textBlock = new TextBlock
                    {
                        Text = "リセット中です...",
                        FontSize = 16,
                        HorizontalAlignment = HorizontalAlignment.Center,
                        Margin = new Thickness(0, 0, 0, 10)
                    };
                    
                    var progressBar = new ProgressBar
                    {
                        IsIndeterminate = true,
                        Height = 20,
                        Width = 250
                    };
                    
                    stackPanel.Children.Add(textBlock);
                    stackPanel.Children.Add(progressBar);
                    waitWindow.Content = stackPanel;
                    
                    // 非同期でリセット処理を実行
                    System.Threading.Tasks.Task.Run(() =>
                    {
                        try
                        {
                            System.Diagnostics.Debug.WriteLine($"[ResetProject] Starting reset process for Group{groupId} projects...");
                            
                            // 選択されているタブのプロジェクトグループを取得
                            if (_viewModel?.ProjectGroups != null && selectedTabIndex < _viewModel.ProjectGroups.Count)
                            {
                                var selectedGroup = _viewModel.ProjectGroups[selectedTabIndex];
                                int successCount = 0;
                                int failCount = 0;
                                var failedProjects = new List<string>();
                                
                                // 選択されているタブのプロジェクトのみをリセット
                                foreach (var project in selectedGroup.Projects)
                                {
                                    try
                                    {
                                        // UIスレッドで待機ウィンドウのテキストを更新
                                        Dispatcher.Invoke(() =>
                                        {
                                            textBlock.Text = $"リセット中です...\nプロジェクト{project.ProjectId}をリセット中";
                                        });
                                        
                                        System.Diagnostics.Debug.WriteLine($"[ResetProject] Resetting Group{selectedGroup.GroupId}, Project{project.ProjectId}");
                                        ResetProject(selectedGroup.GroupId, project.ProjectId, showMessage: false);
                                        successCount++;
                                    }
                                    catch (Exception ex)
                                    {
                                        failCount++;
                                        string projectName = $"プロジェクト{project.ProjectId}";
                                        failedProjects.Add(projectName);
                                        System.Diagnostics.Debug.WriteLine($"[ResetProject] Error resetting {projectName}: {ex.Message}");
                                    }
                                }
                                
                                // UIスレッドで待機ウィンドウを閉じて結果を表示
                                Dispatcher.Invoke(() =>
                                {
                                    waitWindow.Close();
                                    
                                    // すべてのリセットが終わったら1回だけメッセージを表示
                                    if (failCount == 0)
                                    {
                                        MessageBox.Show("すべてリセットしました", 
                                            "完了", 
                                            MessageBoxButton.OK, 
                                            MessageBoxImage.Information);
                                    }
                                    else
                                    {
                                        string message = $"リセットが完了しました。\n\n";
                                        message += $"成功: {successCount} プロジェクト\n";
                                        message += $"失敗: {failCount} プロジェクト\n";
                                        if (failedProjects.Count > 0)
                                        {
                                            message += $"失敗したプロジェクト: {string.Join(", ", failedProjects)}";
                                        }
                                        
                                        MessageBox.Show(message, 
                                            "完了", 
                                            MessageBoxButton.OK, 
                                            MessageBoxImage.Warning);
                                    }
                                });
                            }
                            else
                            {
                                Dispatcher.Invoke(() =>
                                {
                                    waitWindow.Close();
                                    MessageBox.Show("プロジェクト情報を取得できませんでした。", 
                                        "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            Dispatcher.Invoke(() =>
                            {
                                waitWindow.Close();
                                System.Diagnostics.Debug.WriteLine($"[ResetProject] Error in reset process: {ex.Message}\n{ex.StackTrace}");
                                MessageBox.Show($"プロジェクトリセット中にエラーが発生しました: {ex.Message}", 
                                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                            });
                        }
                    });
                    
                    // 待機ウィンドウを表示
                    waitWindow.ShowDialog();
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[ResetProject] Reset cancelled by user");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ResetProject] Error in ProjectResetButton_Click: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"プロジェクトリセット中にエラーが発生しました: {ex.Message}", 
                    "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        public void ResetProject(int groupId, int projectId, bool showMessage = true)
        {
            try
            {
                ExcelLogReader.ClearOperationLogForProject(projectId);
                ExcelLogReader.ClearDestructiveLogForProject(projectId);

                // 実際に開いているファイルパスを取得
                string projectFilePath = null;
                
                // 1. CurrentProjectから実際のファイルパスを取得（最優先）
                if (_viewModel?.CurrentProject != null)
                {
                    projectFilePath = _viewModel.CurrentProject.FilePath;
                    System.Diagnostics.Debug.WriteLine($"[ResetProject] Current project file path: {projectFilePath}");
                }
                
                // 2. CurrentProjectのFilePathがない場合、config.jsonから取得
                if (string.IsNullOrEmpty(projectFilePath))
                {
                    string configPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                    if (File.Exists(configPath))
                    {
                        string jsonContent = File.ReadAllText(configPath);
                        JObject config = JObject.Parse(jsonContent);
                        
                        var projectConfig = config["tabs"]?[groupId.ToString()]?["projects"]?[projectId.ToString()];
                        if (projectConfig != null)
                        {
                            // initialDataFileを優先
                            projectFilePath = projectConfig["initialDataFile"]?.ToString();
                            
                            // initialDataFileがない場合、excelFileを使用
                            if (string.IsNullOrEmpty(projectFilePath))
                            {
                                projectFilePath = projectConfig["excelFile"]?.ToString();
                            }
                            
                            // それでもない場合、Initialフォルダのパスを生成
                            if (string.IsNullOrEmpty(projectFilePath))
                            {
                                projectFilePath = $"C:\\MOSTest\\Excel365\\Tab{groupId}\\Initial\\project{projectId}.xlsx";
                            }
                            
                            System.Diagnostics.Debug.WriteLine($"[ResetProject] File path from config.json: {projectFilePath}");
                        }
                    }
                }
                
                if (string.IsNullOrEmpty(projectFilePath))
                {
                    string errorMsg = $"プロジェクトファイルのパスを取得できませんでした。\nプロジェクト: Group{groupId}, Project{projectId}";
                    if (showMessage)
                    {
                        MessageBox.Show(errorMsg, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    throw new Exception(errorMsg);
                }
                
                System.Diagnostics.Debug.WriteLine($"[ResetProject] Resetting project file: {projectFilePath}");
                
                // Templatesフォルダからテンプレートファイルのパスを生成・検索
                string templatesFolder = $"C:\\MOSTest\\Excel365\\Templates\\Tab{groupId}";
                string templatePath = null;
                
                // まず、Templates\Tab{groupId}フォルダ内のファイルを動的に検索
                if (Directory.Exists(templatesFolder))
                {
                    System.Diagnostics.Debug.WriteLine($"Searching templates folder: {templatesFolder}");
                    var excelFiles = Directory.GetFiles(templatesFolder, "*.xlsx", SearchOption.TopDirectoryOnly);
                    System.Diagnostics.Debug.WriteLine($"Found {excelFiles.Length} Excel files in templates folder");
                    
                    // project{projectId}を含むファイル名を探す（大文字小文字を区別しない）
                    string searchPattern = $"project{projectId}".ToLower();
                    foreach (var file in excelFiles)
                    {
                        string fileName = System.IO.Path.GetFileNameWithoutExtension(file);
                        System.Diagnostics.Debug.WriteLine($"Checking file: {fileName}");
                        
                        // 大文字小文字を区別しない比較
                        string fileNameLower = fileName.ToLower();
                        if (fileNameLower.Contains(searchPattern) || fileNameLower == searchPattern)
                        {
                            templatePath = file;
                            System.Diagnostics.Debug.WriteLine($"Template file found in folder: {templatePath}");
                            break;
                        }
                    }
                }
                
                // 見つからない場合、固定パターンで検索
                if (string.IsNullOrEmpty(templatePath))
                {
                    // パターン1: Templates\Tab{groupId}\project{projectId}.xlsx
                    string templatePath1 = $"C:\\MOSTest\\Excel365\\Templates\\Tab{groupId}\\project{projectId}.xlsx";
                    // パターン2: Templates\project{projectId}.xlsx
                    string templatePath2 = $"C:\\MOSTest\\Excel365\\Templates\\project{projectId}.xlsx";
                    // パターン3: Templates\Tab{groupId}\Tab{groupId}_project{projectId}.xlsx
                    string templatePath3 = $"C:\\MOSTest\\Excel365\\Templates\\Tab{groupId}\\Tab{groupId}_project{projectId}.xlsx";
                    
                    if (File.Exists(templatePath1))
                    {
                        templatePath = templatePath1;
                        System.Diagnostics.Debug.WriteLine($"Template file found (pattern 1): {templatePath}");
                    }
                    else if (File.Exists(templatePath2))
                    {
                        templatePath = templatePath2;
                        System.Diagnostics.Debug.WriteLine($"Template file found (pattern 2): {templatePath}");
                    }
                    else if (File.Exists(templatePath3))
                    {
                        templatePath = templatePath3;
                        System.Diagnostics.Debug.WriteLine($"Template file found (pattern 3): {templatePath}");
                    }
                }
                
                if (string.IsNullOrEmpty(templatePath))
                {
                    string errorMsg = $"テンプレートファイルが見つかりません。\n\nプロジェクト: Group{groupId}, Project{projectId}\n検索フォルダ: {templatesFolder}";
                    
                    if (Directory.Exists(templatesFolder))
                    {
                        var files = Directory.GetFiles(templatesFolder, "*.xlsx");
                        errorMsg += $"\n\nフォルダ内のファイル:\n";
                        foreach (var file in files)
                        {
                            errorMsg += $"  - {System.IO.Path.GetFileName(file)}\n";
                        }
                    }
                    
                    if (showMessage)
                    {
                        MessageBox.Show(errorMsg, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    throw new Exception(errorMsg);
                }
                
                System.Diagnostics.Debug.WriteLine($"Template file found: {templatePath}");
                
                // テンプレートファイル自体の Zone.Identifier を削除（存在する場合）
                RemoveZoneIdentifier(templatePath);

                // テンプレートファイルを読み取り専用で保護（テンプレートを変更されないようにする）
                FileInfo templateFileInfo = new FileInfo(templatePath);
                if (!templateFileInfo.IsReadOnly)
                {
                    templateFileInfo.IsReadOnly = true;
                    System.Diagnostics.Debug.WriteLine($"Template file set to read-only for protection: {templatePath}");
                }
                
                // Excel を全ブック閉じたうえで終了し、ロック・二重オープンを防ぐ（共有 COM 参照もクリア）
                System.Diagnostics.Debug.WriteLine($"[ResetProject] Quitting Excel before reset");
                _viewModel.QuitExcelForProjectReset();
                
                // ファイルがロックされているか確認してからコピー
                int retryCount = 0;
                while (retryCount < 10 && IsFileLocked(projectFilePath))
                {
                    System.Threading.Thread.Sleep(500);
                    retryCount++;
                    System.Diagnostics.Debug.WriteLine($"[ResetProject] File still locked, retry {retryCount}/10");
                }
                
                // テンプレートファイルをプロジェクトファイルにコピー
                try
                {
                    // プロジェクトファイルが存在する場合、読み取り専用属性を解除
                    if (File.Exists(projectFilePath))
                    {
                        FileInfo projectFileInfo = new FileInfo(projectFilePath);
                        if (projectFileInfo.IsReadOnly)
                        {
                            projectFileInfo.IsReadOnly = false;
                        }
                    }
                    
                    // テンプレートファイルをプロジェクトファイルにコピー
                    System.Diagnostics.Debug.WriteLine($"[ResetProject] Copying template file...");
                    System.Diagnostics.Debug.WriteLine($"[ResetProject]   From: {templatePath}");
                    System.Diagnostics.Debug.WriteLine($"[ResetProject]   To: {projectFilePath}");
                    
                    // コピー前にディレクトリが存在することを確認
                    string projectDirectory = System.IO.Path.GetDirectoryName(projectFilePath);
                    if (!Directory.Exists(projectDirectory))
                    {
                        Directory.CreateDirectory(projectDirectory);
                        System.Diagnostics.Debug.WriteLine($"[ResetProject] Created directory: {projectDirectory}");
                    }
                    
                    // テンプレートファイルをプロジェクトファイルにコピー
                    // 読み取り専用ファイルからでもコピー可能
                    File.Copy(templatePath, projectFilePath, overwrite: true);
                    // File.Copy は Zone.Identifier ADS も引き継ぐため、コピー直後に削除して保護ビューを防ぐ
                    RemoveZoneIdentifier(projectFilePath);

                    System.Diagnostics.Debug.WriteLine($"[ResetProject] Project file reset successfully: {projectFilePath}");
                    
                    // ファイルサイズを確認
                    FileInfo newProjectFile = new FileInfo(projectFilePath);
                    FileInfo templateFile = new FileInfo(templatePath);
                    System.Diagnostics.Debug.WriteLine($"[ResetProject] Template file size: {templateFile.Length} bytes");
                    System.Diagnostics.Debug.WriteLine($"[ResetProject] Project file size: {newProjectFile.Length} bytes");
                    
                    // コピー後、すぐに読み取り専用属性を解除
                    if (newProjectFile.IsReadOnly)
                    {
                        newProjectFile.IsReadOnly = false;
                        System.Diagnostics.Debug.WriteLine($"[ResetProject] Removed read-only attribute from project file after copy");
                    }
                    
                    // Initialフォルダのパスを生成
                    string initialFolderPath = $"C:\\MOSTest\\Excel365\\Tab{groupId}\\Initial";
                    string initialFilePath = System.IO.Path.Combine(initialFolderPath, $"project{projectId}.xlsx");
                    
                    System.Diagnostics.Debug.WriteLine($"[ResetProject] Initial folder path: {initialFolderPath}");
                    System.Diagnostics.Debug.WriteLine($"[ResetProject] Initial file path: {initialFilePath}");
                    
                    // ディレクトリが存在しない場合は作成
                    if (!Directory.Exists(initialFolderPath))
                    {
                        Directory.CreateDirectory(initialFolderPath);
                        System.Diagnostics.Debug.WriteLine($"[ResetProject] Created directory: {initialFolderPath}");
                    }
                    
                    // Initialフォルダのファイルが存在する場合、読み取り専用属性を解除
                    if (File.Exists(initialFilePath))
                    {
                        FileInfo initialFileInfo = new FileInfo(initialFilePath);
                        if (initialFileInfo.IsReadOnly)
                        {
                            initialFileInfo.IsReadOnly = false;
                            System.Diagnostics.Debug.WriteLine($"[ResetProject] Removed read-only attribute from Initial folder file");
                        }
                    }
                    
                    // リセットしたプロジェクトファイルをInitialフォルダにもコピー
                    try
                    {
                        File.Copy(projectFilePath, initialFilePath, overwrite: true);
                        // ADS の引き継ぎを防ぐため、Initialフォルダのファイルも Zone.Identifier を削除
                        RemoveZoneIdentifier(initialFilePath);

                        System.Diagnostics.Debug.WriteLine($"[ResetProject] Copied project file to Initial folder: {initialFilePath}");
                        
                        // Initialフォルダのファイルの読み取り専用属性を解除
                        FileInfo initialFileInfo = new FileInfo(initialFilePath);
                        if (initialFileInfo.IsReadOnly)
                        {
                            initialFileInfo.IsReadOnly = false;
                            System.Diagnostics.Debug.WriteLine($"[ResetProject] Removed read-only attribute from Initial folder file after copy");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ResetProject] Error copying to Initial folder: {ex.Message}");
                        // Initialフォルダへのコピーに失敗しても、プロジェクトファイルのリセットは成功しているので続行
                    }
                    
                    // リセット後、現在のプロジェクトなら Excel でブックを開き直す（シェル起動）
                    if (_viewModel?.CurrentProject != null)
                    {
                        var currentProject = _viewModel.CurrentProject;
                        int currentGroupId = int.Parse(currentProject.Group.Replace("Group ", ""));
                        int currentProjectId = currentProject.ProjectNumber;
                        
                        if (currentGroupId == groupId && currentProjectId == projectId)
                        {
                            System.Diagnostics.Debug.WriteLine($"[ResetProject] Processing Excel file after reset: {projectFilePath}");
                            
                            // Initialフォルダのパスは既に外側で定義されているので、そのまま使用
                            System.Diagnostics.Debug.WriteLine($"[ResetProject] Initial folder path: {initialFolderPath}");
                            System.Diagnostics.Debug.WriteLine($"[ResetProject] Initial file path: {initialFilePath}");
                            
                            // ディレクトリが存在しない場合は作成（既に外側で作成済みの可能性があるが、念のため確認）
                            if (!Directory.Exists(initialFolderPath))
                            {
                                Directory.CreateDirectory(initialFolderPath);
                                System.Diagnostics.Debug.WriteLine($"[ResetProject] Created directory: {initialFolderPath}");
                            }
                            
                            // テンプレートコピー済み。COM で空 Excel を先に立てず、シェルで開き直す（アプリバーからのリセットでスタート画面が残るのを防ぐ）
                            FileInfo projectFileInfo = new FileInfo(projectFilePath);
                            if (projectFileInfo.IsReadOnly)
                            {
                                projectFileInfo.IsReadOnly = false;
                                System.Diagnostics.Debug.WriteLine($"[ResetProject] Removed read-only attribute from project file before opening Excel");
                            }

                            string pathToOpen = File.Exists(initialFilePath) ? initialFilePath : projectFilePath;
                            _viewModel.OpenExcelWorkbookAfterResetByShell(pathToOpen);
                            System.Diagnostics.Debug.WriteLine($"[ResetProject] Reopened workbook via shell: {pathToOpen}");
                        }
                    }
                    
                    if (showMessage)
                    {
                        MessageBox.Show($"プロジェクト (Group{groupId}, Project{projectId}) をリセットしました。", 
                            "完了", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    string errorMsg = $"ファイルへのアクセスが拒否されました。\nファイルが他のプログラムで開かれている可能性があります。\n\nプロジェクト: Group{groupId}, Project{projectId}\n{ex.Message}";
                    if (showMessage)
                    {
                        MessageBox.Show(errorMsg, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    System.Diagnostics.Debug.WriteLine($"UnauthorizedAccessException in ResetProject: {ex.Message}");
                    throw new Exception(errorMsg, ex);
                }
                catch (IOException ex)
                {
                    string errorMsg = $"ファイルのコピー中にエラーが発生しました。\nファイルがロックされている可能性があります。\n\nプロジェクト: Group{groupId}, Project{projectId}\n{ex.Message}";
                    if (showMessage)
                    {
                        MessageBox.Show(errorMsg, "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    System.Diagnostics.Debug.WriteLine($"IOException in ResetProject: {ex.Message}");
                    throw new Exception(errorMsg, ex);
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"プロジェクトリセット中にエラーが発生しました。\nプロジェクト: Group{groupId}, Project{projectId}\n{ex.Message}";
                if (showMessage)
                {
                    MessageBox.Show($"{errorMsg}\n\n{ex.StackTrace}", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                System.Diagnostics.Debug.WriteLine($"Error in ResetProject: {ex.Message}\n{ex.StackTrace}");
                throw;
            }
        }
        
        // Zone.Identifier ADS を削除して保護ビューを防ぐ
        [System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
        private static extern bool DeleteFileW(string lpFileName);

        /// <summary>
        /// ファイルの Zone.Identifier ADS（NTFS 代替データストリーム）を削除する。
        /// File.Copy は ADS も引き継ぐため、コピー直後にこのメソッドを呼ぶことで保護ビューを防ぐ。
        /// 読み取り専用ファイルの場合は一時的に属性を外してから削除を試みる。
        /// </summary>
        private static void RemoveZoneIdentifier(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath)) return;
            try
            {
                // 読み取り専用だと ADS 削除が失敗する場合があるため、一時的に属性を外す
                FileInfo fi = new FileInfo(filePath);
                bool wasReadOnly = fi.IsReadOnly;
                if (wasReadOnly)
                {
                    try { fi.IsReadOnly = false; } catch { }
                }

                DeleteFileW(filePath + ":Zone.Identifier");

                // 読み取り専用だったファイルは元に戻す
                if (wasReadOnly)
                {
                    try { new FileInfo(filePath).IsReadOnly = true; } catch { }
                }
            }
            catch { }
        }

        // ファイルがロックされているか確認するメソッド
        private bool IsFileLocked(string filePath)
        {
            if (!File.Exists(filePath))
            {
                return false;
            }
            
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false;
                }
            }
            catch (IOException)
            {
                return true;
            }
            catch
            {
                return false;
            }
        }
        
        // 特定のプロジェクトのExcelファイルを閉じるメソッド
        private void CloseProjectExcelFile(string filePath)
        {
            try
            {
                ExcelApp excelApp = null;
                try
                {
                    excelApp = (ExcelApp)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    // Excelが開いていない場合は無視
                    System.Diagnostics.Debug.WriteLine("[MainWindow] Excel application not found");
                    return;
                }
                
                if (excelApp != null && excelApp.Workbooks != null)
                {
                    string fileName = System.IO.Path.GetFileName(filePath);
                    foreach (ExcelWorkbook wb in excelApp.Workbooks)
                    {
                        try
                        {
                            if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                                wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                            {
                                System.Diagnostics.Debug.WriteLine($"[MainWindow] Closing workbook: {wb.Name}");
                                wb.Close(SaveChanges: false);
                                Marshal.ReleaseComObject(wb);
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[MainWindow] Error accessing workbook: {ex.Message}");
                        }
                    }
                    Marshal.ReleaseComObject(excelApp.Workbooks);
                }
                
                if (excelApp != null)
                {
                    Marshal.ReleaseComObject(excelApp);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[MainWindow] Error closing Excel file: {ex.Message}");
            }
        }
        
        protected override void OnClosed(EventArgs e)
        {
            // イベント購読を解除
            if (_viewModel != null)
            {
                _viewModel.ShowAppBarRequested -= OnShowAppBarRequested;
                _viewModel.HideMainWindowRequested -= OnHideMainWindowRequested;
                _viewModel.ShowMainWindowRequested -= OnShowMainWindowRequested;
                _viewModel.ExamEnded -= OnExamEnded;
                _viewModel.UiTestRequested -= OnUiTestRequested;
            }
            
            // アプリバーウィンドウを閉じる
            _appBarWindow?.Close();
            
            base.OnClosed(e);
        }
    }
}
