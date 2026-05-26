using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Core.Ports.Primary;
using Core.Ports.Secondary;

namespace Infrastructure
{
    public class ExcelCheckerRepository : IExcelCheckerRepository
    {
        private JObject _config;
        
        public ExcelCheckerRepository()
        {
            LoadConfig();
        }
        
        private void LoadConfig()
        {
            string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
            if (File.Exists(configPath))
            {
                string json = File.ReadAllText(configPath);
                _config = JObject.Parse(json);
            }
        }
        
        // インターフェースに合わせたメソッドシグネチャ
        public bool ExecuteCheck(int groupId, int projectId, string filePath)
        {
            try
            {
                // config.jsonから適切なライブラリを検索
                var libraryInfo = FindLibraryByFilePath(filePath);
                if (libraryInfo != null)
                {
                    // config.jsonで見つかった場合はそのライブラリを使用
                    return ExecuteCheckWithLibrary(libraryInfo.Value.LibraryName, filePath);
                }
                else
                {
                    // 従来の方法（groupId, projectIdから決定）
                    string libraryName = $"ExcelChecker{groupId}_{projectId}";
                    return ExecuteCheckWithLibrary(libraryName, filePath);
                }
            }
            catch
            {
                return false;
            }
        }
        
        private bool ExecuteCheckWithLibrary(string libraryName, string filePath)
        {
            try
            {
                string dllPath = GetDllPathByLibraryName(libraryName);
                if (!File.Exists(dllPath))
                    return false;

                Assembly assembly = Assembly.LoadFrom(dllPath);
                string className = GetClassNameFromLibrary(libraryName);
                Type type = assembly.GetType(className);
                
                if (type == null)
                    return false;

                object instance = Activator.CreateInstance(type);
                MethodInfo method = type.GetMethod("CheckExcel");
                
                if (method == null)
                    return false;

                return (bool)method.Invoke(instance, new object[] { filePath });
            }
            catch
            {
                return false;
            }
        }
        
        private (string LibraryName, int TaskCount)? FindLibraryByFilePath(string filePath)
        {
            if (_config == null) return null;
            
            foreach (var tab in _config["tabs"])
            {
                foreach (var project in tab.First["projects"])
                {
                    var projectData = project.First;
                    string configFilePath = projectData["excelFile"]?.ToString();
                    
                    if (string.Equals(configFilePath, filePath, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(Path.GetFileName(configFilePath), Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase))
                    {
                        return (projectData["library"]?.ToString(), projectData["taskCount"]?.ToObject<int>() ?? 0);
                    }
                }
            }
            
            return null;
        }
        
        private string GetDllPathByLibraryName(string libraryName)
        {
            // ExcelChecker1_1 -> Group1/ExcelChecker1_1.dll
            var parts = libraryName.Split('_');
            if (parts.Length >= 2)
            {
                string groupId = parts[0].Substring(parts[0].Length - 1); // "1" from "ExcelChecker1"
                return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, 
                    "Libraries", $"Group{groupId}", $"{libraryName}.dll");
            }
            return string.Empty;
        }
        
        private string GetClassNameFromLibrary(string libraryName)
        {
            // ExcelChecker1_1 -> Libraries.Group1.ExcelChecker1_1
            var parts = libraryName.Split('_');
            if (parts.Length >= 2)
            {
                string groupId = parts[0].Substring(parts[0].Length - 1);
                return $"Libraries.Group{groupId}.{libraryName}";
            }
            return string.Empty;
        }

        public List<ProjectInfo> LoadAllProjects()
        {
            var projects = new List<ProjectInfo>();
            
            if (_config != null)
            {
                foreach (var tab in _config["tabs"])
                {
                    var tabProperty = (JProperty)tab;
                    int tabId = int.Parse(tabProperty.Name);
                    foreach (var project in tabProperty.Value["projects"])
                    {
                        var projectProperty = (JProperty)project;
                        int projectId = int.Parse(projectProperty.Name);
                        var projectData = projectProperty.Value;
                        
                        projects.Add(new ProjectInfo
                        {
                            GroupId = tabId,
                            ProjectId = projectId,
                            Name = $"Project{tabId}_{projectId}",
                            DllPath = GetDllPathByLibraryName(projectData["library"]?.ToString() ?? "")
                        });
                    }
                }
            }
            else
            {
                // config.jsonが読み込めない場合は従来の方法
                for (int groupId = 1; groupId <= 3; groupId++)
                {
                    for (int projectId = 1; projectId <= 10; projectId++)
                    {
                        projects.Add(new ProjectInfo
                        {
                            GroupId = groupId,
                            ProjectId = projectId,
                            Name = $"Project{groupId}_{projectId}",
                            DllPath = GetDllPathByLibraryName($"ExcelChecker{groupId}_{projectId}")
                        });
                    }
                }
            }
            
            return projects;
        }
    }
}