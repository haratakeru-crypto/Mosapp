using System;
using System.Collections.Generic;

namespace Core.Ports.Primary
{
    public interface IExcelCheckerService
    {
        bool CheckExcel(int groupId, int projectId, string filePath);
        List<ProjectInfo> GetAllProjects();
    }

    public class ProjectInfo
    {
        public int GroupId { get; set; }
        public int ProjectId { get; set; }
        public string Name { get; set; }
        public string DllPath { get; set; }
    }
}