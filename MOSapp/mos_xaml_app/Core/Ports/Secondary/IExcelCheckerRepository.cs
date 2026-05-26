using System;
using System.Collections.Generic;
using Core.Ports.Primary;

namespace Core.Ports.Secondary
{
    public interface IExcelCheckerRepository
    {
        bool ExecuteCheck(int groupId, int projectId, string filePath);
        List<ProjectInfo> LoadAllProjects();
    }
}