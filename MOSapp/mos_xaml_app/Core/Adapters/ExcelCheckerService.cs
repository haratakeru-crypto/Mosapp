using System;
using System.Collections.Generic;
using Core.Ports.Primary;
using Core.Ports.Secondary;

namespace Core.Adapters
{
    public class ExcelCheckerService : IExcelCheckerService
    {
        private readonly IExcelCheckerRepository _repository;

        public ExcelCheckerService(IExcelCheckerRepository repository)
        {
            _repository = repository;
        }

        public bool CheckExcel(int groupId, int projectId, string filePath)
        {
            return _repository.ExecuteCheck(groupId, projectId, filePath);
        }

        public List<ProjectInfo> GetAllProjects()
        {
            return _repository.LoadAllProjects();
        }
    }
}