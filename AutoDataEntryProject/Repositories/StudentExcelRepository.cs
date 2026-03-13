using AutoDataEntryProject.Models;
using AutoDataEntryProject.Utilities;
using System;
using System.Collections.Generic;

namespace AutoDataEntryProject.Repositories
{
    
    public class StudentExcelRepository : IStudentRepository
    {
        private readonly ExcelReader _excelReader;

        public StudentExcelRepository()
        {
            _excelReader = new ExcelReader();
        }

       
        public List<Student> GetAllStudents(string sourcePath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
                throw new ArgumentException("مسار الملف غير صحيح", nameof(sourcePath));

            return _excelReader.ReadStudentFile(sourcePath);
        }

        
        public bool ValidateDataSource(string sourcePath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
                return false;

            return _excelReader.ValidateExcelStructure(sourcePath);
        }

       
        public int GetStudentCount(string sourcePath)
        {
            try
            {
                var students = GetAllStudents(sourcePath);
                return students.Count;
            }
            catch
            {
                return 0;
            }
        }
    }
}
