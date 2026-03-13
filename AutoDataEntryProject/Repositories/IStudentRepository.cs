using AutoDataEntryProject.Models;
using System.Collections.Generic;

namespace AutoDataEntryProject.Repositories
{
  
    public interface IStudentRepository
    {
       
        List<Student> GetAllStudents(string sourcePath);

        
        bool ValidateDataSource(string sourcePath);

       
        int GetStudentCount(string sourcePath);
    }
}
