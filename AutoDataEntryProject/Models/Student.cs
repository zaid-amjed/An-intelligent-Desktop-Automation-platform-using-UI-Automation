using System;

namespace AutoDataEntryProject.Models
{
   
    public class Student
    {
        public string Name { get; set; }
        public string StudentId { get; set; }
        public string Email { get; set; }

        public Student()
        {
            Name = string.Empty;
            StudentId = string.Empty;
            Email = string.Empty;
        }

        public Student(string name, string id, string email)
        {
            Name = name ?? string.Empty;
            StudentId = id ?? string.Empty;
            Email = email ?? string.Empty;
        }

        
        public bool IsValid()
        {
            return !string.IsNullOrWhiteSpace(Name) &&
                   !string.IsNullOrWhiteSpace(StudentId) &&
                   !string.IsNullOrWhiteSpace(Email);
        }

        public override string ToString()
        {
            return $"{Name} - {StudentId} - {Email}";
        }
    }
}
