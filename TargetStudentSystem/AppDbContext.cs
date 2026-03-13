using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace TargetStudentSystem
{
    public class AppDbContext : DbContext
    {
        public DbSet<StudentModel> Students { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder options)
            => options.UseSqlite("Data Source=StudentsData.db"); // اسم ملف قاعدة البيانات
    }
}
