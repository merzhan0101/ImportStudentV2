using GeneratorDiplom.Models;
using Microsoft.EntityFrameworkCore;

namespace GeneratorDiplom
{
    public class AppContext : DbContext
    {
        public DbSet<LocalizerModel> Localizers { get; set; }
        public DbSet<StudentModel> Students { get; set; }
        public DbSet<GroupModel> Groups { get; set; }
        public DbSet<GradeModel> Grades { get; set; }
        public DbSet<SubjectModel> Subjects { get; set; }

        private bool IsLocal { get; set; }

        public AppContext()
        {
            IsLocal = true;
            Database.EnsureCreated();
        }
        public AppContext(DbContextOptions options) : base(options) =>
            Database.EnsureCreated();

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (IsLocal)
            {
                optionsBuilder.UseMySql(
                    "server=localhost;user=root;password=;database=college_diplom;"
                );
            }
        }
    }
}
