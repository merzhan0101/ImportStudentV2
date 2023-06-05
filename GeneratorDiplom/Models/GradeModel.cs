namespace GeneratorDiplom.Models
{
    public class GradeModel
    {
        public int Id { get; set; }
        public StudentModel Student { get; set; }
        public int StudentId { get; set; }
        public SubjectModel Subject { get; set; }
        public int SubjectId { get; set; }

        public string Score { get; set; }
    }
}
