using System;
using System.Collections.Generic;

namespace GeneratorDiplom.Models
{
    public class StudentModel
    {
        public int Id { get; set; }

        public LocalizerModel Initials { get; set; }
        public LocalizerModel Initials_Dat { get; set; }

        public GroupModel Group { get; set; }

        public LocalizerModel Topic { get; set; }
        public int? TopicId { get; set; }

        public int GroupId { get; set; }

        public string NumApplication { get; set; }
        public int? DateApplication { get; set; }

        public List<GradeModel> Grades { get; set; }

        public StudentModel() =>
            Grades = new List<GradeModel>();


        public string nameBefore { get; set; }
    }
}
