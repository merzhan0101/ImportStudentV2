using System.Collections.Generic;
using System.ComponentModel;

namespace GeneratorDiplom.Models
{
    public enum LangModel
    {
        RU,
        KZ
    }
    public class GroupModel
    {
        public int Id { get; set; }

        public LocalizerModel Title { get; set; }
        public LocalizerModel Qualification { get; set; }
        public LocalizerModel svidqual1 { get; set; }
        public LocalizerModel svidrazrad1 { get; set; }

        public string Code { get; set; }

        public int StartStudies { get; set; }

        public LangModel Language { get; set; }

        public int EndStudies { get; set; }

        public List<SubjectModel> Subjects { get; set; }
    }
}
