using System.ComponentModel;

namespace GeneratorDiplom.Models
{
    public class SubjectModel
    {
        public int Id { get; set; }

        public LocalizerModel Title { get; set; }

        [DisplayName("Название дисциплины")]
        public int TitleId { get; set; }

        public GroupModel Group { get; set; }
        public int GroupId { get; set; }

        [DisplayName("Кол-во часов")]
        public int Hours { get; set; }

        public int ZIndex { get; set; }

        public SubjectModel() 
            => Title = new LocalizerModel();
    }
}
