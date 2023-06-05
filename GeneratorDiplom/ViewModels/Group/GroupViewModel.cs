using GeneratorDiplom.Models;
using System.ComponentModel;

namespace GeneratorDiplom.ViewModels.Group
{
    public class GroupViewModel
    {
        [DisplayName("Наименование на русском")]
        public string Title_RU { get; set; }

        [DisplayName("Наименование на казахском")]
        public string Title_KZ { get; set; }

        [DisplayName("Квалификация на русском")]
        public string Qualification_RU { get; set; }

        [DisplayName("Квалификация на казахском")]
        public string Qualification_KZ { get; set; }

        [DisplayName("Код группы")]
        public string Code { get; set; }

        [DisplayName("Язык обучения")]
        public LangModel Lang { get; set; }

        [DisplayName("Начало обучения")]
        public int StartStudies { get; set; }

        [DisplayName("Конец обучения")]
        public int EndStudies { get; set; }
    }
}
