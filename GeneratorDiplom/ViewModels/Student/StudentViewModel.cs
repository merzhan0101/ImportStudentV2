using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace GeneratorDiplom.ViewModels.Student
{
    public class StudentViewModel
    {
        [Required]
        [DisplayName("Имя на русском")]
        public string Name_RU { get; set; }

        [Required]
        [DisplayName("Фамилия на русском")]
        public string Surname_RU { get; set; }

        [Required]
        [DisplayName("Отчество на русском")]
        public string Middlename_RU { get; set; }

        [Required]
        [DisplayName("Имя на русском в дательном падеже")]
        public string Name_Dat_RU { get; set; }

        [Required]
        [DisplayName("Фамилия на русском в дательном падеже")]
        public string Surname_Dat_RU { get; set; }

        [Required]
        [DisplayName("Отчество на русском в дательном падеже")]
        public string Middlename_Dat_RU { get; set; }

        [Required]
        [DisplayName("Имя на казахском в дательном падеже")]
        public string Name_Dat_KZ { get; set; }

        [Required]
        [DisplayName("Фамилия на казахском в дательном падеже")]
        public string Surname_Dat_KZ { get; set; }

        [Required]
        [DisplayName("Отчество на казахском в дательном падеже")]
        public string Middlename_Dat_KZ { get; set; }


        [DisplayName("Группа")]
        public int GroupId { get; set; }

        [Required]
        [DisplayName("Номер приложения")]
        public string NumApplication { get; set; }
    }
}
