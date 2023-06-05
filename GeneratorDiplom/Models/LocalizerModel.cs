using System.ComponentModel.DataAnnotations.Schema;

namespace GeneratorDiplom.Models
{
    public class LocalizerModel
    {
        public int Id { get; set; }
        public string Title_RU { get; set; }
        public string Title_KZ { get; set; }

        [NotMapped]
        public string Get
        {
            get {
                return Title_RU;
            }
        }

    }
}
