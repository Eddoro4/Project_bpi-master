using System.Collections.Generic;

namespace Project_bpi.Models
{
    public class SubSection
    {
        public int Id { get; set; }
        public int SectionId { get; set; }
        public int? ParentSubsectionId { get; set; }
        public int Number { get; set; }
        public string Title { get; set; }

        public virtual Section Section { get; set; }
        public virtual SubSection ParentSubsection { get; set; }
        public virtual ICollection<SubSection> SubSections { get; set; }
        public virtual ICollection<Table> Tables { get; set; }
        public virtual ICollection<Text> Texts { get; set; }
    }
}
