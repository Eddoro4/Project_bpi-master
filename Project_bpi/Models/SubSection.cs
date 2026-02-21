using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using static System.Net.Mime.MediaTypeNames;

namespace Project_bpi.Models
{
    public class SubSection
    {
        public int Id { get; set; }
        public int SectionId { get; set; }
        public int Number { get; set; }
        public string Title { get; set; }

        public virtual Section Section { get; set; }
        public virtual ICollection<Table> Tables { get; set; }
        public virtual ICollection<Text> Texts { get; set; }
    }
}
