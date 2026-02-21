using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_bpi.Models
{
    public class Section
    {
        public int Id { get; set; }
        public int ReportId { get; set; }
        public int Number { get; set; }
        public string Title { get; set; }

        public virtual Report Report { get; set; }
        public virtual ICollection<SubSection> SubSections { get; set; }
    }
}
