using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_bpi.Models
{
    public class Report
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public int Year { get; set; }
        public int PattarnId { get; set; }

        public virtual PatternFile PatternFile { get; set; }
        public virtual ICollection<Section> Sections { get; set; }
    }
}
