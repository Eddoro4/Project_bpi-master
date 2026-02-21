using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_bpi.Models
{
    public class PatternFile
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public int Year { get; set; }
        public string Path { get; set; }
        public virtual ICollection<Report> Reports { get; set; }
    }
}
