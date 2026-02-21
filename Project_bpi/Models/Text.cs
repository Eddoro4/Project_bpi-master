using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_bpi.Models
{
    public class Text
    {
        public int Id { get; set; }
        public int SubsectionId { get; set; }
        public string Content { get; set; }
        public string PatternName { get; set; }

        public virtual SubSection Subsection { get; set; }
    }
}
