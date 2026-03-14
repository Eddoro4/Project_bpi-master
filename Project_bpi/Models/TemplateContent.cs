using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_bpi.Models
{
    public class TemplateContent
    {
        public int Id { get; set; }
        public string Description { get; set; }
        public string Tag { get; set; }
        public string ContentType { get; set; }
        public int Subsection_Id { get; set; }
    }
}
