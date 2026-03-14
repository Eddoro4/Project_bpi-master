using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_bpi.Models
{
    public class TemplateSection
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public int Section_Number { get; set; }
        public int Template_Id { get; set; }

        public List<TemplateSubsection> Subsections { get; set; }
    }
}
