using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_bpi.Models
{
    public class TemplateSubsection
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public int Subsection_Number { get; set; }
        public int SectionId { get; set; }

        public List<TemplateContent> TemplateContents { get; set; }
    }
}
