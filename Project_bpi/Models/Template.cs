using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_bpi.Models
{
    public class Template
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int Year { get; set; }
        public string Path { get; set; } 
        public List<TemplateSection> Sections { get; set; }
    }
}
