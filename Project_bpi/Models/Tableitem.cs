using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_bpi.Models
{
    public class TableItem
    {
        public int Id { get; set; }
        public int TableId { get; set; }
        public int Column { get; set; }
        public int Row { get; set; }
        public string Header { get; set; }
        public int ColSpan { get; set; } = 1;
        public int RowSpan { get; set; } = 1;
        public bool IsHeader { get; set; }

        public virtual Table Table { get; set; }
    }
}
