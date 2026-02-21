using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Project_bpi.Models
{
    public class Table
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public int SubsectionId { get; set; }
        public string PatternName { get; set; }

        // Навигационные свойства
        public virtual SubSection Subsection { get; set; }
        public virtual ICollection<TableItem> TableItems { get; set; }
        public int RowCount => TableItems?.Max(i => i.Row) ?? 0;
        public int ColumnCount => TableItems?.Max(i => i.Column) ?? 0;
    }
}
