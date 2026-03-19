using System;

namespace Project_bpi.Models
{
    public class HistoryEntry
    {
        public int Id { get; set; }
        public DateTime ChangedAtUtc { get; set; }
        public string ActionType { get; set; }
        public string EntityType { get; set; }
        public string Location { get; set; }
        public string Details { get; set; }
    }
}
