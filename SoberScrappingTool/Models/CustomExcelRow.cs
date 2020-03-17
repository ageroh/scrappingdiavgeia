using System;

namespace SoberScrappingTool.Models
{
    public class CustomExcelRow
    {
        public string Code { get; set; }
        public DateTime LastDateChanged { get; set; }
        public string Pdf { get; set; }
        public string Provider { get; set; }
        public string Description { get; set; }
        public string Type { get; set; }
        public string Categories { get; set; }
        public Status Status { get; set; }
    }
}
