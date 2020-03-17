using System.Collections.Generic;

namespace SoberScrappingTool.Models
{
    public class ScrappedResults
    {
        public Dictionary<int, CustomExcelRow> DataRowsToUpdate { get; set; }
        public List<CustomExcelRow> DataRowsToAdd { get; set; }
        public string SearchKeyword { get; set; }
    }
}
