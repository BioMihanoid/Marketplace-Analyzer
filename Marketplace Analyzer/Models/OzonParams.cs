namespace Marketplace_Analyzer.Models;

public class OzonParams(string sheet, string nameColumn, string qtyColumn, string costColumn, string returnColumn)
{
    public string Sheet { get; set; } = sheet;
    public string NameColumn { get; set; } = nameColumn;
    public string QtyColumn { get; set; } = qtyColumn;
    public string CostColumn { get; set; } = costColumn;
    public string ReturnColumn { get; set; } = returnColumn;
}