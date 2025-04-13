namespace Marketplace_Analyzer;

public class OzonParams
{
    public OzonParams(string sheet, string nameColumn, string qtyColumn, string costColumn, string returnColumn)
    {
        Sheet = sheet;
        NameColumn = nameColumn;
        QtyColumn = qtyColumn;
        CostColumn = costColumn;
        ReturnColumn = returnColumn;
    }

    public string Sheet { get; set; }
    public string NameColumn { get; set; }
    public string QtyColumn { get; set; }
    public string CostColumn { get; set; }
    public string ReturnColumn { get; set; }
}