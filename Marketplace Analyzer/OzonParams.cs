namespace Marketplace_Analyzer;

public class OzonParams
{
    public OzonParams(string sheet, string nameColumn, string qtyColumn, string costColumn)
    {
        Sheet = sheet;
        NameColumn = nameColumn;
        QtyColumn = qtyColumn;
        CostColumn = costColumn;
    }

    public string Sheet { get; set; }
    public string NameColumn { get; set; }
    public string QtyColumn { get; set; }
    public string CostColumn { get; set; }
}