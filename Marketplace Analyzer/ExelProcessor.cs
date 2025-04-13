using ClosedXML.Excel;

namespace Marketplace_Analyzer;


public class ExelProcessor
{
    public static List<ResultItem> ProcessOzonReport(string filePath, OzonParams ozonParams)
    {
        var result = new Dictionary<string, (int qty, double cost)>();
        using var workbook = new XLWorkbook(filePath);
        var ws = workbook.Worksheet(ozonParams.Sheet);

        int skipRows = 0;
        foreach (var row in ws.RowsUsed().Skip(skipRows))
        {
            string name = row.Cell(ozonParams.NameColumn).GetString();
            if (!int.TryParse(row.Cell(ozonParams.QtyColumn).GetString(), out int qty)) continue;
            if (!double.TryParse(row.Cell(ozonParams.CostColumn).GetString(), out double cost)) continue;
            if (double.TryParse(row.Cell(ozonParams.ReturnColumn).GetString(), out double ret)) continue;

            if (result.ContainsKey(name))
            {
                result[name] = (result[name].qty + qty, result[name].cost + cost);
            }
            else
            {
                result[name] = (qty, cost);
            }
        }

        var output = new List<ResultItem>();
        foreach (var kvp in result)
        {
            if (kvp.Key.Length < 25) continue;
            output.Add(new ResultItem
            {
                Name = kvp.Key,
                TotalQty = kvp.Value.qty,
                AvgPrice = Math.Round(kvp.Value.cost / kvp.Value.qty, 2)
            });
        }

        return output;
    }
    
    private class YandexTransactionItem(string productName, int quantity)
    {
        public string ProductName  { get; set; } = productName;
        public int Quantity { get; set; } = quantity;
    }

    private class YandexTotalCountItem(double cost, int totalQty)
    {
        public int TotalQty { get; set; } = totalQty;
        public double Cost { get; set; } = cost;
    }

    public static List<ResultItem> ProcessYandexReport(string filePath, YandexParams yandexParams)
    {
        var result = new Dictionary<string, (int qty, double cost)>();
        using var workbook = new XLWorkbook(filePath);
        var ws1 = workbook.Worksheet(yandexParams.TransactionsSheet);
        var ws2 = workbook.Worksheet(yandexParams.ServiceSheet);
        var map = new Dictionary<string, List<YandexTransactionItem>>();
        
        foreach (var row in ws1.RowsUsed())
        {
            string numberOrder = row.Cell(yandexParams.NumberOrder1).GetString();
            string name = row.Cell(yandexParams.NameProduct).GetString();
            if (!int.TryParse(row.Cell(yandexParams.CountInDelivery).GetString(), out int count)) continue;

            if (map.ContainsKey(numberOrder))
            {
                map[numberOrder].Add(new YandexTransactionItem(name, count));
            }
            else
            {
                map.Add(numberOrder, new List<YandexTransactionItem>());
                map[numberOrder].Add(new YandexTransactionItem(name, count));
            }
        }

        var numberIncomeDict = new Dictionary<string, List<Double>>();
        foreach (var row in ws2.RowsUsed())
        {
            string numberOrder = row.Cell(yandexParams.NumberOrder2).GetString();
            string statusPay = row.Cell(yandexParams.StatusPay).GetString();
            if (!double.TryParse(row.Cell(yandexParams.IncomeFromMarket).GetString(), out double incomeFromMarket)) continue;
            
            if (statusPay != "Переведён")
            {
                continue;
            }

            if (numberIncomeDict.ContainsKey(numberOrder))
            {
                numberIncomeDict[numberOrder].Add(incomeFromMarket);
            }
            else
            {
                numberIncomeDict.Add(numberOrder, new List<double>());
                numberIncomeDict[numberOrder].Add(incomeFromMarket);
            }
        }
        
        foreach (var order in numberIncomeDict)
        {
            string orderNumber = order.Key;
            var incomes = order.Value; // на всякий случай обрабатываем, если в одном заказе несколько доходов
            double totalIncome = incomes.Sum();

            if (!map.ContainsKey(orderNumber)) continue;
            var items = map[orderNumber];

            int totalQty = items.Sum(item => item.Quantity);
            if (totalQty == 0) continue;

            foreach (var item in items)
            {
                double share = (double)item.Quantity / totalQty;
                double itemIncome = totalIncome * share;

                if (result.ContainsKey(item.ProductName))
                {
                    result[item.ProductName] = (
                        result[item.ProductName].qty + item.Quantity,
                        result[item.ProductName].cost + itemIncome
                    );
                }
                else
                {
                    result[item.ProductName] = (item.Quantity, itemIncome);
                }
            }
        }
        
        var output = new List<ResultItem>();
        foreach (var kvp in result)
        {
            if (kvp.Key.Length < 25) continue;
            output.Add(new ResultItem
            {
                Name = kvp.Key,
                TotalQty = kvp.Value.qty,
                AvgPrice = Math.Round(kvp.Value.cost / kvp.Value.qty, 2)
            });
        }

        return output;
    }
}

