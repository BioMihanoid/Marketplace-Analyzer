namespace Marketplace_Analyzer;

public class YandexParams
{
    public YandexParams(string transactionsSheet, string numberOrder1, string nameProduct, string countInDelivery, 
        string serviceSheet, string numberOrder2, string incomeFromMarket, string statusPay)
    {
        TransactionsSheet = transactionsSheet;
        NumberOrder1 = numberOrder1;
        NameProduct = nameProduct;
        CountInDelivery = countInDelivery;
        ServiceSheet = serviceSheet;
        NumberOrder2 = numberOrder2;
        IncomeFromMarket = incomeFromMarket;
        StatusPay = statusPay;
    }

    public string TransactionsSheet { get; set; }
    public string NumberOrder1 { get; set; }
    public string NameProduct { get; set; }
    public string CountInDelivery { get; set; }
    public string ServiceSheet { get; set; }
    public string NumberOrder2 { get; set; }
    public string IncomeFromMarket { get; set; }
    public string StatusPay { get; set; }
}