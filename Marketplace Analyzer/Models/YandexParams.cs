namespace Marketplace_Analyzer.Models;

public class YandexParams(
    string transactionsSheet,
    string numberOrder1,
    string nameProduct,
    string countInDelivery,
    string serviceSheet,
    string numberOrder2,
    string incomeFromMarket,
    string statusPay)
{
    public string TransactionsSheet { get; set; } = transactionsSheet;
    public string NumberOrder1 { get; set; } = numberOrder1;
    public string NameProduct { get; set; } = nameProduct;
    public string CountInDelivery { get; set; } = countInDelivery;
    public string ServiceSheet { get; set; } = serviceSheet;
    public string NumberOrder2 { get; set; } = numberOrder2;
    public string IncomeFromMarket { get; set; } = incomeFromMarket;
    public string StatusPay { get; set; } = statusPay;
}