using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace Marketplace_Analyzer;

public partial class MainWindow : Window
    {
        private string loadedFilePath;
        private List<ResultItem> resultItems = new();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void LoadFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = "Excel файлы (*.xlsx)|*.xlsx" };
            if (openFileDialog.ShowDialog() == true)
            {
                loadedFilePath = openFileDialog.FileName;
                FilePathText.Text = loadedFilePath;
            }
        }

        private void MarketplaceComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(loadedFilePath))
            {
                if (MarketplaceCombo.SelectedIndex == 0)
                {
                    OzonOptions.Visibility = Visibility.Visible;
                    YandexOptions.Visibility = Visibility.Collapsed;
                }
                else if (MarketplaceCombo.SelectedIndex == 1)
                {
                    OzonOptions.Visibility = Visibility.Collapsed;
                    YandexOptions.Visibility = Visibility.Visible;
                }
                CalcButton.IsEnabled = true;
            }
        }

        private void Calculate_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(loadedFilePath)) return;
            
            var ozonParams = new OzonParams(SheetNameBox.Text, NameColBox.Text, 
                QtyColBox.Text, CostColBox.Text);

            var yandexParams = new YandexParams(TransactionsSheet.Text, NumberOrder1.Text, NameProduct.Text, 
                CountInDelivery.Text, ServicesSheet.Text, NumberOrder2.Text, IncomeFromMarket.Text, StatusPay.Text);

            switch (MarketplaceCombo.SelectedIndex)
            {
                case 0:
                    resultItems = ExelProcessor.ProcessOzonReport(loadedFilePath, ozonParams);
                    break;
                case 1: 
                    resultItems = ExelProcessor.ProcessYandexReport(loadedFilePath, yandexParams);
                    break;
            }
            ResultDataGrid.ItemsSource = resultItems;
            SaveButton.IsEnabled = true;
            ResetButton.IsEnabled = true;
        }

        private void SaveToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (resultItems.Count == 0) return;

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel File (*.xlsx)|*.xlsx",
                FileName = "result.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                using var workbook = new XLWorkbook();
                var ws = workbook.AddWorksheet("Результат");
                ws.Cell(1, 1).Value = "Название";
                ws.Cell(1, 2).Value = "Количество";
                ws.Cell(1, 3).Value = "Средняя цена";

                for (int i = 0; i < resultItems.Count; i++)
                {
                    ws.Cell(i + 2, 1).Value = resultItems[i].Name;
                    ws.Cell(i + 2, 2).Value = resultItems[i].TotalQty;
                    ws.Cell(i + 2, 3).Value = resultItems[i].AvgPrice;
                }

                workbook.SaveAs(saveFileDialog.FileName);
                MessageBox.Show("Результат сохранён!");
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            loadedFilePath = string.Empty;
            FilePathText.Text = "Файл не выбран";
            ResultDataGrid.ItemsSource = null;
            SaveButton.IsEnabled = false;
            CalcButton.IsEnabled = false;
            MarketplaceCombo.SelectedIndex = -1;
            OzonOptions.Visibility = Visibility.Collapsed;
            YandexOptions.Visibility = Visibility.Collapsed;
            ResetButton.IsEnabled = false;
        }
    }