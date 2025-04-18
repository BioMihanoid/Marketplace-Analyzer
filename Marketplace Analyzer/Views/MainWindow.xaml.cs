using System.Windows;
using ClosedXML.Excel;
using Marketplace_Analyzer.Models;
using Marketplace_Analyzer.Services;
using Microsoft.Win32;

namespace Marketplace_Analyzer.Views;

public partial class MainWindow : Window
    {
        private string _loadedFilePath;
        private List<ResultItem> _resultItems = new();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void LoadFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog { Filter = "Excel файлы (*.xlsx)|*.xlsx" };
            if (openFileDialog.ShowDialog() == true)
            {
                _loadedFilePath = openFileDialog.FileName;
                FilePathText.Text = _loadedFilePath;
            }
        }

        private void MarketplaceComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_loadedFilePath))
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
            if (string.IsNullOrEmpty(_loadedFilePath)) return;
            
            var ozonParams = new OzonParams(SheetNameBox.Text, NameColBox.Text, 
                QtyColBox.Text, CostColBox.Text, ReturnColBox.Text);

            var yandexParams = new YandexParams(TransactionsSheet.Text, NumberOrder1.Text, NameProduct.Text, 
                CountInDelivery.Text, ServicesSheet.Text, NumberOrder2.Text, IncomeFromMarket.Text, StatusPay.Text);

            var sumValue = 0.0;
            switch (MarketplaceCombo.SelectedIndex)
            {
                case 0:
                    _resultItems = ExelProcessor.ProcessOzonReport(_loadedFilePath, ozonParams, ref sumValue);
                    break;
                case 1: 
                    _resultItems = ExelProcessor.ProcessYandexReport(_loadedFilePath, yandexParams, ref sumValue);
                    break;
            }
            ResultDataGrid.ItemsSource = _resultItems;
            SaveButton.IsEnabled = true;
            ResetButton.IsEnabled = true;
            SumValue.Text = "Итоговая стоимость: " + sumValue.ToString("F2") + " рублей";
            SumValue.IsEnabled = true;
        }

        private void SaveToExcel_Click(object sender, RoutedEventArgs e)
        {
            if (_resultItems.Count == 0) return;

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

                for (int i = 0; i < _resultItems.Count; i++)
                {
                    ws.Cell(i + 2, 1).Value = _resultItems[i].Name;
                    ws.Cell(i + 2, 2).Value = _resultItems[i].TotalQty;
                    ws.Cell(i + 2, 3).Value = _resultItems[i].AvgPrice;
                }

                workbook.SaveAs(saveFileDialog.FileName);
                MessageBox.Show("Результат сохранён!");
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            _loadedFilePath = string.Empty;
            FilePathText.Text = "Файл не выбран";
            ResultDataGrid.ItemsSource = null;
            SaveButton.IsEnabled = false;
            CalcButton.IsEnabled = false;
            MarketplaceCombo.SelectedIndex = -1;
            OzonOptions.Visibility = Visibility.Collapsed;
            YandexOptions.Visibility = Visibility.Collapsed;
            ResetButton.IsEnabled = false;
            SumValue.IsEnabled = false;
            SumValue.Text = "";
        }
    }