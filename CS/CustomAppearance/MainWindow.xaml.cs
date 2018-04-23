using System;
using System.Windows;
using System.Windows.Media;
using CustomAppearance.DataSet1TableAdapters;
using DevExpress.Xpf.PivotGrid;

namespace CustomAppearance {
    public partial class MainWindow : Window {
        DataSet1.SalesPersonDataTable salesPersonDataTable = new DataSet1.SalesPersonDataTable();
        SalesPersonTableAdapter salesPersonDataAdapter = new SalesPersonTableAdapter();
        decimal minValue, maxValue, minTotalValue, maxTotalValue;
        bool maxMinCalculated;
        public MainWindow() {
            InitializeComponent();
            pivotGridControl1.DataSource = salesPersonDataTable;
        }
        void Window_Loaded(object sender, RoutedEventArgs e) {
            salesPersonDataAdapter.Fill(salesPersonDataTable); 
            pivotGridControl1.CollapseAll();
        }

        // Handles the CustomCellAppearance event. Sets the current cell's foreground color
        // to a color generated in a custom manner.
        void OnCustomAppearance(object sender, PivotCustomCellAppearanceEventArgs e) {
            if(!(e.Value is decimal)) return;
            decimal value = (decimal)e.Value;
            bool isGrandTotal = IsGrandTotal(e);
            if(IsValueNotFit(value, isGrandTotal))
                ResetMaxMin();
            EnsureMaxMin();
            if(IsValueNotFit(value, isGrandTotal))
                return;
            e.Foreground =
                new SolidColorBrush(GetColorByValue(value, IsGrandTotal(e)));
        }

        bool IsValueNotFit(decimal value, bool isGrandToatl) {
            if(isGrandToatl)
                return value < minTotalValue || value > maxTotalValue;
            else
                return value < minValue || value > maxValue;
        }
        void OnGridLayout(object sender, RoutedEventArgs e) {
            ResetMaxMin();
        }

        // Generates a custom color for a cell based on the cell value's share 
        // in the maximum summary or total value (for summary and total cells),
        // or in the maximum Grand Total value (for Grand Total cells).
        Color GetColorByValue(decimal value, bool isGrandTotal) {
            int variation;
            if(isGrandTotal) {
                variation = 
                    Convert.ToInt32(510 * (value - minTotalValue) / (maxTotalValue - minTotalValue));
            } else {
                variation = 
                    Convert.ToInt32(510 * (value - minValue) / (maxValue - minValue));
            }
            byte r, b;
            if(variation >= 255) {
                r = Convert.ToByte(510 - variation);
                b = 255;
            } else {
                r = 255;
                b = Convert.ToByte(variation);
            }
            return Color.FromRgb(r, 0, b);
        }
        bool IsGrandTotal(PivotCustomSummaryEventArgs e) {
            return e.RowField == null || e.ColumnField == null;
        }
        bool IsGrandTotal(PivotCellBaseEventArgs e) {
            return e.RowValueType == FieldValueType.GrandTotal || 
                e.ColumnValueType == FieldValueType.GrandTotal;
        }

        // Calculates the maximum and minimum summary and Grand Total values.
        void EnsureMaxMin() {
            if(maxMinCalculated) return;
            for(int i = 0; i < pivotGridControl1.RowCount; i++)
                for(int j = 0; j < pivotGridControl1.ColumnCount; j++) {
                    object val = pivotGridControl1.GetCellValue(j, i);
                    if(!(val is decimal)) continue;
                    decimal value = (decimal)val;
                    bool isGrandTotal =
                      pivotGridControl1.GetFieldValueType(true, j) == FieldValueType.GrandTotal ||
                      pivotGridControl1.GetFieldValueType(false, i) == FieldValueType.GrandTotal;
                    if(isGrandTotal) {
                        if(value > maxTotalValue)
                            maxTotalValue = value;
                        if(value < minTotalValue)
                            minTotalValue = value;
                    } else {
                        if(value > maxValue)
                            maxValue = value;
                        if(value < minValue)
                            minValue = value;
                    }
                }
            if(minTotalValue == maxTotalValue)
                maxTotalValue++;
            if(minValue == maxValue)
                maxValue++;
            maxMinCalculated = true;
        }

        // Resets the maximum and minimum summary and Grand Total values.
        void ResetMaxMin() {
            minValue = decimal.MaxValue;
            maxValue = decimal.MinValue;
            minTotalValue = decimal.MaxValue;
            maxTotalValue = decimal.MinValue;
            maxMinCalculated = false;
        }
    }
}
