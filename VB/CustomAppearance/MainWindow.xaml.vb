Imports System
Imports System.Windows
Imports System.Windows.Media
Imports CustomAppearance.DataSet1TableAdapters
Imports DevExpress.Xpf.PivotGrid

Namespace CustomAppearance

    Public Partial Class MainWindow
        Inherits Window

        Private salesPersonDataTable As DataSet1.SalesPersonDataTable = New DataSet1.SalesPersonDataTable()

        Private salesPersonDataAdapter As SalesPersonTableAdapter = New SalesPersonTableAdapter()

        Private minValue, maxValue, minTotalValue, maxTotalValue As Decimal

        Private maxMinCalculated As Boolean

        Public Sub New()
            Me.InitializeComponent()
            Me.pivotGridControl1.DataSource = salesPersonDataTable
        End Sub

        Private Sub Window_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            salesPersonDataAdapter.Fill(salesPersonDataTable)
            Me.pivotGridControl1.CollapseAll()
        End Sub

        ' Handles the CustomCellAppearance event. Sets the current cell's foreground color
        ' to a color generated in a custom manner.
        Private Sub OnCustomAppearance(ByVal sender As Object, ByVal e As PivotCustomCellAppearanceEventArgs)
            If Not(TypeOf e.Value Is Decimal) Then Return
            Dim value As Decimal = CDec(e.Value)
            Dim isGrandTotal As Boolean = Me.IsGrandTotal(e)
            If IsValueNotFit(value, isGrandTotal) Then ResetMaxMin()
            EnsureMaxMin()
            If IsValueNotFit(value, isGrandTotal) Then Return
            e.Foreground = New SolidColorBrush(GetColorByValue(value, Me.IsGrandTotal(e)))
        End Sub

        Private Function IsValueNotFit(ByVal value As Decimal, ByVal isGrandToatl As Boolean) As Boolean
            If isGrandToatl Then
                Return value < minTotalValue OrElse value > maxTotalValue
            Else
                Return value < minValue OrElse value > maxValue
            End If
        End Function

        Private Sub OnGridLayout(ByVal sender As Object, ByVal e As RoutedEventArgs)
            ResetMaxMin()
        End Sub

        ' Generates a custom color for a cell based on the cell value's share 
        ' in the maximum summary or total value (for summary and total cells),
        ' or in the maximum Grand Total value (for Grand Total cells).
        Private Function GetColorByValue(ByVal value As Decimal, ByVal isGrandTotal As Boolean) As Color
            Dim variation As Integer
            If isGrandTotal Then
                variation = Convert.ToInt32(510 * (value - minTotalValue) / (maxTotalValue - minTotalValue))
            Else
                variation = Convert.ToInt32(510 * (value - minValue) / (maxValue - minValue))
            End If

            Dim r, b As Byte
            If variation >= 255 Then
                r = Convert.ToByte(510 - variation)
                b = 255
            Else
                r = 255
                b = Convert.ToByte(variation)
            End If

            Return Color.FromRgb(r, 0, b)
        End Function

        Private Function IsGrandTotal(ByVal e As PivotCustomSummaryEventArgs) As Boolean
            Return e.RowField Is Nothing OrElse e.ColumnField Is Nothing
        End Function

        Private Function IsGrandTotal(ByVal e As PivotCellBaseEventArgs) As Boolean
            Return e.RowValueType = FieldValueType.GrandTotal OrElse e.ColumnValueType = FieldValueType.GrandTotal
        End Function

        ' Calculates the maximum and minimum summary and Grand Total values.
        Private Sub EnsureMaxMin()
            If maxMinCalculated Then Return
            For i As Integer = 0 To Me.pivotGridControl1.RowCount - 1
                For j As Integer = 0 To Me.pivotGridControl1.ColumnCount - 1
                    Dim val As Object = Me.pivotGridControl1.GetCellValue(j, i)
                    If Not(TypeOf val Is Decimal) Then Continue For
                    Dim value As Decimal = CDec(val)
                    Dim isGrandTotal As Boolean = Me.pivotGridControl1.GetFieldValueType(True, j) = FieldValueType.GrandTotal OrElse Me.pivotGridControl1.GetFieldValueType(False, i) = FieldValueType.GrandTotal
                    If isGrandTotal Then
                        If value > maxTotalValue Then maxTotalValue = value
                        If value < minTotalValue Then minTotalValue = value
                    Else
                        If value > maxValue Then maxValue = value
                        If value < minValue Then minValue = value
                    End If
                Next
            Next

            If minTotalValue = maxTotalValue Then maxTotalValue += 1
            If minValue = maxValue Then maxValue += 1
            maxMinCalculated = True
        End Sub

        ' Resets the maximum and minimum summary and Grand Total values.
        Private Sub ResetMaxMin()
            minValue = Decimal.MaxValue
            maxValue = Decimal.MinValue
            minTotalValue = Decimal.MaxValue
            maxTotalValue = Decimal.MinValue
            maxMinCalculated = False
        End Sub
    End Class
End Namespace
