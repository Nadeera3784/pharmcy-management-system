Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO

Public Class frmCurrentStock

    Public Sub Getdata()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT PID, RTRIM(Product.ProductCode),RTRIM(ProductName),RTRIM(Description),(CostPrice),(SellingPrice),(Discount),(VAT),Qty from Temp_Stock,Product where Product.PID=Temp_Stock.ProductID and Qty > 0  order by ProductName", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmLogs_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Getdata()

        dgw.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        dgw.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        dgw.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        dgw.Columns(7).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        dgw.Columns(8).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                If lblSet.Text = "Billing" Then
                    frmBilling.Show()
                    Me.Hide()
                    frmBilling.txtProductID.Text = dr.Cells(0).Value.ToString()
                    frmBilling.txtProductCode.Text = dr.Cells(1).Value.ToString()
                    frmBilling.txtProductName.Text = dr.Cells(2).Value.ToString()
                    frmBilling.txtCostPrice.Text = dr.Cells(4).Value.ToString()
                    frmBilling.txtSellingPrice.Text = dr.Cells(5).Value.ToString()
                    Dim num As Double
                    num = Val(dr.Cells(5).Value) - Val(dr.Cells(4).Value)
                    num = Math.Round(num, 2)
                    frmBilling.txtMargin.Text = num
                    frmBilling.txtDiscountPer.Text = dr.Cells(6).Value.ToString()
                    frmBilling.txtVAT.Text = dr.Cells(7).Value.ToString()
                    frmBilling.txtQty.Focus()
                    lblSet.Text = ""
                End If
                If lblSet.Text = "Billing1" Then
                    frmBilling1.Show()
                    Me.Hide()
                    frmBilling1.txtProductID.Text = dr.Cells(0).Value.ToString()
                    frmBilling1.txtProductCode.Text = dr.Cells(1).Value.ToString()
                    frmBilling1.txtProductName.Text = dr.Cells(2).Value.ToString()
                    frmBilling1.txtCostPrice.Text = dr.Cells(4).Value.ToString()
                    frmBilling1.txtSellingPrice.Text = dr.Cells(5).Value.ToString()
                    Dim num As Double
                    num = Val(dr.Cells(5).Value) - Val(dr.Cells(4).Value)
                    num = Math.Round(num, 2)
                    frmBilling1.txtMargin.Text = num
                    frmBilling1.txtDiscountPer.Text = dr.Cells(6).Value.ToString()
                    frmBilling1.txtVAT.Text = dr.Cells(7).Value.ToString()
                    frmBilling1.txtQty.Focus()
                    lblSet.Text = ""
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgw_RowPostPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles dgw.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If dgw.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            dgw.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub
    Sub Reset()
        txtProductName.Text = ""
        Getdata()
    End Sub
    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub



    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        Dim rowsTotal, colsTotal As Short
        Dim I, j, iC As Short
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Dim xlApp As New Excel.Application
        Try
            Dim excelBook As Excel.Workbook = xlApp.Workbooks.Add
            Dim excelWorksheet As Excel.Worksheet = CType(excelBook.Worksheets(1), Excel.Worksheet)
            xlApp.Visible = True

            rowsTotal = dgw.RowCount
            colsTotal = dgw.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = dgw.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = dgw.Rows(I).Cells(j).Value
                    Next j
                Next I
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 12

                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'RELEASE ALLOACTED RESOURCES
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            xlApp = Nothing
        End Try
    End Sub

    Private Sub txtProductName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtProductName.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT PID, RTRIM(Product.ProductCode),RTRIM(ProductName),RTRIM(Description),(CostPrice),(SellingPrice),(Discount),(VAT),Qty from Temp_Stock,Product where Product.PID=Temp_Stock.ProductID and Qty > 0 and ProductName like '%" & txtProductName.Text & "%' order by ProductName", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            dgw.Rows.Clear()
            While (rdr.Read() = True)
                dgw.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
