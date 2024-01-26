Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO

Public Class frmQuotationRecord

    Public Sub Getdata()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select Q_ID, RTRIM(QuotationNo), Date, Customer.ID,RTRIM(Customer.CustomerID),RTRIM(Name), RTRIM(ContactNo),GrandTotal, RTRIM(quotation.Remarks) from Customer,quotation where Customer.ID=quotation.CustomerID order by Date", con)
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
        fillQuotationNo()
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        Try
            If dgw.Rows.Count > 0 Then
                Dim dr As DataGridViewRow = dgw.SelectedRows(0)
                frmQuotation.Show()
                Me.Hide()
                frmQuotation.txtID.Text = dr.Cells(0).Value.ToString()
                frmQuotation.txtQuotationNo.Text = dr.Cells(1).Value.ToString()
                frmQuotation.dtpQuotationDate.Text = dr.Cells(2).Value.ToString()
                frmQuotation.txtCustomerID.Text = dr.Cells(4).Value.ToString()
                frmQuotation.txtCID.Text = dr.Cells(3).Value.ToString()
                frmQuotation.txtCustomerID.Text = dr.Cells(4).Value.ToString()
                frmQuotation.txtCustomerName.Text = dr.Cells(5).Value.ToString()
                frmQuotation.txtContactNo.Text = dr.Cells(6).Value.ToString()
                frmQuotation.txtGrandTotal.Text = dr.Cells(7).Value.ToString()
                frmQuotation.txtRemarks.Text = dr.Cells(8).Value.ToString()
                frmQuotation.btnSave.Enabled = False
                frmQuotation.btnUpdate.Enabled = True
                frmQuotation.btnPrint.Enabled = True
                frmQuotation.btnDelete.Enabled = True
                frmQuotation.lblSet.Text = "Not Allowed"
                frmQuotation.btnAdd.Enabled = False
                con = New SqlConnection(cs)
                con.Open()
                Dim sql As String = "SELECT RTRIM(ProductCode),RTRIM(ProductName), Quotation_Join.Cost, Quotation_Join.Qty, Quotation_Join.Amount, Quotation_Join.DiscountPer, Quotation_Join.Discount, Quotation_Join.VATPer, Quotation_Join.VAT, Quotation_Join.TotalAmount,Product.PID from quotation,Quotation_Join,Product where quotation.Q_ID=Quotation_Join.QuotationID and Product.PID=Quotation_Join.ProductID and quotation.Q_ID=@d1"
                cmd = New SqlCommand(sql, con)
                cmd.Parameters.AddWithValue("@d1", dr.Cells(0).Value.ToString())
                rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
                frmQuotation.DataGridView1.Rows.Clear()
                While (rdr.Read() = True)
                    frmQuotation.DataGridView1.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6), rdr(7), rdr(8), rdr(9), rdr(10))
                End While
                con.Close()
                con = New SqlConnection(cs)
                con.Open()
                Dim ct As String = "select RTRIM(CustomerType) from Customer where ID=" & dr.Cells(3).Value & ""
                cmd = New SqlCommand(ct)
                cmd.Connection = con
                rdr = cmd.ExecuteReader()
                If rdr.Read Then
                    frmQuotation.txtCustomerType.Text = rdr.GetValue(0)
                    If Not rdr Is Nothing Then
                        rdr.Close()
                    End If
                    Exit Sub
                End If
                con.Close()
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
    Sub fillQuotationNo()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(QuotationNo) FROM quotation", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbQuotationNo.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbQuotationNo.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Reset()
        cmbQuotationNo.Text = ""
        txtCustomerName.Text = ""
        fillQuotationNo()
        dtpDateFrom.Text = Today
        dtpDateTo.Text = Today
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

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select Q_ID, RTRIM(QuotationNo), Date, Customer.ID,RTRIM(Customer.CustomerID),RTRIM(Name), RTRIM(ContactNo),GrandTotal, RTRIM(quotation.Remarks) from Customer,quotation where Customer.ID=quotation.CustomerID and Date between @d1 and @d2 order by Date", con)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
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

    Private Sub cmbOrderNo_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbQuotationNo.SelectedIndexChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select Q_ID, RTRIM(QuotationNo), Date, Customer.ID,RTRIM(Customer.CustomerID),RTRIM(Name), RTRIM(ContactNo),GrandTotal, RTRIM(quotation.Remarks) from Customer,quotation where Customer.ID=quotation.CustomerID and QuotationNo='" & cmbQuotationNo.Text & "' order by Date", con)
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

    Private Sub txtCustomerName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtCustomerName.TextChanged
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select Q_ID, RTRIM(QuotationNo), Date, Customer.ID,RTRIM(Customer.CustomerID),RTRIM(Name), RTRIM(ContactNo),GrandTotal, RTRIM(quotation.Remarks) from Customer,quotation where Customer.ID=quotation.CustomerID and Name like '%" & txtCustomerName.Text & "%' order by Date", con)
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

    Private Sub cmbQuotationNo_Format(sender As System.Object, e As System.Windows.Forms.ListControlConvertEventArgs) Handles cmbQuotationNo.Format
        If (e.DesiredType Is GetType(String)) Then
            e.Value = e.Value.ToString.Trim
        End If
    End Sub
End Class
