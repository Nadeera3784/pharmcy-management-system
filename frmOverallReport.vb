Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmOverallReport

    Dim a, b, c, d, f, g, h, i As Decimal

    Sub Reset()
        dtpDateFrom.Text = Today
        dtpDateTo.Text = Today
    End Sub
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub


    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub


    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptOverall 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand, MyCommand1, MyCommand2, MyCommand3 As New SqlCommand()
            Dim myDA, myDA1, myDA2, myDA3 As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand2.Connection = myConnection
            MyCommand3.Connection = myConnection
            MyCommand.CommandText = "SELECT Customer.ID, Customer.Name, Customer.Gender, Customer.Address, Customer.City, Customer.State, Customer.ZipCode, Customer.ContactNo, Customer.EmailID, Customer.Remarks,Customer.Photo, InvoiceInfo.Inv_ID, InvoiceInfo.InvoiceNo, InvoiceInfo.InvoiceDate, InvoiceInfo.CustomerID , InvoiceInfo.GrandTotal, InvoiceInfo.TotalPaid, InvoiceInfo.Balance, Invoice_Product.IPo_ID, Invoice_Product.InvoiceID, Invoice_Product.ProductID, Invoice_Product.CostPrice, Invoice_Product.SellingPrice, Invoice_Product.Margin,Invoice_Product.Qty, Invoice_Product.Amount, Invoice_Product.DiscountPer, Invoice_Product.Discount, Invoice_Product.VATPer, Invoice_Product.VAT, Invoice_Product.TotalAmount, Product.PID,Product.ProductCode, Product.ProductName FROM Customer INNER JOIN InvoiceInfo ON Customer.ID = InvoiceInfo.CustomerID INNER JOIN Invoice_Product ON InvoiceInfo.Inv_ID = Invoice_Product.InvoiceID INNER JOIN Product ON Invoice_Product.ProductID = Product.PID where InvoiceDate between @d1 and @d2 order by InvoiceDate"
            MyCommand.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            MyCommand.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            MyCommand1.CommandText = "SELECT * FROM Service INNER JOIN Customer ON Service.CustomerID = Customer.ID INNER JOIN InvoiceInfo1 ON Service.S_ID = InvoiceInfo1.ServiceID INNER JOIN Invoice1_Product ON InvoiceInfo1.Inv_ID = Invoice1_Product.InvoiceID where InvoiceInfo1.InvoiceDate between @d1 and @d2 order by invoiceDate"
            MyCommand1.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            MyCommand1.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            MyCommand2.CommandText = "SELECT Distinct Stock.ST_ID, Stock.Stock_ID, Stock.Date, Stock.SupplierID, Stock.GrandTotal, Stock.TotalPayment, Stock.PaymentDue, Stock.Remarks, Stock_Product.SP_ID, Stock_Product.StockID, Stock_Product.ProductID,Stock_Product.Qty, Stock_Product.Price, Stock_Product.TotalAmount, Supplier.ID, Supplier.SupplierID AS Expr1, Supplier.Name, Supplier.Address, Supplier.City, Supplier.State, Supplier.ZipCode,Supplier.ContactNo, Supplier.EmailID, Supplier.Remarks AS Expr2, Product.PID, Product.ProductCode, Product.ProductName, Product.SubCategoryID, Product.Description, Product.CostPrice, Product.SellingPrice,Product.Discount, Product.VAT, Product.ReorderPoint FROM Stock INNER JOIN Stock_Product ON Stock.ST_ID = Stock_Product.StockID INNER JOIN Supplier ON Stock.SupplierID = Supplier.ID INNER JOIN Product ON Stock_Product.ProductID = Product.PID where Stock.Date between @d1 and @d2 order by Stock.Date"
            MyCommand2.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            MyCommand2.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            MyCommand3.CommandText = "SELECT Voucher.Id, Voucher.VoucherNo, Voucher.Name, Voucher.Date, Voucher.Details, Voucher.GrandTotal, Voucher_OtherDetails.VD_ID, Voucher_OtherDetails.VoucherID, Voucher_OtherDetails.Particulars,Voucher_OtherDetails.Amount, Voucher_OtherDetails.Note FROM Voucher INNER JOIN Voucher_OtherDetails ON Voucher.Id = Voucher_OtherDetails.VoucherID where Voucher.Date between @d1 and @d2 order by Voucher.Date"
            MyCommand3.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            MyCommand3.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            MyCommand2.CommandType = CommandType.Text
            MyCommand3.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA2.SelectCommand = MyCommand2
            myDA3.SelectCommand = MyCommand3
            myDA.Fill(myDS, "InvoiceInfo")
            myDA.Fill(myDS, "Invoice_Product")
            myDA.Fill(myDS, "Customer")
            myDA.Fill(myDS, "Product")
            myDA1.Fill(myDS, "InvoiceInfo1")
            myDA1.Fill(myDS, "Invoice1_Product")
            myDA1.Fill(myDS, "Service")
            myDA1.Fill(myDS, "Customer")
            myDA2.Fill(myDS, "Stock")
            myDA2.Fill(myDS, "Stock_Product")
            myDA2.Fill(myDS, "Product")
            myDA2.Fill(myDS, "Supplier")
            myDA3.Fill(myDS, "Voucher")
            myDA3.Fill(myDS, "Voucher_OtherDetails")
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select ISNULL(sum(GrandTotal),0),ISNULL(sum(TotalPaid),0),ISNULL(sum(Balance),0) from InvoiceInfo where InvoiceDate between @d1 and @d2"
            cmd = New SqlCommand(ct)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            cmd.Connection = con
            rdr = cmd.ExecuteReader
            If (rdr.Read()) Then
                a = rdr.GetValue(0)
                b = rdr.GetValue(1)
                c = rdr.GetValue(2)

            Else
                a = 0
                b = 0
                c = 0
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim ct1 As String = "select ISNULL(sum(Margin),0) from InvoiceInfo,Invoice_Product where InvoiceInfo.Inv_ID=Invoice_Product.InvoiceID and InvoiceDate between @d1 and @d2"
            cmd = New SqlCommand(ct1)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            cmd.Connection = con
            rdr = cmd.ExecuteReader
            If (rdr.Read()) Then
                d = rdr.GetValue(0)
            Else
                d = 0
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim ct2 As String = "select ISNULL(sum(GrandTotal),0),ISNULL(sum(TotalPayment),0),ISNULL(sum(PaymentDue),0) from Stock,Supplier where Supplier.ID=Stock.SupplierID and Date between @d3 and @d4"
            cmd = New SqlCommand(ct2)
            cmd.Parameters.Add("@d3", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d4", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            While rdr.Read()
                f = rdr.GetValue(0)
                g = rdr.GetValue(1)
                h = rdr.GetValue(2)
            End While
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim ct3 As String = "select ISNULL(sum(GrandTotal),0) from Voucher where Date between @d1 and @d2"
            cmd = New SqlCommand(ct3)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            While rdr.Read()
                i = rdr.GetValue(0)
            End While
            rpt.Subreports(0).SetDataSource(myDS)
            rpt.Subreports(1).SetDataSource(myDS)
            rpt.Subreports(2).SetDataSource(myDS)
            rpt.Subreports(3).SetDataSource(myDS)
            rpt.SetParameterValue("p1", dtpDateFrom.Value.Date)
            rpt.SetParameterValue("p2", dtpDateTo.Value.Date)
            rpt.SetParameterValue("p3", a)
            rpt.SetParameterValue("p4", b)
            rpt.SetParameterValue("p5", c)
            rpt.SetParameterValue("p6", d)
            rpt.SetParameterValue("p7", Today)
            rpt.SetParameterValue("p8", f)
            rpt.SetParameterValue("p9", g)
            rpt.SetParameterValue("p10", h)
            rpt.SetParameterValue("p11", i)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmSalesReport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
