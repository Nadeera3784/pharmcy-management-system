Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmProfitAndLossReport

    Dim a, b, c, d As Decimal
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
            con = New SqlConnection(cs)
            con.Open()
            Dim ctn As String = "select InvoiceNo from InvoiceInfo where InvoiceDate between @d1 and @d2"
            cmd = New SqlCommand(ctn)
            cmd.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            cmd.Connection = con
            rdr = cmd.ExecuteReader()

            If Not rdr.Read() Then
                MessageBox.Show("Sorry..No record found", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptSales1 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand, MyCommand1 As New SqlCommand()
            Dim myDA, myDA1 As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "SELECT Customer.ID, Customer.Name, Customer.Gender, Customer.Address, Customer.City, Customer.State, Customer.ZipCode, Customer.ContactNo, Customer.EmailID, Customer.Remarks,Customer.Photo, InvoiceInfo.Inv_ID, InvoiceInfo.InvoiceNo, InvoiceInfo.InvoiceDate, InvoiceInfo.CustomerID , InvoiceInfo.GrandTotal, InvoiceInfo.TotalPaid, InvoiceInfo.Balance, Invoice_Product.IPo_ID, Invoice_Product.InvoiceID, Invoice_Product.ProductID, Invoice_Product.CostPrice, Invoice_Product.SellingPrice, Invoice_Product.Margin,Invoice_Product.Qty, Invoice_Product.Amount, Invoice_Product.DiscountPer, Invoice_Product.Discount, Invoice_Product.VATPer, Invoice_Product.VAT, Invoice_Product.TotalAmount, Product.PID,Product.ProductCode, Product.ProductName FROM Customer INNER JOIN InvoiceInfo ON Customer.ID = InvoiceInfo.CustomerID INNER JOIN Invoice_Product ON InvoiceInfo.Inv_ID = Invoice_Product.InvoiceID INNER JOIN Product ON Invoice_Product.ProductID = Product.PID where InvoiceDate between @d1 and @d2 order by InvoiceDate"
            MyCommand.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            MyCommand.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            MyCommand1.CommandText = "SELECT * from Company"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "InvoiceInfo")
            myDA.Fill(myDS, "Invoice_Product")
            myDA.Fill(myDS, "Customer")
            myDA.Fill(myDS, "Product")
            myDA1.Fill(myDS, "Company")
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
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", dtpDateFrom.Value.Date)
            rpt.SetParameterValue("p2", dtpDateTo.Value.Date)
            rpt.SetParameterValue("p3", a)
            rpt.SetParameterValue("p4", b)
            rpt.SetParameterValue("p5", c)
            rpt.SetParameterValue("p6", d)
            rpt.SetParameterValue("p7", Today)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmSalesReport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
