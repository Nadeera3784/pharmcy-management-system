Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO

Public Class frmPurchaseReport
    Dim a, b, c As Decimal

    Sub Reset()
        cmbSupplier.Text = ""
        dtpDateFrom.Value = Today
        dtpDateTo.Value = Today
    End Sub
    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnViewReport_Click(sender As System.Object, e As System.EventArgs) Handles btnViewReport.Click
        Try
            If Len(Trim(cmbSupplier.Text)) = 0 Then
                MessageBox.Show("Please select supplier", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbSupplier.Focus()
                Exit Sub
            End If
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptPurchase 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT Distinct Stock.ST_ID, Stock.Stock_ID, Stock.Date, Stock.SupplierID, Stock.GrandTotal, Stock.TotalPayment, Stock.PaymentDue, Stock.Remarks, Stock_Product.SP_ID, Stock_Product.StockID, Stock_Product.ProductID,Stock_Product.Qty, Stock_Product.Price, Stock_Product.TotalAmount, Supplier.ID, Supplier.SupplierID AS Expr1, Supplier.Name, Supplier.Address, Supplier.City, Supplier.State, Supplier.ZipCode,Supplier.ContactNo, Supplier.EmailID, Supplier.Remarks AS Expr2, Product.PID, Product.ProductCode, Product.ProductName, Product.SubCategoryID, Product.Description, Product.CostPrice, Product.SellingPrice,Product.Discount, Product.VAT, Product.ReorderPoint FROM Stock INNER JOIN Stock_Product ON Stock.ST_ID = Stock_Product.StockID INNER JOIN Supplier ON Stock.SupplierID = Supplier.ID INNER JOIN Product ON Stock_Product.ProductID = Product.PID where Supplier.Name=@d1"
            MyCommand.Parameters.AddWithValue("@d1", cmbSupplier.Text)
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Stock")
            myDA.Fill(myDS, "Stock_Product")
            myDA.Fill(myDS, "Product")
            myDA.Fill(myDS, "Supplier")
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select ISNULL(sum(GrandTotal),0),ISNULL(sum(TotalPayment),0),ISNULL(sum(PaymentDue),0) from Stock,Supplier where Supplier.ID=Stock.SupplierID and Name=@d1"
            cmd = New SqlCommand(ct)
            cmd.Parameters.AddWithValue("@d1", cmbSupplier.Text)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            While rdr.Read()
                a = rdr.GetValue(0)
                b = rdr.GetValue(1)
                c = rdr.GetValue(2)
            End While
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT CONVERT(varchar(10),YEAR(Date)) AS Year, SUM(GrandTotal) AS GrandTotal from Stock,Supplier where Supplier.ID=Stock.SupplierID and Name=@d1 GROUP BY YEAR(Date) ORDER BY Year", con)
            cmd.Parameters.AddWithValue("@d1", cmbSupplier.Text)
            adp = New SqlDataAdapter(cmd)
            dtable = New DataTable()
            adp.Fill(dtable)
            con.Close()
            myDS.Tables.Add(dtable)
            myDS.WriteXmlSchema("TotalPurchase.xml")
            rpt.Subreports(0).SetDataSource(myDS)
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", dtpDateFrom.Value.Date)
            rpt.SetParameterValue("p2", dtpDateTo.Value.Date)
            rpt.SetParameterValue("p3", a)
            rpt.SetParameterValue("p4", b)
            rpt.SetParameterValue("p5", c)
            rpt.SetParameterValue("p6", Today)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptPurchase 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand As New SqlCommand()
            Dim myDA As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand.CommandText = "SELECT Distinct Stock.ST_ID, Stock.Stock_ID, Stock.Date, Stock.SupplierID, Stock.GrandTotal, Stock.TotalPayment, Stock.PaymentDue, Stock.Remarks, Stock_Product.SP_ID, Stock_Product.StockID, Stock_Product.ProductID,Stock_Product.Qty, Stock_Product.Price, Stock_Product.TotalAmount, Supplier.ID, Supplier.SupplierID AS Expr1, Supplier.Name, Supplier.Address, Supplier.City, Supplier.State, Supplier.ZipCode,Supplier.ContactNo, Supplier.EmailID, Supplier.Remarks AS Expr2, Product.PID, Product.ProductCode, Product.ProductName, Product.SubCategoryID, Product.Description, Product.CostPrice, Product.SellingPrice,Product.Discount, Product.VAT, Product.ReorderPoint FROM Stock INNER JOIN Stock_Product ON Stock.ST_ID = Stock_Product.StockID INNER JOIN Supplier ON Stock.SupplierID = Supplier.ID INNER JOIN Product ON Stock_Product.ProductID = Product.PID where Stock.Date between @d1 and @d2 order by Stock.Date"
            MyCommand.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            MyCommand.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            MyCommand.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA.Fill(myDS, "Stock")
            myDA.Fill(myDS, "Stock_Product")
            myDA.Fill(myDS, "Product")
            myDA.Fill(myDS, "Supplier")
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select ISNULL(sum(GrandTotal),0),ISNULL(sum(TotalPayment),0),ISNULL(sum(PaymentDue),0) from Stock,Supplier where Supplier.ID=Stock.SupplierID and Date between @d3 and @d4"
            cmd = New SqlCommand(ct)
            cmd.Parameters.Add("@d3", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d4", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            While rdr.Read()
                a = rdr.GetValue(0)
                b = rdr.GetValue(1)
                c = rdr.GetValue(2)
            End While
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT CONVERT(varchar(10),YEAR(Date)) AS Year, SUM(GrandTotal) AS GrandTotal FROM Stock where date between @d3 and @d4 GROUP BY YEAR(Date) ORDER BY Year", con)
            cmd.Parameters.Add("@d3", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            cmd.Parameters.Add("@d4", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            adp = New SqlDataAdapter(cmd)
            dtable = New DataTable()
            adp.Fill(dtable)
            con.Close()
            myDS.Tables.Add(dtable)
            myDS.WriteXmlSchema("TotalPurchase.xml")
            rpt.Subreports(0).SetDataSource(myDS)
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", dtpDateFrom.Value.Date)
            rpt.SetParameterValue("p2", dtpDateTo.Value.Date)
            rpt.SetParameterValue("p3", a)
            rpt.SetParameterValue("p4", b)
            rpt.SetParameterValue("p5", c)
            rpt.SetParameterValue("p6", Today)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Sub fillSupplier()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT RTRIM(Name) FROM Supplier order by Name", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbSupplier.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbSupplier.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub frmPurchaseReport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillSupplier()
    End Sub

    Private Sub cmbCompany_Format(sender As System.Object, e As System.Windows.Forms.ListControlConvertEventArgs) Handles cmbSupplier.Format
        If (e.DesiredType Is GetType(String)) Then
            e.Value = e.Value.ToString.Trim
        End If
    End Sub
End Class
