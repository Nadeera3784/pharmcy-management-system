Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmCreditorsAndDebtorsReport

    Private Sub btnClose_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnCreditors.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT Supplier.SupplierID,Supplier.Name,Supplier.City,Supplier.ContactNo,Sum(Stock.PaymentDue) from Supplier,Stock where Supplier.ID=Stock.SupplierID group by Supplier.SupplierID,Supplier.Name,Supplier.City,Supplier.ContactNo having (Sum(PaymentDue) > 0) order by Supplier.Name", con)
            adp = New SqlDataAdapter(cmd)
            dtable = New DataTable()
            adp.Fill(dtable)
            con.Close()
            ds = New DataSet()
            ds.Tables.Add(dtable)
            ds.WriteXmlSchema("Creditors.xml")
            Dim rpt As New rptCreditors
            rpt.SetDataSource(ds)
            rpt.SetParameterValue("p1", Today)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles btnDebtors.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("Select CustomerID,Name,City,ContactNo,Sum(Balance) as [Balance] From (Select Customer.CustomerID,Customer.Name,Customer.City,Customer.ContactNo,Sum(Balance) as [Balance] from Customer,InvoiceInfo where Customer.ID=InvoiceInfo.CustomerID group by Customer.CustomerID,Customer.Name,Customer.City,Customer.ContactNo having (Sum(Balance) > 0)  Union All Select Customer.CustomerID,Customer.Name,Customer.City,Customer.ContactNo,Sum(Balance) as [Balance] from Customer,InvoiceInfo1,Service where Customer.ID=Service.CustomerID and Service.S_ID=InvoiceInfo1.ServiceID group by Customer.CustomerID,Customer.Name,Customer.City,Customer.ContactNo having (Sum(Balance) > 0)) As Customer Group By CustomerID,Name,City,ContactNo", con)
            adp = New SqlDataAdapter(cmd)
            dtable = New DataTable()
            adp.Fill(dtable)
            con.Close()
            ds = New DataSet()
            ds.Tables.Add(dtable)
            ds.WriteXmlSchema("Debtors.xml")
            Dim rpt As New rptDebtors
            rpt.SetDataSource(ds)
            rpt.SetParameterValue("p1", Today)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmStockInAndOutReport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
