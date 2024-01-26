Imports System.Data.SqlClient
Imports Excel = Microsoft.Office.Interop.Excel

Public Class frmServiceDoneReport

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


    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptService 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand, MyCommand1 As New SqlCommand()
            Dim myDA, myDA1 As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "SELECT * FROM Service INNER JOIN Customer ON Service.CustomerID = Customer.ID INNER JOIN InvoiceInfo1 ON Service.S_ID = InvoiceInfo1.ServiceID INNER JOIN Invoice1_Product ON InvoiceInfo1.Inv_ID = Invoice1_Product.InvoiceID where InvoiceInfo1.InvoiceDate between @d1 and @d2 order by invoiceDate"
            MyCommand.Parameters.Add("@d1", SqlDbType.DateTime, 30, "Date").Value = dtpDateFrom.Value.Date
            MyCommand.Parameters.Add("@d2", SqlDbType.DateTime, 30, "Date").Value = dtpDateTo.Value.Date
            MyCommand1.CommandText = "SELECT * from Company"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "InvoiceInfo1")
            myDA.Fill(myDS, "Invoice1_Product")
            myDA.Fill(myDS, "Service")
            myDA.Fill(myDS, "Customer")
            myDA1.Fill(myDS, "Company")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", dtpDateFrom.Value.Date)
            rpt.SetParameterValue("p2", dtpDateTo.Value.Date)
            rpt.SetParameterValue("p3", Today)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmServiceDoneReport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
