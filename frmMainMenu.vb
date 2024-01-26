Imports System.Data.SqlClient
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.SqlServer.Management.Smo
Imports System.Globalization

Public Class frmMainMenu
    Dim Filename As String


    Private Sub AboutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        frmAbout.ShowDialog()
    End Sub
    Sub Backup()
        Try
            Dim dt As DateTime = Today
            Dim destdir As String = "SIS_DB " & System.DateTime.Now.ToString("dd-MM-yyyy_h-mm-ss") & ".bak"
            Dim objdlg As New SaveFileDialog
            objdlg.FileName = destdir
            objdlg.ShowDialog()
            Filename = objdlg.FileName
            Cursor = Cursors.WaitCursor
            Timer2.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "backup database SIS_DB to disk='" & Filename & "'with init,stats=10"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub BackupToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BackupToolStripMenuItem.Click
        Backup()
    End Sub

    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Cursor = Cursors.Default
        Timer2.Enabled = False
    End Sub

    Private Sub RestoreToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RestoreToolStripMenuItem.Click
        Try
            With OpenFileDialog1
                .Filter = ("DB Backup File|*.bak;")
                .FilterIndex = 4
            End With
            'Clear the file name
            OpenFileDialog1.FileName = ""

            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                Cursor = Cursors.WaitCursor
                Timer2.Enabled = True
                SqlConnection.ClearAllPools()
                con = New SqlConnection(cs)
                con.Open()
                Dim cb As String = "USE Master ALTER DATABASE SIS_DB SET Single_User WITH Rollback Immediate Restore database SIS_DB FROM disk='" & OpenFileDialog1.FileName & "' WITH REPLACE ALTER DATABASE SIS_DB SET Multi_User "
                cmd = New SqlCommand(cb)
                cmd.Connection = con
                cmd.ExecuteReader()
                con.Close()
                Dim st As String = "Sucessfully performed the restore"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Successfully performed", "Database Restore", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RegistrationToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles RegistrationToolStripMenuItem.Click
        frmRegistration.lblUser.Text = lblUser.Text
        frmRegistration.Reset()
        frmRegistration.ShowDialog()
    End Sub

    Private Sub LogsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LogsToolStripMenuItem.Click
        frmLogs.Reset()
        frmLogs.lblUser.Text = lblUser.Text
        frmLogs.ShowDialog()
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Dim dt As DateTime = Today
        lblDateTime.Text = dt.ToString("dd/MM/yyyy")
        lblTime.Text = TimeOfDay.ToString("h:mm:ss tt")
    End Sub

    Private Sub CalculatorToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CalculatorToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("Calc.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub NotepadToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles NotepadToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("Notepad.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub WordpadToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles WordpadToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("wordpad.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MSWordToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MSWordToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("winword.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub TaskManagerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TaskManagerToolStripMenuItem.Click
        Try
            System.Diagnostics.Process.Start("TaskMgr.exe")
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SystemInfoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SystemInfoToolStripMenuItem.Click
        frmSystemInfo.ShowDialog()
    End Sub
    Sub LogOut()
        frmStock.Hide()
        frmProduct.Hide()
        Dim st As String = "Successfully logged out"
        LogFunc(lblUser.Text, st)
        Me.Hide()
        frmLogin.Show()
        frmLogin.UserID.Text = ""
        frmLogin.Password.Text = ""
        frmLogin.UserID.Focus()
    End Sub
    Private Sub LogoutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LogoutToolStripMenuItem.Click
        Try
            If MessageBox.Show("Do you really want to logout from application?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                If MessageBox.Show("Do you want backup database before logout?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    Backup()
                    LogOut()
                Else
                    LogOut()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmMainMenu_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        e.Cancel = True
    End Sub

    Private Sub CompanyInfoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CompanyInfoToolStripMenuItem.Click
        frmCompany.lblUser.Text = lblUser.Text
        frmCompany.Reset()
        frmCompany.ShowDialog()
    End Sub

    Private Sub CustomerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CustomerToolStripMenuItem.Click
        frmCustomer.lblUser.Text = lblUser.Text
        frmCustomer.Reset()
        frmCustomer.ShowDialog()
    End Sub

    Private Sub CategoryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CategoryToolStripMenuItem.Click
        frmCategory.lblUser.Text = lblUser.Text
        frmCategory.Reset()
        frmCategory.ShowDialog()
    End Sub

    Private Sub SubCategoryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SubCategoryToolStripMenuItem.Click
        frmSubCategory.lblUser.Text = lblUser.Text
        frmSubCategory.Reset()
        frmSubCategory.ShowDialog()
    End Sub

    Private Sub SupplierToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SupplierToolStripMenuItem.Click
        frmSupplier.lblUser.Text = lblUser.Text
        frmSupplier.Reset()
        frmSupplier.ShowDialog()
    End Sub

    Private Sub CustomerToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles CustomerToolStripMenuItem1.Click
        frmCustomerRecord1.Reset()
        frmCustomerRecord1.ShowDialog()
    End Sub

    Private Sub SupplierToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles SupplierToolStripMenuItem1.Click
        frmSupplierRecord.Reset()
        frmSupplierRecord.ShowDialog()
    End Sub


    Private Sub StockToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StockToolStripMenuItem.Click
        frmStock.lblUser.Text = lblUser.Text
        frmStock.lblUserType.Text = lblUserType.Text
        frmStock.Reset()
        frmStock.ShowDialog()
    End Sub

    Public Sub Getdata()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = New SqlCommand("SELECT RTRIM(Product.ProductCode),RTRIM(ProductName),CostPrice,SellingPrice,Discount,VAT,Qty from Temp_Stock,Product where Product.PID=Temp_Stock.ProductID and Qty > 0 order by ProductName", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView1.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView1.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6))
            End While
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                con = New SqlConnection(cs)
                con.Open()
                Dim ct As String = "select ReorderPoint from Product where ProductCode=@d1"
                cmd = New SqlCommand(ct)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", r.Cells(0).Value.ToString())
                rdr = cmd.ExecuteReader()
                If (rdr.Read()) Then
                    Dim i As Integer
                    i = rdr(0)
                    If r.Cells(6).Value < i Then
                        r.DefaultCellStyle.BackColor = Color.Red
                    End If
                End If
            Next
            con.Close()
            DataGridView1.ClearSelection()

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Reset()
        txtProductName.Text = ""
        Getdata()
    End Sub
    Private Function HandleRegistry() As Boolean
        Dim firstRunDate As Date
        Dim st As Date = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\InventorySoft2", "Set", Nothing)
        firstRunDate = st
        If firstRunDate = Nothing Then
            firstRunDate = System.DateTime.Today.Date
            My.Computer.Registry.SetValue("HKEY_LOCAL_MACHINE\SOFTWARE\InventorySoft2", "Set", firstRunDate)
        ElseIf (Now - firstRunDate).Days > 7 Then
            Return False
        End If
        Return True
    End Function
    Private Sub frmMainMenu_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim result As Boolean = HandleRegistry()

        If result = False Then 'something went wrong
            MessageBox.Show("Trial expired" & vbCrLf & "for purchasing the full version of software call us at +919827858191", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End
        End If
        If lblUserType.Text = "Admin" Then
            MasterEntryToolStripMenuItem.Enabled = True
            RegistrationToolStripMenuItem.Enabled = True
            LogsToolStripMenuItem.Enabled = True
            DatabaseToolStripMenuItem.Enabled = True
            CustomerToolStripMenuItem.Enabled = True
            SupplierToolStripMenuItem.Enabled = True
            ProductToolStripMenuItem.Enabled = True
            StockToolStripMenuItem.Enabled = True
            ServiceToolStripMenuItem.Enabled = True
            StockInToolStripMenuItem.Enabled = True
            BillingToolStripMenuItem.Enabled = True
            QuotationToolStripMenuItem.Enabled = True
            RecordToolStripMenuItem.Enabled = True
            ReportsToolStripMenuItem.Enabled = True
            VoucherToolStripMenuItem.Enabled = True
        End If
        If lblUserType.Text = "Sales Person" Then
            MasterEntryToolStripMenuItem.Enabled = False
            RegistrationToolStripMenuItem.Enabled = False
            LogsToolStripMenuItem.Enabled = False
            DatabaseToolStripMenuItem.Enabled = False
            CustomerToolStripMenuItem.Enabled = True
            SupplierToolStripMenuItem.Enabled = False
            ProductToolStripMenuItem.Enabled = False
            StockToolStripMenuItem.Enabled = False
            ServiceToolStripMenuItem.Enabled = True
            StockInToolStripMenuItem.Enabled = True
            BillingToolStripMenuItem.Enabled = True
            QuotationToolStripMenuItem.Enabled = True
            RecordToolStripMenuItem.Enabled = False
            ReportsToolStripMenuItem.Enabled = False
            VoucherToolStripMenuItem.Enabled = False
        End If
        If lblUserType.Text = "Inventory Manager" Then
            MasterEntryToolStripMenuItem.Enabled = False
            RegistrationToolStripMenuItem.Enabled = False
            LogsToolStripMenuItem.Enabled = False
            DatabaseToolStripMenuItem.Enabled = False
            CustomerToolStripMenuItem.Enabled = False
            SupplierToolStripMenuItem.Enabled = False
            ProductToolStripMenuItem.Enabled = True
            StockToolStripMenuItem.Enabled = True
            ServiceToolStripMenuItem.Enabled = False
            StockInToolStripMenuItem.Enabled = True
            BillingToolStripMenuItem.Enabled = False
            QuotationToolStripMenuItem.Enabled = False
            RecordToolStripMenuItem.Enabled = False
            ReportsToolStripMenuItem.Enabled = False
            VoucherToolStripMenuItem.Enabled = False
        End If
        Getdata()
        DataGridView1.ClearSelection()
        DataGridView1.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
        DataGridView1.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

    End Sub

    Private Sub StockToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles StockToolStripMenuItem1.Click
        frmStockRecord1.Reset()
        frmStockRecord1.ShowDialog()
    End Sub

    Private Sub StockInToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StockInToolStripMenuItem.Click
        frmCurrentStock.Reset()
        frmCurrentStock.ShowDialog()
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

            rowsTotal = DataGridView1.RowCount
            colsTotal = DataGridView1.Columns.Count - 1
            With excelWorksheet
                .Cells.Select()
                .Cells.Delete()
                For iC = 0 To colsTotal
                    .Cells(1, iC + 1).Value = DataGridView1.Columns(iC).HeaderText
                Next
                For I = 0 To rowsTotal - 1
                    For j = 0 To colsTotal
                        .Cells(I + 2, j + 1).value = DataGridView1.Rows(I).Cells(j).Value
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
            cmd = New SqlCommand("SELECT RTRIM(Product.ProductCode),RTRIM(ProductName),CostPrice,SellingPrice,Discount,VAT,Qty from Temp_Stock,Product where Product.PID=Temp_Stock.ProductID and qty > 0 and ProductName like '%" & txtProductName.Text & "%' order by ProductName", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            DataGridView1.Rows.Clear()
            While (rdr.Read() = True)
                DataGridView1.Rows.Add(rdr(0), rdr(1), rdr(2), rdr(3), rdr(4), rdr(5), rdr(6))
            End While
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                con = New SqlConnection(cs)
                con.Open()
                Dim ct As String = "select ReorderPoint from Product where ProductCode=@d1"
                cmd = New SqlCommand(ct)
                cmd.Connection = con
                cmd.Parameters.AddWithValue("@d1", r.Cells(0).Value.ToString())
                rdr = cmd.ExecuteReader()
                If (rdr.Read()) Then
                    Dim i As Integer
                    i = rdr(0)
                    If r.Cells(6).Value < i Then
                        r.DefaultCellStyle.BackColor = Color.Red
                    End If
                End If
            Next
            con.Close()
            DataGridView1.ClearSelection()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnReset_Click(sender As System.Object, e As System.EventArgs) Handles btnReset.Click
        Reset()
    End Sub

    Private Sub ContactsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ContactsToolStripMenuItem.Click
        frmContacts.lblUser.Text = lblUser.Text
        frmContacts.Reset()
        frmContacts.ShowDialog()
    End Sub

    Private Sub IndividualToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs)
        frmProductRecord.Reset()
        frmProductRecord.ShowDialog()
    End Sub

    Private Sub ProductToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ProductToolStripMenuItem.Click
        frmProduct.lblUser.Text = lblUser.Text
        frmProduct.lblUserType.Text = lblUserType.Text
        frmProduct.Reset()
        frmProduct.ShowDialog()
    End Sub

    Private Sub ProductToolStripMenuItem2_Click(sender As System.Object, e As System.EventArgs) Handles ProductToolStripMenuItem2.Click
        frmProductRecord.Reset()
        frmProductRecord.ShowDialog()
    End Sub

    Private Sub ServiceToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ServiceToolStripMenuItem.Click
        frmServices.lblUser.Text = lblUser.Text
        frmServices.lblUserType.Text = lblUserType.Text
        frmServices.Reset()
        frmServices.ShowDialog()
    End Sub

    Private Sub ServiceToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles ServiceToolStripMenuItem1.Click
        frmServicesRecord.Reset()
        frmServicesRecord.ShowDialog()
    End Sub

    Private Sub QuotationToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles QuotationToolStripMenuItem.Click
        frmQuotation.lblUser.Text = lblUser.Text
        frmQuotation.lblUserType.Text = lblUserType.Text
        frmQuotation.Reset()
        frmQuotation.ShowDialog()
    End Sub

    Private Sub QuotationToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles QuotationToolStripMenuItem1.Click
        frmQuotationRecord1.Reset()
        frmQuotationRecord1.ShowDialog()
    End Sub

    Private Sub ProductsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ProductsToolStripMenuItem.Click
        frmBilling.lblUser.Text = lblUser.Text
        frmBilling.lblUserType.Text = lblUserType.Text
        frmBilling.Reset()
        frmBilling.ShowDialog()
    End Sub

    Private Sub ProductsRepairToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ProductsRepairToolStripMenuItem.Click
        frmBilling1.lblUser.Text = lblUser.Text
        frmBilling1.lblUserType.Text = lblUserType.Text
        frmBilling1.Reset()
        frmBilling1.ShowDialog()
    End Sub

    Private Sub BillingProductsServiceToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BillingProductsServiceToolStripMenuItem.Click
        frmBillingRecord1.Reset()
        frmBillingRecord1.ShowDialog()
    End Sub

    Private Sub SMSSettingToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SMSSettingToolStripMenuItem.Click
        frmSMSSetting.Reset()
        frmSMSSetting.ShowDialog()
    End Sub

    Private Sub SalesToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SalesToolStripMenuItem.Click
        frmSalesReport.Reset()
        frmSalesReport.ShowDialog()
    End Sub

    Private Sub ServiceBillingToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ServiceBillingToolStripMenuItem.Click
        frmServiceDoneReport.Reset()
        frmServiceDoneReport.ShowDialog()
    End Sub

    Private Sub StockInAndStockOutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StockInAndStockOutToolStripMenuItem.Click
        frmStockInAndOutReport.ShowDialog()
    End Sub

    Private Sub PurchaseToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PurchaseToolStripMenuItem.Click
        frmPurchaseReport.Reset()
        frmPurchaseReport.ShowDialog()
    End Sub

    Private Sub ProfitAndLossToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ProfitAndLossToolStripMenuItem.Click
        frmProfitAndLossReport.Reset()
        frmProfitAndLossReport.ShowDialog()
    End Sub

    Private Sub VoucherToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles VoucherToolStripMenuItem.Click
        frmVoucher.Reset()
        frmVoucher.lblUser.Text = lblUser.Text
        frmVoucher.ShowDialog()
    End Sub

    Private Sub ExpenditureToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExpenditureToolStripMenuItem.Click
        frmVoucherReport.Reset()
        frmVoucherReport.ShowDialog()
    End Sub

    Private Sub CreditorsAndDebtorsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CreditorsAndDebtorsToolStripMenuItem.Click
        frmCreditorsAndDebtorsReport.ShowDialog()
    End Sub

    Private Sub OverallToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles OverallToolStripMenuItem.Click
        frmOverallReport.Reset()
        frmOverallReport.ShowDialog()
    End Sub
End Class