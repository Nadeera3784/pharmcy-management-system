Imports System.Data.SqlClient

Public Class frmServices

    Dim st2 As String
    Sub Reset()
        txtChargesQuote.Text = ""
        txtCID.Text = ""
        txtCustomerID.Text = ""
        txtCustomerName.Text = ""
        txtItemsDescription.Text = ""
        txtProblemDescription.Text = ""
        txtRemarks.Text = ""
        txtUpfront.Text = ""
        cmbServiceType.Text = ""
        cmbStatus.SelectedIndex = 1
        txtContactNo.Text = ""
        dtpServiceCreationDate.Text = Today
        dtpEstimatedRepairDate.Text = Today
        btnPrint.Enabled = False
        btnDelete.Enabled = False
        btnUpdate.Enabled = False
        btnSave.Enabled = True
        auto()
    End Sub
    Private Function GenerateID() As String
        con = New SqlConnection(cs)
        Dim value As String = "0000"
        Try
            ' Fetch the latest ID from the database
            con.Open()
            cmd = New SqlCommand("SELECT TOP 1 S_ID FROM Service ORDER BY S_ID DESC", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.HasRows Then
                rdr.Read()
                value = rdr.Item("S_ID")
            End If
            rdr.Close()
            ' Increase the ID by 1
            value += 1
            ' Because incrementing a string with an integer removes 0's
            ' we need to replace them. If necessary.
            If value <= 9 Then 'Value is between 0 and 10
                value = "000" & value
            ElseIf value <= 99 Then 'Value is between 9 and 100
                value = "00" & value
            ElseIf value <= 999 Then 'Value is between 999 and 1000
                value = "0" & value
            End If
        Catch ex As Exception
            ' If an error occurs, check the connection state and close it if necessary.
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
            value = "0000"
        End Try
        Return value
    End Function
    Sub auto()
        Try
            txtID.Text = GenerateID()
            txtServiceCode.Text = "SC-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub btnSelect_Click(sender As System.Object, e As System.EventArgs) Handles btnSelect.Click
        frmCustomerRecord2.lblSet.Text = "Services"
        frmCustomerRecord2.lblUser.Text = lblUser.Text
        frmCustomerRecord2.Reset()
        frmCustomerRecord2.ShowDialog()
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Sub Print()
        Try

            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptServiceReceipt 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand, MyCommand1 As New SqlCommand()
            Dim myDA, myDA1 As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "SELECT Service.S_ID, Service.ServiceCode, Service.ServiceType, Service.ServiceCreationDate, Service.ItemDescription, Service.ProblemDescription, Service.ChargesQuote,Service.AdvanceDeposit, Service.EstimatedRepairDate, Service.Remarks, Service.Status, Customer.ID, Customer.Name, Customer.Gender, Customer.Address, Customer.City,Customer.State, Customer.ZipCode, Customer.ContactNo, Customer.EmailID, Customer.Remarks AS Expr2, Customer.Photo FROM Service INNER JOIN Customer ON Service.CustomerID = Customer.ID where Service.ServiceCode=@d1"
            MyCommand.Parameters.AddWithValue("@d1", txtServiceCode.Text)
            MyCommand1.CommandText = "SELECT * from Company"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "Service")
            myDA.Fill(myDS, "Customer")
            myDA1.Fill(myDS, "Company")
            rpt.SetDataSource(myDS)
            rpt.SetParameterValue("p1", txtCustomerID.Text)
            rpt.SetParameterValue("p2", Today)
            frmReport.CrystalReportViewer1.ReportSource = rpt
            frmReport.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub


    Private Sub btnDelete_Click(sender As System.Object, e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub DeleteRecord()

        Try
            Dim RowsAffected As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim cl As String = "SELECT S_ID FROM Service INNER JOIN InvoiceInfo1 ON Service.S_ID = InvoiceInfo1.ServiceID where S_ID=@d1"
            cmd = New SqlCommand(cl)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Billing", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from Service where S_ID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                Dim st As String = "deleted the record having service code '" & txtServiceCode.Text & "'"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
                fillServiceType()
            Else
                MessageBox.Show("No Record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtItemsDescription.Text)) = 0 Then
            MessageBox.Show("Please enter items description", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtItemsDescription.Focus()
            Exit Sub
        End If
        If Len(Trim(txtChargesQuote.Text)) = 0 Then
            MessageBox.Show("Please enter charges quote", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtChargesQuote.Focus()
            Exit Sub
        End If
        If Len(Trim(txtUpfront.Text)) = 0 Then
            MessageBox.Show("Please enter upfront", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtUpfront.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCustomerName.Text)) = 0 Then
            MessageBox.Show("Please retrieve customer details", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ctn As String = "select * from Company"
            cmd = New SqlCommand(ctn)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()

            If Not rdr.Read() Then
                MessageBox.Show("Add company profile first in master entry", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into Service(S_ID, ServiceCode, CustomerID, ServiceType, ServiceCreationDate, ItemDescription, ProblemDescription, ChargesQuote, AdvanceDeposit, EstimatedRepairDate, Remarks, Status) Values (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11,@d12)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", txtServiceCode.Text)
            cmd.Parameters.AddWithValue("@d3", txtCID.Text)
            cmd.Parameters.AddWithValue("@d4", cmbServiceType.Text)
            cmd.Parameters.AddWithValue("@d5", dtpServiceCreationDate.Value.Date)
            cmd.Parameters.AddWithValue("@d6", txtItemsDescription.Text)
            cmd.Parameters.AddWithValue("@d7", txtProblemDescription.Text)
            cmd.Parameters.AddWithValue("@d8", txtChargesQuote.Text)
            cmd.Parameters.AddWithValue("@d9", txtUpfront.Text)
            cmd.Parameters.AddWithValue("@d10", dtpEstimatedRepairDate.Value.Date)
            cmd.Parameters.AddWithValue("@d11", txtRemarks.Text)
            cmd.Parameters.AddWithValue("@d12", cmbStatus.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "added the new service having service code '" & txtServiceCode.Text & "'"
            LogFunc(lblUser.Text, st)
            If CheckForInternetConnection() = True Then
                con = New SqlConnection(cs)
                con.Open()
                Dim ctn1 As String = "select RTRIM(APIURL) from SMSSetting where IsDefault='Yes' and IsEnabled='Yes'"
                cmd = New SqlCommand(ctn1)
                cmd.Connection = con
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    st2 = rdr.GetValue(0)
                    Dim st3 As String = "Hello, " & txtCustomerName.Text & " service has been created successfully having service code " & txtServiceCode.Text & ""
                    SMSFunc(txtContactNo.Text, st3, st2)
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                End If
            End If
            con.Close()
            btnSave.Enabled = False
            fillServiceType()
            MessageBox.Show("Successfully created", "Service", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Print()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtItemsDescription.Text)) = 0 Then
            MessageBox.Show("Please enter items description", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtItemsDescription.Focus()
            Exit Sub
        End If
        If Len(Trim(txtChargesQuote.Text)) = 0 Then
            MessageBox.Show("Please enter charges quote", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtChargesQuote.Focus()
            Exit Sub
        End If
        If Len(Trim(txtUpfront.Text)) = 0 Then
            MessageBox.Show("Please enter upfront", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtUpfront.Focus()
            Exit Sub
        End If
        If Len(Trim(txtCustomerName.Text)) = 0 Then
            MessageBox.Show("Please retrieve customer details", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update Service set ServiceCode=@d2, CustomerID=@d3, ServiceType=@d4, ServiceCreationDate=@d5, ItemDescription=@d6, ProblemDescription=@d7, ChargesQuote=@d8, AdvanceDeposit=@d9, EstimatedRepairDate=@d10, Remarks=@d11, Status=@d12 where S_ID=@d1"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", txtServiceCode.Text)
            cmd.Parameters.AddWithValue("@d3", txtCID.Text)
            cmd.Parameters.AddWithValue("@d4", cmbServiceType.Text)
            cmd.Parameters.AddWithValue("@d5", dtpServiceCreationDate.Value.Date)
            cmd.Parameters.AddWithValue("@d6", txtItemsDescription.Text)
            cmd.Parameters.AddWithValue("@d7", txtProblemDescription.Text)
            cmd.Parameters.AddWithValue("@d8", txtChargesQuote.Text)
            cmd.Parameters.AddWithValue("@d9", txtUpfront.Text)
            cmd.Parameters.AddWithValue("@d10", dtpEstimatedRepairDate.Value.Date)
            cmd.Parameters.AddWithValue("@d11", txtRemarks.Text)
            cmd.Parameters.AddWithValue("@d12", cmbStatus.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "updated the service having service code '" & txtServiceCode.Text & "'"
            LogFunc(lblUser.Text, st)
            btnUpdate.Enabled = False
            fillServiceType()
            MessageBox.Show("Successfully updated", "Service", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmServicesRecord.lblSet.Text = "Services"
        frmServicesRecord.Reset()
        frmServicesRecord.ShowDialog()
    End Sub

    Private Sub btnNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub


    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        Print()
    End Sub

    Private Sub txtChargesQuote_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtUpfront.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtUpfront.Text
            Dim selectionStart = Me.txtUpfront.SelectionStart
            Dim selectionLength = Me.txtUpfront.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an integer that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub
    Sub fillServiceType()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(ServiceType) FROM Service", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbServiceType.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbServiceType.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub txtChargesQuote_KeyPress_1(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtChargesQuote.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtChargesQuote.Text
            Dim selectionStart = Me.txtChargesQuote.SelectionStart
            Dim selectionLength = Me.txtChargesQuote.SelectionLength

            text = text.Substring(0, selectionStart) & keyChar & text.Substring(selectionStart + selectionLength)

            If Integer.TryParse(text, New Integer) AndAlso text.Length > 16 Then
                'Reject an integer that is longer than 16 digits.
                e.Handled = True
            ElseIf Double.TryParse(text, New Double) AndAlso text.IndexOf("."c) < text.Length - 3 Then
                'Reject a real number with two many decimal places.
                e.Handled = False
            End If
        Else
            'Reject all other characters.
            e.Handled = True
        End If
    End Sub

    Private Sub frmServices_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillServiceType()
    End Sub

    Private Sub cmbServiceType_Format(sender As System.Object, e As System.Windows.Forms.ListControlConvertEventArgs) Handles cmbServiceType.Format
        If (e.DesiredType Is GetType(String)) Then
            e.Value = e.Value.ToString.Trim
        End If
    End Sub
End Class
