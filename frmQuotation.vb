Imports System.Data.SqlClient
Imports System.IO

Public Class frmQuotation
    Dim st2 As String

    Sub Reset()
        txtCID.Text = ""
        txtRemarks.Text = ""
        txtCustomerName.Text = ""
        txtAmount.Text = ""
        txtSellingPrice.Text = ""
        txtCustomerID.Text = ""
        txtDiscountAmount.Text = ""
        txtDiscountPer.Text = ""
        txtQuotationNo.Text = ""
        txtProductCode.Text = ""
        txtProductName.Text = ""
        txtQty.Text = ""
        txtSellingPrice.Text = ""
        txtTotalAmount.Text = ""
        txtTotalQty.Text = ""
        txtVAT.Text = ""
        txtVATAmount.Text = ""
        txtGrandTotal.Text = ""
        dtpQuotationDate.Text = Today
        btnDelete.Enabled = False
        btnUpdate.Enabled = False
        btnSave.Enabled = True
        btnRemove.Enabled = False
        btnAdd.Enabled = True
        btnPrint.Enabled = False
        txtContactNo.Text = ""
        txtCustomerType.Text = ""
        auto()
        lblSet.Text = "Allowed"
        DataGridView1.Rows.Clear()
        Clear()
    End Sub
    Private Function GenerateID() As String
        con = New SqlConnection(cs)
        Dim value As String = "0000"
        Try
            ' Fetch the latest ID from the database
            con.Open()
            cmd = New SqlCommand("SELECT TOP 1 Q_ID FROM Quotation order BY Q_ID DESC", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.HasRows Then
                rdr.Read()
                value = rdr.Item("Q_ID")
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
            txtQuotationNo.Text = "Q-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub


    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles btnSelectionInv.Click
        frmProductRecord.lblSet.Text = "Quotation"
        frmProductRecord.Reset()
        frmProductRecord.ShowDialog()
    End Sub
    Sub Compute()
        Dim num1, num2, num3, num4, num5 As Double
        num1 = CDbl(Val(txtQty.Text) * Val(txtSellingPrice.Text))
        num1 = Math.Round(num1, 2)
        txtAmount.Text = num1
        num2 = CDbl((Val(txtAmount.Text) * Val(txtDiscountPer.Text)) / 100)
        num2 = Math.Round(num2, 2)
        txtDiscountAmount.Text = num2
        num3 = Val(txtAmount.Text) - Val(txtDiscountAmount.Text)
        num4 = CDbl((Val(txtVAT.Text) * Val(num3)) / 100)
        num4 = Math.Round(num4, 2)
        txtVATAmount.Text = num4
        num5 = CDbl(Val(txtAmount.Text) + Val(txtVATAmount.Text) - Val(txtDiscountAmount.Text))
        num5 = Math.Round(num5, 2)
        txtTotalAmount.Text = num5
    End Sub

    Private Sub txtQty_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtQty.TextChanged
        Compute()
    End Sub

    Private Sub txtQty_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtQty.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub
    Public Function GrandTotal() As Double
        Dim sum As Double = 0
        Try
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(9).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return sum
    End Function
    Sub Print()
        Try
            If txtCustomerType.Text <> "Non Regular" Then
                Cursor = Cursors.WaitCursor
                Timer1.Enabled = True
                Dim rpt As New rptQuotation 'The report you created.
                Dim myConnection As SqlConnection
                Dim MyCommand, MyCommand1 As New SqlCommand()
                Dim myDA, myDA1 As New SqlDataAdapter()
                Dim myDS As New DataSet 'The DataSet you created.
                myConnection = New SqlConnection(cs)
                MyCommand.Connection = myConnection
                MyCommand1.Connection = myConnection
                MyCommand.CommandText = "SELECT Customer.ID, Customer.Name, Customer.Gender, Customer.Address, Customer.City, Customer.State, Customer.ZipCode, Customer.ContactNo, Customer.EmailID,Customer.Photo, Quotation.Q_ID, Quotation.QuotationNo, Quotation.Date, Quotation.GrandTotal, Quotation_Join.QJ_ID, Quotation_Join.QuotationID,Quotation_Join.ProductID, Quotation_Join.Cost, Quotation_Join.Qty, Quotation_Join.Amount, Quotation_Join.DiscountPer, Quotation_Join.Discount, Quotation_Join.VATPer, Quotation_Join.VAT,Quotation_Join.TotalAmount, Product.PID, Product.ProductCode, Product.ProductName FROM Customer INNER JOIN Quotation ON Customer.ID = Quotation.CustomerID INNER JOIN Quotation_Join ON Quotation.Q_ID = Quotation_Join.QuotationID INNER JOIN Product ON Quotation_Join.ProductID = Product.PID where QuotationNo=@d1"
                MyCommand.Parameters.AddWithValue("@d1", txtQuotationNo.Text)
                MyCommand1.CommandText = "SELECT * from Company"
                MyCommand.CommandType = CommandType.Text
                MyCommand1.CommandType = CommandType.Text
                myDA.SelectCommand = MyCommand
                myDA1.SelectCommand = MyCommand1
                myDA.Fill(myDS, "Quotation")
                myDA.Fill(myDS, "Quotation_Join")
                myDA.Fill(myDS, "Customer")
                myDA.Fill(myDS, "Product")
                myDA1.Fill(myDS, "Company")
                rpt.SetDataSource(myDS)
                rpt.SetParameterValue("p1", txtCustomerID.Text)
                rpt.SetParameterValue("p2", Today)
                frmReport.CrystalReportViewer1.ReportSource = rpt
                frmReport.ShowDialog()
            End If
            If txtCustomerType.Text = "Non Regular" Then
                Cursor = Cursors.WaitCursor
                Timer1.Enabled = True
                Dim rpt As New rptQuotation1 'The report you created.
                Dim myConnection As SqlConnection
                Dim MyCommand, MyCommand1 As New SqlCommand()
                Dim myDA, myDA1 As New SqlDataAdapter()
                Dim myDS As New DataSet 'The DataSet you created.
                myConnection = New SqlConnection(cs)
                MyCommand.Connection = myConnection
                MyCommand1.Connection = myConnection
                MyCommand.CommandText = "SELECT Customer.ID, Customer.Name, Customer.Gender, Customer.Address, Customer.City, Customer.State, Customer.ZipCode, Customer.ContactNo, Customer.EmailID,Customer.Photo, Quotation.Q_ID, Quotation.QuotationNo, Quotation.Date, Quotation.GrandTotal, Quotation_Join.QJ_ID, Quotation_Join.QuotationID,Quotation_Join.ProductID, Quotation_Join.Cost, Quotation_Join.Qty, Quotation_Join.Amount, Quotation_Join.DiscountPer, Quotation_Join.Discount, Quotation_Join.VATPer, Quotation_Join.VAT,Quotation_Join.TotalAmount, Product.PID, Product.ProductCode, Product.ProductName FROM Customer INNER JOIN Quotation ON Customer.ID = Quotation.CustomerID INNER JOIN Quotation_Join ON Quotation.Q_ID = Quotation_Join.QuotationID INNER JOIN Product ON Quotation_Join.ProductID = Product.PID where QuotationNo=@d1"
                MyCommand.Parameters.AddWithValue("@d1", txtQuotationNo.Text)
                MyCommand1.CommandText = "SELECT * from Company"
                MyCommand.CommandType = CommandType.Text
                MyCommand1.CommandType = CommandType.Text
                myDA.SelectCommand = MyCommand
                myDA1.SelectCommand = MyCommand1
                myDA.Fill(myDS, "Quotation")
                myDA.Fill(myDS, "Quotation_Join")
                myDA.Fill(myDS, "Customer")
                myDA.Fill(myDS, "Product")
                myDA1.Fill(myDS, "Company")
                rpt.SetDataSource(myDS)
                rpt.SetParameterValue("p1", txtCustomerID.Text)
                rpt.SetParameterValue("p2", Today)
                frmReport.CrystalReportViewer1.ReportSource = rpt
                frmReport.ShowDialog()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    Private Sub frmBilling_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd.Click
        Try
            If txtProductCode.Text = "" Then
                MessageBox.Show("Please retrieve product code", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtProductCode.Focus()
                Exit Sub
            End If
            If txtQty.Text = "" Then
                MessageBox.Show("Please enter quantity", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtQty.Focus()
                Exit Sub
            End If
            If txtQty.Text = 0 Then
                MessageBox.Show("Quantity can not be zero", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtQty.Focus()
                Exit Sub
            End If
            If DataGridView1.Rows.Count = 0 Then
                DataGridView1.Rows.Add(txtProductCode.Text, txtProductName.Text, txtSellingPrice.Text, txtQty.Text, txtAmount.Text, txtDiscountPer.Text, txtDiscountAmount.Text, txtVAT.Text, txtVATAmount.Text, txtTotalAmount.Text, txtProductID.Text)
                Dim k As Double = 0
                k = GrandTotal()
                k = Math.Round(k, 2)
                txtGrandTotal.Text = k
                Clear()
                Exit Sub
            End If
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                If r.Cells(0).Value = txtProductCode.Text Then
                    r.Cells(0).Value = txtProductCode.Text
                    r.Cells(1).Value = txtProductName.Text
                    r.Cells(2).Value = txtSellingPrice.Text
                    r.Cells(3).Value = Val(r.Cells(3).Value) + Val(txtQty.Text)
                    r.Cells(4).Value = Val(r.Cells(4).Value) + Val(txtAmount.Text)
                    r.Cells(5).Value = Val(txtDiscountPer.Text)
                    r.Cells(6).Value = Val(r.Cells(6).Value) + Val(txtDiscountAmount.Text)
                    r.Cells(7).Value = Val(txtVAT.Text)
                    r.Cells(8).Value = Val(r.Cells(8).Value) + Val(txtVATAmount.Text)
                    r.Cells(9).Value = Val(r.Cells(9).Value) + Val(txtTotalAmount.Text)
                    r.Cells(10).Value = txtProductID.Text
                    Dim i As Double = 0
                    i = GrandTotal()
                    i = Math.Round(i, 2)
                    txtGrandTotal.Text = i
                    Clear()
                    Exit Sub
                End If
            Next
            DataGridView1.Rows.Add(txtProductCode.Text, txtProductName.Text, txtSellingPrice.Text, txtQty.Text, txtAmount.Text, txtDiscountPer.Text, txtDiscountAmount.Text, txtVAT.Text, txtVATAmount.Text, txtTotalAmount.Text, txtProductID.Text)
            Dim j As Double = 0
            j = GrandTotal()
            j = Math.Round(j, 2)
            txtGrandTotal.Text = j
            Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub Clear()
        txtProductCode.Text = ""
        txtProductName.Text = ""
        txtSellingPrice.Text = ""
        txtQty.Text = ""
        txtAmount.Text = ""
        txtDiscountPer.Text = ""
        txtDiscountAmount.Text = ""
        txtVAT.Text = ""
        txtVATAmount.Text = ""
        txtTotalAmount.Text = ""
        btnAdd.Enabled = True
        btnRemove.Enabled = False
        btnListUpdate.Enabled = False
    End Sub

    Private Sub btnRemove_Click(sender As System.Object, e As System.EventArgs) Handles btnRemove.Click
        Try
            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                DataGridView1.Rows.Remove(row)
            Next
            Dim k As Double = 0
            k = GrandTotal()
            k = Math.Round(k, 2)
            txtGrandTotal.Text = k
            Compute()
            Clear()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DataGridView1_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick

        If (Me.DataGridView1.Rows.Count > 0) Then
            If lblSet.Text = "Not Allowed" Then
                btnRemove.Enabled = False
                btnListUpdate.Enabled = False
            Else
                btnRemove.Enabled = True
                btnListUpdate.Enabled = True
            End If
            Me.btnAdd.Enabled = False
            Dim row As DataGridViewRow = Me.DataGridView1.SelectedRows.Item(0)
            Me.txtProductCode.Text = (row.Cells.Item(0).Value)
            Me.txtProductName.Text = (row.Cells.Item(1).Value)
            Me.txtSellingPrice.Text = (row.Cells.Item(2).Value)
            Me.txtQty.Text = (row.Cells.Item(3).Value)
            Me.txtAmount.Text = (row.Cells.Item(4).Value)
            Me.txtDiscountPer.Text = (row.Cells.Item(5).Value)
            Me.txtDiscountAmount.Text = (row.Cells.Item(6).Value)
            Me.txtVAT.Text = (row.Cells.Item(7).Value)
            Me.txtVATAmount.Text = (row.Cells.Item(8).Value)
            Me.txtTotalAmount.Text = (row.Cells.Item(9).Value)
            Me.txtProductID.Text = (row.Cells.Item(10).Value)
        End If
    End Sub

    Private Sub DataGridView1_RowPostPaint(sender As Object, e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView1.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView1.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

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
            Dim cq As String = "delete from Quotation where Q_ID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                Dim st As String = "deleted the invoice no. '" & txtQuotationNo.Text & "'"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
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
    Private Function GenerateID1() As String
        con = New SqlConnection(cs)
        Dim value As String = "0000"
        Try
            ' Fetch the latest ID from the database
            con.Open()
            cmd = New SqlCommand("SELECT TOP 1 ID FROM Customer ORDER BY ID DESC", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.HasRows Then
                rdr.Read()
                value = rdr.Item("ID")
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
    Sub auto1()
        Try
            txtCID.Text = GenerateID1()
            txtCustomerID.Text = "C-" + GenerateID1()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtCustomerName.Text)) = 0 Then
            MessageBox.Show("Please retrieve customer details", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If DataGridView1.Rows.Count = 0 Then
            MessageBox.Show("sorry no product added to cart", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ctn1 As String = "select * from Company"
            cmd = New SqlCommand(ctn1)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()

            If Not rdr.Read() Then
                MessageBox.Show("Add company profile first in master entry", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If
            If txtCustomerName.ReadOnly = False Then
                auto1()
                con = New SqlConnection(cs)
                con.Open()
                Dim cbn As String = "insert into Customer(ID, CustomerID, [Name], Gender, Address, City, ContactNo, EmailID,Remarks,State,ZipCode,Photo,CustomerType) Values (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11,@d12,'Non Regular')"
                cmd = New SqlCommand(cbn)
                cmd.Parameters.AddWithValue("@d1", txtCID.Text)
                cmd.Parameters.AddWithValue("@d2", txtCustomerID.Text)
                cmd.Parameters.AddWithValue("@d3", txtCustomerName.Text)
                cmd.Parameters.AddWithValue("@d4", "")
                cmd.Parameters.AddWithValue("@d5", "")
                cmd.Parameters.AddWithValue("@d6", "")
                cmd.Parameters.AddWithValue("@d7", txtContactNo.Text)
                cmd.Parameters.AddWithValue("@d8", "")
                cmd.Parameters.AddWithValue("@d9", "")
                cmd.Parameters.AddWithValue("@d10", "")
                cmd.Parameters.AddWithValue("@d11", "")
                cmd.Connection = con
                Dim ms As New MemoryStream()
                Dim bmpImage As New Bitmap(My.Resources.photo)
                bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
                Dim data As Byte() = ms.GetBuffer()
                Dim p As New SqlParameter("@d12", SqlDbType.Image)
                p.Value = data
                cmd.Parameters.Add(p)
                cmd.ExecuteNonQuery()
                con.Close()
                txtCustomerType.Text = "Non Regular"
            End If
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into Quotation(Q_ID, QuotationNo, Date, CustomerID, GrandTotal, Remarks) Values (@d1,@d2,@d3,@d4,@d5,@d6)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", txtQuotationNo.Text)
            cmd.Parameters.AddWithValue("@d3", dtpQuotationDate.Value.Date)
            cmd.Parameters.AddWithValue("@d4", txtCID.Text)
            cmd.Parameters.AddWithValue("@d5", txtGrandTotal.Text)
            cmd.Parameters.AddWithValue("@d6", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "insert into Quotation_Join(QuotationID, Cost, Qty,Amount, DiscountPer, Discount, VATPer, VAT, TotalAmount,ProductID) VALUES (" & txtID.Text & " ,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11,@d12)"
            cmd = New SqlCommand(cb1)
            cmd.Connection = con
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    cmd.Parameters.AddWithValue("@d4", row.Cells(2).Value)
                    cmd.Parameters.AddWithValue("@d5", row.Cells(3).Value)
                    cmd.Parameters.AddWithValue("@d6", row.Cells(4).Value)
                    cmd.Parameters.AddWithValue("@d7", row.Cells(5).Value)
                    cmd.Parameters.AddWithValue("@d8", row.Cells(6).Value)
                    cmd.Parameters.AddWithValue("@d9", row.Cells(7).Value)
                    cmd.Parameters.AddWithValue("@d10", row.Cells(8).Value)
                    cmd.Parameters.AddWithValue("@d11", row.Cells(9).Value)
                    cmd.Parameters.AddWithValue("@d12", row.Cells(10).Value)
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            Dim st As String = "added the new quotation having quotation no. '" & txtQuotationNo.Text & "'"
            LogFunc(lblUser.Text, st)
            btnSave.Enabled = False
            If CheckForInternetConnection() = True Then
                con = New SqlConnection(cs)
                con.Open()
                Dim ctn As String = "select RTRIM(APIURL) from SMSSetting where IsDefault='Yes' and IsEnabled='Yes'"
                cmd = New SqlCommand(ctn)
                cmd.Connection = con
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    st2 = rdr.GetValue(0)
                    Dim st3 As String = "Hello, " & txtCustomerName.Text & " you have successfully applied for quotation having quotation no. " & txtQuotationNo.Text & ""
                    SMSFunc(txtContactNo.Text, st3, st2)
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                End If
            End If
            con.Close()
            RefreshRecords()
            MessageBox.Show("Successfully saved", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Print()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtCustomerName.Text)) = 0 Then
            MessageBox.Show("Please retrieve customer details", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If DataGridView1.Rows.Count = 0 Then
            MessageBox.Show("sorry no product added to cart", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update Quotation set QuotationNo=@d2, Date=@d3, CustomerID=@d4, GrandTotal=@d5, Remarks=@d6 where Q_ID=@d1"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", txtQuotationNo.Text)
            cmd.Parameters.AddWithValue("@d3", dtpQuotationDate.Value.Date)
            cmd.Parameters.AddWithValue("@d4", txtCID.Text)
            cmd.Parameters.AddWithValue("@d5", txtGrandTotal.Text)
            cmd.Parameters.AddWithValue("@d6", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            Dim st As String = "updated the quotation having invoice no. '" & txtQuotationNo.Text & "'"
            LogFunc(lblUser.Text, st)
            btnUpdate.Enabled = False
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmQuotationRecord.Reset()
        frmQuotationRecord.ShowDialog()
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

    Private Sub btnListReset_Click(sender As System.Object, e As System.EventArgs) Handles btnListReset.Click
        Clear()
    End Sub

    Private Sub btnListUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnListUpdate.Click
        Try
            If txtProductCode.Text = "" Then
                MessageBox.Show("Please retrieve product code", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtProductCode.Focus()
                Exit Sub
            End If
            If txtQty.Text = "" Then
                MessageBox.Show("Please enter quantity", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtQty.Focus()
                Exit Sub
            End If
            If txtQty.Text = 0 Then
                MessageBox.Show("Quantity can not be zero", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtQty.Focus()
                Exit Sub
            End If

            For Each row As DataGridViewRow In DataGridView1.SelectedRows
                DataGridView1.Rows.Remove(row)
            Next
            DataGridView1.Rows.Add(txtProductCode.Text, txtProductName.Text, txtSellingPrice.Text, txtQty.Text, txtAmount.Text, txtDiscountPer.Text, txtDiscountAmount.Text, txtVAT.Text, txtVATAmount.Text, txtTotalAmount.Text, txtProductID.Text)
            Dim k As Double = 0
            k = GrandTotal()
            k = Math.Round(k, 2)
            txtGrandTotal.Text = k
            Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnSelect_Click_1(sender As System.Object, e As System.EventArgs) Handles btnSelect.Click
        frmCustomerRecord2.lblSet.Text = "Quotation"
        frmCustomerRecord2.lblUser.Text = lblUser.Text
        frmCustomerRecord2.Reset()
        frmCustomerRecord2.ShowDialog()
    End Sub
End Class
