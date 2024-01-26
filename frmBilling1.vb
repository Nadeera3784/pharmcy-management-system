Imports System.Data.SqlClient
Imports System.IO

Public Class frmBilling1
    Dim st2 As String

    Sub Reset()
        txtCID.Text = ""
        txtRemarks.Text = ""
        txtCustomerName.Text = ""
        txtAmount.Text = ""
        txtCostPrice.Text = ""
        txtCustomerID.Text = ""
        txtDiscountAmount.Text = ""
        txtDiscountPer.Text = ""
        txtMargin.Text = ""
        txtInvoiceNo.Text = ""
        txtProductCode.Text = ""
        txtProductName.Text = ""
        txtQty.Text = ""
        txtSellingPrice.Text = ""
        txtTotalAmount.Text = ""
        txtTotalQty.Text = ""
        txtVAT.Text = ""
        txtVATAmount.Text = ""
        txtGrandTotal.Text = ""
        txtTotalPayment.Text = ""
        txtPaymentDue.Text = ""
        txtServiceCode.Text = ""
        txtRepairCharges.Text = ""
        txtUpfront.Text = ""
        txtProductCharges.Text = 0
        txtServiceTaxPer.Text = ""
        txtServiceTaxAmount.Text = ""
        dtpInvoiceDate.Text = Today
        btnDelete.Enabled = False
        btnUpdate.Enabled = False
        btnSave.Enabled = True
        btnRemove.Enabled = False
        btnAdd.Enabled = True
        btnRemove1.Enabled = False
        btnAdd1.Enabled = True
        btnPrint.Enabled = False
        txtContactNo.Text = ""
        auto()
        lblSet.Text = "Allowed"
        DataGridView1.Rows.Clear()
        DataGridView2.Rows.Clear()
        Clear()
        Clear1()
    End Sub
    Private Function GenerateID() As String
        con = New SqlConnection(cs)
        Dim value As String = "0000"
        Try
            ' Fetch the latest ID from the database
            con.Open()
            cmd = New SqlCommand("SELECT TOP 1 Inv_ID FROM InvoiceInfo1 ORDER BY Inv_ID DESC", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.HasRows Then
                rdr.Read()
                value = rdr.Item("Inv_ID")
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
            txtInvoiceNo.Text = "INV-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub btnSelect_Click(sender As System.Object, e As System.EventArgs) Handles btnSelect.Click
        frmServicesRecord1.Reset()
        frmServicesRecord1.ShowDialog()
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles btnSelectionInv.Click
        frmCurrentStock.lblSet.Text = "Billing1"
        frmCurrentStock.Reset()
        frmCurrentStock.ShowDialog()
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
                sum = sum + r.Cells(11).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return sum
    End Function
    Public Function TotalPayment() As Double
        Dim sum As Double = 0
        Dim sum1 As Double
        Try
            For Each r As DataGridViewRow In Me.DataGridView2.Rows
                sum = sum + r.Cells(1).Value
            Next
            sum1 = sum + Val(txtUpfront.Text)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return sum1
    End Function
    Sub Print()
        Try
            Cursor = Cursors.WaitCursor
            Timer1.Enabled = True
            Dim rpt As New rptInvoice1 'The report you created.
            Dim myConnection As SqlConnection
            Dim MyCommand, MyCommand1 As New SqlCommand()
            Dim myDA, myDA1 As New SqlDataAdapter()
            Dim myDS As New DataSet 'The DataSet you created.
            myConnection = New SqlConnection(cs)
            MyCommand.Connection = myConnection
            MyCommand1.Connection = myConnection
            MyCommand.CommandText = "SELECT Service.S_ID, Service.ServiceCode, Service.ServiceType, Service.ServiceCreationDate, Service.ItemDescription, Service.ProblemDescription, Service.ChargesQuote,Service.AdvanceDeposit, Service.EstimatedRepairDate, Service.Remarks, Service.Status, Customer.ID, Customer.Name, Customer.Gender, Customer.Address, Customer.City,Customer.State, Customer.ZipCode, Customer.ContactNo, Customer.EmailID, Customer.Remarks AS Expr2, Customer.Photo, InvoiceInfo1.Inv_ID, InvoiceInfo1.InvoiceNo, InvoiceInfo1.InvoiceDate,InvoiceInfo1.ServiceID, InvoiceInfo1.RepairCharges, InvoiceInfo1.Upfront, InvoiceInfo1.ProductCharges, InvoiceInfo1.ServiceTaxPer, InvoiceInfo1.ServiceTax, InvoiceInfo1.GrandTotal, InvoiceInfo1.TotalPaid,InvoiceInfo1.Balance, InvoiceInfo1.Remarks AS Expr3, Invoice1_Product.Ipo_ID, Invoice1_Product.InvoiceID, Invoice1_Product.ProductID, Invoice1_Product.CostPrice, Invoice1_Product.SellingPrice,Invoice1_Product.Margin, Invoice1_Product.Qty, Invoice1_Product.Amount, Invoice1_Product.DiscountPer, Invoice1_Product.Discount, Invoice1_Product.VATPer, Invoice1_Product.VAT,Invoice1_Product.TotalAmount, Product.PID, Product.ProductCode, Product.ProductName, Product.SubCategoryID, Product.Description FROM Service INNER JOIN Customer ON Service.CustomerID = Customer.ID INNER JOIN InvoiceInfo1 ON Service.S_ID = InvoiceInfo1.ServiceID INNER JOIN Invoice1_Product ON InvoiceInfo1.Inv_ID = Invoice1_Product.InvoiceID INNER JOIN Product ON Invoice1_Product.ProductID = Product.PID where InvoiceInfo1.Invoiceno=@d1"
            MyCommand.Parameters.AddWithValue("@d1", txtInvoiceNo.Text)
            MyCommand1.CommandText = "SELECT * from Company"
            MyCommand.CommandType = CommandType.Text
            MyCommand1.CommandType = CommandType.Text
            myDA.SelectCommand = MyCommand
            myDA1.SelectCommand = MyCommand1
            myDA.Fill(myDS, "InvoiceInfo1")
            myDA.Fill(myDS, "Invoice1_Product")
            myDA.Fill(myDS, "Service")
            myDA.Fill(myDS, "Customer")
            myDA.Fill(myDS, "Product")
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
                DataGridView1.Rows.Add(txtProductCode.Text, txtProductName.Text, txtCostPrice.Text, txtSellingPrice.Text, txtMargin.Text, txtQty.Text, txtAmount.Text, txtDiscountPer.Text, txtDiscountAmount.Text, txtVAT.Text, txtVATAmount.Text, txtTotalAmount.Text, txtProductID.Text)
                Dim k As Double = 0
                k = GrandTotal()
                k = Math.Round(k, 2)
                txtProductCharges.Text = k
                Compute1()
                Clear()
                Exit Sub
            End If
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                If r.Cells(0).Value = txtProductCode.Text Then
                    r.Cells(0).Value = txtProductCode.Text
                    r.Cells(1).Value = txtProductName.Text
                    r.Cells(2).Value = txtCostPrice.Text
                    r.Cells(3).Value = txtSellingPrice.Text
                    r.Cells(4).Value = txtMargin.Text
                    r.Cells(5).Value = Val(r.Cells(5).Value) + Val(txtQty.Text)
                    r.Cells(6).Value = Val(r.Cells(6).Value) + Val(txtAmount.Text)
                    r.Cells(7).Value = Val(txtDiscountPer.Text)
                    r.Cells(8).Value = Val(r.Cells(8).Value) + Val(txtDiscountAmount.Text)
                    r.Cells(9).Value = Val(txtVAT.Text)
                    r.Cells(10).Value = Val(r.Cells(10).Value) + Val(txtVATAmount.Text)
                    r.Cells(11).Value = Val(r.Cells(11).Value) + Val(txtTotalAmount.Text)
                    r.Cells(12).Value = txtProductID.Text
                    Dim i As Double = 0
                    i = GrandTotal()
                    i = Math.Round(i, 2)
                    txtProductCharges.Text = i
                    Compute1()
                    Clear()
                    Exit Sub
                End If
            Next
            DataGridView1.Rows.Add(txtProductCode.Text, txtProductName.Text, txtCostPrice.Text, txtSellingPrice.Text, txtMargin.Text, txtQty.Text, txtAmount.Text, txtDiscountPer.Text, txtDiscountAmount.Text, txtVAT.Text, txtVATAmount.Text, txtTotalAmount.Text, txtProductID.Text)
            Dim j As Double = 0
            j = GrandTotal()
            j = Math.Round(j, 2)
            txtProductCharges.Text = j
            Compute1()
            Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub Clear()
        txtProductCode.Text = ""
        txtProductName.Text = ""
        txtCostPrice.Text = ""
        txtSellingPrice.Text = ""
        txtMargin.Text = ""
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
            txtProductCharges.Text = k
            Compute()
            Compute1()
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
            Me.txtCostPrice.Text = (row.Cells.Item(2).Value)
            Me.txtSellingPrice.Text = (row.Cells.Item(3).Value)
            Me.txtMargin.Text = (row.Cells.Item(4).Value)
            Me.txtQty.Text = (row.Cells.Item(5).Value)
            Me.txtAmount.Text = (row.Cells.Item(6).Value)
            Me.txtDiscountPer.Text = (row.Cells.Item(7).Value)
            Me.txtDiscountAmount.Text = (row.Cells.Item(8).Value)
            Me.txtVAT.Text = (row.Cells.Item(9).Value)
            Me.txtVATAmount.Text = (row.Cells.Item(10).Value)
            Me.txtTotalAmount.Text = (row.Cells.Item(11).Value)
            Me.txtProductID.Text = (row.Cells.Item(12).Value)
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
            Dim cq As String = "delete from InvoiceInfo1 where Inv_ID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                Dim st As String = "deleted the bill (Products + Service) having invoice no. '" & txtInvoiceNo.Text & "'"
                LogFunc(lblUser.Text, st)
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
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
    Sub Compute1()
        Dim num1, num2, num3 As Double
        num1 = CDbl((Val(txtRepairCharges.Text) * Val(txtServiceTaxPer.Text)) / 100)
        num1 = Math.Round(num1, 2)
        txtServiceTaxAmount.Text = num1
        num2 = CDbl(Val(txtRepairCharges.Text) + Val(txtServiceTaxAmount.Text) + Val(txtProductCharges.Text))
        num2 = Math.Round(num2, 2)
        txtGrandTotal.Text = num2
        num3 = Val(txtGrandTotal.Text) - Val(txtTotalPayment.Text)
        num3 = Math.Round(num3, 2)
        txtPaymentDue.Text = num3
    End Sub
    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtCustomerName.Text)) = 0 Then
            MessageBox.Show("Please retrieve service details", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If Len(Trim(txtRepairCharges.Text)) = 0 Then
            MessageBox.Show("Please enter service charges", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtRepairCharges.Focus()
            Exit Sub
        End If
        If Len(Trim(txtServiceTaxPer.Text)) = 0 Then
            MessageBox.Show("Please enter service tax %", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtServiceTaxPer.Focus()
            Exit Sub
        End If
        If DataGridView2.Rows.Count = 0 Then
            MessageBox.Show("sorry no payment info added to cart", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If Val(txtTotalPayment.Text) > Val(txtGrandTotal.Text) Then
            MessageBox.Show("Total payment can not be more than grand total", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
            For Each row As DataGridViewRow In DataGridView1.Rows
                Dim con As New SqlConnection(cs)
                con.Open()
                Dim cmd As New SqlCommand("SELECT Qty from Temp_Stock where ProductID=@d1", con)
                cmd.Parameters.AddWithValue("@d1", row.Cells(12).Value.ToString())
                Dim da As New SqlDataAdapter(cmd)
                Dim ds As DataSet = New DataSet()
                da.Fill(ds)
                If ds.Tables(0).Rows.Count > 0 Then
                    txtTotalQty.Text = ds.Tables(0).Rows(0)("Qty")
                    If CInt(Val(row.Cells(5).Value)) > Val(txtTotalQty.Text) Then
                        MessageBox.Show("added qty. to cart are more than" & vbCrLf & "available qty. of product code '" & row.Cells(0).Value.ToString() & "' and Product Name='" & row.Cells(1).Value & "'", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If
                End If
                con.Close()
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into InvoiceInfo1( Inv_ID, InvoiceNo, InvoiceDate, ServiceID, RepairCharges, Upfront, ProductCharges, ServiceTaxPer, ServiceTax, GrandTotal, TotalPaid, Balance, Remarks) Values (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11,@d12,@d13)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", txtInvoiceNo.Text)
            cmd.Parameters.AddWithValue("@d3", dtpInvoiceDate.Value.Date)
            cmd.Parameters.AddWithValue("@d4", txtS_ID.Text)
            cmd.Parameters.AddWithValue("@d5", txtRepairCharges.Text)
            cmd.Parameters.AddWithValue("@d6", txtUpfront.Text)
            cmd.Parameters.AddWithValue("@d7", txtProductCharges.Text)
            cmd.Parameters.AddWithValue("@d8", txtServiceTaxPer.Text)
            cmd.Parameters.AddWithValue("@d9", txtServiceTaxAmount.Text)
            cmd.Parameters.AddWithValue("@d10", txtGrandTotal.Text)
            cmd.Parameters.AddWithValue("@d11", txtTotalPayment.Text)
            cmd.Parameters.AddWithValue("@d12", txtPaymentDue.Text)
            cmd.Parameters.AddWithValue("@d13", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()

            If DataGridView1.Rows.Count = 0 Then
                con = New SqlConnection(cs)
                con.Open()
                Dim cx As String = "insert into Invoice1_Product(InvoiceID, CostPrice, SellingPrice, Margin, Qty, Amount, DiscountPer, Discount, VATPer, VAT, TotalAmount,ProductID) VALUES (" & txtID.Text & " ,0,0,0,0,0,0,0,0,0,0,1)"
                cmd = New SqlCommand(cx)
                cmd.Connection = con
                cmd.ExecuteNonQuery()
                con.Close()
            Else
                con = New SqlConnection(cs)
                con.Open()
                Dim cb1 As String = "insert into Invoice1_Product(InvoiceID, CostPrice, SellingPrice, Margin, Qty, Amount, DiscountPer, Discount, VATPer, VAT, TotalAmount,ProductID) VALUES (" & txtID.Text & " ,@d4,@d5,@d6,@d7,@d8,@d9,@d10,@d11,@d12,@d13,@d14)"
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
                        cmd.Parameters.AddWithValue("@d13", row.Cells(11).Value)
                        cmd.Parameters.AddWithValue("@d14", row.Cells(12).Value)
                        cmd.ExecuteNonQuery()
                        cmd.Parameters.Clear()
                    End If
                Next
                con.Close()
            End If

            con = New SqlConnection(cs)
            con.Open()
            Dim cb2 As String = "insert into Invoice1_Payment(InvoiceID,PaymentMode,TotalPaid,PaymentDate) VALUES (" & txtID.Text & " ,@d4,@d5,@d6)"
            cmd = New SqlCommand(cb2)
            cmd.Connection = con
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
            For Each row As DataGridViewRow In DataGridView2.Rows
                If Not row.IsNewRow Then
                    cmd.Parameters.AddWithValue("@d4", row.Cells(0).Value)
                    cmd.Parameters.AddWithValue("@d5", row.Cells(1).Value)
                    cmd.Parameters.AddWithValue("@d6", row.Cells(2).Value)
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            If DataGridView1.Rows.Count <> 0 Then
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If Not row.IsNewRow Then
                        con = New SqlConnection(cs)
                        con.Open()
                        Dim cb4 As String = "update Temp_stock set qty = qty - (" & row.Cells(5).Value & ") where ProductID=@d1"
                        cmd = New SqlCommand(cb4)
                        cmd.Connection = con
                        cmd.Parameters.AddWithValue("@d1", row.Cells(12).Value)
                        cmd.ExecuteNonQuery()
                        con.Close()
                    End If
                Next
            End If
            con.Close()
            Dim st As String = "added the bill (Products + Service) order having Invoice no. '" & txtInvoiceNo.Text & "'"
            LogFunc(lblUser.Text, st)
            If CheckForInternetConnection() = True Then
                con = New SqlConnection(cs)
                con.Open()
                Dim ctn As String = "select RTRIM(APIURL) from SMSSetting where IsDefault='Yes' and IsEnabled='Yes'"
                cmd = New SqlCommand(ctn)
                cmd.Connection = con
                rdr = cmd.ExecuteReader()
                If rdr.Read() Then
                    st2 = rdr.GetValue(0)
                    Dim st3 As String = "Hello, " & txtCustomerName.Text & " you have successfully received your item having invoice no. " & txtInvoiceNo.Text & ""
                    SMSFunc(txtContactNo.Text, st3, st2)
                    If (rdr IsNot Nothing) Then
                        rdr.Close()
                    End If
                End If
            End If
            con.Close()
            btnSave.Enabled = False
            RefreshRecords()
            MessageBox.Show("Successfully done", "Billing", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Print()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtCustomerName.Text)) = 0 Then
            MessageBox.Show("Please retrieve service details", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If Len(Trim(txtRepairCharges.Text)) = 0 Then
            MessageBox.Show("Please enter service charges", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtRepairCharges.Focus()
            Exit Sub
        End If
        If Len(Trim(txtServiceTaxPer.Text)) = 0 Then
            MessageBox.Show("Please enter service tax %", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtServiceTaxPer.Focus()
            Exit Sub
        End If
        If DataGridView2.Rows.Count = 0 Then
            MessageBox.Show("sorry no payment info added to cart", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If Val(txtTotalPayment.Text) > Val(txtGrandTotal.Text) Then
            MessageBox.Show("Total payment can not be more than grand total", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update InvoiceInfo1 set InvoiceNo=@d2, InvoiceDate=@d3, ServiceID=@d4, RepairCharges=@d5, Upfront=@d6, ProductCharges=@d7, ServiceTaxPer=@d8, ServiceTax=@d9, GrandTotal=@d10, TotalPaid=@d11, Balance=@d12, Remarks=@d13 where Inv_ID=@d1"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Parameters.AddWithValue("@d2", txtInvoiceNo.Text)
            cmd.Parameters.AddWithValue("@d3", dtpInvoiceDate.Value.Date)
            cmd.Parameters.AddWithValue("@d4", txtS_ID.Text)
            cmd.Parameters.AddWithValue("@d5", txtRepairCharges.Text)
            cmd.Parameters.AddWithValue("@d6", txtUpfront.Text)
            cmd.Parameters.AddWithValue("@d7", txtProductCharges.Text)
            cmd.Parameters.AddWithValue("@d8", txtServiceTaxPer.Text)
            cmd.Parameters.AddWithValue("@d9", txtServiceTaxAmount.Text)
            cmd.Parameters.AddWithValue("@d10", txtGrandTotal.Text)
            cmd.Parameters.AddWithValue("@d11", txtTotalPayment.Text)
            cmd.Parameters.AddWithValue("@d12", txtPaymentDue.Text)
            cmd.Parameters.AddWithValue("@d13", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cq As String = "delete from Invoice1_Payment where InvoiceID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb2 As String = "insert into Invoice1_Payment(InvoiceID,PaymentMode,TotalPaid,PaymentDate) VALUES (" & txtID.Text & " ,@d4,@d5,@d6)"
            cmd = New SqlCommand(cb2)
            cmd.Connection = con
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
            For Each row As DataGridViewRow In DataGridView2.Rows
                If Not row.IsNewRow Then
                    cmd.Parameters.AddWithValue("@d4", row.Cells(0).Value)
                    cmd.Parameters.AddWithValue("@d5", row.Cells(1).Value)
                    cmd.Parameters.AddWithValue("@d6", row.Cells(2).Value)
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            Dim st As String = "updated the bill (Products + Service) having invoice no. '" & txtInvoiceNo.Text & "'"
            LogFunc(lblUser.Text, st)
            btnUpdate.Enabled = False
            MessageBox.Show("Successfully updated", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmBillingRecord1.lblSet.Text = "Billing"
        frmBillingRecord1.Reset()
        frmBillingRecord1.ShowDialog()
    End Sub

    Private Sub btnNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
        Reset()
        Reset()
    End Sub


    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Cursor = Cursors.Default
        Timer1.Enabled = False
    End Sub

    Private Sub btnPrint_Click(sender As System.Object, e As System.EventArgs) Handles btnPrint.Click
        Print()
    End Sub

    Private Sub btnAdd1_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd1.Click
        Try
            If cmbPaymentMode.Text = "" Then
                MessageBox.Show("Please select payment mode", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbPaymentMode.Focus()
                Exit Sub
            End If
            If txtPayment.Text = "" Then
                MessageBox.Show("Please enter payment", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPayment.Focus()
                Exit Sub
            End If
            DataGridView2.Rows.Add(cmbPaymentMode.Text, txtPayment.Text, dtpPaymentDate.Value.Date)
            Dim j As Double = 0
            j = TotalPayment()
            j = Math.Round(j, 2)
            txtTotalPayment.Text = j
            Compute1()
            Clear1()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Sub Clear1()
        cmbPaymentMode.SelectedIndex = -1
        txtPayment.Text = ""
        dtpPaymentDate.Text = Today
        btnAdd1.Enabled = True
        btnRemove1.Enabled = False
        btnListUpdate1.Enabled = False
    End Sub
    Private Sub btnRemove1_Click(sender As System.Object, e As System.EventArgs) Handles btnRemove1.Click
        Try
            For Each row As DataGridViewRow In DataGridView2.SelectedRows
                DataGridView2.Rows.Remove(row)
            Next
            Dim k As Double = 0
            k = TotalPayment()
            k = Math.Round(k, 2)
            txtTotalPayment.Text = k
            Compute1()
            Compute()
            Clear1()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtPayment_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPayment.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtPayment.Text
            Dim selectionStart = Me.txtPayment.SelectionStart
            Dim selectionLength = Me.txtPayment.SelectionLength

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

    Private Sub DataGridView2_MouseClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView2.MouseClick
        btnRemove1.Enabled = True
        If (Me.DataGridView2.Rows.Count > 0) Then
            Me.btnRemove1.Enabled = True
            Me.btnListUpdate1.Enabled = True
            Me.btnAdd1.Enabled = False
            Dim row As DataGridViewRow = Me.DataGridView2.SelectedRows.Item(0)
            Me.cmbPaymentMode.Text = (row.Cells.Item(0).Value)
            Me.txtPayment.Text = (row.Cells.Item(1).Value)
            Me.dtpPaymentDate.Text = (row.Cells.Item(2).Value)
        End If
    End Sub

    Private Sub btnListReset1_Click(sender As System.Object, e As System.EventArgs) Handles btnListReset1.Click
        Clear1()
    End Sub

    Private Sub btnListReset_Click(sender As System.Object, e As System.EventArgs) Handles btnListReset.Click
        Clear()
    End Sub

    Private Sub DataGridView2_RowPostPaint(sender As Object, e As System.Windows.Forms.DataGridViewRowPostPaintEventArgs) Handles DataGridView2.RowPostPaint
        Dim strRowNumber As String = (e.RowIndex + 1).ToString()
        Dim size As SizeF = e.Graphics.MeasureString(strRowNumber, Me.Font)
        If DataGridView2.RowHeadersWidth < Convert.ToInt32((size.Width + 20)) Then
            DataGridView2.RowHeadersWidth = Convert.ToInt32((size.Width + 20))
        End If
        Dim b As Brush = SystemBrushes.ControlText
        e.Graphics.DrawString(strRowNumber, Me.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))

    End Sub

    Private Sub btnListUpdate1_Click(sender As System.Object, e As System.EventArgs) Handles btnListUpdate1.Click
        Try
            If cmbPaymentMode.Text = "" Then
                MessageBox.Show("Please select payment mode", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                cmbPaymentMode.Focus()
                Exit Sub
            End If
            If txtPayment.Text = "" Then
                MessageBox.Show("Please enter payment", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPayment.Focus()
                Exit Sub
            End If
            For Each row As DataGridViewRow In DataGridView2.SelectedRows
                DataGridView2.Rows.Remove(row)
            Next
            DataGridView2.Rows.Add(cmbPaymentMode.Text, txtPayment.Text, dtpPaymentDate.Value.Date)
            Dim j As Double = 0
            j = TotalPayment()
            j = Math.Round(j, 2)
            txtTotalPayment.Text = j
            Compute1()
            Clear1()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
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
            DataGridView1.Rows.Add(txtProductCode.Text, txtProductName.Text, txtCostPrice.Text, txtSellingPrice.Text, txtMargin.Text, txtQty.Text, txtAmount.Text, txtDiscountPer.Text, txtDiscountAmount.Text, txtVAT.Text, txtVATAmount.Text, txtTotalAmount.Text, txtProductID.Text)
            Dim k As Double = 0
            k = GrandTotal()
            k = Math.Round(k, 2)
            txtGrandTotal.Text = k
            Compute1()
            Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub txtRepairCharges_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtRepairCharges.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtRepairCharges.Text
            Dim selectionStart = Me.txtRepairCharges.SelectionStart
            Dim selectionLength = Me.txtRepairCharges.SelectionLength

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

    Private Sub txtServiceTaxPer_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtServiceTaxPer.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtServiceTaxPer.Text
            Dim selectionStart = Me.txtServiceTaxPer.SelectionStart
            Dim selectionLength = Me.txtServiceTaxPer.SelectionLength

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

    Private Sub txtRepairCharges_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtRepairCharges.TextChanged
        Compute1()
    End Sub

    Private Sub txtServiceTaxPer_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtServiceTaxPer.TextChanged
        Compute1()
    End Sub
End Class
