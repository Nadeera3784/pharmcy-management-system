Imports System.Data.SqlClient
Imports System.IO

Public Class frmStock

    Private Sub auto()
        Try
            Dim Num As Integer = 0
            con = New SqlConnection(cs)
            con.Open()
            Dim Sql As String = ("SELECT MAX(ST_ID) FROM Stock")
            cmd = New SqlCommand(Sql)
            cmd.Connection = con
            If (IsDBNull(cmd.ExecuteScalar)) Then
                Num = 1
                txtST_ID.Text = Num.ToString
                txtStockID.Text = "ST" + Num.ToString
            Else
                Num = cmd.ExecuteScalar + 1
                txtST_ID.Text = Num.ToString
                txtStockID.Text = "ST" + Num.ToString
            End If
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtSupplierID.Text)) = 0 Then
            MessageBox.Show("Please retrieve supplier id", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSupplierID.Focus()
            Exit Sub
        End If
        If DataGridView1.Rows.Count = 0 Then
            MessageBox.Show("Sorry no product info added to grid", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If Len(Trim(txtTotalPayment.Text)) = 0 Then
            MessageBox.Show("Please enter total payment", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTotalPayment.Focus()
            Exit Sub
        End If
        If Val(txtTotalPayment.Text) > Val(txtGrandTotal.Text) Then
            MessageBox.Show("Total payment can not be more than grand total", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtTotalPayment.Focus()
            Exit Sub
        End If
        Try
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    con = New SqlConnection(cs)
                    con.Open()
                    Dim ct As String = "select ProductID from Temp_Stock where ProductID=@d1"
                    cmd = New SqlCommand(ct)
                    cmd.Connection = con
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value.ToString())
                    rdr = cmd.ExecuteReader()
                    If (rdr.Read()) Then

                        con = New SqlConnection(cs)
                        con.Open()
                        Dim cb2 As String = "Update Temp_Stock set Qty = Qty + " & row.Cells(3).Value & " where ProductID=@d1"
                        cmd = New SqlCommand(cb2)
                        cmd.Connection = con
                        cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value.ToString())
                        cmd.ExecuteReader()
                        con.Close()

                    Else
                        con = New SqlConnection(cs)
                        con.Open()
                        Dim cb3 As String = "insert into Temp_Stock(ProductID,Qty) VALUES (@d1,@d2)"
                        cmd = New SqlCommand(cb3)
                        cmd.Connection = con
                        cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value.ToString())
                        cmd.Parameters.AddWithValue("@d2", row.Cells(3).Value)
                        cmd.ExecuteReader()
                        con.Close()
                    End If
                End If
            Next
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into Stock(ST_ID, Stock_ID, [Date], SupplierID, GrandTotal, TotalPayment, PaymentDue, Remarks) VALUES (@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtST_ID.Text)
            cmd.Parameters.AddWithValue("@d2", txtStockID.Text)
            cmd.Parameters.AddWithValue("@d3", dtpDate.Value.Date)
            cmd.Parameters.AddWithValue("@d4", txtSup_ID.Text)
            cmd.Parameters.AddWithValue("@d5", txtGrandTotal.Text)
            cmd.Parameters.AddWithValue("@d6", txtTotalPayment.Text)
            cmd.Parameters.AddWithValue("@d7", txtPaymentDue.Text)
            cmd.Parameters.AddWithValue("@d8", txtRemarks.Text)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "insert into Stock_Product(StockID,ProductID,Qty,Price,TotalAmount) VALUES (" & txtST_ID.Text & ",@d1,@d2,@d3,@d4)"
            cmd = New SqlCommand(cb1)
            cmd.Connection = con
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    cmd.Parameters.AddWithValue("@d1", row.Cells(0).Value)
                    cmd.Parameters.AddWithValue("@d2", row.Cells(3).Value)
                    cmd.Parameters.AddWithValue("@d3", row.Cells(4).Value)
                    cmd.Parameters.AddWithValue("@d4", row.Cells(5).Value)
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            RefreshRecords()
            LogFunc(lblUser.Text, "added the new stock having stock id '" & txtStockID.Text & "'")
            MessageBox.Show("Successfully saved", "Stock", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnSave.Enabled = False
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        frmSupplierRecord.lblSet.Text = "Stock Entry"
        frmSupplierRecord.Reset()
        frmSupplierRecord.ShowDialog()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        frmProductRecord.lblSet.Text = "Stock"
        frmProductRecord.Reset()
        frmProductRecord.ShowDialog()
    End Sub
    Sub Reset()
        txtGrandTotal.Text = ""
        txtPricePerQty.Text = ""
        txtProductCode.Text = ""
        txtProductName.Text = ""
        txtQty.Text = ""
        txtRemarks.Text = ""
        txtSupplierID.Text = ""
        txtSupplierName.Text = ""
        txtTotalAmount.Text = ""
        txtTotalPayment.Text = ""
        dtpDate.Text = Today
        DataGridView1.Rows.Clear()
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnRemove.Enabled = False
        txtPaymentDue.Text = ""
        dtpDate.Enabled = True
        DataGridView1.Enabled = True
        btnAdd.Enabled = True
        auto()
    End Sub
    Private Sub btnNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNew.Click
        Reset()
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
            If txtPricePerQty.Text = "" Then
                MessageBox.Show("Please enter price per qty.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                txtPricePerQty.Focus()
                Exit Sub
            End If
            If DataGridView1.Rows.Count = 0 Then
                DataGridView1.Rows.Add(txtProductID.Text, txtProductCode.Text, txtProductName.Text, txtQty.Text, txtPricePerQty.Text, txtTotalAmount.Text)
                Dim k As Double = 0
                k = GrandTotal()
                k = Math.Round(k, 2)
                txtGrandTotal.Text = k
                Clear()
                Exit Sub
            End If
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                If r.Cells(1).Value = txtProductCode.Text Then
                    r.Cells(0).Value = txtProductID.Text
                    r.Cells(1).Value = txtProductCode.Text
                    r.Cells(2).Value = txtProductName.Text
                    r.Cells(3).Value = Val(r.Cells(3).Value) + Val(txtQty.Text)
                    r.Cells(4).Value = txtPricePerQty.Text
                    r.Cells(5).Value = Val(r.Cells(5).Value) + Val(txtTotalAmount.Text)
                    Dim i As Double = 0
                    i = GrandTotal()
                    i = Math.Round(i, 2)
                    txtGrandTotal.Text = i
                    Clear()
                    Exit Sub
                End If
            Next
            DataGridView1.Rows.Add(txtProductID.Text, txtProductCode.Text, txtProductName.Text, txtQty.Text, txtPricePerQty.Text, txtTotalAmount.Text)
            Dim j As Double = 0
            j = GrandTotal()
            j = Math.Round(j, 2)
            txtGrandTotal.Text = j
            Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Function GrandTotal() As Double
        Dim sum As Double = 0
        Try
            For Each r As DataGridViewRow In Me.DataGridView1.Rows
                sum = sum + r.Cells(5).Value
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return sum
    End Function
    Sub Clear()
        txtProductCode.Text = ""
        txtProductName.Text = ""
        txtQty.Text = ""
        txtPricePerQty.Text = ""
        txtTotalAmount.Text = ""
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
            btnRemove.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub Compute()
        Dim i As Double = 0
        i = Val(txtGrandTotal.Text) - Val(txtTotalPayment.Text)
        i = Math.Round(i, 2)
        txtPaymentDue.Text = i
    End Sub

    Private Sub DataGridView1_MouseClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick
        If DataGridView1.Rows.Count > 0 Then
            btnRemove.Enabled = True
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

    Private Sub frmStock_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub txtPricePerQty_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtPricePerQty.TextChanged
        Dim i As Double = 0
        i = CDbl(Val(txtQty.Text) * Val(txtPricePerQty.Text))
        i = Math.Round(i, 2)
        txtTotalAmount.Text = i
    End Sub


    Private Sub txtQty_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtQty.TextChanged
        Dim i As Double = 0
        i = CDbl(Val(txtQty.Text) * Val(txtPricePerQty.Text))
        i = Math.Round(i, 2)
        txtTotalAmount.Text = i
    End Sub

    Private Sub txtQty_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtQty.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtPricePerQty_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPricePerQty.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtPricePerQty.Text
            Dim selectionStart = Me.txtPricePerQty.SelectionStart
            Dim selectionLength = Me.txtPricePerQty.SelectionLength

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

    Private Sub txtTotalPayment_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtTotalPayment.TextChanged
        Compute()
    End Sub

    Private Sub txtTotalPayment_Validating(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles txtTotalPayment.Validating
        If Val(txtTotalPayment.Text) > Val(txtGrandTotal.Text) Then
            MessageBox.Show("Total payment can not be more than grand total", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        Exit Sub
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        frmStockRecord.Reset()
        frmStockRecord.ShowDialog()
    End Sub

    Private Sub btnUpdate_Click(sender As System.Object, e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtSupplierID.Text)) = 0 Then
            MessageBox.Show("Please retrieve supplier id", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSupplierID.Focus()
            Exit Sub
        End If
        If DataGridView1.Rows.Count = 0 Then
            MessageBox.Show("Sorry no product info added to grid", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If Len(Trim(txtTotalPayment.Text)) = 0 Then
            MessageBox.Show("Please enter total payment", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtTotalPayment.Focus()
            Exit Sub
        End If
        If Val(txtTotalPayment.Text) > Val(txtGrandTotal.Text) Then
            MessageBox.Show("Total payment can not be more than grand total", "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            txtTotalPayment.Focus()
            Exit Sub
        End If
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update Stock set Stock_ID=@d2, [Date]=@d3, SupplierID=@d4, GrandTotal=@d5, TotalPayment=@d6, PaymentDue=@d7, Remarks=@d8 where ST_ID=@d1"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d2", txtStockID.Text)
            cmd.Parameters.AddWithValue("@d3", dtpDate.Value)
            cmd.Parameters.AddWithValue("@d4", txtSup_ID.Text)
            cmd.Parameters.AddWithValue("@d5", txtGrandTotal.Text)
            cmd.Parameters.AddWithValue("@d6", txtTotalPayment.Text)
            cmd.Parameters.AddWithValue("@d7", txtPaymentDue.Text)
            cmd.Parameters.AddWithValue("@d8", txtRemarks.Text)
            cmd.Parameters.AddWithValue("@d1", txtST_ID.Text)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            RefreshRecords()
            LogFunc(lblUser.Text, "updated stock having stock id '" & txtStockID.Text & "'")
            MessageBox.Show("Successfully updated", "Stock", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
