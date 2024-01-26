Imports System.Data.SqlClient
Imports System.IO

Public Class frmProduct
    Private Function GenerateID() As String
        con = New SqlConnection(cs)
        Dim value As String = "0000"
        Try
            ' Fetch the latest ID from the database
            con.Open()
            cmd = New SqlCommand("SELECT TOP 1 PID FROM Product ORDER BY PID DESC", con)
            rdr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
            If rdr.HasRows Then
                rdr.Read()
                value = rdr.Item("PID")
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
            txtProductCode.Text = "P-" + GenerateID()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Sub Reset()
        txtCostPrice.Text = ""
        txtProductCode.Text = ""
        txtDiscount.Text = 0.0
        txtSellingPrice.Text = ""
        txtVAT.Text = 0.0
        txtReorderPoint.Text = ""
        txtFeatures.Text = ""
        txtProductName.Text = ""
        cmbCategory.SelectedIndex = -1
        cmbSubCategory.SelectedIndex = -1
        cmbSubCategory.Enabled = False
        txtProductCode.Focus()
        btnSave.Enabled = True
        btnUpdate.Enabled = False
        btnDelete.Enabled = False
        Picture.Image = My.Resources._12
        dgw.Rows.Clear()
        btnRemove.Enabled = False
        auto()
    End Sub
    Sub fillCategory()
        Try
            con = New SqlConnection(cs)
            con.Open()
            adp = New SqlDataAdapter()
            adp.SelectCommand = New SqlCommand("SELECT distinct RTRIM(CategoryName) FROM Category", con)
            ds = New DataSet("ds")
            adp.Fill(ds)
            dtable = ds.Tables(0)
            cmbCategory.Items.Clear()
            For Each drow As DataRow In dtable.Rows
                cmbCategory.Items.Add(drow(0).ToString())
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub


    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If Len(Trim(txtProductCode.Text)) = 0 Then
            MessageBox.Show("Please enter charges quote", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtProductCode.Focus()
            Exit Sub
        End If
        If Len(Trim(txtProductName.Text)) = 0 Then
            MessageBox.Show("Please enter upfront", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtProductName.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbCategory.Text)) = 0 Then
            MessageBox.Show("Please select category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbCategory.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbSubCategory.Text)) = 0 Then
            MessageBox.Show("Please select sub category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbSubCategory.Focus()
            Exit Sub
        End If

        If Len(Trim(txtCostPrice.Text)) = 0 Then
            MessageBox.Show("Please enter cost price", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCostPrice.Focus()
            Exit Sub
        End If
        If Len(Trim(txtDiscount.Text)) = 0 Then
            MessageBox.Show("Please enter discount", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtDiscount.Focus()
            Exit Sub
        End If
        If Len(Trim(txtSellingPrice.Text)) = 0 Then
            MessageBox.Show("Please enter selling price", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSellingPrice.Focus()
            Exit Sub
        End If
        If Len(Trim(txtVAT.Text)) = 0 Then
            MessageBox.Show("Please enter VAT", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtVAT.Focus()
            Exit Sub
        End If
        If Len(Trim(txtReorderPoint.Text)) = 0 Then
            MessageBox.Show("Please enter reorder point", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtReorderPoint.Focus()
            Exit Sub
        End If
        If dgw.Rows.Count = 0 Then
            MessageBox.Show("Please add product images in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            Fill()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "insert into Product(PID,ProductCode, Productname, SubCategoryID, Description, CostPrice, SellingPrice, Discount, VAT, ReorderPoint) VALUES (" & txtID.Text & ",@d1,@d2,@d3,@d4,@d5,@d6,@d7,@d8,@d9)"
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d1", txtProductCode.Text)
            cmd.Parameters.AddWithValue("@d2", txtProductName.Text)
            cmd.Parameters.AddWithValue("@d3", txtSubCategoryID.Text)
            cmd.Parameters.AddWithValue("@d4", txtFeatures.Text)
            cmd.Parameters.AddWithValue("@d5", txtCostPrice.Text)
            cmd.Parameters.AddWithValue("@d6", txtSellingPrice.Text)
            cmd.Parameters.AddWithValue("@d7", txtDiscount.Text)
            cmd.Parameters.AddWithValue("@d8", txtVAT.Text)
            cmd.Parameters.AddWithValue("@d9", txtReorderPoint.Text)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim ck As String = "insert into Product_Join(ProductID,photo) VALUES (" & txtID.Text & ",@d2)"
            cmd = New SqlCommand(ck)
            cmd.Connection = con
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                    Dim ms As New MemoryStream()
                    Dim img As Image = row.Cells(0).Value
                    Dim bmpImage As New Bitmap(img)
                    bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
                    Dim data As Byte() = ms.GetBuffer()
                    Dim p As New SqlParameter("@d2", SqlDbType.Image)
                    p.Value = data
                    cmd.Parameters.Add(p)
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            LogFunc(lblUser.Text, "added the new Product '" & txtProductName.Text & "' having Product code '" & txtProductCode.Text & "'")
            RefreshRecords()
            MessageBox.Show("Successfully saved", "Product Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            auto()
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If MessageBox.Show("Do you really want to delete this record?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.Yes Then
                DeleteRecord()
                RefreshRecords()
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
            Dim cl2 As String = "SELECT PID FROM Product INNER JOIN Stock_Product ON Product.PID = Stock_Product.ProductID where PID=@d1"
            cmd = New SqlCommand(cl2)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Stock Entry", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cl As String = "SELECT PID FROM Product INNER JOIN Invoice_Product ON Product.PID = Invoice_Product.ProductID where PID=@d1"
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
            Dim cl1 As String = "SELECT PID FROM Product INNER JOIN Quotation_Join ON Product.PID = Quotation_Join.ProductID where PID=@d1"
            cmd = New SqlCommand(cl1)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read Then
                MessageBox.Show("Unable to delete..Already in use in Quotation", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                If Not rdr Is Nothing Then
                    rdr.Close()
                End If
                Exit Sub
            End If
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cl3 As String = "SELECT PID FROM Product INNER JOIN Invoice1_Product ON Product.PID = Invoice1_Product.ProductID where PID=@d1"
            cmd = New SqlCommand(cl3)
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
            Dim cq As String = "delete from Product where PID=@d1"
            cmd = New SqlCommand(cq)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Connection = con
            RowsAffected = cmd.ExecuteNonQuery()
            If RowsAffected > 0 Then
                LogFunc(lblUser.Text, "deleted the Product '" & txtProductName.Text & "' having Product code '" & txtProductCode.Text & "'")
                MessageBox.Show("Successfully deleted", "Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
            Else
                MessageBox.Show("No record found", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Reset()
                If con.State = ConnectionState.Open Then
                    con.Close()
                End If
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        If Len(Trim(txtProductCode.Text)) = 0 Then
            MessageBox.Show("Please enter product code", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtProductCode.Focus()
            Exit Sub
        End If
        If Len(Trim(txtProductName.Text)) = 0 Then
            MessageBox.Show("Please enter product name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtProductName.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbCategory.Text)) = 0 Then
            MessageBox.Show("Please select category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbCategory.Focus()
            Exit Sub
        End If
        If Len(Trim(cmbSubCategory.Text)) = 0 Then
            MessageBox.Show("Please select sub category", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            cmbSubCategory.Focus()
            Exit Sub
        End If

        If Len(Trim(txtCostPrice.Text)) = 0 Then
            MessageBox.Show("Please enter cost price", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtCostPrice.Focus()
            Exit Sub
        End If
        If Len(Trim(txtDiscount.Text)) = 0 Then
            MessageBox.Show("Please enter discount", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtDiscount.Focus()
            Exit Sub
        End If
        If Len(Trim(txtSellingPrice.Text)) = 0 Then
            MessageBox.Show("Please enter selling price", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtSellingPrice.Focus()
            Exit Sub
        End If
        If Len(Trim(txtVAT.Text)) = 0 Then
            MessageBox.Show("Please enter VAT", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtVAT.Focus()
            Exit Sub
        End If
        If Len(Trim(txtReorderPoint.Text)) = 0 Then
            MessageBox.Show("Please enter reorder point", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtReorderPoint.Focus()
            Exit Sub
        End If
        If dgw.Rows.Count = 0 Then
            MessageBox.Show("Please add product images in datagridview", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            Fill()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb As String = "Update Product set Productname=@d2, SubCategoryID=@d3, Description=@d4, CostPrice=@d5, SellingPrice=@d6, Discount=@d7, VAT=@d8, ReorderPoint=@d9,ProductCode=@d1 where PID=" & txtID.Text & ""
            cmd = New SqlCommand(cb)
            cmd.Parameters.AddWithValue("@d2", txtProductName.Text)
            cmd.Parameters.AddWithValue("@d3", txtSubCategoryID.Text)
            cmd.Parameters.AddWithValue("@d4", txtFeatures.Text)
            cmd.Parameters.AddWithValue("@d5", txtCostPrice.Text)
            cmd.Parameters.AddWithValue("@d6", txtSellingPrice.Text)
            cmd.Parameters.AddWithValue("@d7", txtDiscount.Text)
            cmd.Parameters.AddWithValue("@d8", txtVAT.Text)
            cmd.Parameters.AddWithValue("@d9", txtReorderPoint.Text)
            cmd.Parameters.AddWithValue("@d1", txtProductCode.Text)
            cmd.Connection = con
            cmd.ExecuteNonQuery()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim cb1 As String = "delete from Product_Join where ProductID=@d1"
            cmd = New SqlCommand(cb1)
            cmd.Parameters.AddWithValue("@d1", txtID.Text)
            cmd.Connection = con
            cmd.ExecuteReader()
            con.Close()
            con = New SqlConnection(cs)
            con.Open()
            Dim ck As String = "insert into Product_Join(ProductID,Photo) VALUES (" & txtID.Text & ",@d2)"
            cmd = New SqlCommand(ck)
            cmd.Connection = con
            ' Prepare command for repeated execution
            cmd.Prepare()
            ' Data to be inserted
            For Each row As DataGridViewRow In dgw.Rows
                If Not row.IsNewRow Then
                    Dim ms As New MemoryStream()
                    Dim img As Image = row.Cells(0).Value
                    Dim bmpImage As New Bitmap(img)
                    bmpImage.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
                    Dim data As Byte() = ms.GetBuffer()
                    Dim p As New SqlParameter("@d2", SqlDbType.Image)
                    p.Value = data
                    cmd.Parameters.Add(p)
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End If
            Next
            con.Close()
            LogFunc(lblUser.Text, "updated the Product '" & txtProductName.Text & "' having Product code '" & txtProductCode.Text & "'")
            RefreshRecords()
            MessageBox.Show("Successfully updated", "Product Record", MessageBoxButtons.OK, MessageBoxIcon.Information)
            btnUpdate.Enabled = False
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Reset()
    End Sub

    Private Sub Browse_Click(sender As System.Object, e As System.EventArgs) Handles Browse.Click
        Try
            With OpenFileDialog1
                .Filter = ("Images |*.png; *.bmp; *.jpg;*.jpeg; *.gif;")
                .FilterIndex = 4
            End With
            'Clear the file name
            OpenFileDialog1.FileName = ""
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                Picture.Image = Image.FromFile(OpenFileDialog1.FileName)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
    End Sub

    Private Sub BRemove_Click(sender As System.Object, e As System.EventArgs) Handles BRemove.Click
        Picture.Image = My.Resources._12
    End Sub

    Private Sub btnGetData_Click(sender As System.Object, e As System.EventArgs) Handles btnGetData.Click
        Dim frm As New frmProductRecord
        frm.lblSet.Text = "Product Entry"
        frm.Reset()
        frm.ShowDialog()
    End Sub
    Sub Fill()
        Try
            con = New SqlConnection(cs)
            con.Open()
            cmd = con.CreateCommand()
            cmd.CommandText = "SELECT ID from SubCategory where Category=@d1 and SubCategoryName=@d2"
            cmd.Parameters.AddWithValue("@d1", cmbCategory.Text)
            cmd.Parameters.AddWithValue("@d2", cmbSubCategory.Text)
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                txtSubCategoryID.Text = rdr.GetValue(0)
            End If
            If (rdr IsNot Nothing) Then
                rdr.Close()
            End If
            If con.State = ConnectionState.Open Then
                con.Close()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub
    Private Sub frmProduct_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        fillCategory()
    End Sub

    Private Sub cmbCategory_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbCategory.SelectedIndexChanged
        Try
            cmbSubCategory.Enabled = True
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "SELECT distinct RTRIM(SubCategoryName) FROM SubCategory,Category where SubCategory.Category=Category.CategoryName and CategoryName=@d1"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", cmbCategory.Text)
            rdr = cmd.ExecuteReader()
            cmbSubCategory.Items.Clear()
            While rdr.Read
                cmbSubCategory.Items.Add(rdr(0))
            End While
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub txtPrice_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCostPrice.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtCostPrice.Text
            Dim selectionStart = Me.txtCostPrice.SelectionStart
            Dim selectionLength = Me.txtCostPrice.SelectionLength

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

    Private Sub txtReorderPoint_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtReorderPoint.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtDiscount_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtDiscount.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtDiscount.Text
            Dim selectionStart = Me.txtDiscount.SelectionStart
            Dim selectionLength = Me.txtDiscount.SelectionLength

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

    Private Sub txtVAT_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtVAT.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtVAT.Text
            Dim selectionStart = Me.txtVAT.SelectionStart
            Dim selectionLength = Me.txtVAT.SelectionLength

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

    Private Sub txtSellingPrice_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtSellingPrice.KeyPress
        Dim keyChar = e.KeyChar

        If Char.IsControl(keyChar) Then
            'Allow all control characters.
        ElseIf Char.IsDigit(keyChar) OrElse keyChar = "."c Then
            Dim text = Me.txtSellingPrice.Text
            Dim selectionStart = Me.txtSellingPrice.SelectionStart
            Dim selectionLength = Me.txtSellingPrice.SelectionLength

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

    Private Sub btnAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd.Click
        dgw.Rows.Add(Picture.Image)
    End Sub

    Private Sub btnRemove_Click(sender As System.Object, e As System.EventArgs) Handles btnRemove.Click
        Try
            For Each row As DataGridViewRow In dgw.SelectedRows
                dgw.Rows.Remove(row)
            Next
            btnRemove.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dgw_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgw.MouseClick
        If dgw.Rows.Count > 0 Then
            btnRemove.Enabled = True
        End If
    End Sub
End Class
