Imports System.Data.SqlClient

Public Class frmRegistration1

    Sub Reset()
        txtContactNo.Text = ""
        txtEmailID.Text = ""
        txtName.Text = ""
        txtPassword.Text = ""
        txtUserID.Text = ""
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        End
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If txtUserID.Text = "" Then
            MessageBox.Show("Please enter user id", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtUserID.Focus()
            Return
        End If
        If txtPassword.Text = "" Then
            MessageBox.Show("Please enter password", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtPassword.Focus()
            Return
        End If
        If txtName.Text = "" Then
            MessageBox.Show("Please enter name", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtName.Focus()
            Return
        End If
        If txtContactNo.Text = "" Then
            MessageBox.Show("Please enter contact no.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            txtContactNo.Focus()
            Return
        End If

        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select userid from registration where userid=@d1"
            cmd = New SqlCommand(ct)
            cmd.Parameters.AddWithValue("@d1", txtUserID.Text)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()

            If rdr.Read() Then
                MessageBox.Show("user id Already Exists", "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
                txtUserID.Text = ""
                txtUserID.Focus()
                If (rdr IsNot Nothing) Then
                    rdr.Close()
                End If
                Return
            End If

            con = New SqlConnection(cs)
            con.Open()

            Dim cb As String = "insert into Registration(userid, UserType, Password, Name, ContactNo, EmailID,JoiningDate,Active) VALUES (@d1,'Admin',@d3,@d4,@d5,@d6,@d7,'Yes')"
            cmd = New SqlCommand(cb)
            cmd.Connection = con
            cmd.Parameters.AddWithValue("@d1", txtUserID.Text)
            cmd.Parameters.AddWithValue("@d3", Encrypt(txtPassword.Text.Trim()))
            cmd.Parameters.AddWithValue("@d4", txtName.Text)
            cmd.Parameters.AddWithValue("@d5", txtContactNo.Text)
            cmd.Parameters.AddWithValue("@d6", txtEmailID.Text)
            cmd.Parameters.AddWithValue("@d7", Now)
            cmd.ExecuteReader()
            con.Close()
            MessageBox.Show("Successfully Registered", "User", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Hide()
            frmLogin.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.[Error])
        End Try
    End Sub

    Private Sub frmRegistration1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Reset()
    End Sub
End Class
