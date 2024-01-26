Imports System.Data.SqlClient
Public Class frmSplash

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Try
            con = New SqlConnection(cs)
            con.Open()
            Dim ct As String = "select * from Registration"
            cmd = New SqlCommand(ct)
            cmd.Connection = con
            rdr = cmd.ExecuteReader()
            If rdr.Read() Then
                ProgressBar1.Visible = True
                ProgressBar1.Value = ProgressBar1.Value + 2
                If (ProgressBar1.Value = 10) Then
                    Label1.Text = "Reading modules.."
                ElseIf (ProgressBar1.Value = 20) Then
                    Label1.Text = "Turning on modules."
                ElseIf (ProgressBar1.Value = 40) Then
                    Label1.Text = "Starting modules.."
                ElseIf (ProgressBar1.Value = 60) Then
                    Label1.Text = "Loading modules.."
                ElseIf (ProgressBar1.Value = 80) Then
                    Label1.Text = "Done Loading modules.."
                ElseIf (ProgressBar1.Value = 100) Then
                    Me.Hide()
                    frmLogin.Show()
                    Timer1.Enabled = False
                End If
                Else
                    ProgressBar1.Visible = True
                    ProgressBar1.Value = ProgressBar1.Value + 2
                    If (ProgressBar1.Value = 10) Then
                        Label1.Text = "Reading modules.."
                    ElseIf (ProgressBar1.Value = 20) Then
                        Label1.Text = "Turning on modules."
                    ElseIf (ProgressBar1.Value = 40) Then
                        Label1.Text = "Starting modules.."
                    ElseIf (ProgressBar1.Value = 60) Then
                        Label1.Text = "Loading modules.."
                    ElseIf (ProgressBar1.Value = 80) Then
                        Label1.Text = "Done Loading modules.."
                    ElseIf (ProgressBar1.Value = 100) Then
                        Me.Hide()
                    frmRegistration1.Show()
                    Timer1.Enabled = False
                End If
                    End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error!")
        End Try
    End Sub

    Private Sub frmSplash1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
      
    End Sub
End Class