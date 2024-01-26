Imports System.Data.SqlClient
Imports System.Security.Cryptography
Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports System.Net
Imports System.Web

Module ModFunc
    Public Function CheckForInternetConnection() As Boolean
        Try
            Using client = New WebClient()
                Using stream = client.OpenRead("http://www.google.com")
                    Return True
                End Using
            End Using
        Catch
            Return False
        End Try
    End Function

    Sub LogFunc(ByVal st1 As String, ByVal st2 As String)
        con = New SqlConnection(cs)
        con.Open()
        Dim cb As String = "insert into Logs(UserID,Date,Operation) VALUES (@d1,@d2,@d3)"
        cmd = New SqlCommand(cb)
        cmd.Connection = con
        cmd.Parameters.AddWithValue("@d1", st1)
        cmd.Parameters.AddWithValue("@d2", System.DateTime.Now)
        cmd.Parameters.AddWithValue("@d3", st2)
        cmd.ExecuteReader()
        con.Close()
    End Sub
    Sub SMSFunc(ByVal st1 As String, ByVal st2 As String, ByVal st3 As String)
        st3 = st3.Replace("@MobileNo", st1).Replace("@Message", st2)
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim myUri As New Uri(st3)
        request = DirectCast(WebRequest.Create(myUri), HttpWebRequest)
        response = DirectCast(request.GetResponse(), HttpWebResponse)
    End Sub
    Public Function Encrypt(password As String) As String
        Dim strmsg As String = String.Empty
        Dim encode As Byte() = New Byte(password.Length - 1) {}
        encode = Encoding.UTF8.GetBytes(password)
        strmsg = Convert.ToBase64String(encode)
        Return strmsg
    End Function

    Public Function Decrypt(encryptpwd As String) As String
        Dim decryptpwd As String = String.Empty
        Dim encodepwd As New UTF8Encoding()
        Dim Decode As Decoder = encodepwd.GetDecoder()
        Dim todecode_byte As Byte() = Convert.FromBase64String(encryptpwd)
        Dim charCount As Integer = Decode.GetCharCount(todecode_byte, 0, todecode_byte.Length)
        Dim decoded_char As Char() = New Char(charCount - 1) {}
        Decode.GetChars(todecode_byte, 0, todecode_byte.Length, decoded_char, 0)
        decryptpwd = New [String](decoded_char)
        Return decryptpwd
    End Function
    Sub RefreshRecords()
        Dim obj As frmMainMenu = DirectCast(Application.OpenForms("frmMainMenu"), frmMainMenu)
        obj.Getdata()
        obj.DataGridView1.Refresh()
        obj.DataGridView1.Update()
    End Sub
End Module
