Imports System.Security.Cryptography
Imports System.Text
Imports System.IO

Public Class SaltedHash
    Public Function GetMd5Hash(ByVal md5Hash As MD5, ByVal input As String, ByVal Salt As String, ByVal HashLen As Integer) As String
        Try
            If HashLen < 8 Or HashLen > 32 Then
                HashLen = 32
            End If
            Dim data As Byte() = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(input & Salt))

            Dim sBuilder As New StringBuilder()

            Dim i As Integer
            For i = 0 To data.Length - 1
                sBuilder.Append(data(i).ToString("x2"))
            Next i

            Return Truncate(sBuilder.ToString(), HashLen)
        Catch ex As Exception
            Return "Error"
        End Try

    End Function

    Public Function VerifyMd5Hash(ByVal md5Hash As MD5, ByVal input As String, ByVal hash As String, ByVal Salt As String, ByVal HashLen As Integer) As Boolean
        Try
            If HashLen < 8 Or HashLen > 32 Then
                HashLen = 32
            End If
            Dim hashOfInput As String = GetMd5Hash(md5Hash, input, Salt, HashLen)

            Dim comparer As StringComparer = StringComparer.OrdinalIgnoreCase

            If 0 = comparer.Compare(hashOfInput, hash) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Function Truncate(value As String, length As Integer) As String
        If length > value.Length Then
            Return value
        Else
            Return value.Substring(0, length)
        End If
    End Function

    'Start Hash Generator
    Function GetMd5FileHash(ByVal md5Hash As MD5, ByVal file_name As String, ByVal HashLen As Integer) As String
        Try
            If HashLen < 8 Or HashLen > 32 Then
                HashLen = 32
            End If
            Dim hashValue() As Byte
            Dim fileStream As FileStream = File.OpenRead(file_name)
            fileStream.Position = 0
            hashValue = md5Hash.ComputeHash(fileStream)
            Dim hash_hex = PrintByteArray(hashValue)
            fileStream.Close()
            Return Truncate(hash_hex, HashLen)

        Catch ex As Exception
            Return "Error"
        End Try
    End Function

    Function VerifyMd5FileHash(ByVal md5Hash As MD5, ByVal file_name As String, ByVal hash As String, ByVal Salt As String, ByVal HashLen As Integer) As Boolean
        Try
            If HashLen < 8 Or HashLen > 32 Then
                HashLen = 32
            End If
            Dim hashOfInput As String = GetMd5FileHash(md5Hash, file_name, HashLen)
            hashOfInput = GetMd5Hash(md5Hash, hashOfInput, Salt, HashLen)
            Dim comparer As StringComparer = StringComparer.OrdinalIgnoreCase

            If 0 = comparer.Compare(hashOfInput, hash) Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function PrintByteArray(ByVal array() As Byte) As String
        Dim hex_value As String = ""
        Dim i As Integer
        For i = 0 To array.Length - 1
            hex_value += array(i).ToString("X2")
        Next i
        Return hex_value.ToLower
    End Function

End Class
