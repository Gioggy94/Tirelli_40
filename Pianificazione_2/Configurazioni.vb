Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows.Controls

Public Class Configurazioni

    Sub CaricaConfigurazioni()

        DataGridConfigurazioni.Rows.Clear()

        If Not DataGridConfigurazioni.Columns.Contains("Key") Then
            DataGridConfigurazioni.Columns.Add("Key", "Chiave")
        End If
        If Not DataGridConfigurazioni.Columns.Contains("Value") Then
            DataGridConfigurazioni.Columns.Add("Value", "Valore")
        End If

        For Each element As KeyValuePair(Of String, String) In Pianificazione.ConfigDictionary
            DataGridConfigurazioni.Rows.Add(element.Key, element.Value)
        Next

    End Sub

    Private Sub BTNSaveConfig_Click(sender As Object, e As EventArgs) Handles BTNSaveConfig.Click

        Dim stringbuilder As StringBuilder = New StringBuilder()
        Dim stringa As String = ""
        Dim i As Int32 = 0
        For Each row In DataGridConfigurazioni.Rows
            If Not String.IsNullOrEmpty(DataGridConfigurazioni.Rows(i).Cells(0).Value) Then
                stringbuilder.AppendLine(String.Format("[{0}]={1}", DataGridConfigurazioni.Rows(i).Cells(0).Value, DataGridConfigurazioni.Rows(i).Cells(1).Value))
                i = i + 1
            End If
        Next
        stringa = stringbuilder.ToString()
        Dim buffer As Byte() = Encoding.ASCII.GetBytes(stringa)
        Dim ms As New MemoryStream(buffer)

        Dim out_stream = CryptStream("T1r3l11@4zero!?", ms, True)

        'Dim filetest As New FileStream(".\test", FileMode.Create, FileAccess.Write)
        'ms.WriteTo(filetest)
        'filetest.Close()

        Dim file As New FileStream(".\MES.INI", FileMode.Create, FileAccess.Write)
        out_stream.WriteTo(file)
        file.Close()

        DialogResult = DialogResult.OK

        Close()

    End Sub

    Public Function CryptStream(ByVal password As String,
        ByVal in_stream As Stream, ByVal encrypt As Boolean) As MemoryStream

        Dim ms As MemoryStream = New MemoryStream
        Dim retVal As Stream

        ' Make an AES service provider.
        Dim aes_provider As New AesCryptoServiceProvider()

        ' Find a valid key size for this provider.
        Dim key_size_bits As Integer = 0
        For i As Integer = 1024 To 1 Step -1
            If (aes_provider.ValidKeySize(i)) Then
                key_size_bits = i
                Exit For
            End If
        Next i
        Debug.Assert(key_size_bits > 0)
        Console.WriteLine("Key size: " & key_size_bits)

        ' Get the block size for this provider.
        Dim block_size_bits As Integer = aes_provider.BlockSize

        ' Generate the key and initialization vector.
        Dim key() As Byte = Nothing
        Dim iv() As Byte = Nothing
        Dim salt() As Byte = {&H0, &H0, &H1, &H2, &H3, &H4,
        &H5, &H6, &HF1, &HF0, &HEE, &H21, &H22, &H45}
        MakeKeyAndIV(password, salt, key_size_bits,
        block_size_bits, key, iv)

        ' Make the encryptor or decryptor.
        Dim crypto_transform As ICryptoTransform
        If (encrypt) Then
            crypto_transform =
            aes_provider.CreateEncryptor(key, iv)
        Else
            crypto_transform =
            aes_provider.CreateDecryptor(key, iv)
        End If

        ' Attach a crypto stream to the output stream.
        ' Closing crypto_stream sometimes throws an
        ' exception if the decryption didn't work
        ' (e.g. if we use the wrong password).
        Try
            Using crypto_stream As New CryptoStream(ms,
            crypto_transform, CryptoStreamMode.Write)
                ' Encrypt or decrypt the file.
                Const block_size As Integer = 1024
                Dim buffer(block_size) As Byte
                Dim bytes_read As Integer
                Do
                    ' Read some bytes.
                    bytes_read = in_stream.Read(buffer, 0,
                    block_size)
                    If (bytes_read = 0) Then Exit Do

                    ' Write the bytes into the CryptoStream.
                    crypto_stream.Write(buffer, 0, bytes_read)
                Loop
            End Using
        Catch
        End Try

        crypto_transform.Dispose()

        retVal = New MemoryStream(ms.ToArray())
        retVal.Seek(0, SeekOrigin.Begin)

        Return retVal

    End Function

    Private Sub MakeKeyAndIV(ByVal password As String, ByVal _
  salt() As Byte, ByVal key_size_bits As Integer, ByVal _
  block_size_bits As Integer, ByRef key() As Byte, ByRef _
  iv() As Byte)
        Dim derive_bytes As New Rfc2898DeriveBytes(password,
        salt, 1000)

        key = derive_bytes.GetBytes(key_size_bits / 8)
        iv = derive_bytes.GetBytes(block_size_bits / 8)
    End Sub

    Private Sub BTNClose_Click(sender As Object, e As EventArgs) Handles BTNClose.Click
        Close()
    End Sub
End Class


