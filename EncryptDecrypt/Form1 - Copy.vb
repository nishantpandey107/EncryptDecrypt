Imports System
Imports System.IO
Imports System.Security
Imports System.Security.Cryptography

Public Class Form1
    Dim strFileToEncrypt As String
    Dim strFileToDecrypt As String
    Dim strOutputEncrypt As String
    Dim strOutputDecrypt As String
    Dim fsInput As System.IO.FileStream
    Dim fsOutput As System.IO.FileStream
    Dim inputFile As String

    Dim sFileArray(0) As String
    Dim sInt As Integer = 0
    Dim sOriginalFolderPath As String

    Dim sFileProgressOldValue As Integer = 0
    Dim sFileProgressNewValue As Integer = 0
    Dim sFolderProgressOldValue As Integer = 0
    Dim sFolderProgressNewValue As Integer = 0

    Private Function CreateKey(ByVal strPassword As String) As Byte()
        'Convert strPassword to an array and store in chrData.
        Dim chrData() As Char = strPassword.ToCharArray
        'Use intLength to get strPassword size.
        Dim intLength As Integer = chrData.GetUpperBound(0)
        'Declare bytDataToHash and make it the same size as chrData.
        Dim bytDataToHash(intLength) As Byte

        'Use For Next to convert and store chrData into bytDataToHash.
        For i As Integer = 0 To chrData.GetUpperBound(0)
            bytDataToHash(i) = CByte(Asc(chrData(i)))
        Next

        'Declare what hash to use.
        Dim SHA512 As New System.Security.Cryptography.SHA512Managed
        'Declare bytResult, Hash bytDataToHash and store it in bytResult.
        Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)
        'Declare bytKey(31).  It will hold 256 bits.
        Dim bytKey(31) As Byte

        'Use For Next to put a specific size (256 bits) of 
        'bytResult into bytKey. The 0 To 31 will put the first 256 bits
        'of 512 bits into bytKey.
        For i As Integer = 0 To 31
            bytKey(i) = bytResult(i)
        Next

        Return bytKey 'Return the key.
    End Function

    Private Function CreateIV(ByVal strPassword As String) As Byte()
        'Convert strPassword to an array and store in chrData.
        Dim chrData() As Char = strPassword.ToCharArray
        'Use intLength to get strPassword size.
        Dim intLength As Integer = chrData.GetUpperBound(0)
        'Declare bytDataToHash and make it the same size as chrData.
        Dim bytDataToHash(intLength) As Byte

        'Use For Next to convert and store chrData into bytDataToHash.
        For i As Integer = 0 To chrData.GetUpperBound(0)
            bytDataToHash(i) = CByte(Asc(chrData(i)))
        Next

        'Declare what hash to use.
        Dim SHA512 As New System.Security.Cryptography.SHA512Managed
        'Declare bytResult, Hash bytDataToHash and store it in bytResult.
        Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)

        'Declare bytIV(15).  It will hold 128 bits.
        Dim bytIV(15) As Byte

        'Use For Next to put a specific size (128 bits) of bytResult into bytIV.
        'The 0 To 30 for bytKey used the first 256 bits of the hashed password.
        'The 32 To 47 will put the next 128 bits into bytIV.
        For i As Integer = 32 To 47
            bytIV(i - 32) = bytResult(i)
        Next

        Return bytIV 'Return the IV.
    End Function

    Private Enum CryptoAction
        'Define the enumeration for CryptoAction.
        ActionEncrypt = 1
        ActionDecrypt = 2
    End Enum

    Private Sub EncryptOrDecryptFile(ByVal strInputFile As String, ByVal strOutputFile As String, ByVal bytKey() As Byte, _
                                     ByVal bytIV() As Byte, ByVal Direction As CryptoAction)

        Try

            'Setup file streams to handle input and output.
            fsInput = New System.IO.FileStream(strInputFile, FileMode.Open, FileAccess.Read)
            fsOutput = New System.IO.FileStream(strOutputFile, FileMode.OpenOrCreate, FileAccess.Write)
            fsOutput.SetLength(0) 'make sure fsOutput is empty

            'Declare variables for encrypt/decrypt process.
            Dim bytBuffer(4096) As Byte 'holds a block of bytes for processing
            Dim lngBytesProcessed As Long = 0 'running count of bytes processed
            Dim inputFileLength As Long = fsInput.Length 'the input file's length
            Dim intBytesInCurrentBlock As Integer 'current bytes being processed
            Dim csCryptoStream As CryptoStream
            Dim cspRijndael As New System.Security.Cryptography.RijndaelManaged  'Declare your CryptoServiceProvider.

            'Determine if ecryption or decryption and setup CryptoStream.
            Select Case Direction
                Case CryptoAction.ActionEncrypt
                    csCryptoStream = New CryptoStream(fsOutput, cspRijndael.CreateEncryptor(bytKey, bytIV), CryptoStreamMode.Write)
                Case CryptoAction.ActionDecrypt
                    csCryptoStream = New CryptoStream(fsOutput, cspRijndael.CreateDecryptor(bytKey, bytIV), CryptoStreamMode.Write)
            End Select

            'Use While to loop until all of the file is processed.
            Console.WriteLine("Start ------->    " & DateTime.Now)


            Dim nTotalLoopRequired = CInt(inputFileLength / 4096)
            If nTotalLoopRequired = 0 Then
                nTotalLoopRequired = 1
            End If
            Dim nCurrentLoop = 0
            Dim nOverAllProgress As Integer

            While lngBytesProcessed < inputFileLength

                'Read file with the input filestream.
                intBytesInCurrentBlock = fsInput.Read(bytBuffer, 0, 4096)

                'Write output file with the cryptostream.
                csCryptoStream.Write(bytBuffer, 0, intBytesInCurrentBlock)

                'Update lngBytesProcessed
                lngBytesProcessed = lngBytesProcessed + CLng(intBytesInCurrentBlock)

                nOverAllProgress = CInt((nCurrentLoop * 100) / nTotalLoopRequired)
                nCurrentLoop = nCurrentLoop + 1

                sFileProgressNewValue = nOverAllProgress
                If sFileProgressOldValue <> sFileProgressNewValue Then
                    Label2.Text = sFileProgressNewValue & " % File Encryption Done."
                    Label2.Refresh()
                End If
                sFileProgressOldValue = sFileProgressNewValue

            End While
            Console.WriteLine("End ------->    " & DateTime.Now)
            'Close FileStreams and CryptoStream.
            csCryptoStream.Close()
            fsInput.Close()
            fsOutput.Close()
            sFileProgressNewValue = 0
            sFileProgressOldValue = 0

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    'Browse Button Logic
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Label1.Text = "Processing...."

        If RadioButton1.Checked = True Then
            'If File is selected
            Dim FileBrowse As New OpenFileDialog()
            FileBrowse.ValidateNames = True
            FileBrowse.Title = "Select the File"
            FileBrowse.RestoreDirectory = True
            FileBrowse.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            If FileBrowse.ShowDialog = Windows.Forms.DialogResult.OK Then
                inputFile = FileBrowse.FileName
                Dim sFileinfo As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FileBrowse.FileName)
                Label1.Text = "You have selected the follwoing files" & vbCrLf & _
                              "Folder Path = " & sFileinfo.DirectoryName.Replace("\", " > ") & vbCrLf & _
                              "  File Name = " & sFileinfo.Name & vbCrLf & _
                              "     File Size = " & GetSize(sFileinfo.Length)
                If sFileinfo.Name.Substring(sFileinfo.Name.Length - 5) = "crypt" Then
                    Button2.Enabled = True
                    Button1.Enabled = False
                Else
                    Button1.Enabled = True
                    Button2.Enabled = False
                End If
                'ElseIf FileBrowse.ShowDialog = Windows.Forms.DialogResult.Cancel Then
            Else
                Label1.Text = "Click On the Browse Button to select the file"
            End If

        Else
            'If folder is selected
            Dim folderBrowse As New FolderBrowserDialog()
            folderBrowse.Description = "Select the Folder"
            folderBrowse.RootFolder = Environment.SpecialFolder.Desktop
            If folderBrowse.ShowDialog = Windows.Forms.DialogResult.OK Then


                ''Get All the files within the folder and subfolder
                sOriginalFolderPath = folderBrowse.SelectedPath
                GetFileFolder(folderBrowse.SelectedPath)

                'No need for run async process, as first we want the user to select all the files and then perform the action.
                'BackgroundWorker1.RunWorkerAsync()

                Label1.Text = sFileArray.Length & " File(s) selected"

                If folderBrowse.SelectedPath.Split("\")(folderBrowse.SelectedPath.Split("\").Length - 1).Contains(".crypt") = True Then
                    Button2.Enabled = True
                    Button1.Enabled = False
                Else
                    Button1.Enabled = True
                    Button2.Enabled = False
                End If
            Else
                Label1.Text = "Click On the Browse Button to select the file"
                RadioButton1.Select()
            End If

        End If
    End Sub

    'Encrypt Button
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False

        'Declare variables for the key and iv.
        'The key needs to hold 256 bits and the iv 128 bits.
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        'Send the password to the CreateKey function.
        If TextBox1.Text.Length = 0 Then
            bytKey = CreateKey("Nishant")
        Else
            bytKey = CreateKey(TextBox1.Text)
        End If
        'Send the password to the CreateIV function.
        If TextBox1.Text.Length = 0 Then
            bytIV = CreateIV("Nishant")
        Else
            bytIV = CreateIV(TextBox1.Text)
        End If


        If RadioButton1.Checked = True Then
            Dim strInputFile As String = inputFile
            Dim strOutputFile As String = inputFile & ".crypt"
            
            EncryptOrDecryptFile(strInputFile, strOutputFile, bytKey, bytIV, CryptoAction.ActionEncrypt)
        Else
            Label4.Text = "0 % Files( " & 0 & " Out Of " & sFileArray.Length & " ) Encrypted"
            Label4.Refresh()

            If sFileArray.Length > 0 Then
                'create the parallel selected crypted folder
                If Not My.Computer.FileSystem.DirectoryExists(sOriginalFolderPath & ".crypt") Then
                    My.Computer.FileSystem.CreateDirectory(sOriginalFolderPath & ".crypt")
                End If

                Dim nCurrentFile = 0

                'loop through each folder and files
                For Each tempFile In sFileArray

                    Dim strInputFile As String = tempFile
                    'output file should be in the crypted folder(selected once)
                    Dim strOutputFile As String = tempFile.Replace(sOriginalFolderPath, sOriginalFolderPath & ".crypt") & ".crypt"

                    'now create the inside folder
                    If Not My.Computer.FileSystem.DirectoryExists(strOutputFile.Substring(0, strOutputFile.Length - strOutputFile.Split("\")(strOutputFile.Split("\").Length - 1).Length - 1)) Then
                        My.Computer.FileSystem.CreateDirectory(strOutputFile.Substring(0, strOutputFile.Length - strOutputFile.Split("\")(strOutputFile.Split("\").Length - 1).Length - 1))
                    End If

                    EncryptOrDecryptFile(strInputFile, strOutputFile, bytKey, bytIV, CryptoAction.ActionEncrypt)
                    nCurrentFile = nCurrentFile + 1
                    BackgroundWorker1.ReportProgress(nCurrentFile, "Encrypt")
                Next
            End If
        End If
        Button4.Visible = True
    End Sub

    'Decrypt button
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Button1.Enabled = False
        Button2.Enabled = False
        Button3.Enabled = False

        'Declare variables for the key and iv.
        'The key needs to hold 256 bits and the iv 128 bits.
        Dim bytKey As Byte()
        Dim bytIV As Byte()
        'Send the password to the CreateKey function.
        If TextBox1.Text.Length = 0 Then
            bytKey = CreateKey("Nishant")
        Else
            bytKey = CreateKey(TextBox1.Text)
        End If
        'Send the password to the CreateIV function.
        If TextBox1.Text.Length = 0 Then
            bytIV = CreateIV("Nishant")
        Else
            bytIV = CreateIV(TextBox1.Text)
        End If

        If RadioButton1.Checked = True Then
            Dim strInputFile As String = inputFile
            Dim strOutputFile As String = inputFile.Substring(0, inputFile.Length - 6)

            EncryptOrDecryptFile(strInputFile, strOutputFile, bytKey, bytIV, CryptoAction.ActionDecrypt)
        Else
            Label4.Text = "0 % Files( " & 0 & " Out Of " & sFileArray.Length & " ) Decrypted"
            Label4.Refresh()

            If sFileArray.Length > 0 Then

                'create the parallel selected crypted folder
                If Not My.Computer.FileSystem.DirectoryExists(sOriginalFolderPath.Substring(0, (sOriginalFolderPath.Length - 6))) Then
                    My.Computer.FileSystem.CreateDirectory(sOriginalFolderPath.Substring(0, (sOriginalFolderPath.Length - 6)))
                End If

                Dim nCurrentFile = 0

                'loop through each folder and files
                For Each tempFile In sFileArray
                    Dim strInputFile As String = tempFile

                    'output file should be in the crypted folder(selected once)
                    Dim strOutputFile As String = strInputFile.Replace(".crypt", "")

                    'now create the inside folder
                    If Not My.Computer.FileSystem.DirectoryExists(strOutputFile.Substring(0, strOutputFile.Length - strOutputFile.Split("\")(strOutputFile.Split("\").Length - 1).Length - 1)) Then
                        My.Computer.FileSystem.CreateDirectory(strOutputFile.Substring(0, strOutputFile.Length - strOutputFile.Split("\")(strOutputFile.Split("\").Length - 1).Length - 1))
                    End If

                    EncryptOrDecryptFile(strInputFile, strOutputFile, bytKey, bytIV, CryptoAction.ActionDecrypt)
                    nCurrentFile = nCurrentFile + 1
                    BackgroundWorker1.ReportProgress(nCurrentFile, "Decrypt")
                Next
            End If
        End If
        Button4.Visible = True
    End Sub

    Private Sub GetFileFolder(ByVal sSourceFolderpath)

        Dim sFile As String() = System.IO.Directory.GetFiles(sSourceFolderpath)
        For Each Filename In sFile
            ReDim Preserve sFileArray(sInt)
            sFileArray(sInt) = Filename
            BackgroundWorker1.ReportProgress(sFileArray.Length - 1, "GetFileFolder")
            sInt = sInt + 1
        Next

        Dim sFolder As String() = System.IO.Directory.GetDirectories(sSourceFolderpath)
        For Each Folder In sFolder
            GetFileFolder(Folder)
        Next Folder
    End Sub

    Public Function GetSize(ByVal size As Long) As String
        Dim sSize As String = ","
        Select Case size
            Case 0 To 1024
                sSize = size & "," & "Byte"
            Case 1024 To 1048576
                sSize = Format(size / 1024, "0.00").ToString() & " " & "KB"
            Case 1048576 To 1073741824
                sSize = Format(size / 1048576, "0.00").ToString() & " " & "MB"
            Case 1073741824 To 1099511627776
                sSize = Format(size / 1073741824, "0.00").ToString() & " " & "GB"
            Case Else
                sSize = "Size capacity more than GB"
        End Select
        Return sSize
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Button1.Enabled = False
        Button2.Enabled = False
        TextBox1.Visible = False
        Button4.Visible = False

        Dim version As Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
        Dim buildDate As DateTime = New DateTime(2000, 1, 1).AddDays(version.Build).AddSeconds(version.Revision * 2)
        Label3.Text = "Product Version Date = " & buildDate.GetDateTimeFormats()(107)

        BackgroundWorker1.WorkerReportsProgress = True
        'BackgroundWorker1.WorkerSupportsCancellation = True

        Label2.Text = ""
        Label4.Text = ""

        RadioButton2.Checked = CheckState.Checked

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = False Then
            TextBox1.Visible = True
        Else
            TextBox1.Visible = False
        End If
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        If e.UserState = "GetFileFolder" Then

            Label1.Text = sFileArray(e.ProgressPercentage) & vbCrLf & e.ProgressPercentage & " Files selected"
            Label1.Refresh()

        ElseIf e.UserState = "Encrypt" Then

            Dim nCurrentFile As Integer = e.ProgressPercentage
            Dim nOverAllProgress As Integer = CInt((nCurrentFile * 100) / sFileArray.Length)
            sFolderProgressNewValue = nOverAllProgress
            If sFolderProgressOldValue <> sFolderProgressNewValue Then
                Label4.Text = sFolderProgressNewValue & " % Files( " & nCurrentFile & " Out Of " & sFileArray.Length & " ) Encrypted"
                Label4.Refresh()
            End If
            sFolderProgressOldValue = sFolderProgressNewValue

        ElseIf e.UserState = "Decrypt" Then

            Dim nCurrentFile As Integer = e.ProgressPercentage
            Dim nOverAllProgress As Integer = CInt((nCurrentFile * 100) / sFileArray.Length)
            sFolderProgressNewValue = nOverAllProgress
            If sFolderProgressOldValue <> sFolderProgressNewValue Then
                Label4.Text = sFolderProgressNewValue & " % Files( " & nCurrentFile & " Out Of " & sFileArray.Length & " ) Decrypted"
                Label4.Refresh()
            End If
            sFolderProgressOldValue = sFolderProgressNewValue

        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Label4.Text = ""
        Label2.Text = ""
        Label1.Text = "Click On the Browse Button to select the file"
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
End Class

