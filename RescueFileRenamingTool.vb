Public Class RescueFileRenamingTool

    Property DataFolder As String
    Property OutputFolder As String

    Private _DataFiles As DataTable
    Property DataFiles As DataTable
        Get
            If _DataFiles Is Nothing Then
                _DataFiles = New DataTable
                _DataFiles.Columns.Add("FileName", GetType(String))
                _DataFiles.Columns.Add("DateTimeStamp", GetType(DateTime))
                _DataFiles.Columns.Add("NewFileName", GetType(String))
                _DataFiles.Columns.Add("FileSize", GetType(Long))
                _DataFiles.Columns.Add("SummarizedSize", GetType(Long))
            End If
            Return _DataFiles
        End Get
        Set(value As DataTable)
            _DataFiles = value
        End Set
    End Property

    Sub LoadDataFolderAndResortAndRefreshGrid(directory As String)
        Try
            Me.Cursor = Cursors.WaitCursor
            My.Application.DoEvents()
            'Load list of data files
            Dim Dir As New System.IO.DirectoryInfo(directory)
            Dim Files As System.IO.FileInfo() = Dir.GetFiles("*.au", IO.SearchOption.AllDirectories)
            DataFiles = Nothing
            For MyCounter As Integer = 0 To Files.Length - 1
                ShowProgress("Reading data directory", MyCounter + 1, Files.Length)
                Dim row As DataRow = DataFiles.NewRow
                row("FileName") = Files(MyCounter).FullName.Substring(directory.Length + 1)
                row("DateTimeStamp") = Files(MyCounter).LastWriteTimeUtc
                row("FileSize") = Files(MyCounter).Length
                DataFiles.Rows.Add(row)
            Next
            'Sort everything by write time
            DataFiles = CompuMaster.Data.DataTables.CreateDataTableClone(DataFiles, "", "DateTimeStamp ASC")
            'Assign new filename (consider file size limit < 1 GB for each sub output directory)
            Dim OutputSize As Long = 0
            Dim SubDirCounterFor1GBLimitation As Integer = 0
            For MyCounter As Integer = 0 To DataFiles.Rows.Count - 1
                ShowProgress("Preparing data", MyCounter + 1, DataFiles.Rows.Count)
                If Fix((OutputSize + CType(DataFiles.Rows(MyCounter)("FileSize"), Long)) / 1024 ^ 3) > Fix(OutputSize / 1024 ^ 3) Then
                    'next 1 GB sub directory required
                    SubDirCounterFor1GBLimitation += 1
                End If
                OutputSize += CType(DataFiles.Rows(MyCounter)("FileSize"), Long)
                DataFiles(MyCounter)("NewFileName") = "data" & SubDirCounterFor1GBLimitation.ToString("0000") & System.IO.Path.DirectorySeparatorChar &
                    "b" & MyCounter.ToString("00000") &
                    System.IO.Path.GetExtension(CType(DataFiles(MyCounter)("FileName"), String))
                DataFiles(MyCounter)("SummarizedSize") = OutputSize
            Next
            Me.Cursor = Cursors.Default
            ResetStatusLine()
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            ResetStatusLine()
            MsgBox("Files operations failed: " & vbNewLine & ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub

    ''' <summary>
    ''' Ask the user for an input data directory
    ''' </summary>
    ''' <returns>True if the user has selected a valid directory, False if user cancelled dialog</returns>
    Function UserSelectionInputFolder() As Boolean
        Dim folderDlg As New FolderBrowserDialog
        folderDlg.ShowNewFolderButton = False
        folderDlg.RootFolder = Environment.SpecialFolder.MyComputer
        folderDlg.SelectedPath = CompuMaster.Data.Utils.StringNotEmptyOrAlternativeValue(DataFolder, My.Application.Info.DirectoryPath)
        If (folderDlg.ShowDialog(Me) = DialogResult.OK) Then
            If System.IO.Directory.GetFiles(folderDlg.SelectedPath, "*.au", IO.SearchOption.AllDirectories).Length = 0 Then
                MsgBox("No .au files found - this is not an Audacity data directory", MsgBoxStyle.Critical)
                Return False
            Else
                DataFolder = folderDlg.SelectedPath
                LoadDataFolderAndResortAndRefreshGrid(DataFolder)
                Me.DataGridView.DataSource = DataFiles
                Return True
            End If
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Ask the user for an output data directory
    ''' </summary>
    ''' <returns>True if the user has selected a valid directory, False if user cancelled dialog</returns>
    Function UserSelectionOutputFolder() As Boolean
        Dim folderDlg As New FolderBrowserDialog
        folderDlg.ShowNewFolderButton = True
        folderDlg.RootFolder = Environment.SpecialFolder.MyComputer
        folderDlg.SelectedPath = My.Application.Info.DirectoryPath
        If (folderDlg.ShowDialog(Me) = DialogResult.OK) Then
            If DataFolder = folderDlg.SelectedPath Then
                MsgBox("Input data folder and output data folder must be different directories", MsgBoxStyle.Critical)
                Return False
            ElseIf System.IO.Directory.GetFileSystemEntries(folderDlg.SelectedPath).Length > 0 Then
                MsgBox("Output data folder must be an empty directory", MsgBoxStyle.Critical)
                Return False
            Else
                OutputFolder = folderDlg.SelectedPath
                Return True
            End If
        Else
            Return False
        End If
    End Function

    Private Sub ToolStripButtonSelectDataFolder_Click(sender As Object, e As EventArgs) Handles ToolStripButtonSelectDataFolder.Click
        UserSelectionInputFolder()
    End Sub

    Private Sub ToolStripButtonSelectOutputFolder_Click(sender As Object, e As EventArgs) Handles ToolStripButtonSelectOutputFolder.Click
        UserSelectionOutputFolder()
    End Sub

    ''' <summary>
    ''' Open default folder dialogs
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub RescueFileRenamingTool_Load(sender As Object, e As EventArgs) Handles MyBase.Shown
        ResetStatusLine()
        If UserSelectionInputFolder() AndAlso UserSelectionOutputFolder() Then
            MsgBox("Please review the intended file operations. If everything is okay, start the file copy process by clicking the button ""Go!""", MsgBoxStyle.Information)
        Else
            MsgBox("Please select input and output data directory for the intended file operations. If everything is okay, start the file copy process by clicking the button ""Go!""", MsgBoxStyle.Information)
        End If
    End Sub

    Private Sub ToolStripButtonGo_Click(sender As Object, e As EventArgs) Handles ToolStripButtonGo.Click
        Try
            If System.IO.Directory.GetFileSystemEntries(OutputFolder).Length > 0 Then
                MsgBox("Output data folder must be an empty directory", MsgBoxStyle.Critical)
                Return
            End If

            Me.Cursor = Cursors.WaitCursor
            My.Application.DoEvents()
            For MyCounter As Integer = 0 To DataFiles.Rows.Count - 1
                ShowProgress("Copying", MyCounter + 1, DataFiles.Rows.Count)
                Dim InputFile As String = System.IO.Path.Combine(DataFolder, CType(DataFiles.Rows(MyCounter)("FileName"), String))
                Dim OutputFile As String = System.IO.Path.Combine(OutputFolder, CType(DataFiles.Rows(MyCounter)("NewFileName"), String))
                If System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(OutputFile)) = False Then
                    System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(OutputFile))
                End If
                System.IO.File.Copy(InputFile, OutputFile)
            Next
            Me.Cursor = Cursors.Default
            ResetStatusLine()
            If MsgBox("Files have been copied and renamed into order of their date/time stamps." & vbNewLine & vbNewLine & "Please run the audacity recovery tool from http://manual.audacityteam.org/man/recovering_crashes_manually.html" & vbNewLine & vbNewLine & "Do you like to open this website, now?", vbInformation + vbYesNoCancel) = MsgBoxResult.Yes Then
                System.Diagnostics.Process.Start("http://manual.audacityteam.org/man/recovering_crashes_manually.html")
            End If
        Catch ex As Exception
            Me.Cursor = Cursors.Default
            ResetStatusLine()
            MsgBox("Files operations failed: " & vbNewLine & ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub ResetStatusLine()
        Me.ToolStripStatusLabel.Text = "Ready"
        Me.ToolStripProgressBar.Visible = False
        Me.ToolStripProgressBar.Value = 1
        Me.ToolStripProgressBar.Maximum = 1
        My.Application.DoEvents()
    End Sub

    Private Sub ShowProgress(text As String, progressValue As Integer, progressMax As Integer)
        Me.ToolStripStatusLabel.Text = text & " " & progressValue.ToString & "/" & progressMax.ToString
        If Me.ToolStripProgressBar.Value > progressMax Then Me.ToolStripProgressBar.Value = 1
        Me.ToolStripProgressBar.Maximum = progressMax
        Me.ToolStripProgressBar.Value = progressValue
        Me.ToolStripProgressBar.Visible = True
        My.Application.DoEvents()
    End Sub

    ''' <summary>
    ''' Ensure date/time stamp display format with seconds and milli-seconds
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DataGridView1_DataSourceChanged(sender As Object, e As EventArgs) Handles DataGridView.DataSourceChanged
        If Me.DataGridView.DataSource IsNot Nothing AndAlso Me.DataGridView.Columns.Contains("DateTimeStamp") Then
            Me.DataGridView.Columns("DateTimeStamp").DefaultCellStyle.Format = "yyyy-MM-dd HH: mm:ss.fff"
        End If
        If Me.DataGridView.DataSource IsNot Nothing AndAlso Me.DataGridView.Columns.Contains("FileSize") Then
            Me.DataGridView.Columns("FileSize").DefaultCellStyle.Format = "#,##0"
        End If
        If Me.DataGridView.DataSource IsNot Nothing AndAlso Me.DataGridView.Columns.Contains("SummarizedSize") Then
            Me.DataGridView.Columns("SummarizedSize").DefaultCellStyle.Format = "#,##0"
        End If
    End Sub

End Class

