Imports System.Xml
Imports System.IO

Public Class InputXMLSerializationForm
    'Todo: Collapse & Expand currently implemened upto one level. 
    '      For up to 'n' level use the text comparison
    Public ExpandStart, ExpandEnd, GroupStart, GroupEnd As Long()
    Private Sub XMLSerializationForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim files = Directory.GetFiles(Path.GetDirectoryName(Serializer.FilePath) + "\", "*_History*" + ".xml", SearchOption.AllDirectories)
        For Each file In files
            ListView1.Items.Add(Path.GetFileName(file))
        Next
        LoadForm()
    End Sub
    ''' <summary>
    ''' Expand and Collapse control handled
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SumdataGridView_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles SumdataGridView.CellClick
        'Hide contol "-"
        If e.ColumnIndex = 4 Then
            'MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)
            If e.RowIndex = -1 Then
                For i As Integer = 0 To UBound(ExpandStart) - 1
                    If Not i = ExpandStart(i) Then SumdataGridView.Rows(i).Visible = False
                Next
            Else 'Hide individual group
                For i As Integer = ExpandStart(e.RowIndex) + 1 To ExpandEnd(e.RowIndex)
                    SumdataGridView.Rows(i).Visible = False
                Next
            End If
        End If
        'Expand contol "+"
        If e.ColumnIndex = 3 Then
            If e.RowIndex = -1 Then
                For i As Integer = 0 To UBound(ExpandStart) - 1
                    If Not i = ExpandStart(i) Then SumdataGridView.Rows(i).Visible = True
                Next
            Else 'Expand individual group
                For i As Integer = ExpandStart(e.RowIndex) To ExpandEnd(e.RowIndex)
                    SumdataGridView.Rows(i).Visible = True
                Next
            End If
        End If
    End Sub

    ''' <summary>
    ''' Saves the called Object(AnyObject) into xml file. 
    ''' </summary>
    ''' <remarks>This can be used in the XMLSerializationForm</remarks>
    Public Sub DefaultValues()
        Dim type As Type = Serializer.XmlObject.GetType
        Dim FilePath = Path.Combine(Path.GetTempPath, type.FullName & ".xml")
        Serializer.SaveXmlObject(AnyObject:=Serializer.XmlObject, FilePath:=FilePath)

        Serializer.init()
        'Serializer.SaveXmlObject(Serializer.XmlObject, Serializer.FilePath)
        LoadForm()
    End Sub
    ''' <summary>
    ''' Save the datagridview format form data to xml file. Later xml file
    ''' will be deserialized in to an object
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SaveForm()

        Dim fsStream As New FileStream(Serializer.FilePath, FileMode.Create, FileAccess.Write)
        Dim swWriter As New StreamWriter(fsStream)
        swWriter.WriteLine("<?xml version=" & """" & "1.0" & """" & " encoding=" & """" & "utf-8" & """" & "?>")
        swWriter.WriteLine("<Settings xmlns:xsi=" & """" & "http://www.w3.org/2001/XMLSchema-instance" & """" & " xmlns:xsd=" & """" & "http://www.w3.org/2001/XMLSchema" & """" & ">")

        Dim VariableName As String = Nothing
        Dim OldVariableName As String = Nothing
        Dim Value As String = Nothing
        Dim Type As String = Nothing
        Dim OldType As String = Nothing
        Dim ExpandId As Boolean = False
        Try
            Dim rowCount As Integer = 0
            Dim ColCount As Integer = 0
            For rowCount = 0 To SumdataGridView.Rows.Count - 2
                VariableName = SumdataGridView.Rows.Item(rowCount).Cells.Item(0).Value
                Value = SumdataGridView.Rows.Item(rowCount).Cells.Item(1).Value
                Type = SumdataGridView.Rows.Item(rowCount).Cells.Item(2).Value
                If SumdataGridView.Rows.Item(rowCount).Cells.Item(3).Value = "+/-" Then ExpandId = True Else ExpandId = False
                If Type = "" Then
                    If OldVariableName <> VariableName Then If OldVariableName <> "" And OldType <> "" Then swWriter.WriteLine("  </" & OldVariableName & ">")
                    swWriter.WriteLine("  <" & VariableName & ">" & Value & "</" & VariableName & ">")
                End If
                If Type <> "" Then
                    If OldVariableName <> VariableName Then
                        If OldVariableName <> "" And OldType <> "" Then swWriter.WriteLine("  </" & OldVariableName & ">")
                        swWriter.WriteLine("  <" & VariableName & ">")
                    Else

                    End If
                    swWriter.WriteLine("    <" & Type & ">" & Value & "</" & Type & ">")
                End If
                OldVariableName = VariableName
                OldType = Type
            Next
            swWriter.WriteLine("</Settings>")
            swWriter.Flush()
            swWriter.Close()

            'Copy file to history
            GetLatestHistFile()

        Catch ex As Exception
        End Try

    End Sub
    ''' <summary>
    ''' Gets the history file name, which contains the form data.
    ''' It takes Serializer.FilePath as an arguement
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub GetLatestHistFile()
        'get the History file count
        Dim HistFileDir As String = Path.GetDirectoryName(Serializer.FilePath)
        Dim HistFilePath As String
        Dim HistCount As Integer = 0 'to 9 for 10 level of history
        Dim MaxHistCount As Integer = 50
        Dim HistCountFilePath As String = Path.Combine(HistFileDir, "HistCount.xml")
        Dim files = Directory.GetFiles(HistFileDir + "\", "*_History*" + ".xml", SearchOption.AllDirectories)
        Dim MatchFlag As Boolean = False
        For Each File In files
            If FileCompare(File, Serializer.FilePath) Then
                MatchFlag = True
                Exit Sub
            End If
        Next

        If (File.Exists(HistCountFilePath)) Then
            Dim objReader As New System.IO.StreamReader(HistCountFilePath)
            If objReader.Peek() <> -1 Then
                HistCount = CInt(objReader.ReadLine()) + 1
            End If
            objReader.Close()
        Else
            Dim objWriter As New System.IO.StreamWriter(HistCountFilePath)
            objWriter.Write("0")
            objWriter.Close()
        End If

        HistFilePath = Path.Combine(HistFileDir, Now.ToString("MMMM_dd_yyyy_HH_mm_ss") + "_History" + (HistCount Mod MaxHistCount).ToString + ".xml")
        files = Directory.GetFiles(HistFileDir + "\", "*_History" + HistCount.ToString + ".xml", SearchOption.AllDirectories)

        If MatchFlag = False Then
            'delete the old file
            For Each File In files
                Kill(File)
            Next
            'Copy file to history
            My.Computer.FileSystem.CopyFile(Serializer.FilePath, HistFilePath, True)
            Dim objWriter As New System.IO.StreamWriter(HistCountFilePath) 'Update the history file count
            objWriter.Write(HistCount)
            objWriter.Close()
        End If

    End Sub
    ' This method accepts two strings that represent two files to 
    ' compare. A return value of 0 indicates that the contents of the files
    ' are the same. A return value of any other value indicates that the 
    ' files are not the same.
    Private Function FileCompare(ByVal file1 As String, ByVal file2 As String) As Boolean
        Dim file1byte As Integer
        Dim file2byte As Integer
        Dim fs1 As FileStream
        Dim fs2 As FileStream

        ' Determine if the same file was referenced two times.
        If (file1 = file2) Then
            ' Return 0 to indicate that the files are the same.
            Return True
        End If

        ' Open the two files.
        fs1 = New FileStream(file1, FileMode.Open)
        fs2 = New FileStream(file2, FileMode.Open)

        ' Check the file sizes. If they are not the same, the files
        ' are not equal.
        If (fs1.Length <> fs2.Length) Then
            ' Close the file
            fs1.Close()
            fs2.Close()

            ' Return a non-zero value to indicate that the files are different.
            Return False
        End If

        ' Read and compare a byte from each file until either a
        ' non-matching set of bytes is found or until the end of
        ' file1 is reached.
        Do
            ' Read one byte from each file.
            file1byte = fs1.ReadByte()
            file2byte = fs2.ReadByte()
        Loop While ((file1byte = file2byte) And (file1byte <> -1))

        ' Close the files.
        fs1.Close()
        fs2.Close()

        ' Return the success of the comparison. "file1byte" is
        ' equal to "file2byte" at this point only if the files are 
        ' the same.
        Return ((file1byte - file2byte) = 0)
    End Function
    ''' <summary>
    ''' Load the xml of an object converted using xml serialization
    ''' into Datagridview. Includes interactive control such as expand and collapse,
    ''' Color the cell for easy view.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub LoadForm(Optional ByVal filepath As String = "")
        'Initialize the variables
        Dim row As String()

        If filepath = "" Then filepath = Serializer.FilePath

        Dim reader As XmlTextReader = New XmlTextReader(filepath)
        Dim VariableName As String = Nothing
        Dim Value As String = Nothing
        Dim Type As String = Nothing
        Dim ExpandId As Boolean = False
        Dim columnIndex As Integer = 1
        Dim rowIndex As Integer = 1

        Dim ExpandS As Long = 0
        Dim ExpandE As Long = 0
        Dim CollapseS As Long = 0
        Dim CollapseE As Long = 0
        Dim RowCount As Long = 0

        'Clear the old data
        SumdataGridView.Rows.Clear()

        'Adjust the size of the cells automatically
        SumdataGridView.AutoSizeRowsMode = _
        DataGridViewAutoSizeColumnsMode.AllCells

        SumdataGridView.ColumnCount = 5
        SumdataGridView.Columns(0).Name = "Variable"
        SumdataGridView.Columns(1).Name = "Value"
        SumdataGridView.Columns(2).Name = "Type"
        SumdataGridView.Columns(3).Name = "+"
        SumdataGridView.Columns(4).Name = "-"
        Dim bcol As DataGridViewButtonColumn = New DataGridViewButtonColumn
        bcol.HeaderText = "Controls"
        bcol.Text = "Click here"
        bcol.Name = "btnClickMe"
        bcol.UseColumnTextForButtonValue = True
        SumdataGridView.Columns.Add(bcol)
        'Skip till the Name of th Object is reached in xml
        Do While (reader.Read())
            If reader.NodeType = XmlNodeType.Element Then 'beginning of element.
                Exit Do
            End If
        Loop

        Dim DataTypeSupported As String = "int:double:string:char:long int:short int:uint:long double"
        Do While (reader.Read())
            Select Case reader.NodeType
                Case XmlNodeType.Element 'beginning of element.
                    If VariableName <> Nothing Then
                        ExpandId = True 'If its an array Set ExpandId
                        If InStr(DataTypeSupported, reader.Name) > 0 Then
                            Type = reader.Name
                        Else
                            VariableName = VariableName & "." & reader.Name
                        End If
                    Else
                        VariableName = reader.Name 'Assign the Variable Name
                        ExpandId = False 'If its Initial value of array or not an array
                    End If
                Case XmlNodeType.Text 'text in each element.
                    Value = reader.Value
                    If ExpandId = True Then
                        row = New String() {VariableName, Value, Type, "+", "-"}
                    Else
                        row = New String() {VariableName, Value, Type, "", ""}
                    End If
                    SumdataGridView.Rows.Add(row)
                    RowCount = RowCount + 1
                Case XmlNodeType.EndElement 'end of element.
                    If Not InStr(DataTypeSupported, reader.Name) > 0 Then
                        Try : VariableName = Microsoft.VisualBasic.Left(VariableName, Len(VariableName) - 1 - Len(reader.Name)) : Catch : End Try  'Remove the last Tag and "."
                    End If
                    If VariableName = reader.Name Then
                        VariableName = Nothing
                        Type = Nothing
                    End If
            End Select
        Loop
        ReDim Preserve ExpandStart(RowCount)
        ReDim Preserve ExpandEnd(RowCount)
        ReDim Preserve GroupStart(RowCount)
        ReDim Preserve GroupEnd(RowCount)
        Dim Toggle As Boolean = True
        'Create expand and collapse control
        For i As Integer = 0 To RowCount - 1
            If Not SumdataGridView.Rows(i).Cells(0).Value = SumdataGridView.Rows(i + 1).Cells(0).Value Then
                For j As Integer = ExpandS To i
                    If j > ExpandS Then SumdataGridView.Rows(j).Visible = False
                    ExpandStart(j) = ExpandS
                    ExpandEnd(j) = i
                    If Toggle Then SumdataGridView.Rows(j).DefaultCellStyle.BackColor = Color.Cyan
                Next
                ExpandS = i + 1
                Toggle = Not Toggle
            End If
        Next
        ' Resize the master DataGridView columns to fit the newly loaded data.
        SumdataGridView.AutoResizeColumns()
        reader.Close()
    End Sub

    Private Sub DefaultButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        RichTextBox1.Text = "To get default values, enter the Values inside the class initialization"
    End Sub

    Private Sub LoadButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        RichTextBox1.Text = "Load your own xml file"
    End Sub
    Private Sub ContinueButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        RichTextBox1.Text = "Saves the data to default XML and continues"
    End Sub

    Private Sub ExitButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        RichTextBox1.Text = "Exit the Program"
    End Sub

    Private Sub SaveButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        RichTextBox1.Text = "Saves the data to default XML"
    End Sub

    Private Sub SumdataGridView_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles SumdataGridView.CellContentClick
        If e.RowIndex >= SumdataGridView.RowCount - 1 Then Exit Sub
        If (e.ColumnIndex = 5) Then
            If (InStr(SumdataGridView.Rows(e.RowIndex).Cells(0).Value.ToString, "Path") > 0) Then
                Using fold As New OpenFileDialog
                    fold.Filter = "All files (*.*)|*.*"
                    fold.Title = "Select file"
                    If fold.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        'fold.FilterIndex = 2
                        fold.RestoreDirectory = True
                        SumdataGridView.Rows(e.RowIndex).Cells(1).Value = fold.FileName
                    End If
                End Using
            End If
        End If
        If (e.ColumnIndex = 2) Then
            If (SumdataGridView.Rows(e.RowIndex).Cells(2).Value.ToString = "") Then
                SumdataGridView.Rows(e.RowIndex).Cells(2).Value = " "
            End If
        End If
    End Sub

    Private Sub LoadToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadToolStripMenuItem.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.Title = "Select Application Configeration Files Path"
        fd.InitialDirectory = "C:\"
        fd.Filter = "All files (*.*)|*.*|Xml files (*.xml)|*.xml"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True
        If fd.ShowDialog() = DialogResult.OK Then
            Serializer.FilePath = fd.FileName
        End If
        LoadForm()

    End Sub

    Private Sub DefaultToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DefaultToolStripMenuItem.Click
        DefaultValues()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        SaveForm()
    End Sub

    Private Sub MenuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuToolStripMenuItem.Click
        SaveForm()
        Me.Close()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        File.Create(System.IO.Path.GetTempPath & "terminate.txt").Dispose()
        Me.Close()
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        Dim filePath As String = Path.Combine(Path.GetDirectoryName(Serializer.FilePath), ListView1.FocusedItem.SubItems(0).Text)

        LoadForm(filePath)

    End Sub

    Private Sub SumdataGridView_DataSourceChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles SumdataGridView.CellValueChanged
        ''''''****cb****
        If (e.ColumnIndex = 1) Then
            If (SumdataGridView.CurrentCell.Value = "") Then
                SumdataGridView.CurrentCell.Value = "null"
            End If
        End If
        ' If IsNothing(SumdataGridView.CurrentCell) Then Exit Sub
        ' If (SumdataGridView.CurrentCell.Value = Nothing) Then SumdataGridView.CurrentCell.Value = "null"
        ' If (SumdataGridView.CurrentCell.Value.ToString = "") Then SumdataGridView.CurrentCell.Value = "null"

    End Sub
End Class