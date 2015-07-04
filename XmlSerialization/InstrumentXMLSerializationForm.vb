Imports System.Xml
Imports System.IO
Imports System.Xml.Serialization

Public Class InstrumentXMLSerializationForm

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        LoadForm()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub InstrumentXMLSerializationForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    End Sub
    Public ExpandStart, ExpandEnd, GroupStart, GroupEnd As Long()

    Private Sub LoadButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadButton.Click
        SumdataGridView.Rows.Clear()
        LoadForm()
    End Sub
    Private Sub AutoFillButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoFillButton.Click
        DefaultValues()
    End Sub
    ''' <summary>
    ''' Saves the called Object(AnyObject) into xml file. 
    ''' </summary>
    ''' <remarks>This can be used in the XMLSerializationForm</remarks>
    Public Sub DefaultValues()
        RichTextBox1.Text = "This Feature is coming soon!"
    End Sub
    ''' <summary>
    ''' Save the datagridview format form data to xml file. Later xml file
    ''' will be deserialized in to an object
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SaveForm()

        Dim RowCount As Integer = 0
        Dim FileDir As String = Directory.GetCurrentDirectory() & "\EquipmentConfig\"

        Dim fileNames = My.Computer.FileSystem.GetFiles(
        FileDir, FileIO.SearchOption.SearchTopLevelOnly, "*.xml")
        Dim InstrInfo As New InstrumentDetails
        For Each Row As Object In SumdataGridView.Rows

            InstrInfo = New InstrumentDetails 'Clear InstrInfo for every Instrument/Test Address write

            Dim fileName As String = "InstrCount" & Row.Cells.Item(2).Value & " " & Row.Cells.Item(0).Value & " Com.SiLabs.Timing.Instrument." & Row.Cells.Item(1).Value & ".xml"
            Dim FilePath As String = FileDir & fileName

            If Row.Cells.Item(3).Value = "" Then 'ComType
                Row.Cells.Item(3).Selected = True
                SumdataGridView.Select()
                RichTextBox1.Text = "Save Failed, Enter the data in the selected Cell to continue!"
                Exit Sub
            End If
            InstrInfo.ComType = Row.Cells.Item(3).Value
            If Row.Cells.Item(4).Value = "" Then 'Address
                Row.Cells.Item(4).Selected = True
                SumdataGridView.Select()
                RichTextBox1.Text = "Save Failed, Enter the data in the selected Cell to continue!"
                Exit Sub
            End If
            InstrInfo.Address = Row.Cells.Item(4).Value

            'Serialize object to a text file.
            Dim objStreamWriter As New StreamWriter(FilePath)
            Dim x As New XmlSerializer(InstrInfo.GetType)
            x.Serialize(objStreamWriter, InstrInfo)
            objStreamWriter.Close()
        Next

    End Sub
    Class InstrumentDetails
        Public ComType As String = "0"
        Public Address As String = ""

    End Class
    ''' <summary>
    ''' Load the xml of an object converted using xml serialization
    ''' into Datagridview. Includes interactive control such as expand and collapse,
    ''' Color the cell for easy view.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub LoadForm()
        Dim EnableForm As Boolean = False
        Dim RowCount As Integer = 0
        Dim FileDir As String = Directory.GetCurrentDirectory() & "\EquipmentConfig\"
        Dim temp

        Dim fileNames = My.Computer.FileSystem.GetFiles(
        FileDir, FileIO.SearchOption.SearchTopLevelOnly, "*.xml")
        Dim InstrInfo As New InstrumentDetails
        For Each fileName As String In fileNames
            Dim objStreamReader As New StreamReader(fileName)
            Dim serializedSettings As New XmlSerializer(InstrInfo.GetType) 'open deserializer
            InstrInfo = Nothing
            InstrInfo = CType(serializedSettings.Deserialize(objStreamReader), InstrumentDetails) 'deserialize
            objStreamReader.Close() 'close settings file

            Dim FileStr As String = Path.GetFileName(fileName)
            temp = Split(FileStr, " ")
            SumdataGridView.Rows.Add(1)
            RowCount = RowCount + 1

            SumdataGridView.Rows.Item(RowCount - 1).Cells.Item(0).Value = temp(1).ToString
            temp(2) = Replace(temp(2).ToString, "Com.SiLabs.Timing.Instrument.", "")
            temp(2) = Replace(temp(2).ToString, ".xml", "")
            SumdataGridView.Rows.Item(RowCount - 1).Cells.Item(1).Value = temp(2).ToString
            SumdataGridView.Rows.Item(RowCount - 1).Cells.Item(2).Value = Replace(temp(0).ToString, "InstrCount", "")
            Select Case InstrInfo.ComType
                Case "GPIB"
                    SumdataGridView.Rows.Item(RowCount - 1).Cells.Item(3).Value = "GPIB"
                Case "USB"
                    SumdataGridView.Rows.Item(RowCount - 1).Cells.Item(3).Value = "USB"
                Case "RS232"
                    SumdataGridView.Rows.Item(RowCount - 1).Cells.Item(3).Value = "RS232"
                Case "USBX"
                    SumdataGridView.Rows.Item(RowCount - 1).Cells.Item(3).Value = "USBX"
                Case "LAN"
                    SumdataGridView.Rows.Item(RowCount - 1).Cells.Item(3).Value = "LAN"
            End Select
            SumdataGridView.Rows.Item(RowCount - 1).Cells.Item(4).Value = InstrInfo.Address
            If InstrInfo.Address = "" Or InstrInfo.ComType = "" Then EnableForm = True
        Next

        'Check it is called from Main.UpdateTestSelect()
        Dim stackInfo As String = Environment.StackTrace.ToString()
        temp = Split(stackInfo, vbCrLf)
        Dim InstrName As String = ""

        InstrName = temp(4).ToString
        'If called from UpdateTestSelect then try to close
        If EnableForm = False And InStr(InstrName, "UpdateTestSelect") > 0 Then Me.Close()
    End Sub


    Private Sub DefaultButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles AutoFillButton.MouseEnter
        RichTextBox1.Text = "To get default values, enter the Values inside the class initialization"
    End Sub

    Private Sub LoadButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles LoadButton.MouseEnter
        RichTextBox1.Text = "Load your own xml file"
    End Sub
    Private Sub ContinueButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        RichTextBox1.Text = "Saves the data to default XML and continues"
    End Sub

    Private Sub ExitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitButton.Click
        File.Create(System.IO.Path.GetTempPath & "terminate.txt").Dispose()
        Me.Close()
    End Sub
    Private Sub ExitButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExitButton.MouseEnter
        RichTextBox1.Text = "Exit the Program"
    End Sub

    Private Sub SaveButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveButton.Click
        SaveForm()
    End Sub
    Private Sub SaveButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles SaveButton.MouseEnter
        RichTextBox1.Text = "Saves the data to default XML"
    End Sub

End Class