Imports System.Windows.Forms
Imports System.IO
Imports DragNDrop

Public Class FormSerilizor
    Public Shared Sub Serialise(c As Control)
        Dim type As Type = c.GetType
        Dim tempFilePath As String = Directory.GetCurrentDirectory() & "\temp" & type.FullName & ".xml"
        Dim FilePath As String = Directory.GetCurrentDirectory() & "\" & type.FullName & ".xml"

        If System.IO.File.Exists(tempFilePath) Then Kill(tempFilePath)
        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter(tempFilePath, True)

        Dim List = New List(Of Control)
        GetChildControls(c, List)
        For Each cntrl As Control In List
            If (TypeOf cntrl Is TextBox) Then
                file.WriteLine(cntrl.Text)
            End If
        Next

        Dim allTxt As New List(Of Control)
        For Each ListBx As ListBox In FindControlRecursive(allTxt, c, GetType(ListBox))
            Dim temp As String = ""
            For Each tempStr In ListBx.Items
                temp = temp & ":" & tempStr
            Next
            file.WriteLine(temp)
        Next
        allTxt = New List(Of Control)
        For Each ListBx As ListView In FindControlRecursive(allTxt, c, GetType(ListView))
            Dim temp As String = ""
            For Each tempStr In ListBx.Items
                temp = temp & ":" & tempStr
            Next
            file.WriteLine(temp)
        Next
        allTxt = New List(Of Control)
        For Each ListBx As DragNDropListView In FindControlRecursive(allTxt, c, GetType(DragNDropListView))
            Dim temp As String = ""
            For i As Integer = 0 To ListBx.Items.Count - 1
                temp = temp & ":" & ListBx.Items(i).Text
                'for (int i = 0; i < ColXaxisDNDLV.Items.Count; i++)
                'Columnfield[i] = ColXaxisDNDLV.Items[i].Text;
            Next
            file.WriteLine(temp)
        Next
        file.Close()
        If Not AreTheyTheSame(tempFilePath, FilePath) Then
            If System.IO.File.Exists(FilePath) Then Kill(FilePath)
            IO.File.Move(tempFilePath, FilePath)
        End If
        If System.IO.File.Exists(tempFilePath) Then Kill(tempFilePath)
    End Sub
    Public Shared Sub DeSerialise(ByVal c As Control)
        Dim type As Type = c.GetType
        Dim FilePath As String = Directory.GetCurrentDirectory() & "\" & type.FullName & ".xml"

        If Not System.IO.File.Exists(FilePath) Then Exit Sub

        Dim file As System.IO.StreamReader
        file = My.Computer.FileSystem.OpenTextFileReader(FilePath)

        Dim List = New List(Of Control)
        GetChildControls(c, List)
        For Each cntrl As Control In List
            If (TypeOf cntrl Is TextBox) Then
                cntrl.Text = file.ReadLine()
            End If
        Next
        Dim allTxt As New List(Of Control)
        For Each ListBx As ListBox In FindControlRecursive(allTxt, c, GetType(ListBox))
            Dim temp() As String = Split(file.ReadLine, ":")
            ListBx.Items.Clear()
            For Each tempStr In temp
                If tempStr <> "" Then ListBx.Items.Add(tempStr)
            Next
        Next
        allTxt = New List(Of Control)
        For Each ListBx As ListView In FindControlRecursive(allTxt, c, GetType(ListView))
            Dim temp() As String = Split(file.ReadLine, ":")
            ListBx.Items.Clear()
            For Each tempStr In temp
                If tempStr <> "" Then ListBx.Items.Add(tempStr)
            Next
        Next
        allTxt = New List(Of Control)
        For Each ListBx As DragNDropListView In FindControlRecursive(allTxt, c, GetType(DragNDropListView))
            Dim temp() As String = Split(file.ReadLine, ":")
            ListBx.Items.Clear()
            For Each tempStr In temp
                If tempStr <> "" Then ListBx.Items.Add(tempStr)
            Next
        Next
        file.Close()
    End Sub
    Public Shared Sub GetChildControls(container As Control, ByRef _list As List(Of Control))
        For Each child As Control In container.Controls
            _list.Add(child)
            If (child.HasChildren) Then
                GetChildControls(child, _list)
            End If
        Next
    End Sub

    Public Shared Sub GetChildListBoxes(container As Control, ByRef _list As List(Of ListBox))
        For Each child As ListBox In container.Controls
            _list.Add(child)
            If (child.HasChildren) Then
                GetChildListBoxes(child, _list)
            End If
        Next
    End Sub
    Public Shared Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, ByVal ctrlType As System.Type) As List(Of Control)
        If parent Is Nothing Then Return list
        If parent.GetType Is ctrlType Then
            list.Add(parent)
        End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType)
        Next
        Return list
    End Function

    Public Shared Function AreTheyTheSame(ByVal File1 As String, _
  ByVal File2 As String) As Boolean
        On Error GoTo ErrorHandler

        If Not IO.File.Exists(File1) Then If Not File.Exists(File2) Then Return True

        Dim A As String = IO.File.ReadAllText(File1)
        Dim B As String = IO.File.ReadAllText(File2)
        If A = B Then Return True
ErrorHandler:
        Return False
    End Function
End Class
