Imports System.IO
Imports System.Xml.Serialization

Public Class Serializer
    Public Shared FilePath As String
    Public Shared XmlObject As Object
    Public Shared frm As InputXMLSerializationForm
    ''' <summary>
    ''' Converts an Object in to xml and generates dynamic GUI.
    ''' Once users selects 
    ''' </summary>
    ''' <param name="AnyObject"></param>
    ''' <remarks></remarks>
    Sub New(ByRef AnyObject As Object)
        XmlObject = AnyObject
        init()
        AnyObject = XmlObject
        'Dim type As Type = AnyObject.GetType
        'FilePath = Directory.GetCurrentDirectory() & "\..\..\" & type.FullName & ".xml"
        'If Not System.IO.File.Exists(FilePath) Then 'and AnyObject file exists
        '    SaveXmlObject(AnyObject, FilePath)
        'Else
        '    LoadXmlObject()
        '    AnyObject = XmlObject
        'End If
        ''Dim frm As New XMLSerializationForm()
        'frm = New InputXMLSerializationForm()
        ''frm.ShowDialog()
    End Sub
    Public Shared Sub init()
        Dim AnyObject = XmlObject
        Dim type As Type = AnyObject.GetType
        FilePath = Path.Combine(Path.GetTempPath, type.FullName & ".xml")
        If Not System.IO.File.Exists(FilePath) Then 'and AnyObject file exists
            SaveXmlObject(AnyObject, FilePath)
        Else
            LoadXmlObject()
            AnyObject = XmlObject
        End If
        'Dim frm As New XMLSerializationForm()
        frm = New InputXMLSerializationForm()
        'frm.ShowDialog()

    End Sub
    Public Sub showForm(Optional ByVal message As String = "")
        If Not message = "" Then frm.RichTextBox1.Text = message
        frm.ShowDialog()
        LoadXmlObject()
    End Sub
    ''' <summary>
    ''' This will load the AnyObject object members with xml file
    ''' being passed in command line
    ''' </summary>
    ''' <remarks>Deserialises the values using XmlSerialiser</remarks>
    Public Shared Sub LoadXmlObject()
        'initialize test AnyObject from xml file specified on the command line
        If System.IO.File.Exists(FilePath) Then 'and AnyObject file exists
            Dim serializedObject As New XmlSerializer(XmlObject.GetType) 'open deserializer
            XmlObject = Nothing 'New Object 'clear manual test AnyObject
            Dim objStreamReader As New StreamReader(FilePath) 'open AnyObject file
            Try
                XmlObject = CType(serializedObject.Deserialize(objStreamReader), Object) 'deserialize
            Catch ex As Exception

            End Try
            objStreamReader.Close() 'close AnyObject file
        End If
    End Sub
    ''' <summary>
    ''' This will save any object members in to a xml file
    ''' </summary>
    ''' <param name="AnyObject">Name of class to be saved</param>
    ''' <remarks>Deserialises the values using XmlSerialiser</remarks>
    Public Shared Sub SaveXmlObject(ByRef AnyObject As Object, ByVal FilePath As String)
        'initialize test AnyObject from xml file specified on the command line
        'If Environment.GetCommandLineArgs.Length > 1 Then 'if AnyObject file specified on command line
        Dim serializedObject As New XmlSerializer(AnyObject.GetType) 'open deserializer
        Dim objStreamWriter As New StreamWriter(FilePath) 'open AnyObject file
        serializedObject.Serialize(objStreamWriter, AnyObject)
        objStreamWriter.Close() 'close AnyObject file
        'End If
    End Sub
End Class
