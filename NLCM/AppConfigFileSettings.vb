
Imports System.Configuration
Imports System.Xml


Public Class AppConfigFileSettings
    Public Sub UpdateAppSettings(ByVal KeyName As String, ByVal KeyValue As String)
        '  AppDomain.CurrentDomain.SetupInformation.ConfigurationFile 
        ' This will get the app.config file path from Current application Domain
        Dim XmlDoc As New XmlDocument()
        ' Load XML Document
        XmlDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile)
        ' Navigate Each XML Element of app.Config file
        For Each xElement As XmlElement In XmlDoc.DocumentElement
            If xElement.Name = "appSettings" Then
                ' Loop each node of appSettings Element 
                ' xNode.Attributes(0).Value , Mean First Attributes of Node , 
                ' KeyName Portion
                ' xNode.Attributes(1).Value , Mean Second Attributes of Node,
                ' KeyValue Portion
                For Each xNode As XmlNode In xElement.ChildNodes
                    If Not IsNothing(xNode.Attributes) Then
                        If xNode.Attributes(0).Value = KeyName Then
                            xNode.Attributes(1).Value = KeyValue
                        End If
                    End If
                Next
            End If
        Next
        ' Save app.config file
        XmlDoc.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile)
        XmlDoc = Nothing
    End Sub
End Class