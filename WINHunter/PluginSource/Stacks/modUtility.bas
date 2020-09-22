Attribute VB_Name = "modUtility"
Public sObjSelections As String

Public Sub GetObjects()
Dim SXML        As New CGoXML
Dim i           As Integer
Dim sObjName    As String
Dim sClsName    As String

    SXML.Initialize (pavAUTO)
    'START INITIAL FILE TEMPLATE
    Call SXML.OpenFromFile(App.Path & "\plugin.xml")
    If SXML.NodeCount("/PLUGINS/PLUGIN") > 0 Then
        For i = 0 To SXML.NodeCount("/PLUGINS/PLUGIN") - 1
            'PluginXML.OpenFromString SXML.ReadNodeXML("/PLUGINS/PLUGIN[" & i & "]")
            'use late binding to create objects
            'this way, it doesnt matter when the object was created
            'or what the object's registry key is
            sObjName = SXML.ReadNode("/PLUGINS/PLUGIN[" & i & "]/OBJECT_NAME")
            sClsName = SXML.ReadNode("/PLUGINS/PLUGIN[" & i & "]/CLASS_NAME")
            Select Case SXML.ReadNode("/PLUGINS/PLUGIN[" & i & "]/TYPE")
                Case "SELECTIONS"
                    'CreateObject(sObjSelections)
                    sObjSelections = sObjName & "." & sClsName
            End Select
        Next
    End If

End Sub

