VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "table2slides_steps_workbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' This function exports all the code from the application and the step definition (test code)
'  as text into the applications directory under /source and /step_definitions.
' To use this function a reference to Microsoft Visual Basic for Applications Extensibility must be set
'  and access to the VBA project object model must be granted in Excels settings(Security for MacOS)
Private Sub export_all_projects()
    Dim project_name As String
    Dim project As Variant
    
    'the application under test must use the same name as this file but without ending the name with '_steps'
    project_name = Left(Application.VBE.ActiveVBProject.Name, Len("_steps"))
    For Each project In Application.VBE.VBProjects
        If Left(project.Name, Len(project_name)) = project_name Then
            exportCode project
        End If
    Next
End Sub

Private Sub exportCode(Optional project)

    Dim vbe_source_object As VBComponent
    Dim base_path As String
    Dim sub_path As String
    Dim file_path As String
    Dim path_separator As String
    Dim file_suffix As String

    path_separator = get_path_separator
    base_path = table2slides_steps_workbook.path
    If IsMissing(project) Then Set project = Application.VBE.VBProjects("table2slides_steps")
    For Each vbe_source_object In project.VBComponents
        Select Case vbe_source_object.Type
            Case vbext_ct_StdModule
                file_suffix = "bas"
            Case vbext_ct_ClassModule
                file_suffix = "cls"
            Case vbext_ct_Document
                file_suffix = "cls"
            Case vbext_ct_MSForm
                file_suffix = "frm"
            Case Else
                file_suffix = "txt"
        End Select
        'save the code of the step definition classes in the step_definitions directory instead of the source directory
        If LCase(Right(project.Name, 6)) = "_steps" Then
            sub_path = "step_definitions"
        Else
            sub_path = "source"
        End If
        file_path = base_path & path_separator & sub_path & path_separator & vbe_source_object.Name & "." & file_suffix
        Debug.Print "export code to " & file_path
        #If Mac Then
            'try not to change forms unless Microsoft offers full support for forms on the Mac!
            If Not file_suffix = "frm" Then
                vbe_source_object.Export file_path
            End If
        #Else
            vbe_source_object.Export file_path
        #End If
    Next
End Sub

Private Function get_path_separator() As String
    ' word and excel return path separator via Application.PathSeparator
    '  but this property is missing in Powerpoint

    #If MAC_OFFICE_VERSION >= 15 Then
        'in Office 2016 MAC M$ switched to / as path separator
        get_path_separator = "/"
    #ElseIf Mac Then
        get_path_separator = ":"
    #Else
        get_path_separator = "\"
    #End If
End Function
