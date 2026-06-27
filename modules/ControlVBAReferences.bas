Attribute VB_Name = "ControlVBAReferences"
Option Compare Database
Option Explicit

Sub AddReference(RefName As String, RefFilePath As String, Optional AddDelete As String)
    If AddDelete = "" Then AddDelete = "Add"
    Dim refx  As Reference
    Dim IsThere As String
    IsThere = refExists(RefName, False)

    If AddDelete = "Add" And IsThere <> "Y" Then Access.References.AddFromFile (RefFilePath)
    If AddDelete = "Delete" And IsThere = "Y" Then
    For Each refx In References
        If refx.Name = RefName Then Access.References.Remove refx
    Next
    End If
Cleanup:
End Sub

Public Function refExists(RefNamex As String, Optional GetAll As Boolean) As String
Dim ref As Reference
refExists = False
For Each ref In References
    If ref.Name = RefNamex Then
        refExists = "Y"
        If GetAll = True Then refExists = "Broke:" & ref.IsBroken & "|Path:" & ref.FullPath & "|Guid:" & ref.Guid
    End If
Next
End Function
