Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SQLz, XLS As String

Private Sub CloseMe_Click()
DoCmd.OpenForm "FAMainMenu"
DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Combo59_Click()
Me.Combo61.Value = ""
Me.Combo61.Requery
Me.Combo63.Value = ""
Me.Combo63.Requery
End Sub
Private Sub Combo61_Click()
Me.Combo63.Value = ""
Me.Combo63.Requery
End Sub
Private Sub Combo63_Click()
Me.Combo63.Requery
End Sub
Private Sub Combo63_GotFocus()
Me.Combo63.Requery
End Sub

Private Sub Command66_Click()



If IsNull(Me.Combo63.Value) Then
MsgBox "please fill out the inputs Customer/Lessee, Master Lease, Schedule Number, and File Type"
Exit Sub

End If

If Me.Combo68 = "Equipment Files" Then

Call transferEQPtoTemp

Dim oExcel As Excel.Application
Dim oWorkBook As Excel.Workbook
Dim oWorksheet As Excel.Worksheet

Set oExcel = New Excel.Application
Set oWorkBook = oExcel.Workbooks.Open("P:\FLEETADVANTAGE\Documentation\FA Title Trust\Equip Files_Sent\Equipment Files Template (Do Not Delete)\FA Title TR_Equip Files_to Vault_Template.xls")
Set oWorksheet = oWorkBook.Worksheets("Equipment_Files") 'The specified worksheet

Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("tbl_TempEQPFiles", dbOpenDynaset, dbSeeChanges) 'The specified query

DoCmd.OpenForm "EFExportProgress", acNormal
Forms!EFExportProgress.Box1.Visible = True
Forms!EFExportProgress.Box2.Visible = False
Forms!EFExportProgress.Box3.Visible = False
Forms!EFExportProgress.Box4.Visible = False
Forms!EFExportProgress.Box5.Visible = False
Forms!EFExportProgress.Box6.Visible = False
Forms!EFExportProgress.Box7.Visible = False
Forms!EFExportProgress.Box8.Visible = False
Forms!EFExportProgress.Box9.Visible = False


oWorksheet.Range("a2:o5000").ClearContents

Forms!EFExportProgress.Box1.Visible = True
Forms!EFExportProgress.Box2.Visible = True
Pause (1)

Forms!EFExportProgress.Box3.Visible = True
'Pause (1)
Forms!EFExportProgress.Box4.Visible = True
Pause (1)
Forms!EFExportProgress.Box5.Visible = True
'Pause (1)
Forms!EFExportProgress.Box6.Visible = True
Pause (1)
Forms!EFExportProgress.Box7.Visible = True
'Pause (1)
Forms!EFExportProgress.Box8.Visible = True
Pause (1)
Forms!EFExportProgress.Box9.Visible = True



oWorksheet.Range("a2").CopyFromRecordset rs
'Pause (1)
Forms!EFExportProgress.Box8.Visible = True
Pause (1)
Forms!EFExportProgress.Box9.Visible = True

rs.close



oWorkBook.Save
oWorkBook.close
oExcel.Quit



Set rs = Nothing
Set oWorksheet = Nothing
Set oWorkBook = Nothing
Set oExcel = Nothing








DoCmd.close acForm, "EFExportProgress"

DoCmd.OpenForm "EQPFilesExport", acNormal

Exit Sub

End If


If Me.Combo68 = "Data Files" Then



Call transferDatatoTemp




Dim oExcel2 As Excel.Application
Dim oWorkBook2 As Excel.Workbook
Dim oWorksheet2 As Excel.Worksheet

Set oExcel2 = New Excel.Application
Set oWorkBook2 = oExcel2.Workbooks.Open("P:\FLEETADVANTAGE\Documentation\FA Title Trust\Data Files_Sent\Data Files Template (Do Not Delete)\FA Title TR_Data_Files_Template.xls")
Set oWorksheet2 = oWorkBook2.Worksheets("Data_Files") 'The specified worksheet

Dim RS2 As DAO.Recordset
Set RS2 = CurrentDb.OpenRecordset("tbl_TempDataFiles", dbOpenDynaset, dbSeeChanges) 'The specified query

DoCmd.OpenForm "EFExportProgress", acNormal
Forms!EFExportProgress.Box1.Visible = True
Forms!EFExportProgress.Box2.Visible = False
Forms!EFExportProgress.Box3.Visible = False
Forms!EFExportProgress.Box4.Visible = False
Forms!EFExportProgress.Box5.Visible = False
Forms!EFExportProgress.Box6.Visible = False
Forms!EFExportProgress.Box7.Visible = False
Forms!EFExportProgress.Box8.Visible = False
Forms!EFExportProgress.Box9.Visible = False


oWorksheet2.Range("a2:o5000").ClearContents

Forms!EFExportProgress.Box1.Visible = True
Forms!EFExportProgress.Box2.Visible = True
Pause (1)

Forms!EFExportProgress.Box3.Visible = True
'Pause (1)
Forms!EFExportProgress.Box4.Visible = True
Pause (1)
Forms!EFExportProgress.Box5.Visible = True
'Pause (1)
Forms!EFExportProgress.Box6.Visible = True
Pause (1)
Forms!EFExportProgress.Box7.Visible = True
'Pause (1)
Forms!EFExportProgress.Box8.Visible = True
Pause (1)
Forms!EFExportProgress.Box9.Visible = True



oWorksheet2.Range("a2").CopyFromRecordset RS2
'Pause (1)
Forms!EFExportProgress.Box8.Visible = True
Pause (1)
Forms!EFExportProgress.Box9.Visible = True

RS2.close



oWorkBook2.Save
oWorkBook2.close
oExcel2.Quit



Set RS2 = Nothing
Set oWorksheet2 = Nothing
Set oWorkBook2 = Nothing
Set oExcel2 = Nothing








DoCmd.close acForm, "EFExportProgress"

DoCmd.OpenForm "DataFilesExport", acNormal

End If



End Sub

Private Sub Command67_Click()
Me.Combo59.Value = ""
Me.Combo61.Value = ""
Me.Combo63.Value = ""
Me.Combo68.Value = ""
End Sub

Sub transferEQPtoTemp()

 Dim strSQL_Delete As String
' Dim strSQL_Append As String
 Dim strTable As String
 strTable = "tbl_TempEQPFiles"
 strSQL_Delete = "Delete * FROM " & "tbl_TempEQPFiles"
 DoCmd.SetWarnings False
 'Delete the data and refresh it from the query
 CurrentDb.Execute strSQL_Delete
DoCmd.OpenQuery "qry_Append_EQPfiles", acViewNormal, acReadOnly

 DoCmd.SetWarnings True
 End Sub


Sub transferDatatoTemp()

 Dim strSQL_Delete As String
' Dim strSQL_Append As String
 Dim strTable As String
 strTable = "tbl_TempDataFiles"
 strSQL_Delete = "Delete * FROM " & "tbl_TempDataFiles"
 DoCmd.SetWarnings False
 'Delete the data and refresh it from the query
 CurrentDb.Execute strSQL_Delete
 DoCmd.OpenQuery "qry_Append_Datafiles", acViewNormal, acReadOnly

 DoCmd.SetWarnings True
 End Sub

Private Sub ExipiringInsExcel_Click()
SQLz = "Select s.ClientShortName as Client"
SQLz = SQLz & ",M.MLNo"
SQLz = SQLz & ",MLInsRec as DateEntered"
SQLz = SQLz & ",MLInsExp as ExpireDate"
SQLz = SQLz & ",iif(Max(MLLiabIns)='S','Self-Insured','Required') As ReqType"
SQLz = SQLz & ",Count(AssetID) as ActiveUnits"
SQLz = SQLz & ",Max(Coalesce(S.OffLeaseProj,  S.LeaseExp))  as ActiveUntil"
SQLz = SQLz & ",M.MLID"
SQLz = SQLz & " From MstrLease as M"
SQLz = SQLz & " Inner Join Vw_SixKeys as s on S.ClientID= M.Client"
SQLz = SQLz & " Where M.MLOrig = 'FA' AND (S.Status ='A' or S.Status='TR')"
SQLz = SQLz & " Group By s.ClientShortName"
SQLz = SQLz & ",M.MLOrig"
SQLz = SQLz & ",M.MLNo"
SQLz = SQLz & ",MLInsRec"
SQLz = SQLz & ",MLInsExp"
SQLz = SQLz & ",M.MLID"
SQLz = SQLz & " Order By ClientShortName"
XLS = "|D,BC|C,TC|F,BC|G,TC|H,HH|"
AddReportLog "ExipiringInsExcel_Click"
SendReport2Excel SQLz, "MstLeases", "Expiring Insurance as of: " & Format(Date, "mm-dd-yy"), True, XLS, 0, True
End Sub

Private Sub OpenUnitLookUp_Click()
DoCmd.OpenForm "UnitReverseLookUpForm"
End Sub

Private Sub ScheduleDashboard_Click()
DoCmd.OpenForm "ScheduleHealthDashboard"
End Sub
