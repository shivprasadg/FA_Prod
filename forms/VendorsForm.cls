Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim RST As DAO.Recordset
Dim SQLz, Msgx, XLS As String

Private Sub AddVendor_Click()
DoCmd.GotoRecord , , acNewRec
Call Unlockit
End Sub

Public Sub SaveMe_Click()

Msgx = "Error:"
If Nz(Me.VendorName, "") = "" Then Msgx = Msgx & vbNewLine & " No Vendor Name"
If Nz(Me.VendorSNm, "") = "" Then Msgx = Msgx & vbNewLine & " No Vendor Short Name"
If Msgx <> "Error:" Then
MsgBox Msgx, vbCritical, "Missing data"
Exit Sub
End If
DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub CloseMe_Click()
Call SaveMe_Click
If Msgx <> "Error:" Then Exit Sub
DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Current()
Call LockIt
End Sub

Private Sub Form_Load()
Me.SearchVendor = ""
End Sub

Private Sub SearchVendor_Change()
SQLz = "SELECT VendorID, VendorSNm, VendorName FROM Vendors " & _
       "WHERE VendorObsolete=No AND VendorSNm Like '*" & Me.SearchVendor.Text & "*' " & _
       "ORDER BY VendorName;"
Me.VendorList.RowSource = SQLz
End Sub

Private Sub VendorList_Click()
Me.Form.FilterOn = False
Set RST = Me.RecordsetClone
       RST.FindFirst "VendorId = " & Me.VendorList
       Me.Bookmark = RST.Bookmark
    Set RST = Nothing
End Sub

Private Sub Unlockit()
Me.VendorName.Locked = False
Me.VendorSNm.Locked = False
Me.VendorObsolete.Locked = False
Me.ObsoleteDate.Locked = False
End Sub
Private Sub LockIt()
Me.VendorName.Locked = True
Me.VendorSNm.Locked = True
Me.VendorObsolete.Locked = True
Me.ObsoleteDate.Locked = True
End Sub

Private Sub VendorName_DblClick(Cancel As Integer)
Call Unlockit
End Sub

Private Sub VendorObsolete_DblClick(Cancel As Integer)
Call Unlockit
End Sub

Private Sub VendorSNm_DblClick(Cancel As Integer)
Call Unlockit
End Sub

Private Sub VensorODt_DblClick(Cancel As Integer)
Call Unlockit
End Sub

Private Sub ExcelVendorsUnits_Click()
    SQLz = "SELECT * FROM VendorUnits;"
    XLS = "|A,BL|"
    SendReport2Excel SQLz, "Units", "Vendor (" & Me.VendorSNm & ") Units", True, XLS, 0, True
End Sub
