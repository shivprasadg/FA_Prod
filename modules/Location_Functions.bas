Attribute VB_Name = "Location_Functions"
Option Compare Database
Option Explicit
Dim SQLz As String
Dim RST As DAO.Recordset

Public Function CheckForUnitsAtAddress(LocationIdx As Long) As String
Dim Parked, Reg, TITLED, Delv As Integer
Dim Tempx As String
Parked = DCount("UnitID", "Units", "UnitGarageLocationRef=" & LocationIdx & "")
Reg = DCount("UnitID", "Units", "UnitRegLocationRef=" & LocationIdx & "")
TITLED = DCount("UnitID", "Units", "UnitTitleLocationRef=" & LocationIdx & "")
Delv = DCount("UnitID", "Units", "UnitDeliveryLocationRef=" & LocationIdx & "")
If Parked + Reg + TITLED + Delv > 0 Then
Tempx = "There are Units are that address" & vbNewLine
If Parked > 0 Then Tempx = Tempx & Parked & " Parked" & vbNewLine
If Reg > 0 Then Tempx = Tempx & Reg & " Reg" & vbNewLine
If TITLED > 0 Then Tempx = Tempx & TITLED & " Titled" & vbNewLine
If Delv > 0 Then Tempx = Tempx & Delv & " Delv" & vbNewLine
'CheckForUnitsAtAddress = "There are " & Parked + Reg + Titled + Delv & " units assocated with that address"
Else
Tempx = "No Units"
End If
CheckForUnitsAtAddress = Tempx
End Function
Public Function UnitsAtAddress(LocationIdx As Long) As String
Dim Parked, Reg, TITLED, Delv As Integer
Parked = DCount("UnitID", "Units", "UnitGarageLocationRef=" & LocationIdx & "")
Reg = DCount("UnitID", "Units", "UnitRegLocationRef=" & LocationIdx & "")
TITLED = DCount("UnitID", "Units", "UnitTitleLocationRef=" & LocationIdx & "")
Delv = DCount("UnitID", "Units", "UnitDeliveryLocationRef=" & LocationIdx & "")
UnitsAtAddress = Parked + Reg + TITLED + Delv
End Function

Public Function ParkedLocationMove(OldLocationIdx As Long, NewLocationIdx As Long, AssetIDx As Long _
, Optional ThePersonX As String _
, Optional TheEmailX As String _
, Optional ThePhoneNumX As Variant _
, Optional SetTheParkedLocation As Boolean)

    Dim Contactx As String
    Dim AsOfDatex As Date
    Dim Emailx As String
    Dim Phonex As Variant
    Dim CTempx, Tempx As Variant
    Set RST = CurrentDb.OpenRecordset("Select * From Contacts WHERE cLocationID=" & OldLocationIdx & ";", dbOpenSnapshot)
InputContact:
    If RST.RecordCount > 0 Then CTempx = RST!contactFirst & " " & RST!contactLast Else CTempx = Nz(ThePersonX, " the person")
    Contactx = InputBox("Contact Person at pickup location for this move?" & vbNewLine & "This is for the Bennet Request ('stop' to cancel)", "Contact Person for move...", CTempx)
    If Contactx = "" Then GoTo InputContact
    If Contactx = "Stop" Or Contactx = "Cancel" Then GoTo Cancel

InputEmail:
    If RST.RecordCount > 0 Then Tempx = RST!contactEmail Else Tempx = Nz(TheEmailX, "")
    Emailx = InputBox("Email for " & Tempx & " for this move?" & vbNewLine & "This is for the Bennet Request ('stop' to cancel)", "Person's Email for move...", Tempx)
    If Emailx = "" Then GoTo InputEmail
    If Emailx = "Stop" Or Emailx = "Cancel" Then GoTo Cancel
InputPhone:
    If RST.RecordCount > 0 Then Tempx = Nz(RST!contactPrimaryPhone, 0) Else Tempx = Nz(ThePhoneNumX, 0)
    Phonex = InputBox("Phone Num for " & CTempx & " for this move?" & vbNewLine & "This is for the Bennet Request ('stop' to cancel)", "Person's Phone for move...", Tempx)
    If Contactx = "" Then GoTo InputPhone
    If Phonex = "Stop" Or Phonex = "Cancel" Then GoTo Cancel
InputDate:
    On Error GoTo InputDate
    AsOfDatex = InputBox("What is the Actual Request Date the Move?" & vbNewLine & "use date format mm/dd/yyyy ('stop' to cancel)", "Date to Use", Format(Date, "mm/dd/yyyy"))
    If CStr(AsOfDatex) = "Stop" Or CStr(AsOfDatex) = "Cancel" Then GoTo Cancel
    If Not IsDate(AsOfDatex) Then GoTo InputDate
    'On Error GoTo FixMe
    SQLz = "INSERT INTO UnitMoveOrder (mAssetID,mOldLocationID,mNewLocationID,mStartDate,mVendorID,mOrderNum,mStatus,mContact,mPhone,mEmail,mFaUserID,mNote) " & _
        "SELECT " & AssetIDx & ", " & OldLocationIdx & ", " & NewLocationIdx & ", #" & AsOfDatex & "#, 517,'TBD','Not Ordered', '" & Contactx & "','" & Phonex & "', '" & Emailx & "'," & FAUserIDx & ",'Move Order created by " & FAUserName & "';"
    DoCmd.SetWarnings False
    DoCmd.RunSQL (SQLz)
    DoCmd.RunSQL ("Update Units Set UnitStatus='EM', LocationParkedID= " & NewLocationIdx & " WHERE UnitID=" & AssetIDx & ";")
    'If SetTheParkedLocation = True Then DoCmd.RunSQL ("Update Units Set LocationParkedID=" & NewLocationIdx & " WHERE UnitID=" & AssetIDx & ";")
    DoCmd.SetWarnings True
Cancel:
    RST.close
    Set RST = Nothing
End Function

Sub LinkBennetOrderTable()
    Dim db As DAO.DataBase
    Dim i As Integer
    Set db = CurrentDb
    On Error Resume Next
    db.TableDefs.Delete "MoveOrderImport"
    On Error GoTo 0
    db.TableDefs.Refresh
    DoCmd.TransferText TransferType:=acLinkDelim, tableName:="MoveOrderImport", _
        fileName:="C:\FAAS\scripts\report.csv", HasFieldNames:=True
    db.close
    Set RST = CurrentDb.OpenRecordset("Select * From UnitMoveOrderUpdate WHERE [Status]='C';", dbOpenSnapshot)
    RST.MoveLast
    RST.MoveFirst
    If RST.RecordCount > 0 Then
        For i = 1 To RST.RecordCount
           SQLz = "Update UnitMoveOrder Set mOrderNum='" & RST![Order#] & "', mEndDate='" & RST!ETADate & "', mStatus='Done' "
           SQLz = SQLz & ",mMiles=" & RST!Miles & ", mNote=mNote & ' - Updated by Atlass Core Import on " & Date & "'  "
           SQLz = SQLz & " WHERE mAssetID=" & RST!UnitID & " AND mStatus <> 'Done';"
           DoCmd.SetWarnings False
           DoCmd.RunSQL (SQLz)

           SQLz = "Update Units Set UnitStatus='EZ' WHERE UnitStatus <> 'EZ' AND UnitID=" & RST!UnitID & ";"
           DoCmd.RunSQL (SQLz)
           DoCmd.SetWarnings True
           RST.MoveNext
        Next i
    End If
    RST.close
    Set RST = Nothing
    Set db = Nothing
End Sub

Sub LinkBennettInvoiceTable()
    Dim db As DAO.DataBase
    Dim i As Integer
    Set db = CurrentDb
    On Error Resume Next
    db.TableDefs.Delete "BennettMoveInvoice"
    On Error GoTo 0
    db.TableDefs.Refresh
    DoCmd.TransferText TransferType:=acLinkDelim, tableName:="BennettMoveInvoice", _
        fileName:="C:\FAAS\scripts\CUSTOMER_518678_Open_INVOICES.csv", HasFieldNames:=True
    db.close
End Sub
