Attribute VB_Name = "PowerShellFunctions"
Option Compare Database
Option Explicit
Public Function Last30MinChahges()
Dim StartTime As Date
Dim EndTime As Date

StartTime = Date & " " & DatePart("h", Now()) & ":00:00 " & IIf(DatePart("h", Now()) < 12, "AM", "PM")
EndTime = Date & " " & DatePart("h", Now()) & ":30:00 " & IIf(DatePart("h", Now()) < 12, "AM", "PM")
End Function
Public Function AllUnitsAccepted(GroupIDx As Long)

    Dim msg1 As String
    Dim ClientX, Schx, Grpx, AcceptDateX, SQLy, BLCDBasis As String
    Dim ClientIDint, SchIDInt As Integer

    Schx = DLookup("Schedule", "vw_SixKeys", "SGrpID=" & GroupIDx)
    Grpx = DLookup("UnitGroup", "vw_SixKeys", "SGrpID=" & GroupIDx)
    ClientX = DLookup("ClientShortName", "vw_SixKeys", "SGrpID=" & GroupIDx)
    AcceptDateX = DMax("UnitAcceptDate", "Units", "GroupRef=" & GroupIDx)
    ClientIDint = DLookup("ClientID", "vw_SixKeys", "SGrpID=" & GroupIDx)
    SchIDInt = DLookup("ClientID", "vw_SixKeys", "SGrpID=" & GroupIDx)
    BLCDBasis = DLookup("StartsOn", "Lookup_LeaseStartsOn", "ID=" & DLookup("SgrpCommInd", "Schgrp", "SGrpID=" & GroupIDx))

    'make email message
    msg1 = msg1 & "<html><body><p style='color:Navy;font-family:Segoe UI;font-size:20px;'>All Units Accepted</p>"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'>To: Docs Dept.<br /><br />"
    msg1 = msg1 & "Client: <B>" & ClientX & "</B><br />"
    msg1 = msg1 & "Master Lease: <B>" & DLookup("MLNo", "vw_SixKeys", "SGrpID=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "Schedule: <B>" & Schx & "</B><br />"
    msg1 = msg1 & "Group: <B>" & Grpx & "</B><br />"
    msg1 = msg1 & "BLCD Basis:<B> " & BLCDBasis & "</B><br />"
    msg1 = msg1 & "BLCD:<B> " & DLookup("BLCD", "SchGrp", "SGrpID=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "Units In Group: <B>" & DLookup("UnitsInGroup", "SchGrp", "SGrpID=" & GroupIDx) & "</B><br />"
    msg1 = msg1 & "All Units Accepted As Of: <B>" & AcceptDateX & "</B><br />"
    msg1 = msg1 & "User who made change: <B>" & FAUserFullName & "</B><br /><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:8px'><br/>[End Msg] "
    msg1 = msg1 & "Email Created: " & Date & " (via ATLAAS Core Data Manager).<br/><br />"
    msg1 = msg1 & "<p style='font-family:Segoe UI;font-size:12px'><br/>"
    msg1 = msg1 & "</body></html>"
    Call SendPowerShellEmail("@(""<ljohnson@fleetadvantage.com>"")", "All Units Accepted", msg1)
End Function
