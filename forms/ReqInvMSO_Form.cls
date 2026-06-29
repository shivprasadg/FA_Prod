Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim SQLz As String

Private Sub CloseMe_Click()
DoCmd.close acForm, Me.Form.Name
End Sub

Private Sub Form_Load()
Me.Detail.Visible = False
Me.DaytoDelivery = 15
'SQLz = "SELECT UnitNeedMSOInvoice.CompanyName AS Client, UnitNeedMSOInvoice.FAID, UnitNeedMSOInvoice.PurchaseOrder AS PO, " & _
        "Count(UnitNeedMSOInvoice.AssetID) AS Units, Inv_MSOContactq2.contactEmail " & _
        "FROM UnitNeedMSOInvoice LEFT JOIN Inv_MSOContactq2 ON UnitNeedMSOInvoice.VendorId = Inv_MSOContactq2.VendorID " & _
        "WHERE UnitNeedMSOInvoice.Days2Delivery < 15 " & _
        "GROUP BY UnitNeedMSOInvoice.CompanyName, UnitNeedMSOInvoice.FAID, UnitNeedMSOInvoice.PurchaseOrder, Inv_MSOContactq2.contactEmail;"

SQLz = "SELECT ClientShNm AS Client, [FAID/(PO)] AS FAID, PONum1 AS PO, Count(ClientName) AS UNITS, contactEmail " & _
       "FROM ReqInvMSO WHERE ReqInvMSO.Days2Delivery <= " & Me.DaytoDelivery & " " & _
       "GROUP BY ClientShNm, [FAID/(PO)], PONum1, contactEmail;"
Me.Client.RowSource = SQLz

End Sub
Private Sub DaytoDelivery_Click()
SQLz = "SELECT ClientShNm AS Client, [FAID/(PO)] AS FAID, PONum1 AS PO, Count(ClientName) AS UNITS, contactEmail " & _
       "FROM ReqInvMSO WHERE Days2Delivery <= " & Me.DaytoDelivery & " " & _
       "GROUP BY ClientShNm, [FAID/(PO)], PONum1, contactEmail;"
       Me.Client.RowSource = SQLz
'Me.Client.Requery
End Sub
Private Sub client_Click()
Me.FAID = Me.Client.Column(3)
SQLz = "SELECT * From ReqInvMSO WHERE [FAID/(PO)] = '" & Me.Client.Column(1) & "' AND Days2Delivery <=" & Me.DaytoDelivery & " ;"
Me.Form.RecordSource = SQLz

Me.Detail.Visible = True
Me.MakeMSOEmail.Enabled = True
'MsgBox Me.FAID
SQLz = "SELECT FAID, Datesent, VehicleCount, Index " & _
       "FROM ReqInvMSO_Tracking WHERE FAID ='" & Me.Client.Column(1) & "' AND ReqType = 'MSOCopies' Order BY Datesent;"
Me.ReqMSOTracking.RowSource = SQLz
Me.ReqMSOTracking.Requery
End Sub
Private Sub Form_Close()
DoCmd.OpenForm " FAMainMenu"
End Sub
Private Sub GoHome_Click()
DoCmd.OpenForm " FAMainMenu"
End Sub


Private Sub MakeMSOEmail_Click()
Dim RST As DAO.Recordset
Dim FAIDx, POx, VendorIDx, VendorNameX, Contactx, ContactEmailX, ClientX, subjectx, UserX, UserEmailX, msg1, Msgx, ItemsX, Linkx As String
Dim x, UnitsX As Integer
SQLz = "SELECT * From ReqInvMSO WHERE [FAID/(PO)] = '" & Me.Client.Column(1) & "'AND Days2Delivery <=" & Me.DaytoDelivery & ";"
     Set RST = CurrentDb.OpenRecordset(SQLz, dbOpenSnapshot)
     If RST.RecordCount < 1 Then
         RST.close
     Set RST = Nothing
     MsgBox "No Matching Records, Mail Not Sent", , "No records"
    Else
    RST.MoveFirst
                 FAIDx = [FAID/(PO)]
                 POx = Nz(RST![PONum1], "TBD")
              Contactx = RST!VendorContact
             VendorIDx = RST!VendorID
         ContactEmailX = RST!contactEmail
                 UserX = Forms!FAMainMenu![USERID]
            UserEmailX = Forms!FAMainMenu![USERID] & "@fleetadvantage.net"
              subjectx = "Request for MSO copies | Client Name: " & RST!ClientName & "| PO: " & POx & ""
                UnitsX = RST.RecordCount

              msg1 = ""
              Msgx = "<table border-size:0><font size=2><font face='Calibri'>"
      x = 1
      Do Until RST.EOF

            Msgx = Msgx & "<tr><td align='right'><b>  " & x & "</td>" & _
                          "<td></b>Year:<b> " & RST!MYr & "</b></td> " & _
                          "<td></b>Make:<b>  " & RST!Make & "</td>" & _
                          "<td></b>Model:<b> " & RST!Model & "</td>" & _
                          "<td></b>Type:<b>  " & RST!Type & "</td>" & _
                          "<td></b>Unit VIN:<b> " & RST!unitvin & "</td></tr>"
                          '"<td></b>AssetID:<b> " & RST!AssetID & vbNewLine & "</b></td></tr>"
             RST.MoveNext
             x = x + 1
      Loop

     Msgx = Msgx & "</font></table>"
 RST.MoveFirst
 'LinkX = "\\TS-VM-TERMINAL\AccessDB\TDM\Deploy\OpenAlert.bat"
'make email message
        msg1 = msg1 & "<html><body><p style='color:Navy;font-family:Calibri;font-size:20pt;'><b>Items Missing MSO</b></p>"
        msg1 = msg1 & "<p style='font-family:Calibri;font-size:11pt'>"
        msg1 = msg1 & "To Contact: <b>" & Contactx & "</b><br />"
        msg1 = msg1 & "Subject:<b> " & subjectx & "</b><br />"
        msg1 = msg1 & "Priority:<b> High </b><br/><br/>"
        msg1 = msg1 & "Dear " & Contactx & ",<br/><br/>"
        msg1 = msg1 & "Please forward the respective MSO copies so that we can begin the administrative paperwork required to make payment.<br /><br />"
        'Msg1 = Msg1 & "<a  href=" & LinkX & ">Open Alert Form</a><br /><br />"
        msg1 = msg1 & "<p style='font-family:Calibri; font-size:11pt; margin-left:20px'>"
        msg1 = msg1 & "MSO assignment should be to:<br /><b>FA Title Trust<br />401 E. Las Olas Blvd., #1720<br />Fort Lauderdale, FL 33301<br /></b><br /></p>"
        msg1 = msg1 & "<p style='font-family:Calibri; font-size:11pt; margin-left:0px'>"
        'Msg1 = Msg1 & "Client: <B>" & RST!ClientName & "</B> FAID: <B>" & RST![FAID/(PO)] & " <br />"
        msg1 = msg1 & "Vehicle List follows:"
        msg1 = msg1 & "<p style='font-family:Calibri; font-size:10pt; margin-left:30px'>" & Msgx & "</P>"
        msg1 = msg1 & "<p style='font-family:Calibri;font-size:8pt'><br/>[End List] "
        msg1 = msg1 & "Created: " & Date & " (via Fleet Advantage Data Manager).<br/><br />"
        msg1 = msg1 & "<p style='font-family:Calibri;font-size:11pt'><br/>"
        msg1 = msg1 & "</b>Regards,<br /><br /><br />"
        msg1 = msg1 & "</body></html>"
     Call SendEmailMessage(ContactEmailX, subjectx, msg1, False, True, False, "")
        RST.close
        Set RST = Nothing
    End If
            Set RST = Me.RecordsetClone
            RST.MoveLast
            RST.MoveFirst
                SQLz = "INSERT INTO reqinvMSO_Tracking ( Client, Contact, ContactEmail, Datesent, VehicleCount, FAID, VendorID, ReqType, RequestedBy, Subject) " & _
                       "SELECT '" & RST!ClientShNm & "' AS Client, '" & Contactx & "' AS Contact, '" & ContactEmailX & "' AS ContactEmail, " & _
                       "Date() AS Datesent, " & UnitsX & " AS VehicleCount, '" & FAIDx & "' AS FAID, '" & VendorIDx & "' as VendorID, 'MSOCopies' as ReqType, " & _
                       "'" & UserX & "' as RequestedBy, '" & subjectx & "' as Subject;"
                DoCmd.SetWarnings False
                DoCmd.RunSQL (SQLz)
                DoCmd.SetWarnings True
           RST.close
           Set RST = Nothing
           Me.ReqMSOTracking.Requery
End Sub

Private Sub ReqMSOTracking_DblClick(Cancel As Integer)
If IsNull(Me.ReqMSOTracking.Column(3)) Then Exit Sub
   SQLz = "DELETE ReqInvMSO_Tracking.*, ReqInvMSO_Tracking.Index " & _
          "FROM ReqInvMSO_Tracking WHERE Index = " & Me.ReqMSOTracking.Column(3) & ";"
  DoCmd.RunSQL (SQLz)
  Me.ReqMSOTracking.Requery
End Sub
