Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Public pMLType As String
Private Sub Form_Close()
If IsOpen("ATF_SFImporter") Then
Forms!ATF_SFImporter!MLID.Requery
Forms!ATF_SFImporter!MLID = Me.MLID
Forms!ATF_SFImporter!MLID2.Requery
Forms!ATF_SFImporter!MLID2 = Me.MLID
End If
End Sub

Private Sub Form_Current()
Dim SQLz As String
SQLz = "SELECT contactID ,FullName, Title, ClientID FROM ClientContacts_Base WHERE ClientID= " & Me.Client & " Order by FullName;"
Me.MLSigner.RowSource = SQLz
Me.MLSigner2.RowSource = SQLz
Me.MLSigner3.RowSource = SQLz
SQLz = "SELECT contactID,FullName, Title, ClientID FROM ClientContacts_Base WHERE ClientID= " & Me.MLGuarantor & " Order by FullName;"
Me.GuarantorSigner.RowSource = SQLz
Me.MLSignerTitle1.Requery
Me.MLSignerTitle2.Requery
Me.MLSignerTitle3.Requery
Me.GuarantorTitle.Requery
End Sub

Private Sub Form_Load()
If IsOpen("Ops_Lessee_Setup") Then
    Client = Forms!ops_lessee_Setup.ClientSN
End If

End Sub

Private Sub GuarantorSigner_Click()
Me.GuarantorTitle.Requery
End Sub

Private Sub MLGuarantor_Click()
Me.MLGuarantor.Requery
End Sub

Private Sub MLLessor_Click()
'Me.MLLessorID = Me.MLLessor.Column(1)
End Sub

Private Sub MLOrig_BeforeUpdate(Cancel As Integer)
    If pMLType = "FA" Then
        MsgBox "You cnannot change a FA lease Type", vbCritical, "Oh no you don't"
        Me.Undo
    End If
End Sub

Private Sub MLOrig_GotFocus()
pMLType = Me.MLOrig
End Sub

Private Sub MLSigner_Click()
Me.MLSignerTitle1.Requery
End Sub
Private Sub MLSigner2_Click()
Me.MLSignerTitle2.Requery
End Sub
Private Sub MLSigner3_Click()
Me.MLSignerTitle3.Requery
End Sub
