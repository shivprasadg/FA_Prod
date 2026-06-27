Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Load()
    If Me.FSLFeeX3 = 0 Then
        Me.FeeLabel.caption = "    "
        Me.FSLFeeX1.Visible = False
        Me.FSLFeeX2.Visible = False
        Me.FSLFeeX3.Visible = False
    End If
End Sub
