Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Open(Cancel As Integer)
If Me.FSLFeeX2 = 0 Then
Me.FeeLabel = "    "
Me.FSLFeeX1.Visible = False
Me.FSLFeeX2.Visible = False
End If
End Sub
