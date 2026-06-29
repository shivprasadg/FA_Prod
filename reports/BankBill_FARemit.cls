Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Load()
If Me.ClientID = 4209 Then Me.FeeLabel.caption = "Ancillary"
End Sub
