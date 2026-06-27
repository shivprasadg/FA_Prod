Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Current()
On Error GoTo JustMe
Call SetATFDataTable(Me.Form.Name)
On Error GoTo 0
Exit Sub
JustMe:
On Error GoTo 0
End Sub
