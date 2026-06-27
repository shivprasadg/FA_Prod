Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command6_Click()
Me.Undo
Me.Undo
DoCmd.close acForm, "EFExportProgress"
Exit Sub

End Sub

Private Sub Form_Load()
Forms!EFExportProgress.caption = "Equipment Files Excel Template Progress"

Forms!EFExportProgress.Box1.Visible = False
Forms!EFExportProgress.Box2.Visible = False
Forms!EFExportProgress.Box3.Visible = False
Forms!EFExportProgress.Box4.Visible = False
Forms!EFExportProgress.Box5.Visible = False
Forms!EFExportProgress.Box6.Visible = False
Forms!EFExportProgress.Box7.Visible = False
Forms!EFExportProgress.Box8.Visible = False
Forms!EFExportProgress.Box9.Visible = False






End Sub
