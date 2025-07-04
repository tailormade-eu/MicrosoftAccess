Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_DblClick(Cancel As Integer)
    DoCmd.OpenForm "omEmptyTable_Edit", whereCondition:="Id=" & Me.Id, datamode:=acFormEdit, windowMode:=acDialog
    Me.Refresh
End Sub
