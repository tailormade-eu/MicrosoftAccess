﻿Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdNew_Click()
    DoCmd.OpenForm "omEmptyTable_Edit", datamode:=acFormAdd, windowMode:=acDialog
    Form_Search
End Sub

Private Sub txtSearch_AfterUpdate()
    Form_Search
End Sub


Private Sub Form_Search()
    Me.subSearch.Requery
End Sub

Private Sub cmdClearSearch_Click()
    Me.txtSearch = Null
    Form_Search
End Sub
