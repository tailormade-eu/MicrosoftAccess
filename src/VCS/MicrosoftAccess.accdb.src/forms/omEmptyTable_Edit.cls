﻿Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_BeforeInsert(Cancel As Integer)
    omFormFunctions.UpdateCreateTracking Me
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    omFormFunctions.UpdateModifyTracking Me
End Sub
