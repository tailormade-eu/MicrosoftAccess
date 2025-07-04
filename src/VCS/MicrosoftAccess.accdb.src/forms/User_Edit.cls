Attribute VB_GlobalNameSpace = False
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

Private Sub UserRoleId_AfterUpdate()
    omControlFunctions.SetDefaultValue Me.UserRoleId
End Sub

Private Sub WorkHoursPerDay_AfterUpdate()
    omControlFunctions.SetDefaultValue Me.WorkHoursPerDay
End Sub
