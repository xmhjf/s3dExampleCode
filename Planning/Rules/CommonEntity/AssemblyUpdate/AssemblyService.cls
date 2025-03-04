VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DefaultRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MODULE = "AssemblyNotifyRule.DefaultRule"
Implements IJAssemblyCustomNotifier

Private Function IJAssemblyCustomNotifier_GetCustomValue(ByVal pAssembly As Object, ByVal eProcess As GSCADAssemblyBase.AssemblyCustomProcess) As Boolean
Const METHOD As String = "IJAssemblyCustomNotifier_GetCustomValue"
On Error GoTo ErrorHelper

    ' Logic should be added here to determine whether a part can be added/removed or whether Planning Joints can be updated
    If eProcess = UpdatePlanningJoints Then
        ' Update Planning Joints when Assembly properties(Type, Equipment) are changed or assembly orientation changed or assembly base plate is changed
        IJAssemblyCustomNotifier_GetCustomValue = False
    End If

Exit Function
ErrorHelper:
    MsgBox Err.Description
End Function

Private Function IJAssemblyCustomNotifier_Update(ByVal pAssembly As Object) As Boolean
Const METHOD As String = "IJAssemblyCustomNotifier_Update"
On Error GoTo ErrorHelper

    ''logic should be added here to determine whether the assembly should be updated or not when
    '' a part is added or removed from the assembly
    
    IJAssemblyCustomNotifier_Update = True

Exit Function
ErrorHelper:
    MsgBox Err.Description
End Function
