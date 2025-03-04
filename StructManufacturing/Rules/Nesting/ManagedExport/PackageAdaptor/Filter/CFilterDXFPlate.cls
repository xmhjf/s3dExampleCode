VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFilterDXFPlate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'----------------------------------------------------------------------------
' Copyright (C) 2016, Intergraph Corporation.  All rights reserved.
'
' Project
'   MfgManagedExportRule
'
' File
'   CFilterPlate
'
' Description
'   This Filte to filter out the Plate to generate DXF output
'
' Author
'   Nautilus-HSV
'
' History:
'   2016-04-15  Nautilus-HSV  Creation
'----------------------------------------------------------------------------

Option Explicit
Implements IJMfgNestFilter

Private Const MODULE = "CFilterDXFPlate"

Private Function IJMfgNestFilter_Evaluate(ByVal pDispObject As Object) As Boolean
Const sMETHOD As String = "IJMfgNestFilter_Evaluate"
On Error GoTo ErrorHandler
        
    IJMfgNestFilter_Evaluate = False
    
    Dim oAssemblyBase As IJAssemblyBase
    Dim oAssemblyChild As IJAssemblyChild
    
    If TypeOf pDispObject Is IJAssemblyChild Then
        Set oAssemblyChild = pDispObject
        Set oAssemblyBase = oAssemblyChild.Parent
        
         If oAssemblyBase.Type = 0 Then
                IJMfgNestFilter_Evaluate = True
        End If
    End If
    
    Set oAssemblyBase = Nothing
    Set oAssemblyChild = Nothing
        
    Exit Function
ErrorHandler:
End Function



