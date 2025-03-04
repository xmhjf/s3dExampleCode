VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenericDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'----------------------------------------------------------------------------
' Copyright (C) 2007, Intergraph Corporation.  All rights reserved.
'
' Project
'   ISMfgNestDelete
'
' File
'   GNestExport
'
' Description
'   This file contains the delete function for nesting in GNest
'
' Author
'   Kristian Kamph
'
' History:
'   2007-05-15  K. Kamph     Creation date
'----------------------------------------------------------------------------
Option Explicit

Implements IJMfgNestDelete

Private Sub IJMfgNestDelete_Delete(ByVal pMfgObj As Object, ByVal bstrOldFormat As String, ByVal bstrNewFormat As String)
    Const sMETHOD As String = "IJMfgNestDelete_Delete"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
End Sub
