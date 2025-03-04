VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISDetailPartFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   ISDetailPartFilter.cls
'   ProgID:         DetailPartsFilter.ISDetailPartFilter
'   Author:         RamaRao
'   Creation Date:  03.Sep.2004
'   Description:
'
'
'   Change History:
'       03.Sep.2004            RamaRao       Creation
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

'==============================
'Detail Parts Filter
'===============================
Public Function DetailPartFilter(pISPart As Object) As Long
On Error GoTo ErrorHandler

    DetailPartFilter = 0
    
    Dim oModelBody As IJDModelBody
    Dim oIJProfilePart As IJProfilePart
    
    On Error Resume Next
    Set oModelBody = pISPart
    Set oIJProfilePart = pISPart
    On Error GoTo ErrorHandler
    
    Dim bodyType As JS_TOPOLOGY_BODY_TYPE
    If Not oIJProfilePart Is Nothing Then
    If Not (oModelBody Is Nothing) Then
    
        bodyType = oModelBody.bodyType
    
        If bodyType And JS_TOPOLOGY_BODY_SOLID Then
            DetailPartFilter = 1
        End If
    End If
    End If
    Set oIJProfilePart = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number
End Function
