VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StructureFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   StructureFilter.cls
'   ProgID:         GSCADFilterService.StructureFilter
'   Author:         Thakur
'   Creation Date:  18.Nov.2002
'   Description:
'
'
'   Change History:
'       18.Nov.2002             Thakur       Creation
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'----------------------------------------------------------------'
' IsASolidPart: Returns 0 if the  element is Hiull or Light part.
'               Otherwise, returns 1

'----------------------------------------------------------------'

Public Function IsASolidPart(element As Object) As Integer

    Dim oUtils              As New StructUtilities.Utilities
    Dim oStructFun          As New STFilterFunctions.StructFilterFunctions
    Dim oType As Long
    IsASolidPart = 1

    If oStructFun.IsAHull(element) Then

        IsASolidPart = 0
        Exit Function
    End If
    
    On Error Resume Next
    oType = oUtils.ResultType(element)
    If Err.Number = 0 Then
        If (oType And 32768) = 32768 Then IsASolidPart = 0 '32768 = 0x8000 = STRUCT_RESULT_TYPE_LIGHT
        Exit Function
    Else
        Exit Function
    End If
End Function


   


