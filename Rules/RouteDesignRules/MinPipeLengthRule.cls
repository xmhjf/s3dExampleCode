VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MinPipeLengthRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2010, Intergraph Corporation. All rights reserved.
'
'File
'    RouteDesignRule.cls
'
'Author
'    Ravi Kiran Kaki
'
'Description
'    Design Rules implementation.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Implements IJDEngineeringRules      ' Rule implementing Interface.
Private Sub Class_Initialize()
On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    Debug.Assert False
End Sub

Private Sub Class_Terminate()
    'Release the Service.
End Sub

Private Function IJDEngineeringRules_ValidateObject(ByVal pObject As Object, ByVal violationRecords As IJViolationRecords) As Boolean

On Error Resume Next
On Error GoTo ErrorHandler
    Dim length As Double

    If Not pObject Is Nothing Then
        Dim rteStockPartOccur As IJRteStockPartOccur
        Set rteStockPartOccur = Nothing
        Set rteStockPartOccur = pObject
        If Not rteStockPartOccur Is Nothing Then
            length = rteStockPartOccur.length
        Else
            Dim pathLeg As IJRtePathLeg
            Set pathLeg = pObject
            If Not pathLeg Is Nothing Then
                length = pathLeg.length
            End If
        End If
    End If

    If (length - 5) > 0.001 Then
        violationRecords.Clear
        violationRecords.AddViolation vtSevereFoul, "Pipe violates maximum pipe length requirement"
    End If
    IJDEngineeringRules_ValidateObject = True

Exit Function
ErrorHandler:
    Err.Clear
    Debug.Assert (False)
    IJDEngineeringRules_ValidateObject = False
End Function

