VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PartBOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   MQT.cls
'   ProgID:         HS_Hilti_MQParts.MQT
'   Author:         JRK
'   Creation Date:  Wednesday, November 21, 2007

'   Description:
'       Pipe Support for insulated pipes
'
'   Change History:
'   Date              who         change description
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
' Create constant to use with error messages
Private Const MODULE = "PartBOM"
Implements IJHgrSymbolBOMServices

Private Sub IJHgrSymbolBOMServices_EvaluateBOM(ByVal pSupportComp As Object, bstrBOMDesc As String)
    Const METHOD = "IJHgrSymbolBOMServices_EvaluateBOM"
    On Error GoTo ErrorHandler

    ' <Determine BOM description here>
    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Set oPartOcc = pSupportComp    ' The part occurence
    oPartOcc.GetPart oPart         ' The associated catalog part

    Dim sItemNo As String
    Dim dWeight As Double

    Dim sPartNumber As String
    sPartNumber = Trim(oPart.PartNumber)

    sItemNo = GetAttributeFromObject(oPart, "ItemNo")
    
    If InStr(sPartNumber, "MFP") <> 0 Or InStr(sPartNumber, "MP-MI") <> 0 Or InStr(sPartNumber, "MP-MXI") <> 0 Or InStr(sPartNumber, "MP-SNRI") <> 0 Or InStr(sPartNumber, "MV-PI") <> 0 Or InStr(sPartNumber, "MPN") <> 0 Or InStr(sPartNumber, "MSG") <> 0 Or InStr(sPartNumber, "MPSG") <> 0 Then
        bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " " & Replace(sPartNumber, "Hilti ", "")
    ElseIf InStr(sPartNumber, "X-BT M10") <> 0 Then
        bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 1, 52461, Hexagon nut M10 A4-70 + 1, 58042, Flat washer A 10,5/20 A4"
    ElseIf InStr(sPartNumber, "X-BT M8") <> 0 Then
        bstrBOMDesc = Hilti_BuildBom(pSupportComp) & " + 1, 52460, Hexagon nut M8 A4-70 + 1, 58622, Flat washer A 8,4/16 A4"
    Else
        bstrBOMDesc = Hilti_BuildBom(pSupportComp)
    End If
    
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number

End Sub
