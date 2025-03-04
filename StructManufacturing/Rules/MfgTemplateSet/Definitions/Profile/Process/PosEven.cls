VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PosEven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   At Even Position Process Rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit

Implements IJMfgTemplatePositionRule

Private Const MODULE = "MfgTemplateProcessProfile.PosEven"

Private Sub IJMfgTemplatePositionRule_GetPositionData(ByVal pFrameSysDsp As Object, ByVal pPartDisp As Object, ByVal pProcessSettings As Object, ByVal pMfgTemplateSet As Object, dEdgeOffset As Double, lNumOfTemplates As Long, bSupportEdges As Boolean, ppFrames As IMSCoreCollections.IJElements)
Const METHOD = "IJMfgTemplatePositionRule_GetPositionData"
On Error GoTo ErrorHandler
    
    'Get Plate object
    Dim oProfilePart As StructDetailObjects.ProfilePart
    Set oProfilePart = New StructDetailObjects.ProfilePart
    Set oProfilePart.object = pPartDisp
    Dim Length As Double
    
    Dim Offset As Double
    Dim Step As Long
  
    ' *************************************************************
    ' Example for Odense Situation
    ' *************************************************************
    'Set default values
    Offset = 0.1
    Step = 3
    
    'Convert from m to mm
    ' It's Kludge for the time being.. If ProfilePart wrapper classs will support length.. it will be used
    Length = oProfilePart.WebLength * 1000#
   
    If Length < 5000 Then
            Offset = 0.1
            Step = 3
    Else
         If Length < 8500 Then
            Offset = 0.2
            Step = 5
        Else
            Offset = 0.3
            Step = 7
        End If
    End If
  
    ' Send the data back
    dEdgeOffset = Offset
    lNumOfTemplates = Step
    
    bSupportEdges = False
    
    
 Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3010, , "RULES")
End Sub
 
