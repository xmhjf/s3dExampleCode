VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PositionEven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgTemplateProcess
' Module:
'
' Description:  Determines the process settings for the Template
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Implements IJMfgTemplatePositionEvenRule

Private Const MODULE = "MfgTemplateProcess.PositionEven"

Private Sub IJMfgTemplatePositionEvenRule_GetPositionEven(ByVal pPlatePart As GSCADMfgRulesDefinitions.IJPlatePart, dEdgeOffset As Double, lNumOfTemplates As Long)
    Const METHOD = "IJMfgTemplatePositionEvenRule_GetPositionEven"

On Error GoTo ErrorHandler
  
    Dim sERROR As String
    sERROR = ""
    
    'Get Plate object
    Dim oPlatePart As StructDetailObjects.PlatePart
    Set oPlatePart = New StructDetailObjects.PlatePart
    Set oPlatePart.Object = pPlatePart
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
    ' It's Kludge for the time being.. If PlatePart wrapper classs will support length.. it will be used
    Length = oPlatePart.PlateLength * 1000#
   
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
Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3009, , "RULES")
End Sub
 
