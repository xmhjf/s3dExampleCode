VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgPinJigRule
' Module: UserParameters
'
' Description:  Specify various parameters that influence Processing/Remarking of Pin Jigs
'
' Author: Anand Hariharan
'
'*******************************************************************************
Option Explicit
Private Const MODULE As String = "MfgPinJigRule.UserParameters"

Implements IJDPinJigUserParam

'---------------------------------------------------------------------------------------
' Procedure : IJDPinJigProcParam_SpecifyRemarkingTypesForPurpose
' Purpose   : For the "Purpose" specified by the input string, return a list of
'             remarking types that satisfy that particular purpose.
'
' General convention for the "Purpose" string:
'    Prefixed with "IncludeIn_":   Only types specified here will be included.
'    Prefixed with "ExcludeFrom_": Except for types specified here, all others will be included.
'
'---------------------------------------------------------------------------------------

Private Function IJDPinJigUserParam_SpecifyRemarkingTypesForPurpose(ByVal PurposeOfRemarkingTypes As String) As Long()
    Const METHOD = "IJDPinJigProcParam_SpecifyRemarkingTypesForPurpose"
    On Error GoTo ErrorHandler
    
    Dim RemarkTypes() As Long
    Select Case PurposeOfRemarkingTypes
    
        Case "IncludeIn_MarkingCommand"
            'Specify the types of remarking lines that should be displayed
            'within the RAD 2D environment of the marking command.
            ReDim RemarkTypes(1 To 2) As Long
            RemarkTypes(1) = STRMFG_PinJig_Remarking_Plate
            RemarkTypes(2) = STRMFG_PinJig_Remarking_Profile
        
        Case "ExcludeFrom_IntersectionPointCreation"
            'Specify the types of remarking lines that should be excluded
            'from participating in the intersection point creation process.
            ReDim RemarkTypes(1 To 4) As Long
            RemarkTypes(1) = STRMFG_NAVALARCHLINE
            RemarkTypes(2) = STRMFG_PinJig_Remarking_NavalArch
            RemarkTypes(3) = STRMFG_PINJIG_MARK
            RemarkTypes(4) = STRMFG_PINJIG_DIAGONAL
        
        Case Else
            ReDim RemarkTypes(0 To 0) As Long
    
    End Select
    
    IJDPinJigUserParam_SpecifyRemarkingTypesForPurpose = RemarkTypes

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, "Failed specifying Remarking types")
End Function

Private Function IJDPinJigUserParam_SpecifyToleranceForPurpose(ByVal ToleranceContext As String, Optional ByVal ValForRounding As Double = 0#) As Double

End Function
