VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtLowFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Template Only On Lower Side With Fixed Extension Process Rule
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new class
'************************************************************************************************************

Option Explicit
Private Const MODULE = "MfgTemplateProcessProfile.ExtLowFix"
Implements IJMfgTemplateExtensionRule3

'************************************************************************************************************
' Method: IJMfgTemplateExtensionRule3_GetExtension
' Description: Provides the Extension information
' Partdisp -> Profile Part
'
' dLowerSidePortion  -> Portion of Bottom Line needed on the LowerSide of BCL. 1->complete 0->Not needed 0.5->Half of the LowerSide Bottom line
' dLSFixedExtn -> Fixed Extension to be applied on the Lower Side
' dLSOffsetExtn -> Offset Extension to be applied on the Lower Side
' dLSMinExtnForOffset -> Minimum Extension for Offset Extension on the Lower Side
'
' dUpperSidePortion -> Portion of Bottom Line needed on the UpperSide of BCL. 1->complete 0->Not needed 0.5->Half of the UpperSide Bottom line
' dUSFixedExtn -> Fixed Extension to be applied on the Upper Side
' dUSOffsetExtn -> Offset Extension to be applied on the Upper Side
' dUSMinExtnForOffset ->Minimum Extension for Offset Extension on the Upper Side

'************************************************************************************************************
Private Sub IJMfgTemplateExtensionRule3_GetExtension(ByVal Partdisp As Object, dLowerSidePortion As Double, dLSFixedExtn As Double, dLSOffsetExtn As Double, dLSMinExtnForOffset As Double, dUpperSidePortion As Double, dUSFixedExtn As Double, dUSOffsetExtn As Double, dUSMinExtnForOffset As Double)
' Defines the Extension value to be used for the template ebyond the seam defintion
Const METHOD = "IJMfgTemplateExtensionRule3_GetExtension"
On Error GoTo ErrorHandler

    dLowerSidePortion = 1
    dLSFixedExtn = 0.05
    dLSOffsetExtn = 0
    dLSMinExtnForOffset = 0
    
    dUpperSidePortion = 0
    dUSFixedExtn = 0
    dUSOffsetExtn = 0
    dUSMinExtnForOffset = 0

Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3006, , "RULES")
End Sub
