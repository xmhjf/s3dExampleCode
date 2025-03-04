VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IndexNPercent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------------------------------'
'    Copyright (C) 2004 Intergraph Corporation. All rights reserved.
'
'
'Abstract
'    'Index And Percent' Grids Naming Rule.
'
'Notes
'    Extracted from specification:
'
'    Index & Percent Name Rule: This option provides the root label (“GPX” for X-axis, “GPY” for
'    Y-axis) for the primary planes. Sub nesting levels are a percentage (in decimal form) of the
'    bounding and higher nesting levels.
'Example:
'    Grids are usually prefixed with a letter and suffixed with a decimal distance away
'    (A, B, C, C.6, D, E). The number to the right of the decimal place (e.g. C.6) for a secondary
'    plane is the relative position between the previous and next primary plane. Another words it is
'    60% away from the previous primary elevation plane.
'
'    Primary        Secondary       Tertiary
'     GPX1
'     GPX2
'                    GPX2 0.5
'                                    GPX2.5.3
'                                    GPX2.5.8
'                    GPX2 0.6
'     GPX3
'
'History
'    Eric Faivre         07/21/04                Creation.
'--------------------------------------------------------------------------------------------'

Option Explicit

Implements IJNameRule
Implements IJNRAuxiliaryPosition

Private m_IncludeList As New Collection
Private m_ExcludeList As New Collection

Private Sub IJNameRule_ComputeName(ByVal pEntity As Object, ByVal pParents As IMSICDPInterfacesLib.IJElements, ByVal pActiveEntity As Object)
    On Error GoTo ErrorHandler
    
    ApplyNameFromNPercentNR pEntity, m_IncludeList, m_ExcludeList, False
    
    Exit Sub
    
ErrorHandler:
    mg_oErrors.Add Err.Number, "GSNamingRules::IndexNPercent::IJNameRule_ComputeName", Err.Description
    Err.Raise E_FAIL
End Sub

Private Function IJNameRule_GetNamingParents(ByVal pEntity As Object) As IMSICDPInterfacesLib.IJElements
    Set IJNameRule_GetNamingParents = Nothing
End Function

Private Sub IJNRAuxiliaryPosition_ExcludeAuxiliaryPos(ByVal pExcludeList As Object)
    Set m_ExcludeList = pExcludeList
End Sub

Private Sub IJNRAuxiliaryPosition_IncludeAuxiliaryPos(ByVal pIncludeList As Object)
    Set m_IncludeList = pIncludeList
End Sub

Private Sub Class_Initialize()
    Set mg_oErrors = New IMSErrorLog.JServerErrors
End Sub

Private Sub Class_Terminate()
    Set mg_oErrors = Nothing
End Sub

