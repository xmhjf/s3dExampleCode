VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeldCompUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2005, Intergraph Corp.  All rights reserved.
'
' Module:       WeldCompUD
'
' Description:  Apply Weld Compensation with User Defined Parameters
'
' Author:       hkarumug@ingr.com
'
' Comments:
'
'   2005.04.13  hkarumug@ingr.com     Initial Creation.
'
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkRule
 
Private m_oWCFactory As IJMfgWeldCompensationFactory
Private m_oWC As IJMfgWeldCompensation

Const MODULE = "MfgProfileProcess.WeldCompUD"

Private Function IJDMfgSystemMarkRule_CreateAfterUnfold(ByVal Part As Object, ByVal Upside As Long) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
'
End Function

Private Function IJDMfgSystemMarkRule_CreateBeforeUnfold(ByVal Part As Object, ByVal Upside As Long) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "IJDMfgSystemMarkRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set m_oWCFactory = New MfgWeldCompensationFactory
    
    ' Does a compensation already exist ?
    On Error Resume Next
    Set m_oWC = m_oWCFactory.GetWeldCompensationOnProfile(Part)
    
    On Error GoTo ErrorHandler
    If m_oWC Is Nothing Then
        ' Create a WeldCompensation for this ProfilePart with defaults
        Dim oWCDefaults As IJDMfgSystemMarkRule
        Set oWCDefaults = New MfgProfileProcess.WeldCompYes
        
        ' Delegate the call
        Set oGeomCol3d = oWCDefaults.CreateBeforeUnfold(Part, Upside)
        
        ' Do we have one ? must succeed for builtups
        On Error Resume Next
        Set m_oWC = m_oWCFactory.GetWeldCompensationOnProfile(Part)
        If m_oWC Is Nothing Then GoTo CleanUp
    Else
        ' Already have a compensation with some user values
        ' Create Markers
        Dim oGeomCol3dFactory As New GSCADMfgGeometry.MfgGeomCol3dFactory
        Set oGeomCol3d = oGeomCol3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
        Dim oMark As IJMfgGeom3d
        If (m_oWC.CompensationInfo And WELDCOMP_INFO_WC1) = WELDCOMP_INFO_WC1 Then
            Set oMark = m_oWCFactory.CreateWeldCompensationMark( _
                         GetActiveConnection.GetResourceManager(GetActiveConnectionName), _
                         m_oWC, WELDCOMP_INFO_WC1)
            oGeomCol3d.AddGeometry 1, oMark
        End If
        
        If (m_oWC.CompensationInfo And WELDCOMP_INFO_WC2) = WELDCOMP_INFO_WC2 Then
            Set oMark = m_oWCFactory.CreateWeldCompensationMark( _
                         GetActiveConnection.GetResourceManager(GetActiveConnectionName), _
                         m_oWC, WELDCOMP_INFO_WC2)
            oGeomCol3d.AddGeometry 1, oMark
        End If
    End If
    
    Set IJDMfgSystemMarkRule_CreateBeforeUnfold = oGeomCol3d
               
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2037, , "RULES")
End Function
 
