VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeldCompYes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2005, Intergraph Corp.  All rights reserved.
'
' Module:       WeldCompYes
'
' Description:  Apply Weld Compensation with Default Parameters
'
' Author:       hkarumug@ingr.com
'
' Comments:
'
'   2005.04.13  hkarumug@ingr.com     Initial Creation.
'
'*******************************************************************************
' Logic for computing default parameters:
'
'   Defaults:
'
'   Len < 15000     p1 = 30     p2 = 0.5 * p1 * Ht/Len  p3 = -p2
'   Len > 15000     p1 = 45     p2 = 0.5 * p1 * Ht/Len  p3 = -p2
'                   Param1 = Param 2 = 0.5
'
'   Note a,
'
'       If we have a top flange and a bottom flange,
'           apply compensation on the flange with larger height
'
'   Note b,
'
'       If we have a "clear" inflexion along the landing curve,
'           set Param1 = Param2 = tangency point
'
'-------------------------------------------------------------------------------

Option Explicit

Implements IJDMfgSystemMarkRule

Private m_oProfilePart As IJProfilePart

Private m_oWC As IJMfgWeldCompensation
Private m_oWCFactory As IJMfgWeldCompensationFactory

Private m_oProfileHelper As Object
Private m_oProfileSupport As IJProfilePartSupport
Private m_oPartSupport As IJPartSupport

Const MODULE = "MfgProfileProcess.WeldCompYes"

Private Sub Execute(ByVal Part As Object)
    Const METHOD = "Execute"
    On Error GoTo ErrorHandler
        
    Set m_oProfilePart = Part
    
    Call InitializeHelpers
    
    ' Does a compensation already exist ?
    On Error Resume Next
    Set m_oWC = m_oWCFactory.GetWeldCompensationOnProfile(Part)

    If m_oWC Is Nothing Then
        ' Create a WeldCompensation for this ProfilePart
        ' Call will fail for builtups
        Set m_oWC = m_oWCFactory.CreateWeldCompensationOnProfile( _
                    GetActiveConnection.GetResourceManager(GetActiveConnectionName), _
                    Part)
    End If

    On Error GoTo ErrorHandler
    
    ' May be not a builtup ?
    If m_oWC Is Nothing Then GoTo CleanUp
            
    ' Apply defaults
    Dim p1 As Double, p2 As Double, p3 As Double
    If m_oProfileSupport.ApproximateLength < 15 Then         ' meters
        p1 = 0.03
    Else
        p1 = 0.045
    End If
    p2 = 0.5 * p1 * m_oProfileHelper.Height / m_oProfileSupport.ApproximateLength
    p3 = -p2

    With m_oWC
        .CompensationInfo = WELDCOMP_INFO_WEB Or _
                            PickCompensationSide Or _
                            WELDCOMP_INFO_WC1 Or WELDCOMP_INFO_WC2
        
        Call ApplyCompensationPoints
        
        .SetCompensationParametersForPoint1 p2, p1, p3, p1
        .SetCompensationParametersForPoint2 p2, p1, p3, p1
    End With
   
CleanUp:
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

Private Sub InitializeHelpers()
    ' Catch error at caller
    If TypeOf m_oProfilePart Is IJStiffenerPart Then
        Set m_oProfileHelper = New StructDetailObjects.ProfilePart
    ElseIf TypeOf m_oProfilePart Is IJBeamPart Then
        Set m_oProfileHelper = New StructDetailObjects.BeamPart
    End If
    Set m_oProfileHelper.object = m_oProfilePart
    
     
    Set m_oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
    Set m_oPartSupport.Part = m_oProfilePart
    
    Set m_oProfileSupport = m_oPartSupport
    
    Set m_oWCFactory = New MfgWeldCompensationFactory
End Sub

Private Function PickCompensationSide() As MfgWeldCompensationInfo
    On Error GoTo UseDefault
    
    ' See Note a
    Dim dTopThick As Double, dBottomThick As Double
    
    Call m_oProfileSupport.GetThickness(TopFlange, dTopThick)
    Call m_oProfileSupport.GetThickness(BottomFlange, dBottomThick)
             
    If dTopThick > dBottomThick Then
        PickCompensationSide = WELDCOMP_INFO_TOP
    Else
        PickCompensationSide = WELDCOMP_INFO_BOTTOM
    End If
    
CleanUp:
    Exit Function
UseDefault:
    PickCompensationSide = WELDCOMP_INFO_TOP
End Function

Private Sub ApplyCompensationPoints()
    On Error GoTo UseDefault
    
    ' See Note b
    Dim dp1 As Double, dp2 As Double, dp As Double
    dp1 = 0
    dp2 = 0
    dp = 0
    
    Dim oLine As Line3d
    Set oLine = New Line3d
    oLine.DefineBy2Points 0, 0, 0, 100, 100, 100
    
    Dim oCompCurve As IJCurve
    Set oCompCurve = m_oWC.GetWeldCompensationCurve
'Dim oC As BSplineCurve3d
'Set oC = New BSplineCurve3d
'oC.DefineByComplexString oCompCurve
'
'Set oCompCurve = oC
    
    ' Get the inflexion points
    Dim oMfgGeomHelper As New GSCADStrMfgUtilities.MfgGeomHelper
    
    Dim oInflexions As IJElements, oInflexion As IJDPosition
    Set oInflexions = oMfgGeomHelper.GetInflectionPointsOnCurve(oCompCurve)
           
    Dim dLen As Double, dpLen As Double, dr As Double
    dLen = oCompCurve.Length
                
    Dim oStart As IJDPosition, oLast As IJDPosition
    
    Dim x1 As Double, y1 As Double, z1 As Double
    Dim x2 As Double, y2 As Double, z2 As Double
    oCompCurve.EndPoints x1, y1, z1, x2, y2, z2
    
    Set oStart = New DPosition
    oStart.Set x1, y1, z1
    
    Set oLast = New DPosition
    oLast.Set x2, y2, z2
   
    For Each oInflexion In oInflexions
        ' for a tangent-profile we try to find the inflexion points
        ' look for the length as we walk through the curve so we do not
        '   get influenced by local curvature
        
        oMfgGeomHelper.GetLengthBet2Points oCompCurve, oStart, oInflexion, dpLen
        dr = dpLen / dLen
        
        If dr <= 0.5 Then
            ' a good candidate for first point
            oCompCurve.Parameter oInflexion.x, oInflexion.y, oInflexion.z, dp1
            dp2 = dp1
        Else
            oMfgGeomHelper.GetLengthBet2Points oCompCurve, oInflexion, oLast, dpLen
            dr = dpLen / dLen
            
            If dr > 0.3 Then
                ' a good candidate for second point
                oCompCurve.Parameter oInflexion.x, oInflexion.y, oInflexion.z, dp2
            Else
                Exit For
            End If
        End If
        
    Next
    
CleanUp:
    If dp1 < 0.1 And dp2 < 0.1 Then GoTo UseDefault
    If dp1 < 0.1 Then dp1 = dp2
    If dp2 < 0.1 Then dp2 = dp1
    
    m_oWC.CompensationPoint1 = dp1
    m_oWC.CompensationPoint2 = dp2
    Exit Sub
UseDefault:
    dp1 = 0.5
    dp2 = 0.5
    GoTo CleanUp
End Sub

Private Function IJDMfgSystemMarkRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
'
End Function

Private Function IJDMfgSystemMarkRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "IJDMfgSystemMarkRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
        
    ' Add WeldCompensation with defaults
    Call Execute(Part)
    
    If m_oWC Is Nothing Then GoTo CleanUp
    
    'Using the ProfileWrapper, update the relations on Weld Compensation.
    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    oProfileWrapper.UpdateRelationsOnWeldCompensation m_oWC

' Create Markers
    Dim oGeomCol3dFactory As New GSCADMfgGeometry.MfgGeomCol3dFactory
    Dim oGeomCol3d As IJMfgGeomCol3d
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

    Set IJDMfgSystemMarkRule_CreateBeforeUnfold = oGeomCol3d
CleanUp:
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2038, , "RULES")
End Function
