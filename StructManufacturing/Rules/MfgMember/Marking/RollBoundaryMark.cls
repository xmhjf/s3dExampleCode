VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RollBoundaries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgMemberMarking
' Module: Marking Rules
'
' Description:  Determines the marking settings for the mfg profile
'
' Author:
'
' Comments:
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgMemberMarking.RollBoundaries"

Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    Helpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    Helpers.UnInitialize
End Sub
' return the set of roll boundaries that are the ouput of the unfold
' If the profile is a knuckled, return no roll boundaries

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
    Const METHOD = "RollBoundaries: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler

    Dim oProfileWrapper As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part

    Dim oMfgPart As IJMfgProfilePart
    Dim oMfgProfileWrapper As MfgRuleHelpers.MfgProfilePartHlpr
    If oProfileWrapper.ProfileHasMfgPart(oMfgPart) Then
        Set oMfgProfileWrapper = New MfgRuleHelpers.MfgProfilePartHlpr
        Set oMfgProfileWrapper.object = oMfgPart
    Else
        Exit Function
    End If
    
    ' return no roll boundaries for knuckled profiles
    If ((oProfileWrapper.CurvatureType = PROFILE_CURVATURE_BendKnuckleAlongFlange) Or _
        (oProfileWrapper.CurvatureType = PROFILE_CURVATURE_BendKnuckleAlongWeb) Or _
        (oProfileWrapper.CurvatureType = PROFILE_CURVATURE_KnuckledAlongFlange) Or _
        (oProfileWrapper.CurvatureType = PROFILE_CURVATURE_KnuckledAlongWeb)) Then
                Exit Function
    End If
    
    Dim oMfgGeomCol2d As IJMfgGeomCol2d
    Set oMfgGeomCol2d = oMfgProfileWrapper.GetRollBoundaryLines
    
    If oMfgGeomCol2d Is Nothing Then
        'Since there are no roll boundaries we can leave the function
        GoTo CleanUp
    End If
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
      
    Dim oGeomCol2dOut As IJMfgGeomCol2d
    Set oGeomCol2dOut = m_oGeomCol2dFactory.Create(oResourceManager)

    Dim oNew2dGeom As IJMfgGeom2d

    Dim oMathGeom As IJDMfgGeomUtilWrapper
    Set oMathGeom = New GSCADMathGeom.MfgGeomUtilWrapper
    
    Dim oProfilePartSupport As IJProfilePartSupport
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New GSCADSDPartSupport.ProfilePartSupport
    
    Set oPartSupport.Part = oMfgPart.GetDetailedPart
    Set oProfilePartSupport = oPartSupport
    
    Dim dLowestPoint As Double, dHighestPoint As Double, dProfileHeight As Double, dFlangeWidth As Double
    oProfilePartSupport.GetThickness BottomFlange, dLowestPoint
    oProfilePartSupport.GetThickness TopFlange, dHighestPoint
    oProfilePartSupport.GetWebDepth dProfileHeight
    oProfilePartSupport.GetFlangeWidth dFlangeWidth
    
    Dim dWebThickness As Double
    oProfilePartSupport.GetThickness Web, dWebThickness
        
    Dim oGeom2d As IJMfgGeom2d
    Dim GeomType As StrMfgGeometryType
    
    Dim j As Integer
    For j = 1 To oMfgGeomCol2d.Getcount
        Set oGeom2d = oMfgGeomCol2d.GetGeometry(j)
        GeomType = oGeom2d.GetGeometryType
        If GeomType = STRMFG_PROFILE_ROLL_INFO Or _
            GeomType = STRMFG_PROFILE_BEND_INFO Then
    
            Dim oPointCollection As Collection
            Set oPointCollection = oMathGeom.CalculateProfileRollPositionsFromInfo(oGeom2d)
            
            Dim x As Double
            Dim i As Integer
            For i = 1 To oPointCollection.Count
                Dim oPoint As IJDPosition
                Set oPoint = oPointCollection.Item(i)
                
                'check only for odd number
                If i Mod 2 Then
                    'Roll radius 1000 metres or sweep angle 5 degrees ( 0.087 radians)
                    If Abs(oPoint.y) > 1000 Or oPoint.z < 0.0873016 Then
                        i = i + 1   ' skip this and next boundary
                        GoTo NextRollBoundary:
                    End If
                End If
                
                x = oPoint.x
                
                Dim oLine As IJLine
                Set oLine = New Line3d
                If GeomType = STRMFG_PROFILE_BEND_INFO Then
                    oLine.DefineBy2Points x, dLowestPoint, 0, x, (dProfileHeight - dHighestPoint), 0
                ElseIf GeomType = STRMFG_PROFILE_ROLL_INFO Then
                    oLine.DefineBy2Points x, dLowestPoint, 0, x, dFlangeWidth, 0
                End If
            
                Dim oCS As IJComplexString
                Set oCS = New ComplexString3d
                oCS.AddCurve oLine, True
                
                Set oNew2dGeom = m_oGeom2dFactory.Create(oResourceManager)
                oNew2dGeom.PutGeometry oCS
                
                oNew2dGeom.PutGeometrytype STRMFG_ROLL_BOUNDARIES_MARK
                oNew2dGeom.FaceId = UpSide
                
                oGeomCol2dOut.AddGeometry 1, oNew2dGeom

                Dim oSystemMark As IJMfgSystemMark
                Dim oObjSystemMark As IUnknown
                Dim oMarkingInfo As IJMarkingInfo
                Dim oMoniker As IMoniker
                
                'Create a SystemMark object to store additional information
                Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
                Set oObjSystemMark = oSystemMark
                                
                'Set the marking side
                oSystemMark.SetMarkingSide UpSide
                oSystemMark.Set2dGeometry oNew2dGeom
                
                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark
                
                oMarkingInfo.name = "ROLLBOUNDARY-" & i
                oMarkingInfo.Radius = Abs(oPoint.y) - (0.5 * dWebThickness)
                oMarkingInfo.FittingAngle = oPoint.z
                If oPoint.y < 0 Then
                    oMarkingInfo.direction = "down"
                Else
                    oMarkingInfo.direction = "up"
                End If
                
                Set oSystemMark = Nothing
                Set oObjSystemMark = Nothing
                Set oMarkingInfo = Nothing
                Set oMoniker = Nothing

                Set oLine = Nothing
                Set oCS = Nothing
                Set oNew2dGeom = Nothing
NextRollBoundary:
            Next i
        End If
        Set oGeom2d = Nothing
    Next j
    
    Set IJDMfgSystemMarkingRule_CreateAfterUnfold = oGeomCol2dOut

CleanUp:
    Set oMfgGeomCol2d = Nothing
    Set oProfileWrapper = Nothing
    Set oGeomCol2dOut = Nothing

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2014, , "RULES")
    GoTo CleanUp
End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d

End Function

 
