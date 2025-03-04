VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FittingProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgProfileMarking
' Module: Marking Rules
'
' Description:  Determines the marking settings for the mfg profile
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
'*******************************************************************************
Option Explicit

Implements IJDMfgSystemMarkingRule
Private Const MODULE = "MfgProfileMarking.FittingProfile"

Private Const TEEJOINTFITTINGMARK As Boolean = True
Private Const LAPJOINTFITTINGMARK As Boolean = False

Private Sub Class_Initialize()
    'Initialize the most used objects and helpers
    PrMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    'Clean up
    PrMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

End Function

Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "FittingProfile: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error GoTo ErrorHandler
    
    If TypeOf Part Is IJBeamPart Then Exit Function
    
    Dim oResourceManager As IUnknown
    Set oResourceManager = GetActiveConnection.GetResourceManager(GetActiveConnectionName)
    
    'Prepare the output collection of marking line's geometries
    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeomCol3dFactory.Create(oResourceManager)
    
    CreateAPSProfileMarkings STRMFG_FITTING_MARK, ReferenceObjColl, oGeomCol3d
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
    'Prepare collections
    Dim oSDProfileWrapper As Object
    'Create the SD profile Wrapper and initialize it
    If TypeOf Part Is IJStiffenerPart Then
        Set oSDProfileWrapper = New StructDetailObjects.ProfilePart
        Set oSDProfileWrapper.object = Part
    ElseIf TypeOf Part Is IJBeamPart Then
        Set oSDProfileWrapper = New StructDetailObjects.BeamPart
        Set oSDProfileWrapper.object = Part
    End If
    
    Dim oProfileWrapper As New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = Part
    
    Dim oMfgPart As IJMfgProfilePart
    Dim oMfgProfileWrapper As New MfgRuleHelpers.MfgProfilePartHlpr
    If oProfileWrapper.ProfileHasMfgPart(oMfgPart) Then
        Set oMfgProfileWrapper.object = oMfgPart
    Else
        Exit Function
    End If
    
    'Retrieve collection of ConnectionData objects
    Dim oConnectedObjectsColl As Collection
    'Set oConnectedObjectsColl = oSDProfileWrapper.ConnectedObjects
    Set oConnectedObjectsColl = GetPhysicalConnectionData(Part, ReferenceObjColl, True)
    
    If oConnectedObjectsColl Is Nothing Or oConnectedObjectsColl.Count = 0 Then
        'Since there is no connecting structure there will be no fitting marks required, therefore we can leave the function
        GoTo CleanUp
    End If
    
    Dim oConnectionData As ConnectionData
    Dim oLapConnectionData As ConnectionData
    Dim oConnectionDataTmp As ConnectionData
    Dim i As Integer
    
    'Find mounting plate ConnectionData object
    For i = 1 To oConnectedObjectsColl.Count
        oConnectionDataTmp = oConnectedObjectsColl.Item(i)
        
        If TypeOf oConnectionDataTmp.ToConnectable Is IJCollarPart Then GoTo NextItem
        If TypeOf oConnectionDataTmp.ToConnectable Is IJSmartPlate Then GoTo NextItem
        If TypeOf oConnectionDataTmp.ToConnectable Is IJProfilePart Then GoTo NextItem
        If Not TypeOf oConnectionDataTmp.ToConnectable Is IJPlatePart Then GoTo NextItem
        
        Dim res As Integer
        Dim oStructPort As IJStructPort
        Set oStructPort = oConnectionDataTmp.ConnectingPort
        res = oStructPort.ContextID And IMSStructConnection.CTX_LATERAL
        If res <= 0 Then GoTo NextItem
        
        If oStructPort.OperatorID = JXSEC_BOTTOM Then
            oConnectionData = oConnectionDataTmp
        ElseIf oMfgPart.PartType = STRMFG_EDGEREINFORCEMENT Then
            Dim oSDERWrapper As New StructDetailObjects.EdgeReinforcement
            Set oSDERWrapper.object = Part
        
            If oSDERWrapper.ReinforcedObject Is oConnectionDataTmp.ToConnectable Then
                oConnectionData = oConnectionDataTmp
            End If
            Set oSDERWrapper = Nothing
        ElseIf oStructPort.SectionID = -1 Then 'For lap connection port, section ID is returned as -1
            If CheckIfPartIsBracket(oConnectionDataTmp.ToConnectable) = True Then
                oLapConnectionData = oConnectionDataTmp
            End If
        End If
NextItem:
    Next i
    
    ' The below code is customer specific
    If LAPJOINTFITTINGMARK = True Then
        ' Check if there is lap connection data
        If Not oLapConnectionData.AppConnection Is Nothing Then
            Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
            Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
            
            Set oSDPartSupport.Part = oLapConnectionData.ToConnectable
            
            ' Get the PC based connected objects
            Dim oConObjCol As Collection, oConnCol As Collection, oThisPortCol As Collection, oOtherPortCol As Collection
            oSDPartSupport.GetConnectedObjects ConnectionPhysical, oConObjCol, oConnCol, oThisPortCol, oOtherPortCol
            
            ' Get the assembly sequence
            Dim oAssemblySequence As GSCADPlanningInterfaces.IJAssemblySequence
            Set oAssemblySequence = GetAssemblySequenceObject(Part)
            
            ' Get the assembly sequence number for the current part
            Dim lPartSeqIndex As Long
            lPartSeqIndex = oAssemblySequence.ChildIndex(Part)
            
            Dim oConnType       As ContourConnectionType
            Dim bIsCrossOfTee   As Boolean
            
            Dim iCount As Long
            For iCount = 1 To oConnCol.Count
                oSDPartSupport.GetConnectionTypeForContour oConnCol.Item(iCount), oConnType, bIsCrossOfTee
                If oConnType = PARTSUPPORT_CONNTYPE_LAP Then
                    ' Get the current index
                    Dim lChildSeqIndex  As Long
                    lChildSeqIndex = oAssemblySequence.ChildIndex(oConObjCol.Item(iCount))
                    
                    ' if the other connected part sequemce index is less than current part index
                    ' then skip this fitting mark
                    If lChildSeqIndex < lPartSeqIndex Then
                        GoTo NextFittingMark
                    End If
                End If
            Next
            
            Dim oConnection As IJAppConnection
            Set oConnection = oLapConnectionData.AppConnection
            
            oSDPartSupport.GetConnectionTypeForContour oConnection, _
                                                       oConnType, _
                                                       bIsCrossOfTee
    
            ' Check if the connection really is a lap connection
            If oConnType = PARTSUPPORT_CONNTYPE_LAP Then
                
                Dim oWBColl As Collection
                Dim bContourLap As Boolean
                
                Dim oSDPlateWrapper As StructDetailObjects.PlatePart
                Set oSDPlateWrapper = New StructDetailObjects.PlatePart
                Set oSDPlateWrapper.object = oLapConnectionData.ToConnectable
                
                ' Get the plate related lap contour so that fitting mark on Plate and profile match
                bContourLap = oSDPlateWrapper.Connection_ContourLap(oConnection, oWBColl)
                
                If ((bContourLap = True) And Not (oWBColl Is Nothing)) Then
                    
                    If oWBColl.Count = 0 Then
                        GoTo NextFittingMark
                    End If
                    
                    Dim oRefLapCurve       As IJCurve
                    Dim nWBIndex            As Long
                    Dim oLandingCurve       As IJWireBody
                    Dim oLandCurveVec       As IJDVector
                    Dim oTopologyLocate     As New TopologyLocate
                    
                    For nWBIndex = 1 To oWBColl.Count
                        Dim oPCWireBody As IJWireBody
                        Set oPCWireBody = oWBColl.Item(nWBIndex)
                        
                        Dim oCurveStartPos      As IJDPosition
                        Dim oCurveEndPos        As IJDPosition
                        
                        Dim oPCCurveVec         As IJDVector
                        Dim dDotProdVal         As Double
                        
                        ' Get the profile landing curve vector
                        If nWBIndex = 1 Then
                            Set oLandingCurve = Nothing
                            Set oLandCurveVec = Nothing
                            
                            Set oLandingCurve = oTopologyLocate.GetProfileParentWireBody(Part)
                            
                            oLandingCurve.GetEndPoints oCurveStartPos, oCurveEndPos
                            Set oLandCurveVec = oCurveStartPos.Subtract(oCurveEndPos)
                            oLandCurveVec.length = 1
                        End If
                        
                        Set oCurveStartPos = Nothing
                        Set oCurveEndPos = Nothing
                        oPCWireBody.GetEndPoints oCurveStartPos, oCurveEndPos
                        
                        Set oPCCurveVec = oCurveStartPos.Subtract(oCurveEndPos)
                        oPCCurveVec.length = 1
                        
                        dDotProdVal = oLandCurveVec.Dot(oPCCurveVec)
                        
                        ' Get the refernce lap curve(which is parallel to PC) needed to construct fitting mark
                        If Abs(dDotProdVal) > 0.99 Then
                        
                            Dim oPCCSColl   As IJElements
                            Set oPCCSColl = m_oMfgRuleHelper.WireBodyToComplexStrings(oWBColl.Item(nWBIndex))
                            
                            If Not oPCCSColl Is Nothing Then
                                Dim oPCCurve As IJCurve
                                Set oPCCurve = oPCCSColl.Item(1)
                                
                                Set oRefLapCurve = Nothing
                                Set oRefLapCurve = oPCCurve
                            End If
                        End If
                    Next nWBIndex
                    
                    Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double, dAlong As Double
                    oRefLapCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
                    
                    Dim oStartPos   As IJDPosition
                    Set oStartPos = New DPosition
                    
                    oStartPos.Set dStartX, dStartY, dStartZ
                    
                    Dim oRefLapCurveVec As IJDVector
                    Set oRefLapCurveVec = New DVector
                    
                    oRefLapCurveVec.Set dEndX - dStartX, dEndY - dStartY, dEndZ - dStartZ
                    oRefLapCurveVec.length = 1
                    
                    Dim oMfgGeomHelper      As IJMfgGeomHelper
                    Set oMfgGeomHelper = New MfgGeomHelper
                    
                    ' Create fitting mark at 2/3rd of lap length
                    dAlong = oRefLapCurve.length * (2 / 3)
                    
                    Dim oMidPosition    As IJDPosition
                    Set oMidPosition = oMfgGeomHelper.GetPointAtDistAlongCurve(oRefLapCurve, oStartPos, dAlong)
                    
                    Dim oProfileSupport As IJProfilePartSupport
                    Set oProfileSupport = New ProfilePartSupport
                    
                    Set oSDPartSupport = Nothing
                    Set oSDPartSupport = oProfileSupport
                    Set oSDPartSupport.Part = Part
                    
                    Dim pWebContourColl As Collection, pMonikerColl As Collection
                    Dim pWebSB As IJSurfaceBody
                
                    oProfileSupport.GetProfileContours WebLeftFace, pWebSB, pWebContourColl, pMonikerColl
                    
                    Dim oConnPartPort       As IJPort
                    Set oConnPartPort = oLapConnectionData.ToConnectedPort
                    
                    ' Create the fitting mark perpendicular to oRefLapCurveVec
                    Dim oFittingCS          As IJComplexString
                    Set oFittingCS = CreateLapConnFittingMark(oMidPosition, pWebSB, oConnPartPort.Geometry, oSDProfileWrapper.WebLength, oRefLapCurveVec, False)
                    
                    Dim oLapGeom3d As IJMfgGeom3d
                    Set oLapGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                    oLapGeom3d.PutGeometry oFittingCS
                    oLapGeom3d.PutGeometrytype STRMFG_FITTING_MARK
                    oLapGeom3d.PutSubGeometryType STRMFG_LAP_MARK
                    
                    Dim oLapSystemMark As IJMfgSystemMark
                    Set oLapSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
                    oLapSystemMark.SetMarkingSide UpSide
                    oLapSystemMark.Set3dGeometry oLapGeom3d
                    
                    'QI for the MarkingInfo object on the SystemMark
                    Dim oLapMarkingInfo As IJMarkingInfo
                    Set oLapMarkingInfo = oLapSystemMark
                    
                    'Set marking information, if any
                    oLapMarkingInfo.name = "PLATE FITTING"
                    
                    Dim oLapObjSystemMark As IUnknown
                    Dim oLapMoniker As IMoniker
                    
                    Set oLapObjSystemMark = oLapSystemMark
                    Set oLapMoniker = m_oMfgRuleHelper.GetMoniker(oLapConnectionData.AppConnection)
                    oLapGeom3d.PutMoniker oLapMoniker
                    oLapGeom3d.FaceId = UpSide
                    
                    oGeomCol3d.AddGeometry 1, oLapGeom3d
                    
                    Set oLapGeom3d = Nothing
                    Set oLapMoniker = Nothing
                    Set oLapSystemMark = Nothing
                    Set oLapMarkingInfo = Nothing
                    Set oLapObjSystemMark = Nothing
                    
                    Exit Function
                End If
                    
            End If
        End If
    End If

NextFittingMark:


    ' The below code is customer specific
    If TEEJOINTFITTINGMARK = True Then
        If Not oConnectionData.AppConnection Is Nothing Then
            'Get collection of mark positions
            Dim oMarkPosColl As Collection
            On Error Resume Next
            Set oMarkPosColl = GetFittingMarkPositions(Part, oConnectionData, UpSide)
            
            If Not (oMarkPosColl Is Nothing) Then
                Dim oWireBody As IJWireBody
                
                Dim oGeom3d As IJMfgGeom3d
                For i = 1 To oMarkPosColl.Count
                    Set oWireBody = CreateFittingMarkLine(Part, oMarkPosColl.Item(i), oConnectionData, UpSide)
                   
                    If Not (oWireBody Is Nothing) Then
                        Dim oCS As IJComplexString
                        Set oCS = m_oMfgRuleHelper.WireBodyToComplexString(oWireBody)
                        If Not (oCS Is Nothing) Then
                            Set oGeom3d = m_oGeom3dFactory.Create(oResourceManager)
                            oGeom3d.PutGeometry oCS
                            oGeom3d.PutGeometrytype STRMFG_FITTING_MARK
                        
                            Dim oSystemMark As IJMfgSystemMark
                            Set oSystemMark = m_oSystemMarkFactory.Create(oResourceManager)
                            oSystemMark.SetMarkingSide UpSide
                            oSystemMark.Set3dGeometry oGeom3d
                        
                            'QI for the MarkingInfo object on the SystemMark
                            Dim oMarkingInfo As IJMarkingInfo
                            Set oMarkingInfo = oSystemMark
                        
                            'Set marking information, if any
                            oMarkingInfo.name = "FITTING"
                            
                            Dim oObjSystemMark As IUnknown
                            Dim oMoniker As IMoniker
                           
                            Set oObjSystemMark = oSystemMark
                            Set oMoniker = m_oMfgRuleHelper.GetMoniker(oConnectionData.AppConnection)
                            oGeom3d.PutMoniker oMoniker
                            oGeom3d.FaceId = UpSide
                    
                            oGeomCol3d.AddGeometry 1, oGeom3d
                            
                            Set oStructPort = Nothing
                            Set oWireBody = Nothing
                            Set oCS = Nothing
                            Set oGeom3d = Nothing
                            Set oSystemMark = Nothing
                            Set oMarkingInfo = Nothing
                            Set oObjSystemMark = Nothing
                            Set oMoniker = Nothing
                        End If
                    End If
                Next i
            End If
        End If
    End If
        
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
CleanUp:
    Set oSDProfileWrapper = Nothing
    Set oProfileWrapper = Nothing
    Set oMfgPart = Nothing
    Set oMfgProfileWrapper = Nothing
    Set oConnectedObjectsColl = Nothing
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2004, , "RULES")
    GoTo CleanUp
End Function

 
