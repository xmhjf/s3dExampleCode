VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMfgFabMarginRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: GSCADSTRMFGRULES
' Module: CMfgFabMarginRule
'
' Description:  Determines the proposed settings for the Fabrication margin command
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
' 2006.05.18    SumaKiran     CR-67562 Option to have the corner features moved when placing margin
' 2006.08.25    SumaKiran     CR-100253  Add Margin to stiffeners by selecting bounding plate system
'*******************************************************************************
Option Explicit
Const MODULE = "CMfgFabricationMarginRule:"

'- PlateSystem ------------------------------------------------------------'
Private Const IJPlateSystem = "{E0B23CD4-7CEB-11d3-B351-0050040EFC17}"
Private Const m_strFilterNameConnectedPartFilter = "cmbConnectedPartFilter"
Private Const m_strFilterNamePartFilter = "cmbPartFilter"
Private Const Margin_Feature_Check As Long = 1 '0 - FirstFeature, 1 - End Cut Feature

''- StiffenerSystem ----------------------------------------------------------'
'Private Const IJStiffenerSystem = "{E0B23CD5-7CEB-11d3-B351-0050040EFC17}"

Implements IJDFabMarginRule2
Implements IJDFabMarginFilter
Implements IJDMarginByConnRule
Private Enum enumInputColl
    eMrgnPartObject = 1
    eMrgnEdgeObject = 2
    eMrgnType = 3
    eMrgnGeomChange = 4
    eMrgnDirection = 5
    eMrgnStretchOffset = 6
    eMrgnValue1 = 7
    eMrgnValue2 = 8
End Enum
Private m_oStiffPlatePartColl As Collection

Private Sub Class_Initialize()
    Set m_oStiffPlatePartColl = New Collection
End Sub

Private Sub Class_Terminate()
    If Not m_oStiffPlatePartColl Is Nothing Then
        If m_oStiffPlatePartColl.Count <> 0 Then
            While m_oStiffPlatePartColl.Count <> 0
                m_oStiffPlatePartColl.Remove (m_oStiffPlatePartColl.Count)
            Wend
        End If
        Set m_oStiffPlatePartColl = Nothing
    End If
End Sub
Private Function IJDFabMarginFilter_GetCriteria(MarginCmdFilter As GSCADMfgRulesDefinitions.MarginCmdFilter, ByVal MarginFilter As Long, ByVal MarginCmdProcess As GSCADMfgRulesDefinitions.MarginCmdProcess) As String
Const METHOD = "IJDFabMarginFilter_GetCriteria"
    On Error GoTo ErrorHandler

    Dim strFilterCriteria As String
    
    If MarginCmdProcess = ConnectByAssemblies Then
    
        If MarginCmdFilter = MarginAssembly Then
            Select Case MarginFilter
                Case 1
                    strFilterCriteria = "IJAssembly"
            End Select
        ElseIf MarginCmdFilter = MarginConnectedAssembly Then
            Select Case MarginFilter
                Case 1
                    strFilterCriteria = "IJAssembly AND [GSCADSTRMFGRules.FilterFuncs,IsValidAssy]"
            End Select
        End If
    
    End If

    If MarginCmdFilter = MarginConnectedPart Then
        Select Case MarginFilter
    
            Case 1
                If GetProjectRootNestingStage() = PRODUCTION_STAGE Then
                    strFilterCriteria = "IJPlatePart AND NOT [STFilterFunctions.StructFilterFunctions,IsALightPart]" _
                                        & " AND [StrMfgFabMargin.FabMarginFilter,IsAMfgParent]"
                Else
                    ' Filter criteria for prenesting
                    strFilterCriteria = "IJPlatePart AND [StrMfgFabMargin.FabMarginFilter,IsAMfgParent]"
                End If
    
            Case 2
                If GetProjectRootNestingStage() = PRODUCTION_STAGE Then
                    strFilterCriteria = "IJProfilePart AND NOT [STFilterFunctions.StructFilterFunctions,IsALightPart]" _
                                        & " AND [StrMfgFabMargin.FabMarginFilter,IsAMfgParent]"
                Else
                    ' Filter criteria for prenesting
                    strFilterCriteria = "IJProfilePart AND [StrMfgFabMargin.FabMarginFilter,IsAMfgParent]"
                End If
    
            Case 3
                strFilterCriteria = IJPlateSystem
    
    '        Case "Stiffner System"
    '            strFilterCriteria = IJStiffenerSystem
            Case 4
                strFilterCriteria = "ISPSMemberPartPrismatic"
    
        End Select
        
    ElseIf MarginCmdFilter = MarginPart Then
        Select Case MarginFilter
            Case 1 'plate parts only (no brackets and collars)
                If GetProjectRootNestingStage() = PRODUCTION_STAGE Then
                    strFilterCriteria = "IJPlatePart AND NOT [STFilterFunctions.StructFilterFunctions,IsALightPart]" _
                                        & " AND [StrMfgFabMargin.FabMarginFilter,IsAMfgParent]" _
                                        & " AND NOT [StrMfgMarginByConnCmd.FilterFunctions,IsBracketPlate]" _
                                        & " AND NOT IJCollarPart"
                Else
                    ' Filter criteria for prenesting
                    strFilterCriteria = "IJPlatePart AND [StrMfgFabMargin.FabMarginFilter,IsAMfgParent]"
                End If
    
            Case 2 'profile parts
                If GetProjectRootNestingStage() = PRODUCTION_STAGE Then
                    strFilterCriteria = "IJProfilePart AND NOT [STFilterFunctions.StructFilterFunctions,IsALightPart]" _
                                        & " AND [StrMfgFabMargin.FabMarginFilter,IsAMfgParent]"
                Else
                    ' Filter criteria for prenesting
                    strFilterCriteria = "IJProfilePart AND [StrMfgFabMargin.FabMarginFilter,IsAMfgParent]"
                End If
    
            Case 3 'members
                strFilterCriteria = "ISPSMemberPartPrismatic"
                
            Case 4 'brackets
                strFilterCriteria = "IJPlatePart AND NOT [STFilterFunctions.StructFilterFunctions,IsALightPart]" _
                                        & " AND [StrMfgFabMargin.FabMarginFilter,IsAMfgParent]" _
                                        & " AND [StrMfgMarginByConnCmd.FilterFunctions,IsBracketPlate]" _
                
            Case 5 'collars
                strFilterCriteria = "IJCollarPart"
    
        End Select

    End If

    IJDFabMarginFilter_GetCriteria = strFilterCriteria

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 6004, , "RULES")
End Function

Private Function IJDFabMarginRule2_GetAllowedDirections(ByVal pDispPartObj As IJDObject, ByVal pDispEdgeObj As IJDObject, ByVal eMarginMode As GSCADMfgRulesDefinitions.EnumMarginMode, ByVal lMarginType As Long) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
Const METHOD = "IJDFabMarginRule2_GetAllowedDirections"
On Error GoTo ErrorHandler
    
    Dim lTempColl As JCmnShp_CollectionAlias
    Set lTempColl = New Collection
        
    lTempColl.Add (1)
    lTempColl.Add (2)

    Set IJDFabMarginRule2_GetAllowedDirections = lTempColl
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 6005, , "RULES")
End Function

Private Function IJDFabMarginRule2_GetFabMarginFeatureMove(ByVal pDispObj As Object, ByVal pFeatureObj As Object, ByVal MarginType As Long, ByVal MarginMode As GSCADMfgRulesDefinitions.EnumMarginMode, ByVal MarginVal1 As Double, ByVal EndVal1 As Double) As GSCADMfgRulesDefinitions.EnumStrMfgMarginFeature
Const METHOD = "IJDFabMarginRule_GetFabMarginFeatureMove"
    On Error GoTo ErrorHandler
    
    Dim lFeature As EnumStrMfgMarginFeature
    lFeature = FixFeature
    
    'If sketched feature is supplied then code should always return "Fixed Feature" and it should
    'not call get_StructFeatureType as Query Interface for IJStructFeature is not supported
    
    If Not TypeOf pFeatureObj Is IJSketchFeature Then
        
        Dim oFeature As IJStructFeature
        Set oFeature = pFeatureObj
        
        ' Move Corner features only.
        If oFeature.get_structfeaturetype = SF_CornerFeature Then
            lFeature = MoveFeature
        End If
        
    End If
    
    IJDFabMarginRule2_GetFabMarginFeatureMove = lFeature

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 6006, , "RULES")
End Function
Private Sub IJDFabMarginRule2_GetFabMarginParameters(ByVal pDispPartObj As GSCADMfgRulesDefinitions.IJDObject, _
                                                        ByVal pDispEdgeObj As GSCADMfgRulesDefinitions.IJDObject, _
                                                        MarginType As Long, _
                                                        MarginMode As GSCADMfgRulesDefinitions.EnumMarginMode, _
                                                        eMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                                        eMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                                        dStretchOffset As Double, _
                                                        MarginVal1 As Double, _
                                                        Optional MarginVal2 As Double)
                                                        
Const METHOD = "IJDFabMarginRule_GetFabMarginParameters"
On Error GoTo ErrorHandler

Dim oSDPlatePart As New StructDetailObjects.PlatePart
Dim oPartObj As Object
Dim oCollObj As Object
Dim oCurvature  As PlateCurvature
Dim oWorkcenter As String
Dim oCenterline As Boolean
Dim oAssembly As MfgRuleHelpers.MfgAssemblyHelper
Dim oMfgPlatePart As New MfgRuleHelpers.PlatePartHlpr
Const strConstMargin = "ConstMargin"
Const strObliqueMargin = "ObliqueMargin"
Const m_strJigFloor = "JigFloor"

Set oPartObj = pDispPartObj

If TypeOf oPartObj Is IJPlatePart Then
    
    If Not pDispEdgeObj Is Nothing Then
        
        Dim oPort As IJStructPort
        Set oPort = pDispEdgeObj
        
        Dim oStructDetailHelper As IJStructDetailHelper
        Set oStructDetailHelper = New StructDetailHelper
    
        Dim oOperator As Object
        Dim oOperation As IJStructOperation
        
        'check if the plate is a cylindrical plate
        If CheckIfPlateIsCanPlate(pDispPartObj) Then
            
            Dim oPlatePart As IJPlatePart
            Set oPlatePart = pDispPartObj
            
            'Dim oSDPlatePart As New StructDetailObjects.PlatePart
            Set oSDPlatePart.object = oPlatePart
                
            Dim dPlateThickness As Double
            dPlateThickness = oSDPlatePart.PlateThickness 'oPlate1.thickness

            Dim oPortColl As IJElements
            GetEligiblePortCollectionForCan pDispPartObj, oPortColl
            
            If dPlateThickness >= 0.03 Then
                Dim oFacePort As IJPort
                Set oFacePort = EdgeToFacePort(pDispEdgeObj)
                If oPortColl.Contains(oFacePort) Then
                    MarginType = 7 'clamp
                    MarginVal1 = 2 * dPlateThickness * 1000 'converting dplatethickness into mm
                    Exit Sub
                End If
            End If
        End If
        'The operator has to be a seam
        oStructDetailHelper.FindOperatorForOperationInGraphByID oPartObj, oPort.OperationID, oPort.OperatorID, oOperation, oOperator
        
        If Not oOperator Is Nothing Then
            'if edge is from a planning seam, return 6mm as the margin value
            If TypeOf oOperator Is IJSeam Then
                Dim oSeamType As IJDSeamType
                Set oSeamType = oOperator
                Dim eSeamType As StructSeamType
                eSeamType = -1
                eSeamType = oSeamType.SeamType
                
                If eSeamType = sstPlanningSeam Then
                    MarginMode = ConstMargin
                    MarginType = 1
                    eMarginGeomChange = 1
                    eMarginDirection = 1 ' PerpendicularToEdge
                    dStretchOffset = 0#
                    MarginVal1 = 6 'Start Value
                    Exit Sub
                End If
            End If
        End If
        
        'Need to get the PC from the edge and check for the type
        Dim oEdgePort As IJPort
        Set oEdgePort = oPort
                 
        'Dim oFacePort As IJPort
        Set oFacePort = EdgeToFacePort(oEdgePort)
                 
        'get all the physical connections from the face port
        Dim oConnectionsColl As IJElements
        oFacePort.enumConnections oConnectionsColl, ConnectionPhysical, ConnectionStandard
        
        If Not oConnectionsColl Is Nothing Then
           If oConnectionsColl.Count > 0 Then
                Dim oPhysicalConnection As Object
                Set oPhysicalConnection = oConnectionsColl.Item(1) 'Ideally should be only one
                               
                'check for the bevel type
                Dim oPartSupport   As IJPartSupport
                Set oPartSupport = New PartSupport
                Set oPartSupport.Part = oPartObj 'plate part
        
                Dim bIsCrossOfTee As Boolean
                Dim oConnType As ContourConnectionType
                
                oPartSupport.GetConnectionTypeForContour oPhysicalConnection, _
                                                         oConnType, _
                                                         bIsCrossOfTee
                
                      
                If oConnType = PARTSUPPORT_CONNTYPE_BUTT Then
                    MarginMode = ConstMargin
                    MarginType = 1
                    eMarginGeomChange = 1
                    eMarginDirection = 1 ' PerpendicularToEdge
                    dStretchOffset = 0#
                    MarginVal1 = 6 'Start Value
                    Exit Sub
                ElseIf oConnType = PARTSUPPORT_CONNTYPE_TEE Then
                    MarginMode = ConstMargin
                    MarginType = 1
                    eMarginGeomChange = 1
                    eMarginDirection = 1 ' PerpendicularToEdge
                    dStretchOffset = 0#
                    MarginVal1 = 15 'Start Value
                    Exit Sub
                End If
                
            End If 'oConnectionsColl.Count
         End If 'oConnectionsColl Is Nothing check
        
    End If 'pDispEdgeObj Is Nothing

    Set oSDPlatePart.object = oPartObj
    oCenterline = oSDPlatePart.HasCenterLineIntersection()
    Err.Clear 'If the plate part doesnt have a center line intersection, this method is
              'given an error. So clearing the error.

    Set oMfgPlatePart.object = oPartObj

    oCurvature = oMfgPlatePart.CurvatureType
    Set oAssembly = New MfgRuleHelpers.MfgAssemblyHelper
    oWorkcenter = oAssembly.WorkCenterFromPart(oPartObj)

    Select Case oCurvature
    Case PLATE_CURVATURE_Flat
        MarginMode = ConstMargin
        eMarginGeomChange = 1
        MarginType = 1
        If oWorkcenter = m_strJigFloor Then
            MarginVal1 = 10 'Start Value
        Else
            MarginVal1 = 15 'Start Value
        End If

    Case PLATE_CURVATURE_Knuckled
        MarginMode = ConstMargin
        MarginType = 2
        eMarginGeomChange = 1
        If oWorkcenter = m_strJigFloor Then
            MarginVal1 = 10 'Start Value
        Else
            MarginVal1 = 15 'Start Value
        End If
    Case PLATE_CURVATURE_SingleCurvature
        MarginMode = ConstMargin
        MarginType = 2
        eMarginGeomChange = 1
        If oWorkcenter = m_strJigFloor Then
            MarginVal1 = 10 'Start Value
        Else
            MarginVal1 = 15 'Start Value
        End If
    Case PLATE_CURVATURE_SingleCurvature_Knuckled
            MarginMode = ConstMargin
            MarginType = 2
            eMarginGeomChange = 1
            MarginVal1 = 20 'Start Value
    Case PLATE_CURVATURE_DoubleCurvature
        If oCenterline = True Then
            MarginMode = ObliqueMargin
            MarginType = 1
            eMarginGeomChange = 1
            MarginVal1 = 5 'Start Value
            MarginVal2 = 22 'End Value
        Else
            MarginMode = ConstMargin
            MarginType = 1
            eMarginGeomChange = 1
            MarginVal1 = 20 'Start Value
        End If
    Case PLATE_CURVATURE_DoubleCurvature_Knuckled
        If oCenterline = True Then
            MarginMode = ObliqueMargin
            MarginType = 1
            eMarginGeomChange = 1
            MarginVal1 = 6 'Start Value
            MarginVal2 = 28 'End Value
        Else
            MarginMode = ConstMargin
            MarginType = 1
            eMarginGeomChange = 1
            MarginVal1 = 25 'Start Value
        End If
    Case PLATE_CURVATURE_Curved
        If oCenterline = True Then
            MarginMode = ObliqueMargin
            MarginType = 1
            eMarginGeomChange = 1
            MarginVal1 = 5 'Start Value
            MarginVal2 = 22 'End Value
        Else
            MarginMode = ConstMargin
            MarginType = 2
            eMarginGeomChange = 1
            MarginVal1 = 20 'Start Value
        End If
    End Select
    eMarginDirection = 1 ' PerpendicularToEdge
ElseIf TypeOf oPartObj Is IJStructProfilePart Then
    'Not sure about the rules values when profile is selected
    MarginMode = ConstMargin
    MarginType = 1
    eMarginGeomChange = 1
    MarginVal1 = 10 'Start Value
    eMarginDirection = 2     'global
Else
    MarginMode = ConstMargin
    MarginType = 1
    eMarginGeomChange = 1
    eMarginDirection = 2
    MarginVal1 = 10 'Start Value
End If

If eMarginGeomChange <> 2 Then
    dStretchOffset = 0#
Else
    dStretchOffset = 1#
End If

Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 6007, , "RULES")
End Sub
Private Function IJDFabMarginRule2_GetFabMarginTypes(ByVal pDispPart As Object, _
                                                        ByVal MarginMode As GSCADMfgRulesDefinitions.EnumMarginMode) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
Const METHOD = "IJDFabMarginRule2_GetFabMarginTypes"
On Error GoTo ErrorHandler

    Dim lTempColl As JCmnShp_CollectionAlias
    Set lTempColl = New Collection
    
    If (TypeOf pDispPart Is IJPlatePart) Then
       If (MarginMode = ConstMargin) Then
            lTempColl.Add (1) 'Heating Margin
            lTempColl.Add (2) 'Bending Margin
            lTempColl.Add (4) 'Intercoastal Margin
            lTempColl.Add (8) 'Coaming Margin
            lTempColl.Add (3) 'EndFace Margin
                        lTempColl.Add (9) 'PanelEnd Margin
       ElseIf (MarginMode = ObliqueMargin) Then
            lTempColl.Add (1) 'Heating Margin
            lTempColl.Add (2) 'Bending Margin
            lTempColl.Add (4) 'Intercoastal Margin
            lTempColl.Add (8) 'Coaming Margin
            lTempColl.Add (3) 'EndFace Margin
                        lTempColl.Add (9) 'PanelEnd Margin
       End If
    ElseIf (TypeOf pDispPart Is IJStructProfilePart) Then
           lTempColl.Add (7) 'Clamp Margin
           lTempColl.Add (3) 'EndFace Margin
           lTempColl.Add (8) 'Coaming Margin
           lTempColl.Add (1) 'Heating Margin
           lTempColl.Add (2) 'Bending Margin
           lTempColl.Add (4) 'Intercoastal Margin
    End If
    
    lTempColl.Add (100) 'Generic Margin - Non Trimmed
    lTempColl.Add (101) 'Generic Margin - Trimmed
    
    Set IJDFabMarginRule2_GetFabMarginTypes = lTempColl

Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 6008, , "RULES")
End Function

Private Sub IJDMarginByConnRule_GetConnByAssyGroupParameters(ByVal pDispAssyObj As Object, ByVal pDispConnAssyObj As Object, GroupMarginType As Long, eGroupMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, eGroupMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, dGroupStretchOffset As Double, dGroupMarginVal As Double, eMarginOffsetDirection As GSCADMfgRulesDefinitions.EnumMarginOffsetDirection)
    On Error GoTo ErrorHandler
    Const METHOD = "IJDMarginByConnRule_GetConnByAssyGroupParameters"
    
    GroupMarginType = 1 'Heating
    eGroupMarginGeomChange = AsMargin
    eGroupMarginDirection = PerpendicularToEdge
    dGroupStretchOffset = 0# 'In mm
     
    dGroupMarginVal = 15 'Start Value(In mm)
   
    eMarginOffsetDirection = GetGroupOffsetDirection(pDispAssyObj, pDispConnAssyObj)

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub IJDMarginByConnRule_GetMarginParametersForPart(ByVal pDispPartObj As GSCADMfgRulesDefinitions.IJDObject, _
                                                           ByVal pDispEdgeObj As GSCADMfgRulesDefinitions.IJDObject, _
                                                           ByVal pDispAssyObj As Object, ByVal pDispConnAssyObj As Object, _
                                                           GroupMarginType As Long, _
                                                           eGroupMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                                           eGroupMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                                           dGroupStretchOffset As Double, dGroupMarginVal As Double, _
                                                           eMarginOffsetDirection As GSCADMfgRulesDefinitions.EnumMarginOffsetDirection, _
                                                           MarginType As Long, _
                                                           MarginMode As GSCADMfgRulesDefinitions.EnumMarginMode, _
                                                           eMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                                           eMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                                           dStretchOffset As Double, MarginVal1 As Double, Optional MarginVal2 As Double)


    On Error GoTo ErrorHandler
    Const METHOD = "IJDMarginByConnRule_GetMarginParametersForPart"

    Dim oAssemblyChild As IJAssemblyChild
    Set oAssemblyChild = pDispPartObj
    
    'Initializing the values.
    MarginMode = -1
    MarginType = 1
    MarginVal1 = 0#
    eMarginGeomChange = AsMargin
    eMarginDirection = PerpendicularToEdge
    dStretchOffset = 0#
    
    Dim oAssyParent As Object
    Set oAssyParent = oAssemblyChild.Parent
    If Not oAssyParent Is Nothing Then
        If Not TypeOf oAssemblyChild.Parent Is IJAssemblyBase Then 'Margin for parts under "Unassigned Parts" folder.
             Exit Sub
        End If
    End If
    
    'Exclude brackets and collar plates
    If TypeOf pDispPartObj Is IJPlatePart Then
            
        Dim oPlateUtils As IJPlateAttributes
        Set oPlateUtils = New PlateUtils
        
        Dim oObjUtil As IJDMfgGeomUtilWrapper
        Set oObjUtil = New MfgGeomUtilWrapper
        
        Dim oSys As IJSystem
        Set oSys = oObjUtil.GetRootSystem(pDispPartObj)
        
        'Check if it is a bracket plate
        If Not oSys Is Nothing Then
            If (oPlateUtils.IsBracketByPlane(oSys) Or oPlateUtils.IsTrippingBracket(oSys)) Then
                Exit Sub
            End If
        End If
        
        'Check for collar plates
        If (TypeOf pDispPartObj Is IJSmartPlate) Or (TypeOf pDispPartObj Is IJCollarPart) Then
           Exit Sub
        End If
    
    End If
        
    MarginMode = AssemblyMargin
  
    'Margin value would be computed based on orientation
    If TypeOf pDispPartObj Is IJPlatePart Then
        
        GetPlateMarginParameters pDispPartObj, pDispEdgeObj, GroupMarginType, eGroupMarginGeomChange, eGroupMarginDirection, dGroupStretchOffset, dGroupMarginVal, eMarginOffsetDirection, _
                                 MarginType, eMarginGeomChange, eMarginDirection, dStretchOffset, MarginVal1, MarginVal2
       
    ElseIf TypeOf pDispPartObj Is IJProfilePart Then
        
        Dim oStiffenedPlatePart As Object
        Dim oStiffenedPlatePartEdge As Object
        GetStiffenedPlatePart pDispPartObj, pDispAssyObj, pDispConnAssyObj, oStiffenedPlatePart, oStiffenedPlatePartEdge
        
        If oStiffenedPlatePart Is Nothing Then ' This is done as part of TR258032
            MarginType = 1
            eMarginGeomChange = 1
            MarginVal1 = 10 'Start Value
            eMarginDirection = 2     'global
            Exit Sub
        End If
        
        'Get the rule based values for the stiffened plate
        Dim lPlateMarginType As Long, lPlateGeomChange As Long, lPlateDirection As Long
        Dim dPlateStretchOffset As Double, dPlateMarginVal1 As Double, dPlateMarginVal2 As Double
        
        GetPlateMarginParameters oStiffenedPlatePart, oStiffenedPlatePartEdge, GroupMarginType, eGroupMarginGeomChange, eGroupMarginDirection, dGroupStretchOffset, dGroupMarginVal, eMarginOffsetDirection, _
                                 lPlateMarginType, lPlateGeomChange, lPlateDirection, dPlateStretchOffset, dPlateMarginVal1, dPlateMarginVal2
                
        GetProfileParametersFromPlate pDispPartObj, oStiffenedPlatePartEdge, lPlateMarginType, lPlateGeomChange, lPlateDirection, dPlateStretchOffset, dPlateMarginVal1, dPlateMarginVal2, _
                                      MarginType, eMarginGeomChange, eMarginDirection, dStretchOffset, MarginVal1, MarginVal2
        
    Else ''Member case
    '' exit with default values
        eMarginDirection = GlobalDirection
        MarginVal1 = 15
    End If
        
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub AddStiffenedPlatePartToCache(ByVal pDispPartObj As Object, _
                                     ByVal pDispEdgeObj As Object, _
                                     MarginType As Long, _
                                     eMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                     eMarginDir As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                     dStretchOffset As Double, MarginVal1 As Double, Optional MarginVal2 As Double)
                                     
    Dim oInputColl As Dictionary
    Set oInputColl = New Dictionary
    
    oInputColl.Add eMrgnPartObject, pDispPartObj
    oInputColl.Add eMrgnEdgeObject, pDispEdgeObj
    oInputColl.Add eMrgnType, MarginType
    oInputColl.Add eMrgnGeomChange, eMarginGeomChange
    oInputColl.Add eMrgnDirection, eMarginDir
    oInputColl.Add eMrgnStretchOffset, dStretchOffset
    oInputColl.Add eMrgnValue1, MarginVal1
    oInputColl.Add eMrgnValue2, MarginVal2

    m_oStiffPlatePartColl.Add oInputColl
    
End Sub

Private Sub CheckAndGetThePlateParametersFromCacheIfExists(ByVal pDispPartObj As Object, _
                                                           ByVal pDispEdgeObj As Object, _
                                                           bExistsInCache As Boolean, _
                                                           MarginType As Long, _
                                                           eMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                                           eMarginDir As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                                           dStretchOffset As Double, MarginVal1 As Double, Optional MarginVal2 As Double)

    bExistsInCache = False
    
    Dim lCount As Long
    Dim lIndx As Long
    lCount = m_oStiffPlatePartColl.Count
    lIndx = 1
    
    For lIndx = 1 To lCount
        Dim oTempDict As Dictionary
        Set oTempDict = m_oStiffPlatePartColl.Item(lIndx)
        If oTempDict.Item(eMrgnPartObject) Is pDispPartObj Then
            If oTempDict.Item(eMrgnEdgeObject) Is pDispEdgeObj Then
                MarginType = oTempDict.Item(eMrgnType)
                eMarginGeomChange = oTempDict.Item(eMrgnGeomChange)
                eMarginDir = oTempDict.Item(eMrgnDirection)
                dStretchOffset = oTempDict.Item(eMrgnStretchOffset)
                MarginVal1 = oTempDict.Item(eMrgnValue1)
                MarginVal2 = oTempDict.Item(eMrgnValue2)
                bExistsInCache = True
                Exit For
            End If
        End If
    Next
End Sub

Private Function IJDMarginByConnRule_GetConnectedObjectsForPart(ByVal pDispObj As Object, _
                                                                ByVal pDispSelAssy As Object, ByVal pDispConnAssy As Object) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
    
    On Error GoTo ErrorHandler
    Const METHOD = "IJDMarginByConnRule_GetConnectedObjectsForPart"
    
    Set IJDMarginByConnRule_GetConnectedObjectsForPart = New Collection
    
    Dim iCnt As Integer
    
    If TypeOf pDispObj Is IJPlatePart Then
        
        'Return the connected stiffeners at the assembly connections
            'a. Get all the connected stiffener parts
            'b. Get the stiffeners at the connections
            'c. Filter the connected stiffeners
        
        'a. Get all the connected stiffener parts
        Dim oStiffenerColl As IJElements
        Set oStiffenerColl = GetConnectedStiffeners(pDispObj)
        
        'b. Get the stiffeners at the conenctions
        Dim oMfgUtilAssyMargin As IJMfgUtilAssyMargin
        Set oMfgUtilAssyMargin = New MfgUtilAssyMargin
                    
        Dim oConnections As IJElements
        Set oConnections = GetConnectionsBetweenAssemblies(pDispSelAssy, pDispConnAssy)  'Gives the face ports
        
        'c. Filter the connected stiffeners
        For iCnt = 1 To oConnections.Count
            Dim oPort As IJPort
            Set oPort = oConnections.Item(iCnt)
            
            If oStiffenerColl.Contains(oPort.Connectable) Then
                IJDMarginByConnRule_GetConnectedObjectsForPart.Add oPort.Connectable
            End If
            
            Set oPort = Nothing
            
        Next iCnt
               
    ElseIf TypeOf pDispObj Is IJStructProfilePart Then
        'Return Nothing
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Private Sub IJDMarginByConnRule_GetDependentObjectParameters(ByVal pDispPartObj As GSCADMfgRulesDefinitions.IJDObject, _
                                                             ByVal pDispEdgeObj As GSCADMfgRulesDefinitions.IJDObject, _
                                                             ByVal pDispParentEdgeObj As GSCADMfgRulesDefinitions.IJDObject, _
                                                             ParentMarginType As Long, _
                                                             eParentMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                                             eParentMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                                             dParentStretchOffset As Double, dParentMarginVal1 As Double, dParentMarginVal2 As Double, _
                                                             MarginType As Long, _
                                                             eMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                                             eMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                                             dStretchOffset As Double, MarginVal1 As Double, MarginVal2 As Double)

    On Error GoTo ErrorHandler
    Const METHOD = "IJDMarginByConnRule_GetDependentObjectParameters"
    
    GetProfileParametersFromPlate pDispPartObj, pDispParentEdgeObj, ParentMarginType, eParentMarginGeomChange, eParentMarginDirection, dParentStretchOffset, dParentMarginVal1, dParentMarginVal2, _
                                  MarginType, eMarginGeomChange, eMarginDirection, dStretchOffset, MarginVal1, MarginVal2
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
Private Function IJDMarginByConnRule_GetMarginByAssyConnTypes() As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
    On Error GoTo ErrorHandler
    Const METHOD = "IJDMarginByConnRule_GetMarginByAssyConnTypes"
    
    Dim lTempColl As JCmnShp_CollectionAlias
    Set lTempColl = New Collection

    lTempColl.Add (1) 'Heating Margin
    lTempColl.Add (2) 'Bending Margin
    lTempColl.Add (4) 'Intercoastal Margin
    lTempColl.Add (8) 'Coaming Margin
    lTempColl.Add (3) 'EndFace Margin
        lTempColl.Add (9) 'PanelEnd Margin
  
    lTempColl.Add (100) 'Non Trimmed
    lTempColl.Add (101) 'Trimmed

    Set IJDMarginByConnRule_GetMarginByAssyConnTypes = lTempColl

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub IJDMarginByConnRule_GetConnByPartsGroupParameters(ByVal pDispSelObj As Object, ByVal pDispConnObj As Object, GroupMarginType As Long, eGroupMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, eGroupMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, dGroupStretchOffset As Double, dGroupMarginVal As Double)
    On Error GoTo ErrorHandler
    Const METHOD = "IJDMarginByConnRule_GetConnByPartsGroupParameters"
    
    GroupMarginType = 1 'Heating
    eGroupMarginGeomChange = AsMargin
    eGroupMarginDirection = PerpendicularToEdge ' PerpendicularToEdge
    dGroupStretchOffset = 0#
    
    dGroupMarginVal = 15 'Start Value
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub IJDMarginByConnRule_GetMarginParametersForConnPart(ByVal pDispPartObj As GSCADMfgRulesDefinitions.IJDObject, _
                                                               ByVal pDispEdgeObj As GSCADMfgRulesDefinitions.IJDObject, _
                                                               ByVal pDispSelObj As Object, ByVal pDispConnObj As Object, _
                                                               GroupMarginType As Long, _
                                                               eGroupMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                                               eGroupMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                                               dGroupStretchOffset As Double, dGroupMarginVal As Double, _
                                                               MarginType As Long, _
                                                               MarginMode As GSCADMfgRulesDefinitions.EnumMarginMode, _
                                                               eMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                                               eMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                                               dStretchOffset As Double, MarginVal1 As Double, Optional MarginVal2 As Double)
                                                               
                                                               
    On Error GoTo ErrorHandler
    Const METHOD = "IJDMarginByConnRule_GetMarginParametersForConnPart"
    
    'Exclude brackets and collar plates
    If TypeOf pDispPartObj Is IJPlatePart Then
            
        Dim oPlateUtils As IJPlateAttributes
        Set oPlateUtils = New PlateUtils
        
        Dim oObjUtil As IJDMfgGeomUtilWrapper
        Set oObjUtil = New MfgGeomUtilWrapper
        
        Dim oSys As IJSystem
        Set oSys = oObjUtil.GetRootSystem(pDispPartObj)
        
        'Check if it is a bracket plate
        If Not oSys Is Nothing Then
            If (oPlateUtils.IsBracketByPlane(oSys) Or oPlateUtils.IsTrippingBracket(oSys)) Then
                MarginMode = ConstMargin
                MarginType = 1
                eMarginGeomChange = AsMargin
                eMarginDirection = PerpendicularToEdge
                dStretchOffset = 0#
                MarginVal1 = 10 'Start Value
                
                Exit Sub
            End If
        End If
        
        'Check for collar plates
        If (TypeOf pDispPartObj Is IJSmartPlate) Or (TypeOf pDispPartObj Is IJCollarPart) Then
            MarginMode = ConstMargin
            MarginType = 1
            eMarginGeomChange = AsMargin
            eMarginDirection = PerpendicularToEdge
            dStretchOffset = 0#
            MarginVal1 = 10 'Start Value
            
            Exit Sub
        End If
    
    End If
    
    MarginMode = ConstMargin
                
    If TypeOf pDispPartObj Is IJPlatePart Then
        
        GetConnPlateMarginParameters pDispPartObj, pDispEdgeObj, GroupMarginType, eGroupMarginGeomChange, eGroupMarginDirection, dGroupStretchOffset, dGroupMarginVal, _
                                     MarginType, eMarginGeomChange, eMarginDirection, dStretchOffset, MarginVal1, MarginVal2
       
            
    ElseIf TypeOf pDispPartObj Is IJStructProfilePart Then
                  
        Dim oStiffenedPlatePart As Object
        Dim oStiffenedPlatePartEdge As Object
        GetStiffenedPlatePartFromConnPlates pDispPartObj, pDispSelObj, pDispConnObj, oStiffenedPlatePart, oStiffenedPlatePartEdge

        If (oStiffenedPlatePart Is Nothing) Or (oStiffenedPlatePartEdge Is Nothing) Then
            'Return the group parameters as the margin parameters for the part
            MarginType = GroupMarginType
            
            eMarginGeomChange = eGroupMarginGeomChange
            
            eMarginDirection = PerpendicularToEdge
            
            dStretchOffset = dGroupStretchOffset
            
            MarginVal1 = dGroupMarginVal
            
            MarginVal2 = 0#
        Else
            'Get the rule based values for the stiffened plate
             Dim lPlateMarginType As Long, lPlateGeomChange As Long, lPlateDirection As Long
             Dim dPlateStretchOffset As Double, dPlateMarginVal1 As Double, dPlateMarginVal2 As Double
             
             GetConnPlateMarginParameters oStiffenedPlatePart, oStiffenedPlatePartEdge, GroupMarginType, eGroupMarginGeomChange, eGroupMarginDirection, dGroupStretchOffset, dGroupMarginVal, _
                                 lPlateMarginType, lPlateGeomChange, lPlateDirection, dPlateStretchOffset, dPlateMarginVal1, dPlateMarginVal2
                                 
             GetConnProfileParametersFromConnPlate pDispPartObj, oStiffenedPlatePartEdge, lPlateMarginType, lPlateGeomChange, lPlateDirection, dPlateStretchOffset, dPlateMarginVal1, dPlateMarginVal2, _
                                          MarginType, eMarginGeomChange, eMarginDirection, dStretchOffset, MarginVal1, MarginVal2
             
        End If
        
        Set oStiffenedPlatePart = Nothing
        Set oStiffenedPlatePartEdge = Nothing

    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
Private Function IJDMarginByConnRule_GetConnectedObjectsForConnPart(ByVal pDispObj As Object, _
                                                                    ByVal pDispSelObj As Object, ByVal pDispConnObj As Object) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
                                                                    
    On Error GoTo ErrorHandler
    Const METHOD = "IJDMarginByConnRule_GetConnectedObjectsForConnPart"
    
    Set IJDMarginByConnRule_GetConnectedObjectsForConnPart = New Collection
    
    Dim iCnt As Integer
                                                                    
    If TypeOf pDispObj Is IJPlatePart Then
    
        'Return the stiffeners that are present in the input connected objects collection
        Dim oConnObjEle As IJElements
        Set oConnObjEle = pDispConnObj
        
        Dim oConnStiffeners As IJElements
        Set oConnStiffeners = GetConnectedStiffeners(pDispObj)
                    
        For iCnt = 1 To oConnStiffeners.Count
            If oConnObjEle.Contains(oConnStiffeners.Item(iCnt)) Then
                IJDMarginByConnRule_GetConnectedObjectsForConnPart.Add oConnStiffeners.Item(iCnt)
            End If
        Next
    
    ElseIf TypeOf pDispObj Is IJStructProfilePart Then
        'return Nothing
    End If
                                                                    
     
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
                     
End Function
Private Sub IJDMarginByConnRule_GetDependentObjectParametersFromConnPart(ByVal pDispPartObj As GSCADMfgRulesDefinitions.IJDObject, _
                                                                         ByVal pDispEdgeObj As GSCADMfgRulesDefinitions.IJDObject, _
                                                                         ByVal pDispParentEdgeObj As GSCADMfgRulesDefinitions.IJDObject, _
                                                                         ParentMarginType As Long, _
                                                                         eParentMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                                                         eParentMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                                                         dParentStretchOffset As Double, dParentMarginVal1 As Double, dParentMarginVal2 As Double, _
                                                                         MarginType As Long, _
                                                                         eMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                                                         eMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                                                         dStretchOffset As Double, MarginVal1 As Double, MarginVal2 As Double)
    On Error GoTo ErrorHandler
    Const METHOD = "IJDMarginByConnRule_GetDependentObjectParametersFromConnPart"
    
   GetConnProfileParametersFromConnPlate pDispPartObj, pDispParentEdgeObj, ParentMarginType, eParentMarginGeomChange, eParentMarginDirection, dParentStretchOffset, dParentMarginVal1, dParentMarginVal2, _
                                          MarginType, eMarginGeomChange, eMarginDirection, dStretchOffset, MarginVal1, MarginVal2
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function IJDMarginByConnRule_GetMarginByConnTypes(ByVal pDispPart As Object, ByVal MarginMode As GSCADMfgRulesDefinitions.EnumMarginMode) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
      On Error GoTo ErrorHandler
      Const METHOD = "IJDMarginByConnRule_GetMarginByConnTypes"
          
      Dim lTempColl As JCmnShp_CollectionAlias
      Set lTempColl = New Collection

      lTempColl.Add (1) 'Heating Margin
      lTempColl.Add (2) 'Bending Margin
      lTempColl.Add (4) 'Intercoastal Margin
      lTempColl.Add (8) 'Coaming Margin
      lTempColl.Add (3) 'EndFace Margin
          lTempColl.Add (9) 'PanelEnd Margin
      
      lTempColl.Add (100) 'Non Trimmed
      lTempColl.Add (101) 'Trimmed
    
      Set IJDMarginByConnRule_GetMarginByConnTypes = lTempColl
      
      Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function IJDMarginByConnRule_GetMarginByConnAllowedDirections(ByVal eMarginCmdProcess As GSCADMfgRulesDefinitions.MarginCmdProcess) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias

    On Error GoTo ErrorHandler
    Const METHOD = "IJDMarginByConnRule_GetMarginByConnAllowedDirections"
    
    Dim lTempColl As JCmnShp_CollectionAlias
    Set lTempColl = New Collection
        
    lTempColl.Add (1) 'PerpendicularToEdge
    lTempColl.Add (2) 'GlobalDirection

    Set IJDMarginByConnRule_GetMarginByConnAllowedDirections = lTempColl
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function EdgeToFacePort(ByVal oElem As Object) As Object
    On Error GoTo ErrorHandler
    Const METHOD = "EdgeToFacePort"
    
    Dim pMfgHelper As IJMfgMGHelper
    Set pMfgHelper = New MfgMGHelper
    Dim pEdgePort As IJPort
    Dim pFacePort As IJPort
    
    Set pEdgePort = oElem
    Dim oStructPort As IJStructPort
    Set oStructPort = pEdgePort
    
    If (oStructPort.ContextID = CTX_BASE_EXTERNAL_LEDGE) Or _
        (oStructPort.ContextID = CTX_3DSPLIT) Or _
        (oStructPort.ContextID = CInt(CTX_KNUCKLE) + CInt(CTX_BASE_EXTERNAL_LEDGE)) Or _
        (oStructPort.ContextID = CInt(CTX_3DCUTOUT) + CInt(CTX_BASE_EXTERNAL_LEDGE)) Or _
        (oStructPort.ContextID = CInt(CTX_CHAMFER) + CInt(CTX_BASE_EXTERNAL_LEDGE)) Then

        If pEdgePort.Type = PortEdge And _
            (oStructPort.ContextID <> CInt(CTX_CHAMFER) + CInt(CTX_BASE_EXTERNAL_LEDGE)) Then
                
                pMfgHelper.SolidEdgeToLateralFacePort pEdgePort, pFacePort
                Set EdgeToFacePort = pFacePort
                
        ElseIf pEdgePort.Type = PortEdge And _
            (oStructPort.ContextID = CInt(CTX_CHAMFER) + CInt(CTX_BASE_EXTERNAL_LEDGE)) Then

                Dim oEntityHelper As New MfgEntityHelper
                Set EdgeToFacePort = oEntityHelper.GetConnectedLateralFaceOfChampher(oElem)
                
        ElseIf pEdgePort.Type = PortFace Then
            Set EdgeToFacePort = pEdgePort
        End If
    Else
        Set EdgeToFacePort = pEdgePort
    End If

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetEdgeFromFacePort(ByVal oFacePort As IJPort) As Object
    On Error GoTo ErrorHandler
    Const METHOD = "GetEdgeFromFacePort"
    
    Dim oEntityHelper As New MfgEntityHelper
    Dim oTempPort     As IJPort
             
    On Error Resume Next
    Set oTempPort = oEntityHelper.GetEdgePortGivenFacePort(oFacePort, CTX_BASE) 'gives latest edge geometry
    On Error GoTo ErrorHandler
     
    If oTempPort Is Nothing Then
        Dim oChamferPort As IJPort
        Set oChamferPort = oEntityHelper.GetChamferEdgeFromFace(oFacePort)
        Set GetEdgeFromFacePort = oChamferPort
    Else
        Set GetEdgeFromFacePort = oTempPort
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
'Gives the offset direction depending the  enumOffsetMethod.
Private Function GetGroupOffsetDirection(pDispAssy As Object, pDispConnAssy As Object) As EnumMarginOffsetDirection
    On Error GoTo ErrorHandler
    Const METHOD = "GetGroupOffsetDirection"
    
    GetGroupOffsetDirection = MarginGlobalX
    
    Dim oMfgUtilAssyMargin As IJMfgUtilAssyMargin
    Set oMfgUtilAssyMargin = New MfgUtilAssyMargin

    'Ports Normal (Old logic): Gets the common face ports of the parts having connections between the selected and the
                'connected assemblies. Returns the first face port's normal as the group offset direction vector.
    'Parts Range: Gets the parts having connections between the selected and the connected assemblies. Calculates the
                'range boxes of these parts of selected and connected assemblies. If the intersection of the range
                'boxes results into a plane, the normal of the plane is returned as the group offset director vector.
                'If the intersection does not result into a plane, no offset direction vector is returned.
    'Assembly Range: Gets the assembly range. Intersection of selected and connected assemblies range boxes is obtained
                'and normal to the intersection plane returned is retuned as offset direction.
    'Parts COG: Gets the parts having connections between the selected and connected assemblies. Calculates the
                'Center of Gravity (CoG) of these parts of selected and connected assemblies.
                'Returns the vector joining both the CoGs as the offset direction vector.
    'Assembly COG: Gets the Center of Gravity (CoG) of selected and connected assemblies.
                'Returns the vector joining both the CoGs as the offset direction vector.
    
    On Error Resume Next

    Dim oOffsetDirection As IJDVector
    Set oOffsetDirection = oMfgUtilAssyMargin.GetOffsetDirection(pDispAssy, pDispConnAssy, PortNormal)
    
    'Based on the number of connected ports offset direction method(mentoned below) can be chosen.
'    Dim oConnectedPorts As IJElements
'    oMfgUtilAssyMargin.GetConnectedChildren pDispAssy, pDispConnAssy, oConnectedPorts
'
'    If (oConnectedPorts.Count > 5) Then
'     Set oOffsetDirection = oMfgUtilAssyMargin.GetOffsetDirection(pDispAssy, pDispConnAssy, PartsRange)
'    End If

    On Error GoTo ErrorHandler
    
    If Not oOffsetDirection Is Nothing Then
        GetGroupOffsetDirection = GetParallelAxisDirection(oOffsetDirection)
    Else
        GetGroupOffsetDirection = MarginPerpendicular
    End If

    Set oMfgUtilAssyMargin = Nothing
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetMarginValueBasedOnOffsetDirection(dGroupMarginVal As Double, ByVal eMarginOffsetDirection As EnumMarginOffsetDirection, ByVal pDispEdgeObj As Object) As Double
    On Error GoTo ErrorHandler
    Const METHOD = "GetMarginValueBasedOnOffsetDirection"
    
    If eMarginOffsetDirection = MarginPerpendicular Then 'apply the group value to the individual part
        GetMarginValueBasedOnOffsetDirection = dGroupMarginVal
        Exit Function
    End If
    
    Dim oEdgePort As IJPort
    Set oEdgePort = pDispEdgeObj

    Dim oFacePort As IJPort
    Set oFacePort = EdgeToFacePort(pDispEdgeObj)
    
    Dim oConnectable As Object
    Set oConnectable = oEdgePort.Connectable
    
    'For profiles, the landing curve vector and offset direction vector would be used to
    'compute the margin value. For plates, the edge normal and offset direction vector
    'would be used to compute the margin value.
    
    'get the offset direction vector
    Dim oOffsetDirVec As IJDVector
    Set oOffsetDirVec = GetOffsetDirectionVector(eMarginOffsetDirection)
    oOffsetDirVec.length = 1
    
    If TypeOf oConnectable Is IJStructProfilePart Then 'Need to compute value for the profile
        
        Dim oSDProfilePart As New StructDetailObjects.ProfilePart
        Set oSDProfilePart.object = oConnectable 'profile part
        
        Dim oLandingCurve  As IJWireBody
        Dim oThicknessDir  As IJDVector
        Dim bThicknessCentered As Boolean
        
        On Error Resume Next
        oSDProfilePart.LandingCurve oLandingCurve, oThicknessDir, bThicknessCentered
        On Error GoTo ErrorHandler
        
        If oLandingCurve Is Nothing Then Exit Function
        
        Dim oStartPos As IJDPosition
        Dim oEndPos   As IJDPosition
        
        oLandingCurve.GetEndPoints oStartPos, oEndPos
        
        Dim oProfileDir As IJDVector
        Set oProfileDir = New DVector
        oProfileDir.Set oEndPos.x - oStartPos.x, oEndPos.y - oStartPos.y, oEndPos.z - oStartPos.z
        oProfileDir.length = 1
        
        If Abs(oOffsetDirVec.Dot(oProfileDir)) <> 0 Then
            GetMarginValueBasedOnOffsetDirection = dGroupMarginVal / Abs(oOffsetDirVec.Dot(oProfileDir))
        Else
            GetMarginValueBasedOnOffsetDirection = 0#
        End If
                
        Set oStartPos = Nothing
        Set oEndPos = Nothing
        Set oSDProfilePart = Nothing
        Set oLandingCurve = Nothing
        Set oThicknessDir = Nothing
        Set oProfileDir = Nothing
        
    Else 'Need to compute value for the plate
        
        Dim oPartInfo As IJDPartInfo
        Set oPartInfo = New PartInfo
        
        Dim oPortNormalVec As IJDVector
        Dim bApproximationUsed As Boolean
        Set oPortNormalVec = oPartInfo.GetPortNormal(oFacePort, bApproximationUsed)
        oPortNormalVec.length = 1
        
        If Abs(oOffsetDirVec.Dot(oPortNormalVec)) <> 0 Then
            GetMarginValueBasedOnOffsetDirection = dGroupMarginVal / Abs(oOffsetDirVec.Dot(oPortNormalVec))
        Else
            GetMarginValueBasedOnOffsetDirection = 0#
        End If
        
        Set oPortNormalVec = Nothing
        Set oPartInfo = Nothing
    End If
    
    Set oOffsetDirVec = Nothing
    Set oEdgePort = Nothing
    Set oConnectable = Nothing
    Set oFacePort = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetOffsetDirectionVector(ByVal eMarginOffsetDirection As EnumMarginOffsetDirection) As IJDVector
    On Error GoTo ErrorHandler
    Const METHOD = "GetOffsetDirectionVector"
    
    Set GetOffsetDirectionVector = New DVector
    
    If eMarginOffsetDirection = MarginGlobalX Then
        GetOffsetDirectionVector.Set 1, 0, 0
    ElseIf eMarginOffsetDirection = MarginGlobalY Then
        GetOffsetDirectionVector.Set 0, 1, 0
    ElseIf eMarginOffsetDirection = MarginGlobalZ Then
        GetOffsetDirectionVector.Set 0, 0, 1
    End If
        
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetParallelAxisDirection(ByVal oNormal As IJDVector) As EnumMarginOffsetDirection
     Const METHOD = "GetParallelAxis"
     On Error GoTo ErrorHandler
     
     GetParallelAxisDirection = MarginGlobalX
    
     Dim oXVector As IJDVector
     Dim oYVector As IJDVector
     Dim oZVector As IJDVector
     
     Set oXVector = New DVector
     Set oYVector = New DVector
     Set oZVector = New DVector
     
     oXVector.Set 1, 0, 0
     oYVector.Set 0, 1, 0
     oZVector.Set 0, 0, 1
     
     'normalize the vectors
     oXVector.length = 1#
     oYVector.length = 1#
     oZVector.length = 1#
     
     Dim dX As Double
     Dim dY As Double
     Dim dZ As Double
     
     dX = Abs(oXVector.Dot(oNormal))
     dY = Abs(oYVector.Dot(oNormal))
     dZ = Abs(oZVector.Dot(oNormal))
        
     If (dX > dY) And (dX > dZ) Then
         GetParallelAxisDirection = MarginGlobalX
     ElseIf (dY > dX) And (dY > dZ) Then
         GetParallelAxisDirection = MarginGlobalY
     ElseIf (dZ > dX) And (dZ > dY) Then
         GetParallelAxisDirection = MarginGlobalZ
     End If

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub GetPlateMarginParameters(ByVal pDispPartObj As Object, _
                                     ByVal pDispEdgeObj As Object, _
                                     ByVal GroupMarginType As Long, _
                                     ByVal eGroupMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                     ByVal eGroupMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                     ByVal dGroupStretchOffset As Double, ByVal dGroupMarginVal As Double, _
                                     ByVal eMarginOffsetDirection As GSCADMfgRulesDefinitions.EnumMarginOffsetDirection, _
                                     MarginType As Long, _
                                     eMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                     eMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                     dStretchOffset As Double, MarginVal1 As Double, Optional MarginVal2 As Double)
                                     
    Const METHOD = "GetPlateMarginParameters"
    On Error GoTo ErrorHandler
    Dim bIsExistsInCache As Boolean
    bIsExistsInCache = False
    
    CheckAndGetThePlateParametersFromCacheIfExists pDispPartObj, pDispEdgeObj, bIsExistsInCache, MarginType, _
                                                    eMarginGeomChange, eMarginDirection, dStretchOffset, MarginVal1, MarginVal2

    If bIsExistsInCache = False Then
        'Margin Type would be same as the Group Margin Type
        MarginType = GroupMarginType
        
        'Offset direction for Plate Margins is computed based on the 'Group' Offset direction.
        If eMarginOffsetDirection = MarginPerpendicular Then
            eMarginDirection = PerpendicularToEdge
        Else
            eMarginDirection = GlobalDirection
        End If
                   
        MarginVal1 = GetMarginValueBasedOnOffsetDirection(dGroupMarginVal, eMarginOffsetDirection, pDispEdgeObj)
        
        MarginVal2 = 0#
              
        'Margin Geometry Change would be same as the Group Margin Geometry Change
        eMarginGeomChange = eGroupMarginGeomChange
        
        'Margin StretchOffset would be same as the Group Margin StretchOffset
        If eMarginGeomChange = ByStretch Then
            dStretchOffset = ComputeStretchOffsetForPlate(pDispPartObj, pDispEdgeObj, dGroupStretchOffset)
        Else
            dStretchOffset = 0#
        End If
        AddStiffenedPlatePartToCache pDispPartObj, pDispEdgeObj, MarginType, eMarginGeomChange, _
                                  eMarginDirection, dStretchOffset, MarginVal1, MarginVal2
    End If
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub GetProfileParametersFromPlate(pDispProfile As Object, pDispPlateEdge As Object, lPlateMarginType As Long, lPlateGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                          lPlateDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                          dPlateStretchOffset As Double, dPlateMarginVal1 As Double, dPlateMarginVal2 As Double, _
                                          MarginType As Long, eMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, eMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                          dStretchOffset As Double, MarginVal1 As Double, MarginVal2 As Double)
     Const METHOD = "GetProfileParametersFromPlate"
     On Error GoTo ErrorHandler
    
      'Margin Type, Margin Geometry Change and Margin StretchOffset of profile would be same as stiffened plate
      MarginType = lPlateMarginType
                 
      'Margin Direction is Global As Margin value is computed along Global direction
      eMarginDirection = GlobalDirection
      
      'Margin Geometry change would be same as that of the plate
      eMarginGeomChange = lPlateGeomChange
      
      Dim oMfgGeomHelper As IJMfgGeomHelper
      Set oMfgGeomHelper = New MfgGeomHelper
        
      Dim oMfgMGHelper As MfgMGHelper
      Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
      
      Dim oMfgRuleHelper As New MfgRuleHelpers.Helper
            
      Dim oSDProfilePart As New StructDetailObjects.ProfilePart
      Set oSDProfilePart.object = pDispProfile 'profile part
    
      Dim oLandingCurve  As IJWireBody
      Dim oThicknessDir  As IJDVector
      Dim bThicknessCentered As Boolean
    
      On Error Resume Next
      oSDProfilePart.LandingCurve oLandingCurve, oThicknessDir, bThicknessCentered
      On Error GoTo ErrorHandler
    
      Dim oStartPos As IJDPosition
      Dim oEndPos   As IJDPosition
    
      oLandingCurve.GetEndPoints oStartPos, oEndPos
    
      Dim oProfileDir As IJDVector
      Set oProfileDir = New DVector
      oProfileDir.Set oEndPos.x - oStartPos.x, oEndPos.y - oStartPos.y, oEndPos.z - oStartPos.z
      oProfileDir.length = 1
      
      Dim oPlatePort As IJPort
      Set oPlatePort = EdgeToFacePort(pDispPlateEdge)
        
      'Get the plate port's normal
      Dim oPartInfo As IJDPartInfo
      Set oPartInfo = New PartInfo
        
      Dim oPlatePortNormalVec As IJDVector
      Dim bApproximationUsed As Boolean
      Set oPlatePortNormalVec = oPartInfo.GetPortNormal(oPlatePort, bApproximationUsed)
      oPlatePortNormalVec.length = 1
        
      If eMarginGeomChange = ByStretch Then
        
        'Computation of Margin Value:
        '1.Get the profile start and end points
        '2.Get the profile end point which is close to the plate edge
        '3.Offset the closer end point in the direction of the plate margin
        '4.The distance from the newly obtained point(offseted point) and the profile end point gives the new profile length
        '5.Difference between the old and new profile lengths gives the margin value for the profile
                
        'Get the end point of the profile closest to the plate edge
        Dim oModelBody As IJDModelBody
        Set oModelBody = oPlatePort.Geometry
        
        Dim oPos As IJDPosition
        Dim oClosestPos As IJDPosition
        Dim dStartPosDist As Double, dEndPosDist As Double
        
        oModelBody.GetMinimumDistanceFromPosition oStartPos, oClosestPos, dStartPosDist
        oModelBody.GetMinimumDistanceFromPosition oEndPos, oClosestPos, dEndPosDist
        
        If dStartPosDist < dEndPosDist Then
            Set oClosestPos = oStartPos
            Set oPos = oEndPos
        Else
            Set oClosestPos = oEndPos
            Set oPos = oStartPos
        End If
        
        '3.Offset the closer end point in the direction of the plate margin
        oPlatePortNormalVec.length = dPlateMarginVal1
        
        Dim oNewPos As IJDPosition
        Set oNewPos = oClosestPos.Offset(oPlatePortNormalVec)
         
        '4.The distance from the newly obtained point(offseted point) and the profile end point gives the new profile length
        '5.Difference between the old and new profile lengths gives the margin value for the profile
        Dim dOldProfileLen As Double, dNewProfileLen As Double
        dOldProfileLen = oStartPos.DistPt(oEndPos)
        dNewProfileLen = oNewPos.DistPt(oPos)
       
        MarginVal1 = Abs(dNewProfileLen - dOldProfileLen)
        MarginVal2 = 0#
        
        'Computation of profile stretchoffset (relative to plate stretch offset)
        '1. Get the port adjacent to the plate margin port
        '2. Get the point along the adjacent port at a girth distance of dPlateMarginVal1
        '3. Construct plane at that point with plate margin edge normal as the pane normal
        '4. Intersect the plane with profile landing curve
        '5. Distance between this intersection point and profile closest point is the stretch offset
        '6. If there is no intersection with the plane(shorter profiles), propose profile length as stretch offset
        
        '1. Get the port adjacent to the plate margin port
        Dim oAdjPortWB As IJWireBody
        Dim oIntPos    As IJDPosition
        Dim oOtherPos  As IJDPosition
        GetAdjacentPortAndIntersection oPlatePort, oAdjPortWB, oIntPos, oOtherPos
        
        If dPlateStretchOffset <> 0 Then
         '2. Get the point along the adjacent port at a girth distance of dPlateMarginVal1
         Dim oPlanePos As IJDPosition
         Set oPlanePos = oMfgRuleHelper.GetPointAlongCurveAtDistance(oAdjPortWB, oIntPos, dPlateStretchOffset / 1000, oOtherPos)
        
         '3. Construct plane at that point with plate margin edge normal as the pane normal
          oPlatePortNormalVec.length = 1
          Dim oRefPlane As IJPlane
          oMfgGeomHelper.MakeTransientPlane oPlanePos.x, oPlanePos.y, oPlanePos.z, oPlatePortNormalVec.x, oPlatePortNormalVec.y, oPlatePortNormalVec.z, oRefPlane
         
         '4. Intersect the plane with profile landing curve
          Dim oLandingCurveGeom As IJCurve
          oMfgMGHelper.WireBodyToComplexString oLandingCurve, oLandingCurveGeom
         
          Dim oPlaneIntPos As IJDPosition
          Set oPlaneIntPos = oMfgGeomHelper.IntersectCurveWithPlane(oLandingCurveGeom, oRefPlane)
          
          Dim oSGOWireBodyUtilities As SGOWireBodyUtilities
          Set oSGOWireBodyUtilities = New SGOWireBodyUtilities
          
          If Not oPlaneIntPos Is Nothing Then
             '5. Girth Distance between this intersection point and profile closest point is the stretch offset
             oSGOWireBodyUtilities.GetDistanceAlongCurve oLandingCurve, oClosestPos, oPlaneIntPos, dStretchOffset
             
          Else
             '6. If there is no intersection with the plane(shorter profiles), propose profile length as stretch offset
             oSGOWireBodyUtilities.GetDistanceAlongCurve oLandingCurve, oStartPos, oEndPos, dStretchOffset
          End If
          
          dStretchOffset = dStretchOffset * 1000 'return in mm
        Else
            dStretchOffset = dPlateStretchOffset '0mm
        End If
        
     Else 'As Margin
            
          If Abs(oPlatePortNormalVec.Dot(oProfileDir)) <> 0 Then
            MarginVal1 = dPlateMarginVal1 / Abs(oPlatePortNormalVec.Dot(oProfileDir))
          Else
            MarginVal1 = 0#
          End If
          MarginVal2 = 0#
          
          dStretchOffset = 0#
          
     End If
      
     Set oMfgGeomHelper = Nothing
     Set oMfgMGHelper = Nothing
     Set oMfgRuleHelper = Nothing
     Set oStartPos = Nothing
     Set oEndPos = Nothing
     Set oSDProfilePart = Nothing
     Set oLandingCurve = Nothing
     Set oThicknessDir = Nothing
     Set oProfileDir = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub GetConnPlateMarginParameters(ByVal pDispPartObj As Object, _
                                         ByVal pDispEdgeObj As Object, _
                                         ByVal GroupMarginType As Long, _
                                         ByVal eGroupMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                         ByVal eGroupMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                         ByVal dGroupStretchOffset As Double, ByVal dGroupMarginVal As Double, _
                                         MarginType As Long, _
                                         eMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                         eMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                         dStretchOffset As Double, MarginVal1 As Double, Optional MarginVal2 As Double)
    Const METHOD = "GetConnPlateMarginParameters"
    On Error GoTo ErrorHandler
         
    'Margin Type would be same as group margin type
    MarginType = GroupMarginType
        
    'Margin Direction would be "Perpendicular To Edge"
    eMarginDirection = PerpendicularToEdge
    
    'Margin value would be same as the group margin value
    MarginVal1 = dGroupMarginVal
    
    MarginVal2 = 0#
    
    'Margin Geometry Change would be same as Group Geometry Change
    eMarginGeomChange = eGroupMarginGeomChange
        
    'compute stretchoffset
    If eMarginGeomChange = ByStretch Then
        dStretchOffset = ComputeStretchOffsetForPlate(pDispPartObj, pDispEdgeObj, dGroupStretchOffset)
    Else
        dStretchOffset = 0#
    End If
                     
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub GetConnProfileParametersFromConnPlate(pDispProfile As Object, pDispPlateEdge As Object, lPlateMarginType As Long, lPlateGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, _
                                          lPlateDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                          dPlateStretchOffset As Double, dPlateMarginVal1 As Double, dPlateMarginVal2 As Double, _
                                          MarginType As Long, eMarginGeomChange As GSCADMfgRulesDefinitions.EnumMarginGeometryChange, eMarginDirection As GSCADMfgRulesDefinitions.EnumMarginDirection, _
                                          dStretchOffset As Double, MarginVal1 As Double, MarginVal2 As Double)
      Const METHOD = "GetConnProfileParametersFromConnPlate"
      On Error GoTo ErrorHandler
    
      'Margin Type, Margin Geometry Change and Margin StretchOffset of profile would be same as stiffened plate
      MarginType = lPlateMarginType
                 
      'Margin Direction is Global As Margin value is computed along Global direction
      eMarginDirection = GlobalDirection
      
      'Margin Geometry change would be same as that of the plate
      eMarginGeomChange = lPlateGeomChange
      
      Dim oMfgGeomHelper As IJMfgGeomHelper
      Set oMfgGeomHelper = New MfgGeomHelper
        
      Dim oMfgMGHelper As MfgMGHelper
      Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
      
      Dim oMfgRuleHelper As New MfgRuleHelpers.Helper
      
      Dim oSDProfilePart As New StructDetailObjects.ProfilePart
      Set oSDProfilePart.object = pDispProfile 'profile part
    
      Dim oLandingCurve  As IJWireBody
      Dim oThicknessDir  As IJDVector
      Dim bThicknessCentered As Boolean
    
      On Error Resume Next
      oSDProfilePart.LandingCurve oLandingCurve, oThicknessDir, bThicknessCentered
      On Error GoTo ErrorHandler
    
      Dim oStartPos As IJDPosition
      Dim oEndPos   As IJDPosition
    
      oLandingCurve.GetEndPoints oStartPos, oEndPos
    
      Dim oProfileDir As IJDVector
      Set oProfileDir = New DVector
      oProfileDir.Set oEndPos.x - oStartPos.x, oEndPos.y - oStartPos.y, oEndPos.z - oStartPos.z
      oProfileDir.length = 1
      
      Dim oPlatePort As IJPort
      Set oPlatePort = EdgeToFacePort(pDispPlateEdge)
        
      'Get the plate port's normal
      Dim oPartInfo As IJDPartInfo
      Set oPartInfo = New PartInfo
        
      Dim oPlatePortNormalVec As IJDVector
      Dim bApproximationUsed As Boolean
      Set oPlatePortNormalVec = oPartInfo.GetPortNormal(oPlatePort, bApproximationUsed)
      oPlatePortNormalVec.length = 1
        
      If (eMarginGeomChange = ByStretch) Then
                        
        'Computation of Margin Value:
        '1.Get the profile start and end points
        '2.Get the profile end point which is close to the plate edge
        '3.Offset the closer end point in the direction of the plate margin
        '4.The distance from the newly obtained point(offseted point) and the profile end point gives the new profile length
        '5.Difference between the old and new profile lengths gives the margin value for the profile
                        
        '2.Get the profile end point which is close to the plate edge
        Dim oModelBody As IJDModelBody
        Set oModelBody = oPlatePort.Geometry
        
        Dim oPos As IJDPosition
        Dim oClosestPos As IJDPosition
        Dim dStartPosDist As Double, dEndPosDist As Double
        
        oModelBody.GetMinimumDistanceFromPosition oStartPos, oClosestPos, dStartPosDist
        oModelBody.GetMinimumDistanceFromPosition oEndPos, oClosestPos, dEndPosDist
        
        If dStartPosDist < dEndPosDist Then
            Set oClosestPos = oStartPos
            Set oPos = oEndPos
        Else
            Set oClosestPos = oEndPos
            Set oPos = oStartPos
        End If
        
        '3.Offset the closer end point in the direction of the plate margin
        oPlatePortNormalVec.length = dPlateMarginVal1
        
        Dim oNewPos As IJDPosition
        Set oNewPos = oClosestPos.Offset(oPlatePortNormalVec)
                 
        '4.The distance from the newly obtained point(offseted point) and the profile end point gives the new profile length
        '5.Difference between the old and new profile lengths gives the margin value for the profile
        Dim dOldProfileLen As Double, dNewProfileLen As Double
        dOldProfileLen = oStartPos.DistPt(oEndPos)
        dNewProfileLen = oNewPos.DistPt(oPos)
       
        MarginVal1 = Abs(dNewProfileLen - dOldProfileLen)
        MarginVal2 = 0#
        dStretchOffset = 0#

     Else 'As Margin
     
          MarginVal1 = dPlateMarginVal1 / Abs(oPlatePortNormalVec.Dot(oProfileDir))
          MarginVal2 = 0#
          
          dStretchOffset = 0#
          
     End If
     
     Set oMfgGeomHelper = Nothing
     Set oMfgMGHelper = Nothing
     Set oMfgRuleHelper = Nothing
     Set oStartPos = Nothing
     Set oEndPos = Nothing
     Set oSDProfilePart = Nothing
     Set oLandingCurve = Nothing
     Set oThicknessDir = Nothing
     Set oProfileDir = Nothing
     
     Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub GetStiffenedPlatePart(ByVal pDispPartObj As IJStructProfilePart, ByVal pDispAssyObj As Object, ByVal pDispConnAssyObj As Object, oStiffenedPlatePart As Object, oStiffenedPlatePartEdge As Object)
    Const METHOD = "GetStiffenedPlatePart"
    On Error GoTo ErrorHandler
           
    'Get the stiffened plate collection for the input profile
    Dim oProfileWrapper  As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = pDispPartObj

    'Get all the stiffened plates.
    On Error Resume Next
    Dim oPlatePartCollection As IJElements
    Set oPlatePartCollection = oProfileWrapper.GetStiffenedPlates
    On Error GoTo ErrorHandler
    
    If Not oPlatePartCollection Is Nothing Then
       If oPlatePartCollection.Count > 0 Then
            
            'Get the connections between the assemblies
    
            Dim oConnections As IJElements
            Set oConnections = GetConnectionsBetweenAssemblies(pDispAssyObj, pDispConnAssyObj) 'Gives the face ports
                        
            'Create a collection of plates and their corresponding edges at the connections between the assemblies
            Dim oPlatePartsAtAssyConn As IJElements
            Set oPlatePartsAtAssyConn = New JObjectCollection
            
            Dim oPlatePartEdgesAtAssyConn As IJElements
            Set oPlatePartEdgesAtAssyConn = New JObjectCollection
            
            Dim iCount As Integer
            If Not oConnections Is Nothing Then
                For iCount = 1 To oConnections.Count
                    Dim oConn As IJPort
                    Set oConn = oConnections.Item(iCount)
                    
                    If oPlatePartCollection.Contains(oConn.Connectable) Then
                        oPlatePartsAtAssyConn.Add oConn.Connectable
                        oPlatePartEdgesAtAssyConn.Add oConn
                        Exit For
                    End If
                    
                    Set oConn = Nothing
                Next iCount
            End If
            
            'return the first plate part
            If oPlatePartsAtAssyConn.Count > 0 Then
                Set oStiffenedPlatePart = oPlatePartsAtAssyConn.Item(1)
                Set oStiffenedPlatePartEdge = GetEdgeFromFacePort(oPlatePartEdgesAtAssyConn.Item(1))
            Else
                Set oStiffenedPlatePart = Nothing
                Set oConnections = Nothing
                Set oPlatePartCollection = Nothing
                Exit Sub
            End If
            
            Set oConnections = Nothing
       Else
            Exit Sub
       End If
    Else
        Exit Sub
    End If
    
    Set oPlatePartCollection = Nothing
        
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Sub
Private Sub GetStiffenedPlatePartFromConnPlates(ByVal pDispPartObj As IJStructProfilePart, ByVal pDispSelObj As Object, ByVal pDispConnObjs As Object, oStiffenedPlatePart As Object, oStiffenedPlatePartEdge As Object)
    Const METHOD = "GetStiffenedPlatePartFromConnPlates"
    On Error GoTo ErrorHandler
                 
    'Get the stiffened plate collection for the input profile
    Dim oProfileWrapper  As MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = pDispPartObj

    'Get all the stiffened plates.
    On Error Resume Next
    Dim oPlatePartCollection As IJElements
    Set oPlatePartCollection = oProfileWrapper.GetStiffenedPlates
    On Error GoTo ErrorHandler
    
    If Not oPlatePartCollection Is Nothing Then
      If oPlatePartCollection.Count > 0 Then
        
            Dim oConnObjEle As IJElements
            Set oConnObjEle = pDispConnObjs
                 
            'Return the plate present in the connected object collection as the stiffened plate
            Dim iCnt As Integer
            For iCnt = 1 To oPlatePartCollection.Count
               If oConnObjEle.Contains(oPlatePartCollection.Item(iCnt)) Then
                   Set oStiffenedPlatePart = oPlatePartCollection.Item(iCnt)
                   Exit For
               End If
            Next iCnt
               
            Set oConnObjEle = Nothing
              
            'Get the edge port of the stiffened plate that is connected to the selected object
            If Not oStiffenedPlatePart Is Nothing Then
                If TypeOf pDispSelObj Is IJPlateSystem Then
                    Set oStiffenedPlatePartEdge = GetConnPlateEdgeFromPlateSystem(pDispSelObj, oStiffenedPlatePart)
                Else
                    Dim oConnectable As IJConnectable
                    Set oConnectable = pDispSelObj
                    
                    Dim bIsConnected As Boolean
                    Dim oConnections As IJElements
                    oConnectable.isConnectedTo oStiffenedPlatePart, bIsConnected, oConnections
                    
                    For iCnt = 1 To oConnections.Count
                        If TypeOf oConnections.Item(iCnt) Is IJStructPhysicalConnection Then
                            Dim oAppConn As IJAppConnection
                            Set oAppConn = oConnections.Item(iCnt)
                            
                            Dim oPorts As IJElements
                            oAppConn.enumPorts oPorts
                            
                            Dim oThisPort As IJPort
                            Set oThisPort = oPorts.Item(1)
                            
                            If oThisPort.Connectable Is oStiffenedPlatePart Then
                                Set oStiffenedPlatePartEdge = GetEdgeFromFacePort(oThisPort)
                            Else
                                Set oStiffenedPlatePartEdge = GetEdgeFromFacePort(oPorts.Item(2))
                            End If
                            
                            Set oAppConn = Nothing
                            Set oPorts = Nothing
                            Set oThisPort = Nothing
                            
                            Exit For
                        End If
                    Next iCnt
                    
                End If
            End If
        
      End If
    End If
    
    Set oPlatePartCollection = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function ComputeStretchOffsetForPlate(ByVal pDispPartObj As Object, _
                                         ByVal pDispEdgeObj As Object, _
                                         dGroupStretchOffset As Double) As Double
    Const METHOD = "ComputeStretchOffsetForPlate"
    On Error GoTo ErrorHandler
    
    ComputeStretchOffsetForPlate = dGroupStretchOffset
    
    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = pDispPartObj
    
    Dim oPartPort As IJStructPort, oSystemPort As IJStructPort
    Dim oConObjsCol As Collection
    Set oConObjsCol = GetConnectedObjectsData(pDispPartObj)
    
    Dim oMfgMGHelper As MfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    Dim oMfgGeomHelper As MfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
            
    Dim oMfgRuleHelper As New MfgRuleHelpers.Helper

    Dim oSGOWireBodyUtilities As SGOWireBodyUtilities
    Set oSGOWireBodyUtilities = New SGOWireBodyUtilities
    
    'Check if there is a port bounded to hull.
    If oPlateWrapper.IsConnectedToHull(oPartPort, oSystemPort, oConObjsCol) = True Then
        'Get the features on the port bounded to the hull
            'a. get face port from input edge port
            'b. get the primary port of the face port
            'c. get the features using primary port as the input
        Dim oStructPort As IJStructPortEx
        Set oStructPort = oPartPort
     
        Dim oSelFacePort As IJPort
        Set oSelFacePort = EdgeToFacePort(pDispEdgeObj)
        
        Dim oSelStructPort As IJStructPortEx
        Set oSelStructPort = oSelFacePort
        
        'Check if the port connected to hull is not the input port.
        If Not oSelStructPort.PrimaryPort Is oStructPort.PrimaryPort Then
            Dim oSelPortGeometry As IJWireBody, oAdjPortGeometry As IJWireBody
            Set oSelPortGeometry = GetBaseEdgeFromFace(oSelFacePort)
            Set oAdjPortGeometry = GetBaseEdgeFromFace(oPartPort)
            
            Dim oSelPortCurve As IJCurve
            Dim oAdjPortCurve As IJCurve
            oMfgMGHelper.WireBodyToComplexString oSelPortGeometry, oSelPortCurve
            oMfgMGHelper.WireBodyToComplexString oAdjPortGeometry, oAdjPortCurve
                     
            Dim oIntersectionObject As Object
            Dim oIntersectionPos As IJDPosition
            Dim oIntersectionElems As IJElements
            
            Set oIntersectionObject = oMfgGeomHelper.IntersectCurveWithCurve(oSelPortCurve, oAdjPortCurve)
            
            If Not oIntersectionObject Is Nothing Then
                
                If TypeOf oIntersectionObject Is IJDPosition Then
                    Set oIntersectionPos = oIntersectionObject
                Else
                    Set oIntersectionElems = oIntersectionObject
                    Set oIntersectionPos = oIntersectionElems.Item(1)
                    Set oIntersectionElems = Nothing
                End If
                
                'Check if the port connected to the hull has any features
                Dim oEntityHelper As New MfgEntityHelper
                Dim oFeaturesColl As IJElements
                Set oFeaturesColl = oEntityHelper.GetAdjacentFeaturesFromPort(oStructPort.PrimaryPort)
                        
                'Get the least distance from the feature(s) and port corner point
                If Not oFeaturesColl Is Nothing Then
                    If oFeaturesColl.Count > 0 Then
                        Dim oSDPartSupport As GSCADSDPartSupport.IJPartSupport
                        Set oSDPartSupport = New GSCADSDPartSupport.PartSupport
                        Set oSDPartSupport.Part = pDispPartObj
                        
                        Dim dMinDist As Double
                        dMinDist = 1E+100 'Initialize to some large number
                                
                        'Get the feature geometry
                        Dim iCnt As Integer
                        For iCnt = 1 To oFeaturesColl.Count
                            
                            Dim bMeasureDistToFeature As Boolean
                            
                            If Margin_Feature_Check = 0 Then
                                bMeasureDistToFeature = True
                            ElseIf Margin_Feature_Check = 1 Then
                                If TypeOf oFeaturesColl.Item(iCnt) Is IJStructFeature Then
                                    Dim oStructfeature  As IJStructFeature
                                    Set oStructfeature = oFeaturesColl.Item(iCnt)
                                    
                                    If oStructfeature.get_structfeaturetype = SF_Slot Then
                                        bMeasureDistToFeature = True
                                    Else
                                        bMeasureDistToFeature = False
                                    End If
                                End If
                            End If
                            
                            'Check if only cutout features are intended.
                            If (bMeasureDistToFeature = True) Then
                                Dim oContourColl As Collection
                                Dim oPortMonikerColl As Collection
                                
                                oSDPartSupport.GetFeatureInfo oFeaturesColl.Item(iCnt), SideA, oContourColl, oPortMonikerColl
                                
                                Dim oFeatureWireBody As IJWireBody
                                
                                If oContourColl.Count > 1 Then
                                    Dim oMergedWireBody As IJWireBody
                                    oMfgMGHelper.MergeWireBodyCollection oContourColl, oMergedWireBody
                                    
                                    Set oFeatureWireBody = oMergedWireBody
                                Else
                                    Set oFeatureWireBody = oContourColl.Item(1)
                                End If
                                
                                'Get the end points of the feature wirebody
                                Dim oFeatureWBStartPos As IJDPosition, oFeatureWBEndPos As IJDPosition
                                oFeatureWireBody.GetEndPoints oFeatureWBStartPos, oFeatureWBEndPos
                                                            
                                'Get the girth distance along the edge from the Intersection point to the feature end points
                                Dim dStartPtDist As Double, dEndPtDist As Double, dDist As Double
                                Dim oProjIntersectionPos As IJDPosition
                                Set oProjIntersectionPos = oMfgRuleHelper.ProjectPointOnWire(oAdjPortGeometry, oIntersectionPos) 'To ensure point is on the wire
                                
                                oSGOWireBodyUtilities.GetDistanceAlongCurve oAdjPortGeometry, oProjIntersectionPos, oFeatureWBStartPos, dStartPtDist
                                oSGOWireBodyUtilities.GetDistanceAlongCurve oAdjPortGeometry, oProjIntersectionPos, oFeatureWBEndPos, dEndPtDist
                                              
                                If dStartPtDist < dEndPtDist Then
                                    dDist = dStartPtDist
                                Else
                                    dDist = dEndPtDist
                                End If
                                                               
                                If dDist < dMinDist Then
                                    dMinDist = dDist
                                End If
                               
                                Set oContourColl = Nothing
                                Set oPortMonikerColl = Nothing
                                                                
                             End If
                        Next iCnt
                        
                        'return the least distance
                        If dMinDist <> 0 Then
                            ComputeStretchOffsetForPlate = dMinDist * 1000 'return in mm
                        End If
                                                
                        Set oSDPartSupport = Nothing
                            
                    End If 'oFeatureColl.Count
                End If 'oFeatureColl is Nothing
                              
                Set oFeaturesColl = Nothing
                Set oEntityHelper = Nothing
                              
            End If 'oIntersectionObject is nothing
            
            Set oSelPortGeometry = Nothing
            Set oAdjPortGeometry = Nothing
            Set oSelPortCurve = Nothing
            Set oAdjPortCurve = Nothing
            
        End If
           
        Set oStructPort = Nothing
        Set oSelFacePort = Nothing
        Set oSelStructPort = Nothing
    
    End If
    
    Set oPlateWrapper = Nothing
    Set oMfgMGHelper = Nothing
    Set oMfgGeomHelper = Nothing
    Set oMfgRuleHelper = Nothing
    Set oSGOWireBodyUtilities = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetConnectedObjectsData(ByVal oPlate As Object) As Collection
    Const METHOD = "GetConnectedObjectsData"
    On Error GoTo ErrorHandler
        
    Set GetConnectedObjectsData = New Collection
        
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New PlatePartSupport
    Set oPartSupport.Part = oPlate
    
    Dim ThisPortColl As Collection
    Dim OtherPortColl As Collection
    Dim ConnectedObjColl As Collection
    Dim ConnectionsColl As Collection
    oPartSupport.GetConnectedObjects ConnectionPhysical, _
                                         ConnectedObjColl, ConnectionsColl, _
                                         ThisPortColl, OtherPortColl
                                         
    Dim iCnt As Integer
    For iCnt = 1 To ConnectedObjColl.Count
    
        Dim oConnectionData As ConnectionData
        
        Set oConnectionData.AppConnection = ConnectionsColl.Item(iCnt)
        Set oConnectionData.ConnectingPort = ThisPortColl.Item(iCnt)
        Set oConnectionData.ToConnectedPort = OtherPortColl.Item(iCnt)
        Set oConnectionData.ToConnectable = ConnectedObjColl.Item(iCnt)
        
        GetConnectedObjectsData.Add oConnectionData
        
    Next iCnt

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetConnPlateEdgeFromPlateSystem(ByVal pDispSelObj As IJPlateSystem, ByVal oConnPlatePart As IJPlatePart) As Object
    Const METHOD = "GetConnPlateEdgeFromPlateSystem"
    On Error GoTo ErrorHandler
    
    'Get the plate parts from the plate system
    Dim oElems As IJElements
    Set oElems = GetPlatesFromPlateSystem(pDispSelObj)
    
    'Identify which plate part among the above collection which is connected to input plate part
    Dim iCnt As Integer
    For iCnt = 1 To oElems.Count
        
        Dim oConnectable As IJConnectable
        Set oConnectable = oElems.Item(iCnt)
        
        Dim oConnectionColl As IJElements
        Dim bIsConnected As Boolean
        oConnectable.isConnectedTo oConnPlatePart, bIsConnected, oConnectionColl
        
        If bIsConnected = True Then
            Dim jCnt As Integer
            For jCnt = 1 To oConnectionColl.Count
                If TypeOf oConnectionColl.Item(jCnt) Is IJStructPhysicalConnection Then
                
                    Dim oAppConn As IJAppConnection
                    Set oAppConn = oConnectionColl.Item(jCnt)
                    
                    Dim oPorts As IJElements
                    oAppConn.enumPorts oPorts
                    
                    Dim oThisPort As IJPort
                    Set oThisPort = oPorts.Item(1)
                    
                    If oThisPort.Connectable Is oConnPlatePart Then
                        Set GetConnPlateEdgeFromPlateSystem = GetEdgeFromFacePort(oThisPort)
                    Else
                        Set GetConnPlateEdgeFromPlateSystem = GetEdgeFromFacePort(oPorts.Item(2))
                    End If
                    
                    Set oAppConn = Nothing
                    Set oPorts = Nothing
                    Set oThisPort = Nothing
                
                    Exit Function
                End If
            Next jCnt
        End If
        
        Set oConnectable = Nothing
        
    Next iCnt
    
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function GetPlatesFromPlateSystem(oSys As Object) As IJElements
    Const METHOD = "GetPlatesFromPlateSystem"
    On Error GoTo ErrorHandler

    Dim oParts              As IJElements
        
    Dim oStructDetailHelper As StructDetailHelper
    Dim oEnumPartsUnk       As IEnumUnknown
    Dim oCollectionOfParts  As Collection
    Dim ConvertUtils        As CCollectionConversions
    
    If TypeOf oSys Is IJPlateSystem Then
        ' Retrieve all Parts dereived from the Base Plate System

        Set oStructDetailHelper = New StructDetailHelper
        oStructDetailHelper.GetPartsDerivedFromSystem oSys, oEnumPartsUnk, True

        If Not oEnumPartsUnk Is Nothing Then
            'Convert the IEnumUnknown to a VB collection that we can use in VB
            Set ConvertUtils = New CCollectionConversions
            ConvertUtils.CreateVBCollectionFromIEnumUnknown oEnumPartsUnk, oCollectionOfParts
            Set oParts = ConvertUtils.CreateIJElementsCollectionFromVBCollection(oCollectionOfParts)
        End If
        
    End If
    
    Set GetPlatesFromPlateSystem = oParts
    
CleanUp:
    
    Set oParts = Nothing
    Set oStructDetailHelper = Nothing
    Set oEnumPartsUnk = Nothing
    Set oCollectionOfParts = Nothing
    Set ConvertUtils = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function


Private Function GetConnectedStiffeners(oPart As Object) As IJElements
    Const METHOD = "GetConnectedStiffeners"
    On Error GoTo ErrorHandler
    
    Set GetConnectedStiffeners = New JObjectCollection
        
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New PlatePartSupport
    Set oPartSupport.Part = oPart
    
    Dim ThisPortColl As Collection
    Dim OtherPortColl As Collection
    Dim ConnectedObjColl As Collection
    Dim ConnectionsColl As Collection
    oPartSupport.GetConnectedObjects ConnectionPhysical, _
                                         ConnectedObjColl, ConnectionsColl, _
                                         ThisPortColl, OtherPortColl
                                         
    Dim iCnt As Integer
    For iCnt = 1 To ConnectedObjColl.Count
        If TypeOf ConnectedObjColl.Item(iCnt) Is IJProfilePart Then
            GetConnectedStiffeners.Add ConnectedObjColl.Item(iCnt)
        End If
    Next iCnt
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Private Function GetBaseEdgeFromFace(oFacePort As IJPort) As Object
    Const METHOD = "GetBaseEdgeFromFace"
    On Error GoTo ErrorHandler
        
    Const strGeneratedPlatePartAEProgId = "CreatePlatePart.GeneratePlatePart_AE.1"
    Const strStandAlonePlatePartAEProgId = "CreatePlatePart.CreatePlatePart_AE.1"
    
    Dim oStructPort As IJStructPortEx
    Set oStructPort = oFacePort
    
    Dim oRelatedFacePort As IJPort
    Set oRelatedFacePort = oStructPort.RelatedPort(strStandAlonePlatePartAEProgId, False)
        
    If (oFacePort Is oRelatedFacePort) Then
        Set oRelatedFacePort = oStructPort.RelatedPort(strGeneratedPlatePartAEProgId, False)
    End If
    
    Dim oEnumPorts                  As IEnumUnknown
    Dim oCollectionOfPorts          As Collection
    Dim oSGOModelBodyUtilities      As SGOModelBodyUtilities
    Dim oConvertUtils               As CONVERTUTILITIESLib.CCollectionConversions
    
    '- instantiate the services
    Set oSGOModelBodyUtilities = New SGOModelBodyUtilities
    Set oConvertUtils = New CONVERTUTILITIESLib.CCollectionConversions
        
    '- get the enum of base edges
    oSGOModelBodyUtilities.GetEdgesByContextOrOpt oRelatedFacePort.Geometry, _
                                                CTX_BASE, _
                                                oEnumPorts, _
                                                True, _
                                                True, False
    
    '- convert the enum to VB collection
    If Not oEnumPorts Is Nothing Then
        oConvertUtils.CreateVBCollectionFromIEnumUnknown oEnumPorts, oCollectionOfPorts
        Set GetBaseEdgeFromFace = oCollectionOfPorts.Item(1)
    End If
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub GetAdjacentPortAndIntersection(ByVal oPlateFacePort As IJPort, ByRef oAdjPortWB As IJWireBody, ByRef oIntPos As IJDPosition, ByRef oOtherPos As IJDPosition)
    Const METHOD = "GetAdjacentPortAndIntersection"
    On Error GoTo ErrorHandler
        
    'Get all the lateral edge ports of the plate
    Dim oConnectable As IJConnectable
    Set oConnectable = oPlateFacePort.Connectable

    Dim oEnumPortsColl As IJElements
    Set oEnumPortsColl = New JObjectCollection
    
    oConnectable.enumPorts oEnumPortsColl, PortFace
    
    'Get the intersection of each port with the input port
    Dim oMfgMGHelper As MfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    Dim oMfgGeomHelper As MfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    Dim oMfgRuleHelper As New MfgRuleHelpers.Helper
    
    Dim oSelPortGeometry As IJWireBody
    Set oSelPortGeometry = GetBaseEdgeFromFace(oPlateFacePort)
    
    Dim iCnt As Integer
    For iCnt = 1 To oEnumPortsColl.Count
         
        Dim oOtherPort         As IJPort
        Dim oOtherPortGeometry As IJWireBody
        
        Set oOtherPort = oEnumPortsColl.Item(iCnt)
        
        Dim oStructPort As IJStructPort
        Set oStructPort = oOtherPort
        
        If oStructPort.ContextID = CTX_LATERAL_LFACE Then
            Set oOtherPortGeometry = GetBaseEdgeFromFace(oOtherPort)
           
            If Not oOtherPort Is oPlateFacePort Then
                
                Dim oSelPortCurve As IJCurve
                Dim oAdjPortCurve As IJCurve
                oMfgMGHelper.WireBodyToComplexString oSelPortGeometry, oSelPortCurve
                oMfgMGHelper.WireBodyToComplexString oOtherPortGeometry, oAdjPortCurve
                         
                Dim oIntersectionObject As Object
                Dim oIntersectionPos As IJDPosition
                Dim oIntersectionElems As IJElements
                
                On Error Resume Next
                Set oIntersectionObject = oMfgGeomHelper.IntersectCurveWithCurve(oSelPortCurve, oAdjPortCurve)
                On Error GoTo ErrorHandler
                    
                If Not oIntersectionObject Is Nothing Then
                    'If there is an intersection, return that edge port as the adjacent port
                    'Also, return the intersection point
                    If TypeOf oIntersectionObject Is IJDPosition Then
                        Set oIntersectionPos = oIntersectionObject
                    Else
                        Set oIntersectionElems = oIntersectionObject
                        Set oIntersectionPos = oIntersectionElems.Item(1)
                        Set oIntersectionElems = Nothing
                    End If
                    
                    Dim oProjIntersectionPos As IJDPosition
                    Set oProjIntersectionPos = oMfgRuleHelper.ProjectPointOnWire(oOtherPortGeometry, oIntersectionPos) 'To ensure point is on the wire
                    
                    Dim oWB As IJWireBody
                    Set oWB = oOtherPortGeometry
                    Dim oStartPos As IJDPosition, oEndPos As IJDPosition
                    oWB.GetEndPoints oStartPos, oEndPos
                    
                    Dim dDist1 As Double, dDist2 As Double
                    dDist1 = oStartPos.DistPt(oProjIntersectionPos)
                    dDist2 = oEndPos.DistPt(oProjIntersectionPos)
                    
                    If dDist1 < dDist2 Then
                        Set oOtherPos = oEndPos
                        Set oIntPos = oStartPos
                    Else
                        Set oOtherPos = oStartPos
                        Set oIntPos = oEndPos
                    End If
                    
                    Set oAdjPortWB = oOtherPortGeometry
                    
                    Exit For
                End If 'oIntersectionObject Is Nothing
                
          End If 'oOtherPort Is oPlateFacePort
          
        End If 'oStructPort.ContextID = CTX_LATERAL_LFACE
    Next iCnt
        
CleanUp:
    
    Set oConnectable = Nothing
    Set oMfgMGHelper = Nothing
    Set oMfgGeomHelper = Nothing
    Set oMfgRuleHelper = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

Private Function CheckIfPlateIsCanPlate(pCanPlate As Object) As Boolean
    Dim oPlatePart As IJPlatePart
    Set oPlatePart = pCanPlate
    
    Dim oSDPlatePart As New StructDetailObjects.PlatePart
    Set oSDPlatePart.object = oPlatePart
    
    If (oSDPlatePart.PlateType = TransverseTube) Or (oSDPlatePart.PlateType = LongitudinalTube) Or (oSDPlatePart.PlateType = VerticalTube) Or (oSDPlatePart.PlateType = TubePlate) Then
        CheckIfPlateIsCanPlate = True
    Else
        CheckIfPlateIsCanPlate = False
    End If
    
End Function

