VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssyPartIncrement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgShrinkageRules
' Module: GrandBlockRule
'
' Description:  Determines the proposed settings for the GrandBlock type assembly
'
' Author:
'
' Comments:
' 2012.JUNE.21    Raman     New desgin of the class
'*******************************************************************************
Option Explicit

Const MODULE = "Assy Part Rule - By Increments: "

Implements IJDShrinkageRule

Private Const IJScalingShr = "{DE77050C-3300-11D5-BA1A-0090276F4279}"
Private Const TOLERANCE = 0.01

Private m_oShrinkageHelper As MfgShrinkageHelper

Private Function IJDShrinkageRule_GetConnectedObjectsForPart(ByVal pDispObj As Object) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
    Const METHOD = "IJDShrinkageRule_GetConnectedObjectsForPart"
    On Error GoTo ErrorHandler
    
    Set IJDShrinkageRule_GetConnectedObjectsForPart = New Collection
    
    If TypeOf pDispObj Is IJPlatePart Then
    
        'Get the stiffener systems connected to the plate part
        Dim oPartSupport As IJPartSupport
        Set oPartSupport = New PlatePartSupport
        Set oPartSupport.Part = pDispObj
        
        Dim ThisPortColl As Collection
        Dim OtherPortColl As Collection
        Dim ConnectedObjColl As Collection
        Dim ConnectionsColl As Collection
        oPartSupport.GetConnectedObjects ConnectionLogical, _
                                             ConnectedObjColl, ConnectionsColl, _
                                             ThisPortColl, OtherPortColl
                                             
        
        'Return the stiffener systems conencted to the input part
        Dim iCount As Integer
        For iCount = 1 To ConnectedObjColl.Count
            If TypeOf ConnectedObjColl.Item(iCount) Is IJStiffenerSystem Then
                IJDShrinkageRule_GetConnectedObjectsForPart.Add ConnectedObjColl.Item(iCount)
            End If
        Next
        
        Set oPartSupport = Nothing
        Set ThisPortColl = Nothing
        Set OtherPortColl = Nothing
        Set ConnectionsColl = Nothing
    
    End If
        
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub IJDShrinkageRule_GetDependentProfileShrinkageParameters(ByVal pDispObj As Object, ByVal pPlateColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal pPlatePrimaryFactorColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal pPlateSecondaryFactorColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal pPlatePrimaryAxisColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, ByVal pPlateSecondaryAxisColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, PrimaryFactor As Double)
    Const METHOD = "IJDShrinkageRule_GetDependentProfileShrinkageParameters"
    On Error GoTo ErrorHandler
        
    'No implementation
    ' Values for profiles are also calculated using Parent Assembly shrinkages
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IJDShrinkageRule_GetShrinkageParameters(ByVal pDispObj As Object, ShrinkageType As StrMfgShrinkageType, PrimaryAxis As Object, PrimaryFactor As Double, SecondaryAxis As Object, SecondaryFactor As Double, Optional TertiaryAxis As Object, Optional TertiaryFactor As Double)
Const METHOD = "IJDShrinkageRule_GetShrinkageParameters"
On Error GoTo ErrorHandler
    
    Set m_oShrinkageHelper = New MfgShrinkageHelper
    
    'Set properties on shrinkage helper object
    m_oShrinkageHelper.ShrinkageRule = 52 ' By Increments (StrMfgShrinkageRules.AssyPartIncrement) - will be useful to dump the rule name and rule progid
    m_oShrinkageHelper.ShrinkageMode = AssemblyPartShrinkage
    m_oShrinkageHelper.ShrinkageParent = pDispObj
    
    'Start the logging
    m_oShrinkageHelper.WriteHeader
    
    'Log the rule log start time
    m_oShrinkageHelper.WriteTimeStamp
    
    'Log rule basic information
    m_oShrinkageHelper.WriteObjectRuleInfo
    m_oShrinkageHelper.WriteBlankLine
    
    If TypeOf pDispObj Is IJPlatePart Or TypeOf pDispObj Is IJStiffenerPart Or TypeOf pDispObj Is ISPSMemberPartPrismatic Then
    
        If TypeOf pDispObj Is IJPlatePart Then
            'Exclude bracket and collar plates
            Dim oPlateUtils As IJPlateAttributes
            Set oPlateUtils = New PlateUtils
    
            Dim oObjUtil As IJDMfgGeomUtilWrapper
            Set oObjUtil = New MfgGeomUtilWrapper
    
            Dim oSys As IJSystem
            Set oSys = oObjUtil.GetRootSystem(pDispObj)
    
            'Check if it is a bracket plate
            If Not oSys Is Nothing Then
                If (oPlateUtils.IsBracketByPlane(oSys) Or oPlateUtils.IsTrippingBracket(oSys)) Then
                    ShrinkageType = ShrinkageUndefined
                    Exit Sub
                End If
            End If
    
            'Check for collar plates
            If (TypeOf pDispObj Is IJSmartPlate) Or (TypeOf pDispObj Is IJCollarPart) Then
                ShrinkageType = ShrinkageUndefined
                Exit Sub
            End If
        End If
        
        Dim oParentAssyColl As IJElements
        Dim oParentShrColl As IJElements

        m_oShrinkageHelper.GetParentAssemblyShrinkages pDispObj, oParentAssyColl, oParentShrColl
    Else
        ShrinkageType = ShrinkageUndefined
        m_oShrinkageHelper.WriteInvalidObject
        m_oShrinkageHelper.WriteExitRule
        GoTo CleanUp
    End If
    
    Dim oFrameSystem As IHFrameSystem
    Set oFrameSystem = GetGlobalCoordinateSystem(pDispObj)
    
    Dim oEntityHelper As MfgEntityHelper
    Set oEntityHelper = New MfgEntityHelper
    
    If TypeOf pDispObj Is IJPlatePart Then
        'get the directions
        Dim oNumberL As Integer
        Dim oNumberT As Integer
        Dim oNumberV As Integer
        Dim oTempObj As Variant

        'Initialize values
        oNumberL = 0
        oNumberT = 0
        oNumberV = 0

        Dim oSDPlatePart As New StructDetailObjects.PlatePart
        Set oSDPlatePart.object = pDispObj
        Dim collConnectedObj As Collection

        Set collConnectedObj = oSDPlatePart.ConnectedObjects

        For Each oTempObj In collConnectedObj
            Dim oConData As ConnectionData
            oConData = oTempObj
            Set oTempObj = oConData.ToConnectable
            Dim oStiffner  As IJStiffener
            
            If Not oTempObj Is Nothing Then
                If TypeOf oTempObj Is IJStiffener Then
                    Set oStiffner = oTempObj
                    If oStiffner.pType = sptLongitudinal Then
                        oNumberL = oNumberL + 1
                    ElseIf oStiffner.pType = sptTransversal Then
                        oNumberT = oNumberT + 1
                    ElseIf oStiffner.pType = sptVertical Then
                        oNumberV = oNumberV + 1
                    End If
                End If
            End If
        Next
        
        ' Write the number of profiles in each direction
        m_oShrinkageHelper.WriteBlankLine
        m_oShrinkageHelper.WriteNumberOfProfiles oNumberL, oNumberT, oNumberV
                
        Select Case oSDPlatePart.PlateType
        Case DeckPlate
            If oNumberL > oNumberT Then
                'deck stiffened along longitudinal
                ShrinkageType = GlobalType

                If Not oFrameSystem Is Nothing Then
                    Set PrimaryAxis = oFrameSystem.PrincipalXAxis
                    Set SecondaryAxis = oFrameSystem.PrincipalYAxis
                End If
            Else
                'deck stiffened along longitudinal
                ShrinkageType = GlobalType

                If Not oFrameSystem Is Nothing Then
                    Set PrimaryAxis = oFrameSystem.PrincipalYAxis
                    Set SecondaryAxis = oFrameSystem.PrincipalXAxis
                End If
            End If
        Case TBulkheadPlate
            If oNumberT > oNumberV Then
                'Transversal BulkHead stiffened along Transeversal
                ShrinkageType = GlobalType

                If Not oFrameSystem Is Nothing Then
                    Set PrimaryAxis = oFrameSystem.PrincipalYAxis
                    Set SecondaryAxis = oFrameSystem.PrincipalZAxis
                End If
            Else
                'Transversal BulkHead stiffened along Vertical
                ShrinkageType = GlobalType

                If Not oFrameSystem Is Nothing Then
                    Set PrimaryAxis = oFrameSystem.PrincipalZAxis
                    Set SecondaryAxis = oFrameSystem.PrincipalYAxis
                End If
            End If
        Case LBulkheadPlate
            If oNumberL > oNumberV Then
                'Longitudinal Bulkhead stiffened along Longitudinal
                ShrinkageType = GlobalType

                If Not oFrameSystem Is Nothing Then
                    Set PrimaryAxis = oFrameSystem.PrincipalXAxis
                    Set SecondaryAxis = oFrameSystem.PrincipalZAxis
                End If
            Else
                'Longitudinal Bulkhead stiffened along Vertical
                ShrinkageType = GlobalType

                If Not oFrameSystem Is Nothing Then
                    Set PrimaryAxis = oFrameSystem.PrincipalZAxis
                    Set SecondaryAxis = oFrameSystem.PrincipalXAxis
                End If
            End If
        End Select
        
    ElseIf TypeOf pDispObj Is IJProfilePart Then
        Dim oSDProfilePart As New StructDetailObjects.ProfilePart
        Set oSDProfilePart.object = pDispObj
            
        Dim oPrimaryPortAsFacePort As IJPort
        Dim oStiffenerPart As IJStiffenerPart
        Set oStiffenerPart = pDispObj
        oStiffenerPart.GetMountingFacePort oPrimaryPortAsFacePort, vbNullString
        ShrinkageType = ScalingType

        Set PrimaryAxis = oPrimaryPortAsFacePort
        Set SecondaryAxis = Nothing
    Else
        ShrinkageType = ShrinkageUndefined 'No shrinkage needed
    End If
    
    'Log the final shrinkage directions
    m_oShrinkageHelper.WritePartAssyBasedShrDirections PrimaryAxis, SecondaryAxis
    
    'Get the vectors along the primary and secondary directions
    Dim oPrimVector As IJDVector
    Set oPrimVector = GetVectorFromAxisOrPort(PrimaryAxis)
    
    Dim oSecVector As IJDVector
    Set oSecVector = GetVectorFromAxisOrPort(SecondaryAxis)
    
    'Log data retlated to parent assembly shrinkages
    m_oShrinkageHelper.WriteDataOfParentShrinkages oParentShrColl
    
    AccumulateShrinkageParameters oParentShrColl, oPrimVector, oSecVector, _
                MULTIPLY_SHRINKAGES, PrimaryFactor, SecondaryFactor
                
    PrimaryFactor = (PrimaryFactor - 1) * 1000 'return in mm
    SecondaryFactor = (SecondaryFactor - 1) * 1000 'return in mm
    
    'Log the final data returned by the rule
    m_oShrinkageHelper.WriteRuleOutputData ShrinkageType, PrimaryAxis, SecondaryAxis, PrimaryFactor, SecondaryFactor, TertiaryFactor
    
CleanUp:
    ' Write the final log time
    m_oShrinkageHelper.WriteBlankLine
    m_oShrinkageHelper.WriteTimeStamp
    m_oShrinkageHelper.WriteHeader
    
    Set m_oShrinkageHelper = Nothing
   
    Exit Sub
ErrorHandler:
    m_oShrinkageHelper.WriteErrorDescription Err.Description
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 7004, , "RULES")
    GoTo CleanUp
End Sub
