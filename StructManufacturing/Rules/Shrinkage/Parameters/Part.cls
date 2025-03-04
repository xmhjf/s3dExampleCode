VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Part"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgShrinkageRules
' Module: PartRule
'
' Description:  Determines the proposed settings for the GrandBlock type assembly
'
' Author:
'
' Comments:
' 2012.JUNE.21    Raman     New desgin of the class
'*******************************************************************************
Option Explicit

Const MODULE = "Part Rule : "

Implements IJDShrinkageRule

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
        
    ComputeProfileShrinkageValues pDispObj, pPlateColl.Item(1), pPlatePrimaryFactorColl.Item(1), pPlateSecondaryFactorColl.Item(1), pPlatePrimaryAxisColl.Item(1), pPlateSecondaryAxisColl.Item(1), PrimaryFactor
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub GetShrinkageValuesForPlate(oPlatePart As IJPlatePart, ByRef ShrinkageType As StrMfgShrinkageType, ByRef PrimaryFactor As Double, ByRef SecondaryFactor As Double, Optional ByRef PrimaryAxis As Object, Optional ByRef SecondaryAxis As Object)
        Const METHOD = "GetShrinkageValuesForPlate"
        On Error GoTo ErrorHandler
        
        Dim oNumberL As Integer
        Dim oNumberT As Integer
        Dim oTempObj As Variant
        
        'Initialize values
        oNumberL = 0
        oNumberT = 0
    
        Dim oSDPlatePart As New StructDetailObjects.PlatePart
        Set oSDPlatePart.object = oPlatePart
                
        Dim oFrameSystem As IHFrameSystem
        
        Set oFrameSystem = GetGlobalCoordinateSystem(oPlatePart)
        
        Select Case oSDPlatePart.PlateType
            Case DeckPlate
                ShrinkageType = GlobalType
                If Not oFrameSystem Is Nothing Then
                    Set PrimaryAxis = oFrameSystem.PrincipalXAxis
                    Set SecondaryAxis = oFrameSystem.PrincipalYAxis
                End If
            
            Case TBulkheadPlate
                ShrinkageType = GlobalType
                If Not oFrameSystem Is Nothing Then
                    Set PrimaryAxis = oFrameSystem.PrincipalYAxis
                    Set SecondaryAxis = oFrameSystem.PrincipalZAxis
                End If
            
            Case LBulkheadPlate
                ShrinkageType = GlobalType
                If Not oFrameSystem Is Nothing Then
                    Set PrimaryAxis = oFrameSystem.PrincipalXAxis
                    Set SecondaryAxis = oFrameSystem.PrincipalZAxis
                End If
    
            Case Hull
                ShrinkageType = ScalingType
    
            Case UntypedPlate
                ShrinkageType = ScalingType
        End Select
        Set oFrameSystem = Nothing
        
        Dim collConnectedObj As Collection
        Set collConnectedObj = oSDPlatePart.ConnectedObjects
    
    '    'Give me the number of physical connections
        
        ''Count the number of attached profiles in a certain direction, if they are not longitudenal then assume that they are transversal
        '
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
                    Else
                        oNumberT = oNumberT + 1
                    End If
                End If
            End If
        Next
        
        If oNumberL >= 3 And oNumberT < 3 Then
            PrimaryFactor = 1
            SecondaryFactor = 0#
        End If
        
        If oNumberL >= 3 And oNumberT > 3 Then
            PrimaryFactor = 1
            SecondaryFactor = 1
        End If
        
        If oNumberL < 3 And oNumberT >= 3 Then
            PrimaryFactor = 0#
            SecondaryFactor = 1
        End If
        
        If oNumberL < 3 And oNumberT < 3 Then
            PrimaryFactor = 0#
            SecondaryFactor = 0#
        End If
        
        Exit Sub
        
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub GetShrinkageValuesForProfile(oProfilePart As IJStructProfilePart, ByRef ShrinkageType As StrMfgShrinkageType, ByRef PrimaryFactor As Double, ByRef SecondaryFactor As Double, Optional ByRef PrimaryAxis As Object, Optional ByRef SecondaryAxis As Object)
    Const METHOD = "GetShrinkageValuesForProfile"
    On Error GoTo ErrorHandler
    
    ShrinkageType = ScalingType
    SecondaryFactor = 0#
    Set PrimaryAxis = Nothing
    Set SecondaryAxis = Nothing
        
    'if part is member, exit with current values set
    If TypeOf oProfilePart Is ISPSMemberPartPrismatic Then
        PrimaryFactor = 1#
        SecondaryFactor = 1#
        
        ' Return member lateral face port as the primary and secondary axis
        Dim oMfgShrinkageHelper As IJMfgShrinkageHelper
        Set oMfgShrinkageHelper = New MfgShrinkageHelper
        
        Set PrimaryAxis = oMfgShrinkageHelper.GetWebLeftFacePortOnMember(oProfilePart)
        Set SecondaryAxis = PrimaryAxis
        
        Exit Sub
    End If
    
    Dim dPlatePrimaryFactor         As Double
    Dim dPlateSecondaryFactor       As Double
    Dim oPlatePrimaryAxis           As Object
    Dim oPlateSecondaryAxis         As Object
    Dim PlateShrinkageType          As StrMfgShrinkageType
    Dim oPlatePart                  As IJPlatePart
    Dim oPlatePartCollection        As IJElements
    Dim nCounter                    As Integer
    Dim oMfgUpdateService           As MfgUpdateService
    Dim oShrinkage                  As IJScalingShr
    Dim oObjColl                    As IJDObjectCollection
    Dim varMoniker                  As Variant
    Dim oProfileWrapper             As MfgRuleHelpers.ProfilePartHlpr
    
    ' return mounting face port as the primary axis
    Dim oPrimaryPortAsFacePort As IJPort
    Dim oStiffenerPart As IJStiffenerPart
    Set oStiffenerPart = oProfilePart
    
    oStiffenerPart.GetMountingFacePort oPrimaryPortAsFacePort, vbNullString
    Set PrimaryAxis = oPrimaryPortAsFacePort
                    
    Set oProfileWrapper = New MfgRuleHelpers.ProfilePartHlpr
    Set oProfileWrapper.object = oProfilePart

    'Get the stiffened plate.
    On Error Resume Next
    Set oPlatePartCollection = oProfileWrapper.GetStiffenedPlates
    On Error GoTo ErrorHandler
    
    If Not oPlatePartCollection Is Nothing Then
       ' Get the first stiffened plate.
       If oPlatePartCollection.Count > 0 Then
            Set oPlatePart = oPlatePartCollection.Item(1)
       Else
            Exit Sub
       End If
    Else
        Exit Sub
    End If
            
    'Get any existing Shrinkage on the stiffened plate
    Set oMfgUpdateService = New MfgUpdateService
    Set oObjColl = oMfgUpdateService.GetShrinkages(oPlatePart, AllShrinkages)
     
    Dim oIJDObject As IJDObject
    Set oIJDObject = oProfilePart
    
    Dim oIJDPOM As IJDPOM
    Set oIJDPOM = oIJDObject.ResourceManager
    
    For Each varMoniker In oObjColl
        Set oShrinkage = oIJDPOM.GetObject(varMoniker)
        Exit For
    Next
    
    If Not oShrinkage Is Nothing Then
        'Check if the shrinkage type is "Global Axis".
        If oShrinkage.ShrinkageType = GlobalType Then
            'Get the primary and secondary axis and factors of the existing shrinkage of stiffened plate.
            PlateShrinkageType = GlobalType
            dPlatePrimaryFactor = (oShrinkage.PrimaryFactor - 1) * 1000
            dPlateSecondaryFactor = (oShrinkage.SecondaryFactor - 1) * 1000
            
            Dim oPart As Object
            oShrinkage.GetInputs oPart, oPlatePrimaryAxis, oPlateSecondaryAxis
        Else
            ' Return if the type is "Scaling by Edge"
            Exit Sub
        End If
    Else
        'If there is no existing shrinkage, get the primary and secondary factors considering the stiffened plate.
        GetShrinkageValuesForPlate oPlatePart, PlateShrinkageType, dPlatePrimaryFactor, dPlateSecondaryFactor, oPlatePrimaryAxis, oPlateSecondaryAxis
        
        ' Return if the type is "Scaling by Edge"
        If PlateShrinkageType = ScalingType Then Exit Sub
    End If
      
    ComputeProfileShrinkageValues oProfilePart, oPlatePart, dPlatePrimaryFactor, dPlateSecondaryFactor, oPlatePrimaryAxis, oPlateSecondaryAxis, PrimaryFactor

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
'This method returns the profile shrinkage values based on its orientation with respect to the input plate's primary
'and secondary direction.
Private Sub ComputeProfileShrinkageValues(ByVal oProfilePart As Object, _
                                          ByVal oPlatePart As Object, _
                                          ByVal dPlatePrimaryFactor As Double, _
                                          ByVal dPlateSecondaryFactor As Double, _
                                          ByVal oPlatePrimaryAxis As Object, _
                                          ByVal oPlateSecondaryAxis As Object, _
                                          PrimaryFactor As Double)

        Const METHOD = "ComputeProfileShrinkageValues"
        On Error GoTo ErrorHandler
                
        If oPlatePrimaryAxis Is Nothing Then
            Exit Sub
        End If
        
        Dim oLandingCurve               As IJWireBody
        Dim oThicknessDir               As IJDVector
        Dim bThicknessCentered          As Boolean
    
        Dim oSDProfilePart As New StructDetailObjects.ProfilePart
        Set oSDProfilePart.object = oProfilePart
        On Error Resume Next
        oSDProfilePart.LandingCurve oLandingCurve, oThicknessDir, bThicknessCentered
        On Error GoTo ErrorHandler
        If oLandingCurve Is Nothing Then Exit Sub
        
        Dim oStartPos                   As IJDPosition
        Dim oEndPos                     As IJDPosition
        oLandingCurve.GetEndPoints oStartPos, oEndPos
        
        Dim oProfileDir                 As IJDVector
        Dim oPrimaryDir                 As IJDVector
        Dim oSecondaryDir               As IJDVector
        
        Set oProfileDir = New DVector
        Set oPrimaryDir = New DVector
        Set oSecondaryDir = New DVector
        
        ' Find direction vector of the profile
        oProfileDir.Set oEndPos.x - oStartPos.x, oEndPos.y - oStartPos.y, oEndPos.z - oStartPos.z
        oProfileDir.length = 1
        
        Dim oPrimaryAxisInfo    As IJRAxisInfo
        Set oPrimaryAxisInfo = oPlatePrimaryAxis
        Set oPrimaryDir = oPrimaryAxisInfo.UnitVector
        oPrimaryDir.length = 1

        'If theta is the angle between profile landing curve vector and plate primary direction vector
        'and Sx is the plate primary factor and Sy is the plate secondary factor, then
        'the factor for profile would be equal to Sx * cosine(theta) * cosine(theta) + Sy * sin(theta) * sin(theta)
        Dim dProduct1 As Double, dProduct2 As Double
        dProduct1 = Abs(oPrimaryDir.Dot(oProfileDir))
        dProduct2 = Sqr(1 - (dProduct1 * dProduct1)) 'finding sin(theta)
        PrimaryFactor = (dPlatePrimaryFactor * dProduct1 * dProduct1) + (dPlateSecondaryFactor * dProduct2 * dProduct2)
      
CleanUp:
        Set oSDProfilePart = Nothing
        Set oLandingCurve = Nothing
        Set oStartPos = Nothing
        Set oEndPos = Nothing
        Set oProfileDir = Nothing
        Set oPrimaryDir = Nothing
        Set oSecondaryDir = Nothing
        Set oPrimaryAxisInfo = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub IJDShrinkageRule_GetShrinkageParameters(ByVal pDispObj As Object, ShrinkageType As StrMfgShrinkageType, PrimaryAxis As Object, PrimaryFactor As Double, SecondaryAxis As Object, SecondaryFactor As Double, Optional TertiaryAxis As Object, Optional TertiaryFactor As Double)
    Const METHOD = "IJDShrinkageRule_GetShrinkageParameters"
    On Error GoTo ErrorHandler
    
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

        GetShrinkageValuesForPlate pDispObj, ShrinkageType, PrimaryFactor, SecondaryFactor, PrimaryAxis, SecondaryAxis


    ElseIf TypeOf pDispObj Is IJStructProfilePart Then

       'Get the profile shrinkage values based on orientation with respect to stiffened plate's primary and secondary direction
        GetShrinkageValuesForProfile pDispObj, ShrinkageType, PrimaryFactor, SecondaryFactor, PrimaryAxis, SecondaryAxis

    Else

        ShrinkageType = ScalingType
        PrimaryFactor = 0#
        SecondaryFactor = 0#
        Set PrimaryAxis = Nothing
        Set SecondaryAxis = Nothing

    End If

    Exit Sub
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 7004, , "RULES")
End Sub




