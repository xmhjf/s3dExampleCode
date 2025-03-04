VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BalanceRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2011, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgPinJigOrientation
' Module: BalanceRule
'
' Author: Himaja Siddabatula
'
' Comments: This rule drives the PinJig orientation. It proposes the pinjig balance.
'*******************************************************************************
Option Explicit
Private Const MODULE = "StrMfgPinJigOrientation.BalanceRule"
Implements IJDPinJigBalanceRule

'The balance options can be customised by adding through bulkload and to the below enum.
Private Enum PinjigNature

    MostPlanarNatural = 0
    TrueNatural = 1
    AssemblyOrientation = 2
    UserDefined = 3
    ParallelAxis = 4
    
    AverageOfCornersPlane = 51
    AverageOfLowerAftCornersPlane = 52
    AverageOfUpperAftCornersPlane = 53
    AverageOfLowerForeCornersPlane = 54
    AverageOfUpperForeCornersPlane = 55
    
    LowerAftCorners = 101
    UpperAftCorners = 102
    LowerForeCorners = 103
    UpperForeCorners = 104
    
    AverageOfFourPointsPlane = 151
    ByThreepointsPlane = 152
    
    AverageOfReferenceLineAndTwoCornersPlane = 201
    ReferenceLineAndOneCornerPlane = 202
    
'    Balance_CustomType1 = 10001 'Fur further customization
'    Balance_CustomType2 = 10002
    
End Enum

' This method proposes the allowable pinjig balances for the selected supported plates.
Private Function IJDPinJigBalanceRule_GetAllowablePinjigBalanceTypes(ByVal pPinJig As Object) As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias
Const METHOD = "IJDPinJigBalanceRule_GetAllowablePinjigBalanceTypes"
    On Error GoTo ErrorHandler

    Dim oBalanceTypes As Collection
    Set oBalanceTypes = New Collection

    oBalanceTypes.Add MostPlanarNatural
    oBalanceTypes.Add TrueNatural
    oBalanceTypes.Add AssemblyOrientation
    oBalanceTypes.Add UserDefined
    oBalanceTypes.Add ParallelAxis

    oBalanceTypes.Add AverageOfCornersPlane
    oBalanceTypes.Add AverageOfLowerAftCornersPlane
    oBalanceTypes.Add AverageOfUpperAftCornersPlane
    oBalanceTypes.Add AverageOfLowerForeCornersPlane
    oBalanceTypes.Add AverageOfUpperForeCornersPlane
    
    oBalanceTypes.Add LowerAftCorners
    oBalanceTypes.Add UpperAftCorners
    oBalanceTypes.Add LowerForeCorners
    oBalanceTypes.Add UpperForeCorners
    
    oBalanceTypes.Add AverageOfFourPointsPlane
    oBalanceTypes.Add ByThreepointsPlane
    
    oBalanceTypes.Add AverageOfReferenceLineAndTwoCornersPlane
    oBalanceTypes.Add ReferenceLineAndOneCornerPlane

    Set IJDPinJigBalanceRule_GetAllowablePinjigBalanceTypes = oBalanceTypes

 Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5030, , "RULES")
End Function

' This method proposes the default pinjig balance for the selected supported plates.
Private Function IJDPinJigBalanceRule_GetDefaultBalanceType(ByVal pPinJig As Object) As Long
Const METHOD = "IJDPinJigBalanceRule_GetDefaultBalanceType"
    On Error GoTo ErrorHandler

     IJDPinJigBalanceRule_GetDefaultBalanceType = TrueNatural

Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5030, , "RULES")
End Function

' This method proposes the eligible inputs/criteria for user selection to define the pinjig balance.
Private Function IJDPinJigBalanceRule_GetEligiblePinJigBalanceInputs(ByVal pPinJig As Object, ByVal lActivePinJigBalance As Long) As GSCADPinJigEntity.JCmnShp_CollectionAlias
    Const METHOD = "IJDPinJigOrientationRule_GetEligiblePinJigBalanceInputs"
    On Error GoTo ErrorHandler
    
    If lActivePinJigBalance = MostPlanarNatural Or _
        lActivePinJigBalance = TrueNatural Or _
        lActivePinJigBalance = AssemblyOrientation Or _
        lActivePinJigBalance = UserDefined Or _
        lActivePinJigBalance = ParallelAxis Then
        
        'NOT APPLICABLE
        Exit Function
    End If
    
    Dim oPinJig                             As IJPinJig
    Dim oBalanceInputColl                   As Collection
    Dim oMfgUtilSurf                        As IJMfgGeomUtil
    Dim oOuterCornerPtsColl                 As Collection
    Dim oOuterCornerPts                     As IJElements
    Dim oContours                           As IJElements
    Dim oExtremePlateCorners                As IJElements
    Dim oExtremeEligiblePlateCorners        As IJElements
    Dim oEntitiesPoints                     As Collection
    Dim oEntities                           As Collection
    Dim iCnt                                As Integer

    Set oBalanceInputColl = New Collection
    Set oMfgUtilSurf = New MfgUtilSurface
    Set oOuterCornerPtsColl = New Collection
    
    Set oPinJig = pPinJig
    
    'Get All the outer corners of the supported plates.
    oMfgUtilSurf.GetOuterCornersFromPlateCollection oPinJig.SupportedPlates, oOuterCornerPts, oContours
    
    For iCnt = 1 To oOuterCornerPts.Count
        oOuterCornerPtsColl.Add oOuterCornerPts.Item(iCnt)
    Next
    
    Select Case lActivePinJigBalance
                    
        Case AverageOfCornersPlane
        
             'All Outer Plates Corners
             Set oBalanceInputColl = oOuterCornerPtsColl
            
        Case AverageOfLowerAftCornersPlane, AverageOfUpperAftCornersPlane, AverageOfLowerForeCornersPlane, AverageOfUpperForeCornersPlane

            'Get Eligible outer corners of the supported plates.
            ' oExtremePlateCorners will be ordered
            oMfgUtilSurf.GetExtremeCornersFromPlateCollection oPinJig.SupportedPlates, lActivePinJigBalance, oExtremePlateCorners, oExtremeEligiblePlateCorners
            
            For iCnt = 1 To oExtremeEligiblePlateCorners.Count
                oBalanceInputColl.Add oExtremeEligiblePlateCorners.Item(iCnt)
            Next
            
        Case LowerAftCorners, UpperAftCorners, LowerForeCorners, UpperForeCorners
        
            'All Outer Plates Corners
            Set oBalanceInputColl = oOuterCornerPtsColl
            
        Case AverageOfFourPointsPlane, ByThreepointsPlane
        
            'All Outer Plates Corners
            Set oBalanceInputColl = oOuterCornerPtsColl
          
            'End points of Seams, Knuckles, Reference curves, Center Line, User Defined Marking Lines
            GetEntitiesAndEndPoints oPinJig, oEntities, oEntitiesPoints
            
            For iCnt = 1 To oEntitiesPoints.Count
                oBalanceInputColl.Add oEntitiesPoints.Item(iCnt)
            Next
        
        Case AverageOfReferenceLineAndTwoCornersPlane, ReferenceLineAndOneCornerPlane
           
            'All Outer Plates Corners
            Set oBalanceInputColl = oOuterCornerPtsColl
            
            'End points of Seams, Knuckles, Reference curves, Center Line, User Defined Marking Lines
            GetEntitiesAndEndPoints oPinJig, oEntities, oEntitiesPoints
            
            For iCnt = 1 To oEntitiesPoints.Count
                oBalanceInputColl.Add oEntitiesPoints.Item(iCnt)
            Next
            
            'Seams, Knuckles, Reference curves, Center Line, User Defined Marking Lines
            For iCnt = 1 To oEntities.Count
                oBalanceInputColl.Add oEntities.Item(iCnt)
            Next
             
    End Select
    
    Set IJDPinJigBalanceRule_GetEligiblePinJigBalanceInputs = oBalanceInputColl
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5030, , "RULES")
End Function

' This method proposes the base plane for given balance type and user inputs.
Private Function IJDPinJigBalanceRule_GetPinJigBasePlane(ByVal pPinJig As Object, ByVal lPinJigBalanceType As Long, ByVal lMinHeight As Double, ByVal pBalanceInputColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias, Optional pPointsColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As Object
    Const METHOD = "IJDPinJigOrientationRule_GetPinJigBasePlane"
    On Error GoTo ErrorHandler
    
    Dim oPinJig                         As IJPinJig
    Dim oBasePlane                      As IJPlane
    Dim iCnt                            As Integer
    Dim oPlaneUtil                      As New PinJigFloorUtils
    Dim oPointsColl                     As Collection
    Dim oExtremePlateCorners            As IJElements
    Dim oExtremeEligiblePlateCorners    As IJElements
    Dim oMfgUtilSurf                    As IJMfgGeomUtil
    Dim dPercentagePlanarity            As Double
    
    Set oMfgUtilSurf = New MfgUtilSurface
    Set oPointsColl = New Collection
    
    Set oPinJig = pPinJig
    
    Select Case lPinJigBalanceType
    
        Case MostPlanarNatural
            
            If pBalanceInputColl Is Nothing Then
                dPercentagePlanarity = 60
            Else
                dPercentagePlanarity = pBalanceInputColl.Item(1)
            End If
            
            If dPercentagePlanarity < 0.001 Then
                dPercentagePlanarity = 60
            End If
            
            Set oBasePlane = oPlaneUtil.MostPlanarNaturalBalance(oPinJig.SupportedPlates, dPercentagePlanarity, lMinHeight)
        
        Case TrueNatural
     
            Set oBasePlane = oPlaneUtil.TrueNaturalBalanceMethod(oPinJig.SupportedPlates, lMinHeight)

        Case AssemblyOrientation

            Dim oAssm           As IJAssembly
            Dim oPlatePart      As IJPlatePart
            Dim oMfgChild       As IJMfgChild

            Set oPlatePart = oPinJig.SupportedPlates.Item(1)

            Set oMfgChild = oPlatePart
            Set oAssm = oMfgChild.GetParent

            Set oBasePlane = oPlaneUtil.AssemblyBalanceMethod(oAssm, oPinJig.SupportedPlates, lMinHeight)
            
        Case ParallelAxis
        
             Set oBasePlane = oPlaneUtil.ParallelAxisBalance(oPinJig.SupportedPlates, lMinHeight)
             
        Case AverageOfCornersPlane
            
            If pBalanceInputColl Is Nothing Then
              
              ' Get the four extreme corners of supported plates
                oMfgUtilSurf.GetExtremeCornersFromPlateCollection oPinJig.SupportedPlates, lPinJigBalanceType, oExtremePlateCorners, oExtremeEligiblePlateCorners

                For iCnt = 1 To oExtremePlateCorners.Count
                    oPointsColl.Add oExtremePlateCorners.Item(iCnt)
                Next
                
            Else
                'User Selection: Picked Four Points
                For iCnt = 1 To pBalanceInputColl.Count
                    If TypeOf pBalanceInputColl.Item(iCnt) Is IJDPosition Then
                        oPointsColl.Add pBalanceInputColl.Item(iCnt)
                    End If
                Next

            End If
            
            If oPointsColl.Count <> 4 Then
                StrMfgLogError Err, MODULE, METHOD, "Failed to get the four corners to define base plane", , , , "RULES"
            End If
             
            Set oBasePlane = oPlaneUtil.PlateCornersBalance(oPinJig.SupportedPlates, oPointsColl, lMinHeight, lPinJigBalanceType)
            
        Case AverageOfLowerAftCornersPlane, AverageOfUpperAftCornersPlane, AverageOfLowerForeCornersPlane, AverageOfUpperForeCornersPlane
        
            If pBalanceInputColl Is Nothing Then
                ' Get the four extreme corners of supported plates based on the balance type
                oMfgUtilSurf.GetExtremeCornersFromPlateCollection oPinJig.SupportedPlates, lPinJigBalanceType, oExtremePlateCorners, oExtremeEligiblePlateCorners

                For iCnt = 1 To oExtremePlateCorners.Count
                    oPointsColl.Add oExtremePlateCorners.Item(iCnt)
                Next
                
            Else
                'User Selection: Picked Four Points
                For iCnt = 1 To pBalanceInputColl.Count
                    If TypeOf pBalanceInputColl.Item(iCnt) Is IJDPosition Then
                        oPointsColl.Add pBalanceInputColl.Item(iCnt)
                    End If
                Next
            End If
            
            If oPointsColl.Count <> 4 Then
                StrMfgLogError Err, MODULE, METHOD, "Failed to get the four corners to define base plane", , , , "RULES"
            End If
            
            Set oBasePlane = oPlaneUtil.PlateCornersBalance(oPinJig.SupportedPlates, oPointsColl, lMinHeight, lPinJigBalanceType)

         Case LowerAftCorners, UpperAftCorners, LowerForeCorners, UpperForeCorners

            If pBalanceInputColl Is Nothing Then
            
                ' Get the four extreme corners of supported plates based on the balance type
                oMfgUtilSurf.GetExtremeCornersFromPlateCollection oPinJig.SupportedPlates, lPinJigBalanceType, oExtremePlateCorners, oExtremeEligiblePlateCorners

                Dim oAftLowerPos    As IJDPosition
                Dim oLowerForePos   As IJDPosition
                Dim oForeUpperPos   As IJDPosition
                Dim oUpperAftPos    As IJDPosition
                
                Set oAftLowerPos = oExtremePlateCorners.Item(1)
                Set oLowerForePos = oExtremePlateCorners.Item(2)
                Set oForeUpperPos = oExtremePlateCorners.Item(3)
                Set oUpperAftPos = oExtremePlateCorners.Item(4)
                
                If lPinJigBalanceType = LowerAftCorners Then
                    oPointsColl.Add oUpperAftPos
                    oPointsColl.Add oAftLowerPos
                    oPointsColl.Add oLowerForePos
                ElseIf lPinJigBalanceType = LowerForeCorners Then
                    oPointsColl.Add oAftLowerPos
                    oPointsColl.Add oLowerForePos
                    oPointsColl.Add oForeUpperPos
                ElseIf lPinJigBalanceType = UpperForeCorners Then
                    oPointsColl.Add oLowerForePos
                    oPointsColl.Add oForeUpperPos
                    oPointsColl.Add oUpperAftPos
                ElseIf lPinJigBalanceType = UpperAftCorners Then
                    oPointsColl.Add oForeUpperPos
                    oPointsColl.Add oUpperAftPos
                    oPointsColl.Add oAftLowerPos
                End If
               
            Else
                For iCnt = 1 To pBalanceInputColl.Count
                    If TypeOf pBalanceInputColl.Item(iCnt) Is IJDPosition Then
                        oPointsColl.Add pBalanceInputColl.Item(iCnt)
                    End If
                Next
                
            End If
            
            If oPointsColl.Count <> 3 Then
                StrMfgLogError Err, MODULE, METHOD, "Failed to get the Three corners to define base plane", , , , "RULES"
            End If

            Set oBasePlane = oPlaneUtil.PlateCornersBalance(oPinJig.SupportedPlates, oPointsColl, lMinHeight, lPinJigBalanceType)
            
       Case AverageOfFourPointsPlane

            If pBalanceInputColl Is Nothing Then
                ' Not Applicable
            Else
                'User Selection: Picked Four Points
                For iCnt = 1 To pBalanceInputColl.Count
                    If TypeOf pBalanceInputColl.Item(iCnt) Is IJDPosition Then
                        oPointsColl.Add pBalanceInputColl.Item(iCnt)
                    End If
                Next
                If oPointsColl.Count <> 4 Then
                    StrMfgLogError Err, MODULE, METHOD, "Failed to get the Four corners to define base plane", , , , "RULES"
                End If
                
                Set oBasePlane = oPlaneUtil.PlateCornersBalance(oPinJig.SupportedPlates, oPointsColl, lMinHeight, lPinJigBalanceType)

            End If
            
        Case ByThreepointsPlane
            
            If pBalanceInputColl Is Nothing Then
                ' Not Applicable
            Else
                'User Selection: Picked Three Points
                For iCnt = 1 To pBalanceInputColl.Count
                    If TypeOf pBalanceInputColl.Item(iCnt) Is IJDPosition Then
                        oPointsColl.Add pBalanceInputColl.Item(iCnt)
                    End If
                Next
                
                If oPointsColl.Count <> 3 Then
                    StrMfgLogError Err, MODULE, METHOD, "Failed to get the Three corners to define base plane", , , , "RULES"
                End If
                
                Set oBasePlane = oPlaneUtil.PlateCornersBalance(oPinJig.SupportedPlates, oPointsColl, lMinHeight, lPinJigBalanceType)

            End If

        Case AverageOfReferenceLineAndTwoCornersPlane, ReferenceLineAndOneCornerPlane
        
            If pBalanceInputColl Is Nothing Then
                ' Not Applicable
            Else
                
                Set oPointsColl = GetPointsBasedOnRefLine(oPinJig, pBalanceInputColl)
                Set oBasePlane = oPlaneUtil.PlateCornersBalance(oPinJig.SupportedPlates, oPointsColl, lMinHeight, lPinJigBalanceType)
               
            End If
             
    End Select
    
    Set IJDPinJigBalanceRule_GetPinJigBasePlane = oBasePlane
    Set pPointsColl = oPointsColl
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5030, , "RULES")
End Function

