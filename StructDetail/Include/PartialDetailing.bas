Attribute VB_Name = "PartialDetailing"
'*******************************************************************
'
'Copyright (C) 2015-2016 Intergraph Corporation. All rights reserved.
'
'File : PartialDetailing.bas
'
'Author : Alligators
'
'Description :
'
'
'History :
'    02/Nov/2015 - GHM- DI-281291: Support Partial Detailing
'*****************************************************************************
Option Explicit

Private Const MODULE = "S:\StructDetail\Data\Include\PartialDetailing.bas"

Public Enum eObjectType
'
    e_AssemblyConnection = 0
    e_WebCut = 1
    e_FlangeCut = 2
    e_Slot = 3
    
    e_CornerFeature = 4
    e_StiffenerEndCutCornerFeature = 5
    e_MemberEndCutCornerFeature = 6
    e_SlotCornerFeature = 7
    e_CollarCornerFeature = 8

    e_EdgeFeature = 9
    e_Chamfer = 10
    e_PhysicalConnection = 11
    e_FreeEdgeTreatment = 12

    e_Bracket = 13
    e_BearingPlate = 14
    e_CornerCollar = 15
    e_Collar = 16
    e_InsertPlate = 17
'
End Enum

'**********************************************************************************************
' Method      : ExcludeObjectBasedOnDetailedState
' Description :
'
'
'
' oInputObject: Input parent custom assembly
' eObjectType: Object type which needs to be considered in case of partial detailing state.
'               eObjectType is new enum with all the object types
'**********************************************************************************************
Public Function ExcludeObjectBasedOnDetailedState(oParentCustomAssembly As Object, ObjectType As eObjectType) As Boolean

 Const sMETHOD = "ExcludeObjectBasedOnDetailedState"
 On Error GoTo ErrHandler

    ExcludeObjectBasedOnDetailedState = False

    If ObjectType <> e_PhysicalConnection And ObjectType <> e_FreeEdgeTreatment Then
        Exit Function
    End If

    If Not TypeOf oParentCustomAssembly Is IJSmartOccurrence Then
        Exit Function
    End If

    Dim oBoundedOrPenetratingPart As Object
    Dim oBoundingOrPenetratedPart As Object

    Dim oBoundedOrPenetratingPort As IJPort
    Dim oBoundingOrPenetratedPort  As IJPort

    'Get Inputs of the given custom assembly parent
    'custom assembly parent is stiffener assembly connection
    If TypeOf oParentCustomAssembly Is IJAssemblyConnection Then

        Dim oSD_AssemblyConnection As StructDetailObjects.AssemblyConn
        Set oSD_AssemblyConnection = New StructDetailObjects.AssemblyConn

        Set oSD_AssemblyConnection.object = oParentCustomAssembly

        Set oBoundedOrPenetratingPart = oSD_AssemblyConnection.ConnectedObject1
        Set oBoundingOrPenetratedPart = oSD_AssemblyConnection.ConnectedObject2

    'custom assembly parent is member assembly connection
    ElseIf TypeOf oParentCustomAssembly Is IJAppConnection Then

        Dim oAppConnection As IJAppConnection
        Set oAppConnection = oParentCustomAssembly

        Dim oInputElements As IJElements
        oAppConnection.enumPorts oInputElements

        Set oBoundedOrPenetratingPart = oInputElements.Item(1)
        Set oBoundingOrPenetratedPart = oInputElements.Item(2)
    
    ElseIf TypeOf oParentCustomAssembly Is IJFreeEndCut Then
    
        Dim oFreeEndCut As IJFreeEndCut
        Set oFreeEndCut = oParentCustomAssembly
        
        oFreeEndCut.get_FreeEndCutInputs oBoundedOrPenetratingPort, Nothing
        Set oBoundedOrPenetratingPart = oBoundedOrPenetratingPort.Connectable
    
    'custom assembly parent is Feature
    ElseIf TypeOf oParentCustomAssembly Is IJStructFeature Then

        Dim oStructFeature As IJStructFeature
        Set oStructFeature = oParentCustomAssembly

        Dim eFeatureType As StructFeatureTypes
        eFeatureType = oStructFeature.get_StructFeatureType

        If eFeatureType = SF_WebCut Then

            Dim oSD_WebCut As StructDetailObjects.WebCut
            Set oSD_WebCut = New StructDetailObjects.WebCut

            Set oSD_WebCut.object = oStructFeature

            If Not oSD_WebCut.IsFreeEndCut Then
                Set oBoundingOrPenetratedPart = oSD_WebCut.Bounding
            End If
            
            Set oBoundedOrPenetratingPart = oSD_WebCut.Bounded
        
        ElseIf eFeatureType = SF_FlangeCut Then
        
            Dim oSD_FlangeCut As StructDetailObjects.FlangeCut
            Set oSD_FlangeCut = New StructDetailObjects.FlangeCut

            Set oSD_FlangeCut.object = oStructFeature

            If Not oSD_FlangeCut.IsFreeEndCut Then
                Set oBoundingOrPenetratedPart = oSD_FlangeCut.Bounding
            End If
            
            Set oBoundedOrPenetratingPart = oSD_FlangeCut.Bounded

        ElseIf eFeatureType = SF_Slot Then

            Dim oSD_Slot As StructDetailObjects.Slot
            Set oSD_Slot = New StructDetailObjects.Slot

            Set oSD_Slot.object = oStructFeature
            Set oBoundingOrPenetratedPart = oSD_Slot.Penetrated
            Set oBoundedOrPenetratingPart = oSD_Slot.Penetrating
            
        ElseIf eFeatureType = SF_EdgeFeature Then

            Dim oSD_EdgeFeature As StructDetailObjects.EdgeFeature
            Set oSD_EdgeFeature = New StructDetailObjects.EdgeFeature

            Set oSD_EdgeFeature.object = oStructFeature
            Set oBoundedOrPenetratingPart = oSD_EdgeFeature.GetPartObject

        ElseIf eFeatureType = SF_CornerFeature Then

            Dim oSD_CornerFeatureEX As StructDetailObjectsEx.CornerFeature
            Set oSD_CornerFeatureEX = New StructDetailObjectsEx.CornerFeature
            
            Set oSD_CornerFeatureEX.object = oStructFeature
            Set oBoundedOrPenetratingPart = oSD_CornerFeatureEX.GetPartObject

        Else
            '''''''
        End If

    'custom assembly parent is chamfer object
    ElseIf TypeOf oParentCustomAssembly Is IJChamfer Then

        Dim oSD_Chamfer As StructDetailObjects.Chamfer
        Set oSD_Chamfer = New StructDetailObjects.Chamfer

        Set oSD_Chamfer.object = oParentCustomAssembly
        Set oBoundedOrPenetratingPart = oSD_Chamfer.ChamferedPart
        Set oBoundingOrPenetratedPart = oSD_Chamfer.DrivesChamferPart

    ElseIf TypeOf oParentCustomAssembly Is IJCollarPart Then

        Dim oSD_CollarPart As StructDetailObjects.Collar
        Set oSD_CollarPart = New StructDetailObjects.Collar

        Set oSD_CollarPart.object = oParentCustomAssembly
        Set oBoundingOrPenetratedPart = oSD_CollarPart.Penetrated
        Set oBoundedOrPenetratingPart = oSD_CollarPart.Penetrating

    ElseIf TypeOf oParentCustomAssembly Is IJSmartPlate Then

        Dim oSmartPlate As IJSmartPlate
        Set oSmartPlate = oParentCustomAssembly

        Dim oSmartPlateAttributes As IJSDSmartPlateAttributes
        Set oSmartPlateAttributes = New SDSmartPlateUtils

        Dim oGraphicInputs As New Collection

        If oSmartPlate.SmartPlateType = spType_INSERT Then
            oSmartPlateAttributes.GetInputs_InsertPlate oSmartPlate, oGraphicInputs

            Dim oFeature As Object

            Dim oObject As Object
            Set oObject = oGraphicInputs.Item(1)

            If TypeOf oObject Is IJStructFeature Then
                Set oStructFeature = oObject

                eFeatureType = oStructFeature.get_StructFeatureType

                If eFeatureType = SF_EdgeFeature Then
                    Set oSD_EdgeFeature = New StructDetailObjects.EdgeFeature

                    Set oSD_EdgeFeature.object = oStructFeature
                    Set oBoundedOrPenetratingPart = oSD_EdgeFeature.GetPartObject
                End If
            End If

        ElseIf oSmartPlate.SmartPlateType = spType_BEARING Then

            oSmartPlateAttributes.GetInputs_BearingPlate oSmartPlate, oGraphicInputs

            Set oBoundedOrPenetratingPort = oGraphicInputs.Item(1)
            Set oBoundingOrPenetratedPort = oGraphicInputs.Item(2)

            Set oBoundedOrPenetratingPart = oBoundedOrPenetratingPort.Connectable
            Set oBoundingOrPenetratedPart = oBoundingOrPenetratedPort.Connectable

        ElseIf oSmartPlate.SmartPlateType = spType_COLLAR Then

            oSmartPlateAttributes.GetInputs_SmartCollar oSmartPlate, oGraphicInputs

            If TypeOf oGraphicInputs.Item(1) Is IJStructFeature Then
                Set oStructFeature = oGraphicInputs.Item(1)
            ElseIf TypeOf oGraphicInputs.Item(2) Is IJStructFeature Then
                Set oStructFeature = oGraphicInputs.Item(2)
            End If

            If oStructFeature.get_StructFeatureType = SF_CornerFeature Then

                Dim oSD_CornerFeature As StructDetailObjects.CornerFeature
                Set oSD_CornerFeature = New StructDetailObjects.CornerFeature
                
                Set oSD_CornerFeature.object = oStructFeature
                Set oBoundedOrPenetratingPart = oSD_CornerFeature.GetPartObject
            End If
        ElseIf oSmartPlate.SmartPlateType = spType_BRACKET Then
            'Not yet supported
        End If

    End If

    Dim bIsPartPartailDetailed As Boolean
    bIsPartPartailDetailed = False

    If Not oBoundedOrPenetratingPart Is Nothing Then
        bIsPartPartailDetailed = CheckIfPartailDetailed(oBoundedOrPenetratingPart)
    End If

    If bIsPartPartailDetailed Then
        ExcludeObjectBasedOnDetailedState = True
    Else
        If Not oBoundingOrPenetratedPart Is Nothing Then
            ExcludeObjectBasedOnDetailedState = CheckIfPartailDetailed(oBoundingOrPenetratedPart)
        End If
    End If

 Exit Function

ErrHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

'**********************************************************************************************
' Method      : CheckIfPartailDetailed
' Description :
'**********************************************************************************************
Public Function CheckIfPartailDetailed(oPart As Object) As Boolean

 Const sMETHOD = "CheckIfPartailDetailed"
 On Error GoTo ErrHandler

    CheckIfPartailDetailed = False

    'Exit with False when the input is member part
    If TypeOf oPart Is ISPSMemberPartCommon Then
        Exit Function
    End If
    
    Dim oHelper As StructDetailHelper
    Set oHelper = New StructDetailHelper

    CheckIfPartailDetailed = oHelper.IsPartialDetailed(oPart)

 Exit Function

ErrHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

