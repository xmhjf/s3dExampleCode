VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BB3PumpDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************
'   Copyright (C) 2007, Intergraph Corporation. All rights reserved.
'
'   File:      BB3PumpDef.cls
'   Author:    VRK
'   Creation Date:  Thursday, April 12 2007
'
'   Description:
'   This is Horizontal Pump Of Type BB3 Assembly.
'   This Assembly has One DatumPoint with Orientations X towards East and Y along North, and
'   One Rectangular FoundationPort
'
'   Change History:
'   dd.mmm.yyyy         who         change description
'   -----------         -----        ------------------
'   14.Dec.2007         VRK         TR-CP·132125 : API pump BB3 symbol is incorrect Suction
'                                   and Discharge ports should be flipped
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "BB3PumpDef"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const IID_IJDGEOMETRY = "{A1732CBF-5136-11D1-9770-080036754203}"

Private m_oEquipCADHelper As IJEquipCADHelper
Private m_oEditErrors As IJEditErrors
Private m_avSymbolArrayOfInputs() As Variant
Private PI As Double

Private m_oNorth  As IJDVector
Private m_oEast   As IJDVector
Private m_oElevation   As IJDVector

Private m_dImpellerDiameter As Double
Private m_dSuctionCLtoDischargeCL As Double
Private m_dPumpCLtoDischargeFace As Double
Private m_dPumpCLtoSuctionFace As Double
Private m_dPumpCLtoDischargeCL As Double
Private m_dPumpCLtoSuctionCL As Double
Private m_dXboltHole As Double
Private m_dYboltHole As Double

'Implement User Symbol Services(USS)
Implements IJDUserSymbolServices
Implements IJEquipUserAttrMgmt

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
    Set m_oEquipCADHelper = New CADServices
    Set m_oEditErrors = New JServerErrors
    
    m_oEquipCADHelper.ProjectName = "SP3DHoriPumpBB3Asm"
    m_oEquipCADHelper.ClassName = "BB3PumpDef"
    m_oEquipCADHelper.OccurrenceRootClass = orcEquipment
    
    Set m_oEast = New DVector
    m_oEast.x = 1
    m_oEast.y = 0
    m_oEast.z = 0
    
    Set m_oNorth = New DVector
    m_oNorth.x = 0
    m_oNorth.y = 1
    m_oNorth.z = 0
    
    Set m_oElevation = New DVector
    m_oElevation.x = 0
    m_oElevation.y = 0
    m_oElevation.z = 1
End Sub

Private Sub Class_Terminate()
    Set m_oEditErrors = Nothing
    Set m_oEquipCADHelper = Nothing
    Set m_oNorth = Nothing
    Set m_oEast = Nothing
    Set m_oElevation = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implementation of IJDUserSymbolServices
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
    'Progid should be unique
    IJDUserSymbolServices_GetDefinitionName = m_oEquipCADHelper.ProjectName & "." & m_oEquipCADHelper.ClassName
End Function

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
                                                             ByVal defParams As Variant, _
                                                             ByVal pResourceMgr As Object) As Object
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    On Error GoTo ErrorHandler
    
    Dim oSymbolDefinition As IJDSymbolDefinition
    Set oSymbolDefinition = m_oEquipCADHelper.InstanciateDefinition(CodeBase, defParams, pResourceMgr)
    IJDUserSymbolServices_InitializeSymbolDefinition oSymbolDefinition
    Set IJDUserSymbolServices_InstanciateDefinition = oSymbolDefinition
    
    Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition)
    Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
    On Error GoTo ErrorHandler
    
    Dim oPropertyDescriptions As IJDPropertyDescriptions
    Dim oMemberDescriptions As IJDMemberDescriptions
    Dim oMemberDescription As IJDMemberDescription
    Dim oAggregatorDescription As IJDAggregatorDescription
    
    'Set up the aggregator
    Set oAggregatorDescription = oSymbolDefinition
    oAggregatorDescription.AggregatorClsid = m_oEquipCADHelper.OccurrenceRootClassGUID()
    oAggregatorDescription.SetCMConstruct imsCOOKIE_ID_USS_LIB, "CMConstructEquipment"
    oAggregatorDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEquipment"
    oAggregatorDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsEquipment"
    oAggregatorDescription.SetCMRemoveInputs imsCOOKIE_ID_USS_LIB, "CMRemoveInputsEquipment"
    
    'Add property to aggregator
    Set oPropertyDescriptions = oAggregatorDescription
    oPropertyDescriptions.RemoveAll
    oPropertyDescriptions.AddProperty "EquipmentProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEquipment", imsCOOKIE_ID_USS_LIB
    
    'Remove all the previous member descriptions
    Set oMemberDescriptions = oSymbolDefinition
    oMemberDescriptions.RemoveAll
    
    'Add your code here for the declaration of the Public Custom Methods used to manage new members
    'DP1
    'Add new member DP1 to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("DP1", 1, "CMConstructDP1", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsDP1"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructDP1"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalDP1"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseDP1"
    
    'Add properties for DP1
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "DP1Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateDP1", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "DP1GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryDP1", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleSuction1) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleSuction1", 2, "CMConstructNozzleSuction1", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleSuction1"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleSuction1"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleSuction1"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleSuction1"
    'Add properties for (NozzleSuction1)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleSuction1Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleSuction1", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleSuction1GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleSuction1", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleDischarge2) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleDischarge2", 3, "CMConstructNozzleDischarge2", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleDischarge2"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleDischarge2"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleDischarge2"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleDischarge2"
    'Add properties for (NozzleDischarge2)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleDischarge2Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleDischarge2", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleDischarge2GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleDischarge2", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleSTFndPort1) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("s", 4, "CMConstructNozzleSTFndPort1", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleSTFndPort1"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleSTFndPort1"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleSTFndPort1"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleSTFndPort1"
    'Add properties for (NozzleSTFndPort1)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleSTFndPort1Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleSTFndPort1", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleSTFndPort1GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleSTFndPort1", imsCOOKIE_ID_USS_LIB

    Set oAggregatorDescription = Nothing
    Set oMemberDescriptions = Nothing
    Set oMemberDescription = Nothing
    Set oPropertyDescriptions = Nothing
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
    'This method is not used by the CAD.
End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    'Obsolete method. Instead you can record your custom command within the definition (see IJDCommandDescription interface)
    IJDUserSymbolServices_EditOccurence = False
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implementation of IJEquipUserAttrMgmt
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IJEquipUserAttrMgmt_OnAttributeChange(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As IJEquipAttrDescriptor, ByVal varNewAttrValue As Variant) As String
    Const METHOD = "IJEquipUserAttrMgmt_OnAttributeChange"
    On Error GoTo ErrorHandler
    
    Dim oMemberDescription As IJDMemberDescription
    Set oMemberDescription = m_oEquipCADHelper.GetMemberDescriptionFromChild(pIJDAttrs)
        
    Select Case oMemberDescription.Name
        Case "DP1", "NozzleSuction1", "NozzleDischarge2", "NozzleSTFndPort1"
            Select Case UCase(pAttrToChange.InterfaceName)
                Case "IJDELETABLEMEMBER"
                    If UCase(pAttrToChange.AttrName) = "CANBEDELETED" Then
                        m_oEquipCADHelper.MakeMemberDeletable oMemberDescription, pIJDAttrs, CBool(varNewAttrValue)
                    End If
                Case Else
                    '
            End Select

        Case Else
            '
    End Select
    
    Set oMemberDescription = Nothing
    IJEquipUserAttrMgmt_OnAttributeChange = ""
    
    Exit Function
ErrorHandler:
    IJEquipUserAttrMgmt_OnAttributeChange = "ERROR"
    HandleError MODULE, METHOD
End Function

Private Function IJEquipUserAttrMgmt_OnPreCommit(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    Const METHOD = "IJEquipUserAttrMgmt_OnPreCommit"
    On Error GoTo ErrorHandler
    
    IJEquipUserAttrMgmt_OnPreCommit = "ERROR"
    IJEquipUserAttrMgmt_OnPreCommit = ""
    
    Exit Function
ErrorHandler:
    IJEquipUserAttrMgmt_OnPreCommit = "ERROR"
    HandleError MODULE, METHOD
End Function

Private Function IJEquipUserAttrMgmt_OnPreLoad(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
    Const METHOD = "IJEquipUserAttrMgmt_OnPreLoad"
    On Error GoTo ErrorHandler
    
    Dim oMemberDescription As IJDMemberDescription
    IJEquipUserAttrMgmt_OnPreLoad = "ERROR"
    
    Set oMemberDescription = m_oEquipCADHelper.GetMemberDescriptionFromChild(pIJDAttrs)
    Dim oAttrCollection As Collection
    Dim oAttributeDescriptor As IJEquipAttrDescriptor
    Dim m As Long
    Set oAttrCollection = CollAllDisplayedValues
    
    Select Case oMemberDescription.Name
        Case "NozzleSuction1", "NozzleDischarge2", "NozzleSTFndPort1"
             For m = 1 To oAttrCollection.Count
                Set oAttributeDescriptor = oAttrCollection.Item(m)
                Select Case UCase(oAttributeDescriptor.InterfaceName)
                Case "IJDELETABLEMEMBER"
                    If UCase(oAttributeDescriptor.AttrName) = "CANBEDELETED" Then
                        oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsChanged
                    End If
                Case Else
                        '
                End Select
            Next
        Case Else
            '
    End Select
        Select Case oMemberDescription.Name
         Case "DP1"
            For m = 1 To oAttrCollection.Count
                Set oAttributeDescriptor = oAttrCollection.Item(m)
                Select Case UCase(oAttributeDescriptor.InterfaceName)
                    Case "IJUADATUMSHAPE"
                        oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsReadOnly
                    Case Else
                        '
                End Select
            Next
        Case Else
            '
    End Select
     Select Case oMemberDescription.Name
             Case "NozzleSTFndPort1"
            For m = 1 To oAttrCollection.Count
                Set oAttributeDescriptor = oAttrCollection.Item(m)
                Select Case UCase(oAttributeDescriptor.InterfaceName)
                Case "IJNOZZLEORIENTATION"
                    'The nozzle is placed by Point
                    oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsReadOnly
                Case Else
                        '
                End Select
            Next
        Case Else
            '
    End Select
    Set oAttributeDescriptor = Nothing
    Set oMemberDescription = Nothing
    Set oAttrCollection = Nothing
    IJEquipUserAttrMgmt_OnPreLoad = ""
    
    Exit Function
ErrorHandler:
    IJEquipUserAttrMgmt_OnPreLoad = "ERROR"
    HandleError MODULE, METHOD
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods of Equipment (Aggregator)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMConstructEquipment"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMFinalConstructEquipment"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMSetInputsEquipment"
    On Error GoTo ErrorHandler
    
    m_oEquipCADHelper.SetSmartItemAsInputToSymbol pAggregatorDescription
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMRemoveInputsEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMRemoveInputsEquipment"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateEquipment(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateEquipment"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for DP1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructDP1(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oDatumShape As IJShape
    Dim oDesignEquipment As IJDesignEquipment
        
    'Create Datum Shape DP1
    Set oDatumShape = m_oEquipCADHelper.CreateShape(pMemberDescription, pResourceManager, "DatumShape 001", "DP1")
    If Not oDatumShape Is Nothing Then
        Set pObject = oDatumShape
        oDatumShape.RepresentationId = ReferenceGeometry

        Set oDesignEquipment = pMemberDescription.CAO
        oDesignEquipment.AddShape oDatumShape

        GetDimensionsFromSymbolArray oDesignEquipment
        PositionAndOrientDP1 oDesignEquipment, oDatumShape
    End If
    
    Set oDesignEquipment = Nothing
    Set oDatumShape = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructDP1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsDP1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateDP1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryDP1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oEquipment As IJEquipment
    Dim oDatumShape As IJShape
    Set oDatumShape = oPropertyDescription.Object
    
    'Test if this DatumShape follows the symbol inputs
      If m_oEquipCADHelper.IsShapeFreeToTransform(oDatumShape) = False Then
        Set oEquipment = oPropertyDescription.CAO
        GetDimensionsFromSymbolArray oEquipment
        PositionAndOrientDP1 oEquipment, oDatumShape
    Else
        'Do nothing, the equipment will transform the shape for us
    End If
    
    Set oDatumShape = Nothing
    Set oEquipment = Nothing
    
    Exit Sub
ErrorHandler:
    Set oDatumShape = Nothing
    Set oEquipment = Nothing
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalDP1(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseDP1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseDP1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleSuction1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleSuction1(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleSuction1"
    On Error GoTo ErrorHandler
    
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 1, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = m_dPumpCLtoSuctionFace - (0.25 * m_dImpellerDiameter) / 2

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Tangential
    
    oOrientation.N1 = -m_dSuctionCLtoDischargeCL
    oOrientation.N2 = m_dPumpCLtoSuctionFace
    oOrientation.N3 = -m_dPumpCLtoSuctionCL
    oOrientation.OR1 = PI
    Set oNozzle = Nothing
    Set oOrientation = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleSuction1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleSuction1"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleSuction1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleSuction1"
    On Error GoTo ErrorHandler
    
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleSuction1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleSuction1"
    On Error GoTo ErrorHandler
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleSuction1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleSuction1"
    On Error GoTo ErrorHandler
    
    Dim oEquipment As IJEquipment
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    
    Set pObject = oPropertyDescription.Object
    Set oNozzle = pObject
    oNozzle.Length = m_dPumpCLtoSuctionFace - (0.25 * m_dImpellerDiameter) / 2
    
    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)

    'Set the default values
    oOrientation.PlacementType = Tangential
    
    oOrientation.N1 = -m_dSuctionCLtoDischargeCL
    oOrientation.N2 = m_dPumpCLtoSuctionFace
    oOrientation.N3 = -m_dPumpCLtoSuctionCL
    oOrientation.OR1 = PI

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Set oEquipment = Nothing
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleSuction1(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleSuction1"
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleSuction1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleSuction1"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleDischarge2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleDischarge2(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleDischarge2"
    On Error GoTo ErrorHandler
    
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO

    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 2, pResourceManager, DistribPortType_PIPE, pObject, False

    Set oNozzle = pObject
    oNozzle.Length = m_dPumpCLtoDischargeFace - (0.25 * m_dImpellerDiameter) / 2
    
    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)

    'Set the default values
    oOrientation.PlacementType = Tangential
    
    oOrientation.N1 = 0
    oOrientation.N2 = m_dPumpCLtoDischargeFace
    oOrientation.N3 = m_dPumpCLtoDischargeCL
    oOrientation.OR1 = 0

    Set oNozzle = Nothing
    Set oOrientation = Nothing
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleDischarge2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleDischarge2"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleDischarge2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleDischarge2"
    On Error GoTo ErrorHandler
    
    Dim l As Long
    Dim oShape As IJShape
    Dim oSmartOcc As IJSmartOccurrence
    Dim oMemberobjects As IJDMemberObjects
    Dim oMemberDesc As IJDMemberDescription
    Dim oOrientation As IJNozzleOrientation
        
    Set oSmartOcc = pMemberDesc.CAO
    'Get the shape member the nozzle is to be related to
    Set oMemberobjects = oSmartOcc
    For l = 1 To oMemberobjects.Count
        Set oMemberDesc = oMemberobjects.MemberDescriptions.Item(l)
        If oMemberDesc.Name = "DP1" Then
            Set oShape = oMemberobjects.Item(l)
            Exit For
        End If
    Next l

    'Establish the relationship with the shape if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.Object)
    
    Set oShape = Nothing
    Set oSmartOcc = Nothing
    Set oMemberDesc = Nothing
    Set oOrientation = Nothing
    Set oMemberobjects = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleDischarge2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleDischarge2"
    On Error GoTo ErrorHandler
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleDischarge2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleDischarge2"
    On Error GoTo ErrorHandler
    
    Dim oEquipment As IJEquipment
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment
    
    Set pObject = oPropertyDescription.Object
    Set oNozzle = pObject
    oNozzle.Length = m_dPumpCLtoDischargeFace - (0.25 * m_dImpellerDiameter) / 2
    
    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)

   'Set the default values
    oOrientation.PlacementType = Tangential
    
    oOrientation.N1 = 0
    oOrientation.N2 = m_dPumpCLtoDischargeFace
    oOrientation.N3 = m_dPumpCLtoDischargeCL
    oOrientation.OR1 = 0

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Set oEquipment = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleDischarge2(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleDischarge2"
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleDischarge2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleDischarge2"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleSTFndPort1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleSTFndPort1(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleSTFndPort1"
    On Error GoTo ErrorHandler
        
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleFromPH pMemberDescription, pResourceManager, pObject, 3
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleSTFndPort1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleSTFndPort1"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleSTFndPort1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleSTFndPort1"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleSTFndPort1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleSTFndPort1"
    On Error GoTo ErrorHandler
    
    Dim objEqpFoundationPort As IJEqpFoundationPort
    Dim holes() As Variant
    
    'Get the Symbol inputs of the Component
    GetDimensionsFromSymbolArray oPropertyDescription.CAO
    
    'Now Construct the holes Array
    Set objEqpFoundationPort = oPropertyDescription.Object
    Call objEqpFoundationPort.GetHoles(holes())
    
    holes(0, 1) = -m_dXboltHole
    holes(0, 2) = -m_dYboltHole
    holes(1, 1) = m_dXboltHole
    holes(1, 2) = -m_dYboltHole
    holes(2, 1) = m_dXboltHole
    holes(2, 2) = m_dYboltHole
    holes(3, 1) = -m_dXboltHole
    holes(3, 2) = m_dYboltHole

    'Set the calculated holes on foundation port
    Call objEqpFoundationPort.SetHoles(holes)
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleSTFndPort1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleSTFndPort1"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
    m_oEquipCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 3
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleSTFndPort1(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleSTFndPort1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleSTFndPort1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleSTFndPort1"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub HandleError(sModule As String, sMethod As String)
    If Not m_oEditErrors Is Nothing Then
        m_oEditErrors.AddFromErr Err, "", sMethod, sModule
    End If
    
    Err.Raise Err.Number, Err.Source & " " & sMethod, Err.Description, _
    Err.HelpFile, Err.HelpContext
End Sub

'The code of this method is specific to each Equipment Component as it converts the array of inputs in a set of global variables
Private Sub GetDimensionsFromSymbolArray(SmartOccurrence As IJSmartOccurrence)
    Const METHOD = "GetDimensionsFromSymbolArray"
    On Error GoTo ErrorHandler
    
    m_avSymbolArrayOfInputs = m_oEquipCADHelper.GetSymbolArrayOfInputs(SmartOccurrence)
    
    'Inputs,  from equipment symbol code
    m_dSuctionCLtoDischargeCL = m_avSymbolArrayOfInputs(4)           'X
    m_dPumpCLtoDischargeFace = m_avSymbolArrayOfInputs(6)            'T
    m_dPumpCLtoSuctionFace = m_avSymbolArrayOfInputs(7)              'U
    m_dPumpCLtoDischargeCL = m_avSymbolArrayOfInputs(8)              'R
    m_dPumpCLtoSuctionCL = m_avSymbolArrayOfInputs(9)                'S
    m_dImpellerDiameter = m_avSymbolArrayOfInputs(10)
    m_dXboltHole = m_avSymbolArrayOfInputs(17)
    m_dYboltHole = m_avSymbolArrayOfInputs(18)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub LogCalls(sMethod As String)
   m_oEquipCADHelper.logError "Entering " & sMethod
End Sub

Private Sub PositionAndOrientDP1(Equipment As IJEquipment, Shape As IJShape)
    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    oPosition.Set 0, 0, 0
    'We want:
    'the X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'the Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oEast, m_oNorth
    
    Set oPosition = Nothing
End Sub



