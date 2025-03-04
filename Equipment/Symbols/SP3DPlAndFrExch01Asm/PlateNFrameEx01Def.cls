VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlateNFrameEx01Def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"CAD"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/2/2005-7:23:05 PM"
Attribute VB_Ext_KEY = "SP3DEqpCADTemplateVersion" ,"1.2"
'******************************************************************************
'   Copyright (c) 2006-07, Intergraph Corporation. All rights reserved.
'
'File
'   PlateNFrameEx01Def.cls
'
'Author
'    KB
'
'Description
'   Defines the behavior of members of an Equipment that make a custom assembly
'   When a symbol is used with the smart occurrence that contain rigid geometry,
'   This class usually, but is not limited to, manages the nozzles of the equipment.
'Notes
'   From Template Version:1.2
'
'History:
'   1.Aug.2006      svsmylav    CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'   9.Mar.2007      svsmylav    TR-86137 Corrected orientation of nozzle A and nozzle D. Also corrected the N2 of nozzles C and D.
'******************************************************************************
Option Explicit

Private Const MODULE = "SP3DPlAndFrExch01Asm:PlateNFrameEx01Def"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const IID_IJDGEOMETRY = "{A1732CBF-5136-11D1-9770-080036754203}"

Private m_oEquipCADHelper As IJEquipCADHelper
Private m_oEditErrors As IJEditErrors

' Implement User Symbol Services(USS)
Implements IJDUserSymbolServices
Implements IJEquipUserAttrMgmt
Private m_avSymbolArrayOfInputs()   As Variant
Private m_dEquipmentXLength As Double
Private m_dEquipmentYLength As Double
Private m_dEquipmentZLength As Double
Private m_dFrontPlateThickness As Double
Private m_dRearPlateThickness As Double
Private m_dPlatePackLength As Double
Private m_dCPtoFaceY As Double
Private m_dCPtoFaceZ As Double
Private m_dSupportLength As Double
Private m_dSupportWidth As Double
Private m_dSupportThickness As Double
Private m_dXboltHole As Double
Private m_dYboltHole As Double


Private m_oNorth                    As IJDVector
Private m_oEast                     As IJDVector
Private m_oElevation                As IJDVector

Private Sub Class_Initialize()
    Set m_oEquipCADHelper = New CADServices
    Set m_oEditErrors = New JServerErrors
    
    m_oEquipCADHelper.ProjectName = "SP3DPlAndFrExch01Asm"
    m_oEquipCADHelper.ClassName = "PlateNFrameEx01Def"
    m_oEquipCADHelper.OccurrenceRootClass = orcEquipment

    LogCalls "Class_Initialize"
    
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
    LogCalls "Class_Terminate"
    Set m_oNorth = Nothing
    Set m_oEast = Nothing
    Set m_oElevation = Nothing

'    Set m_oPartFclt = Nothing

    Set m_oEditErrors = Nothing
    Set m_oEquipCADHelper = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Implementation of IJDUserSymbolServices
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String

    ' Progid should be unique
    IJDUserSymbolServices_GetDefinitionName = m_oEquipCADHelper.ProjectName & "." & m_oEquipCADHelper.ClassName
    
End Function

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, _
                                                             ByVal defParams As Variant, _
                                                             ByVal pResourceMgr As Object) As Object
    Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
    LogCalls METHOD

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
    LogCalls METHOD
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
    
    'Add new member(NozzleA1) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleA1", 2, "CMConstructNozzleA1", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleA1"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleA1"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleA1"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleA1"
    'Add properties for (NozzleA1)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleA1Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleA1", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleA1GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleA1", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleB2) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleB2", 3, "CMConstructNozzleB2", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleB2"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleB2"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleB2"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleB2"
    'Add properties for (NozzleB2)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleB2Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleB2", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleB2GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleB2", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleC3) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleC3", 4, "CMConstructNozzleC3", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleC3"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleC3"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleC3"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleC3"
    'Add properties for (NozzleC3)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleC3Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleC3", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleC3GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleC3", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleD4) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleD4", 5, "CMConstructNozzleD4", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleD4"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleD4"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleD4"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleD4"
    'Add properties for (NozzleD4)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleD4Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleD4", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleD4GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleD4", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleEqpFoundationPort5) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleEqpFoundationPort5", 6, "CMConstructNozzleEqpFoundationPort5", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleEqpFoundationPort5"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleEqpFoundationPort5"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleEqpFoundationPort5"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleEqpFoundationPort5"
    'Add properties for (NozzleEqpFoundationPort5)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort5Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleEqpFoundationPort5", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort5GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleEqpFoundationPort5", imsCOOKIE_ID_USS_LIB
    
        
    Set oAggregatorDescription = Nothing
    Set oMemberDescriptions = Nothing
    Set oMemberDescription = Nothing
    Set oPropertyDescriptions = Nothing
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)
    Const METHOD = "IJDUserSymbolServices_InvokeRepresentation"
    LogCalls METHOD
    
    'This method is not used by the CAD.

End Sub

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
    Const METHOD = "IJDUserSymbolServices_EditOccurence"
    LogCalls METHOD
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
        Case "DP1", "NozzleA1", "NozzleB2", "NozzleC3", "NozzleD4", "NozzleEqpFoundationPort5"
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
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    
    
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
        Case "DP1"
'            m_oEquipCADHelper.SetMemberPropertiesAsReadOnly oMemberDescription, CollAllDisplayedValues, False
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
        Case "NozzleEqpFoundationPort5"
      
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
' Custom Methods for NozzleA1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleA1(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleA1"
    On Error GoTo ErrorHandler
    
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 1, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.0001

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Axial
    oOrientation.N1 = -0.0001 'If we provide zero, nozzle orientation is taken along positive X-axis.
    oOrientation.N2 = 0
    oOrientation.OR1 = 0

    Set oNozzle = Nothing
    Set oOrientation = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleA1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleA1"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleA1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleA1"
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

Public Sub CMEvaluateNozzleA1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleA1"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleA1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleA1"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleA1(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleA1"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleA1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleA1"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleB2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleB2(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleB2"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 2, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.0001

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Axial
    oOrientation.N1 = (m_dFrontPlateThickness + m_dPlatePackLength + m_dRearPlateThickness)
    oOrientation.N2 = 0
    oOrientation.OR1 = 0

    Set oNozzle = Nothing
    Set oOrientation = Nothing

    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleB2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleB2"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleB2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleB2"
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

Public Sub CMEvaluateNozzleB2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleB2"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleB2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleB2"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleB2(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleB2"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleB2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleB2"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleC3
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleC3(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleC3"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 3, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.0001

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Axial
    oOrientation.N1 = (m_dFrontPlateThickness + m_dPlatePackLength + m_dRearPlateThickness)
    oOrientation.N2 = m_dEquipmentYLength - 2 * m_dCPtoFaceY
    oOrientation.OR1 = 0

    Set oNozzle = Nothing
    Set oOrientation = Nothing

    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleC3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleC3"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleC3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleC3"
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

Public Sub CMEvaluateNozzleC3(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleC3"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleC3(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleC3"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
         
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleC3(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleC3"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleC3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleC3"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleD4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleD4(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleD4"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 4, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.001

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Axial
    oOrientation.N1 = -0.0001 'If we provide zero, nozzle orientation is taken along positive X-axis.
    oOrientation.N2 = m_dEquipmentYLength - 2 * m_dCPtoFaceY
    oOrientation.OR1 = 0

    Set oNozzle = Nothing
    Set oOrientation = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleD4(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleD4"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleD4(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleD4"
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

Public Sub CMEvaluateNozzleD4(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleD4"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleD4(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleD4"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleD4(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleD4"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleD4(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleD4"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleEqpFoundationPort5
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleEqpFoundationPort5(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleEqpFoundationPort5"
    On Error GoTo ErrorHandler
        
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleFromPH pMemberDescription, pResourceManager, pObject, 5
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleEqpFoundationPort5(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleEqpFoundationPort5"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleEqpFoundationPort5(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleEqpFoundationPort5"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleEqpFoundationPort5(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateNozzleEqpFoundationPort5"
On Error GoTo ErrorHandler
    Dim objEqpFoundationPort As IJEqpFoundationPort
    Dim holes() As Variant
    Dim dHorSupportLength As Double
    'Get the Symbol inputs of the Component
    GetDimensionsFromSymbolArray oPropertyDescription.CAO
    dHorSupportLength = m_dEquipmentXLength - 2 * m_dXboltHole - _
                                    m_dFrontPlateThickness - m_dRearPlateThickness - m_dPlatePackLength

    'Now Construct the holes Array
    Set objEqpFoundationPort = oPropertyDescription.Object
    Call objEqpFoundationPort.GetHoles(holes())
    holes(0, 1) = -m_dCPtoFaceY + m_dEquipmentYLength / 2 - m_dYboltHole / 2
    holes(0, 2) = -m_dXboltHole
    holes(1, 1) = -m_dCPtoFaceY + m_dEquipmentYLength / 2 + m_dYboltHole / 2
    holes(1, 2) = -m_dXboltHole
    holes(2, 1) = -m_dCPtoFaceY + m_dEquipmentYLength / 2
    holes(2, 2) = m_dFrontPlateThickness + m_dPlatePackLength + m_dRearPlateThickness + dHorSupportLength _
                                - 0.9 * m_dFrontPlateThickness + m_dXboltHole
    'Set the calculated holes on foundation port
    Call objEqpFoundationPort.SetHoles(holes)
        
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleEqpFoundationPort5(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleEqpFoundationPort5"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
    m_oEquipCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 5
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleEqpFoundationPort5(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleEqpFoundationPort5"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleEqpFoundationPort5(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleEqpFoundationPort5"
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
'The code of this method is specific to each Equipment as it converts the array of inputs in a set of global variables
Private Sub GetDimensionsFromSymbolArray(SmartOccurrence As IJSmartOccurrence)
    Const METHOD = "GetDimensionsFromSymbolArray"
    On Error GoTo ErrorHandler

    m_avSymbolArrayOfInputs = m_oEquipCADHelper.GetSymbolArrayOfInputs(SmartOccurrence)

    'Inputs,  from equipment symbol code
    'Set m_oPartFclt = m_avSymbolArrayOfInputs(1)
    m_dEquipmentXLength = m_avSymbolArrayOfInputs(2)
    m_dEquipmentYLength = m_avSymbolArrayOfInputs(3)
    m_dEquipmentZLength = m_avSymbolArrayOfInputs(4)
    m_dFrontPlateThickness = m_avSymbolArrayOfInputs(5)
    m_dRearPlateThickness = m_avSymbolArrayOfInputs(6)
    m_dPlatePackLength = m_avSymbolArrayOfInputs(7)
    m_dCPtoFaceY = m_avSymbolArrayOfInputs(8)
    m_dCPtoFaceZ = m_avSymbolArrayOfInputs(9)
    m_dSupportLength = m_avSymbolArrayOfInputs(10)
    m_dSupportWidth = m_avSymbolArrayOfInputs(11)
    m_dSupportThickness = m_avSymbolArrayOfInputs(12)
    m_dXboltHole = m_avSymbolArrayOfInputs(13)
    m_dYboltHole = m_avSymbolArrayOfInputs(14)

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

Private Sub LogCalls(sMethod As String)
   m_oEquipCADHelper.LogError "Entering " & sMethod
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
