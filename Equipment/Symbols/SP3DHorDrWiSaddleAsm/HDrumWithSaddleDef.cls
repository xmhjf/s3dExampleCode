VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HDrumWithSaddleDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"CAD"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/2/2005-7:22:51 PM"
Attribute VB_Ext_KEY = "SP3DEqpCADTemplateVersion" ,"1.2"
'******************************************************************************
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'File
'   HDrumWithSaddleDef.cls
'
'Author
'       KB
'
'Description
'       Defines the behavior of members of an Equipment that make a custom assembly
'       When a symbol is used with the smart occurrence that contain rigid geometry,
'       This class usually, but is not limited to, manages the nozzles of the equipment.
'Notes
'   From Template Version:1.2
'
'History:
'   1.Aug.2006      svsmylav                CR-89878 Removed reference to Dow Emetl Standards (replaced existing symbol).
'******************************************************************************
Option Explicit

Private Const MODULE = "SP3DHorDrWiSaddleAsm:HDrumWithSaddleDef"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const IID_IJDGEOMETRY = "{A1732CBF-5136-11D1-9770-080036754203}"

Private m_oEquipCADHelper As IJEquipCADHelper
Private m_oEditErrors As IJEditErrors

' Implement User Symbol Services(USS)
Implements IJDUserSymbolServices
Implements IJEquipUserAttrMgmt
Private m_avSymbolArrayOfInputs()   As Variant

Private PI           As Double
Private m_dVesselDiameter As Double
Private m_dVesselTantoTan As Double
Private m_dFirstSupportLocation As Double
Private m_dSecondSupportLocation As Double
Private m_dInsulationThickness As Double
Private m_parSupportLength As Double
Private m_parXboltHole As Double
Private m_parYboltHole As Double


Private m_oNorth                    As IJDVector
Private m_oEast                     As IJDVector
Private m_oElevation                As IJDVector

Private Sub Class_Initialize()
    Set m_oEquipCADHelper = New CADServices
    Set m_oEditErrors = New JServerErrors
    
    m_oEquipCADHelper.ProjectName = "SP3DHorDrWiSaddleAsm"
    m_oEquipCADHelper.ClassName = "HDrumWithSaddleDef"
    m_oEquipCADHelper.OccurrenceRootClass = orcEquipment

LogCalls "Class_Initialize"
    PI = 4 * Atn(1)
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
    
    'Add new member(NozzleB12) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleB12", 3, "CMConstructNozzleB12", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleB12"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleB12"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleB12"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleB12"
    'Add properties for (NozzleB12)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleB12Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleB12", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleB12GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleB12", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleB23) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleB23", 4, "CMConstructNozzleB23", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleB23"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleB23"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleB23"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleB23"
    'Add properties for (NozzleB23)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleB23Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleB23", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleB23GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleB23", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleH4) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleH4", 5, "CMConstructNozzleH4", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleH4"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleH4"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleH4"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleH4"
    'Add properties for (NozzleH4)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleH4Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleH4", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleH4GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleH4", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleL5) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleL5", 6, "CMConstructNozzleL5", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleL5"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleL5"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleL5"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleL5"
    'Add properties for (NozzleL5)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleL5Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleL5", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleL5GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleL5", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleNozzle66) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleNozzle66", 7, "CMConstructNozzleNozzle66", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleNozzle66"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleNozzle66"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleNozzle66"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleNozzle66"
    'Add properties for (NozzleNozzle66)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleNozzle66Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleNozzle66", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleNozzle66GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleNozzle66", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleNozzle77) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleNozzle77", 8, "CMConstructNozzleNozzle77", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleNozzle77"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleNozzle77"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleNozzle77"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleNozzle77"
    'Add properties for (NozzleNozzle77)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleNozzle77Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleNozzle77", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleNozzle77GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleNozzle77", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleEqpFoundationPort8) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleEqpFoundationPort8", 9, "CMConstructNozzleEqpFoundationPort8", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleEqpFoundationPort8"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleEqpFoundationPort8"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleEqpFoundationPort8"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleEqpFoundationPort8"
    'Add properties for (NozzleEqpFoundationPort8)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort8Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleEqpFoundationPort8", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort8GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleEqpFoundationPort8", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleEqpFoundationPort19) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleEqpFoundationPort19", 10, "CMConstructNozzleEqpFoundationPort19", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleEqpFoundationPort19"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleEqpFoundationPort19"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleEqpFoundationPort19"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleEqpFoundationPort19"
    'Add properties for (NozzleEqpFoundationPort19)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort19Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleEqpFoundationPort19", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleEqpFoundationPort19GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleEqpFoundationPort19", imsCOOKIE_ID_USS_LIB
    
        
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
        Case "DP1", "NozzleA1", "NozzleB12", "NozzleB23", "NozzleH4", "NozzleL5", "NozzleNozzle66", "NozzleNozzle77", "NozzleEqpFoundationPort8", "NozzleEqpFoundationPort19"
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
        Case "DP1"
'            m_oEquipCADHelper.SetMemberPropertiesAsReadOnly oMemberDescription, CollAllDisplayedValues, False
            For m = 1 To oAttrCollection.count
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
        Case "NozzleEqpFoundationPort8", "NozzleEqpFoundationPort19"
      
            For m = 1 To oAttrCollection.count
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
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    oOrientation.N1 = m_dFirstSupportLocation
    oOrientation.N2 = m_dVesselDiameter / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = 3 * PI / 2

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
    For l = 1 To oMemberobjects.count
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
' Custom Methods for NozzleB12
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleB12(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleB12"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 2, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    oOrientation.N1 = m_dVesselTantoTan * 0.85
    oOrientation.N2 = m_dVesselDiameter / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleB12(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleB12"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleB12(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleB12"
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
    For l = 1 To oMemberobjects.count
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

Public Sub CMEvaluateNozzleB12(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleB12"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleB12(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleB12"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleB12(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleB12"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleB12(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleB12"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleB23
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleB23(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleB23"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 3, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    
    oOrientation.N1 = m_dVesselTantoTan * 0.85
    oOrientation.N2 = m_dVesselDiameter / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = 3 * PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleB23(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleB23"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleB23(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleB23"
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
    For l = 1 To oMemberobjects.count
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

Public Sub CMEvaluateNozzleB23(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleB23"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleB23(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleB23"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleB23(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleB23"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleB23(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleB23"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleH4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleH4(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleH4"
    On Error GoTo ErrorHandler
        
    'Create Nozzle
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 4, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Radial
    oOrientation.N1 = m_dVesselTantoTan / 2
    oOrientation.N2 = m_dVesselDiameter / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = 3 * PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleH4(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleH4"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleH4(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleH4"
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
    For l = 1 To oMemberobjects.count
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

Public Sub CMEvaluateNozzleH4(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleH4"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleH4(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleH4"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleH4(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleH4"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleH4(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleH4"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleL5
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleL5(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleL5"
    On Error GoTo ErrorHandler
        
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 5, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Axial
    oOrientation.N1 = m_dVesselTantoTan + m_dVesselDiameter / 4 + oNozzle.Length * 0.8
    oOrientation.N2 = m_dVesselDiameter / 4
    oOrientation.OR1 = PI

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleL5(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleL5"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleL5(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleL5"
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
    For l = 1 To oMemberobjects.count
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

Public Sub CMEvaluateNozzleL5(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleL5"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleL5(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleL5"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleL5(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleL5"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleL5(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleL5"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleNozzle66
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleNozzle66(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleNozzle66"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 6, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
     oOrientation.PlacementType = Radial
    
    oOrientation.N1 = m_dVesselTantoTan / 2
    oOrientation.N2 = m_dVesselDiameter / 2 + oNozzle.Length * 0.8
    oOrientation.OR1 = PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleNozzle66(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleNozzle66"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleNozzle66(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleNozzle66"
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
    For l = 1 To oMemberobjects.count
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

Public Sub CMEvaluateNozzleNozzle66(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleNozzle66"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleNozzle66(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleNozzle66"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleNozzle66(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleNozzle66"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleNozzle66(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleNozzle66"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleNozzle77
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleNozzle77(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleNozzle77"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 7, pResourceManager, DistribPortType_PIPE, pObject, False
    
    Set oNozzle = pObject
    oNozzle.Length = 0.254

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Axial
    
    oOrientation.N1 = -m_dVesselDiameter / 8 - oNozzle.Length * 0.8
    oOrientation.N2 = m_dVesselDiameter / 4
    oOrientation.OR1 = PI / 2

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleNozzle77(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleNozzle77"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleNozzle77(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleNozzle77"
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
    For l = 1 To oMemberobjects.count
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

Public Sub CMEvaluateNozzleNozzle77(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleNozzle77"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleNozzle77(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleNozzle77"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleNozzle77(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleNozzle77"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleNozzle77(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleNozzle77"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleEqpFoundationPort8
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleEqpFoundationPort8(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleEqpFoundationPort8"
    On Error GoTo ErrorHandler
        
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleFromPH pMemberDescription, pResourceManager, pObject, 8
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleEqpFoundationPort8(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleEqpFoundationPort8"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleEqpFoundationPort8(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleEqpFoundationPort8"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleEqpFoundationPort8(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleEqpFoundationPort8"
    On Error GoTo ErrorHandler
    Dim objEqpFoundationPort As IJEqpFoundationPort
    Dim holes() As Variant
    'Get the Symbol inputs of the Component
    GetDimensionsFromSymbolArray oPropertyDescription.CAO
    'Now Construct the holes Array
    Set objEqpFoundationPort = oPropertyDescription.Object
    Call objEqpFoundationPort.GetHoles(holes())
    holes(0, 1) = (m_parSupportLength - m_parYboltHole) / 2
    holes(0, 2) = m_parXboltHole
    holes(1, 1) = (m_parSupportLength + m_parYboltHole) / 2
    holes(1, 2) = m_parXboltHole
    holes(2, 1) = (m_parSupportLength + m_parYboltHole) / 2
    holes(2, 2) = -m_parXboltHole
    holes(3, 1) = (m_parSupportLength - m_parYboltHole) / 2
    holes(3, 2) = -m_parXboltHole
    'Set the calculated holes on foundation port
    Call objEqpFoundationPort.SetHoles(holes)
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleEqpFoundationPort8(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleEqpFoundationPort8"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
    m_oEquipCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 8
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleEqpFoundationPort8(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleEqpFoundationPort8"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleEqpFoundationPort8(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleEqpFoundationPort8"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleEqpFoundationPort19
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleEqpFoundationPort19(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleEqpFoundationPort19"
    On Error GoTo ErrorHandler
        
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleFromPH pMemberDescription, pResourceManager, pObject, 9
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleEqpFoundationPort19(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleEqpFoundationPort19"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleEqpFoundationPort19(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleEqpFoundationPort19"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleEqpFoundationPort19(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateNozzleEqpFoundationPort19"
On Error GoTo ErrorHandler
    Dim objEqpFoundationPort As IJEqpFoundationPort
    Dim holes() As Variant
    'Get the Symbol inputs of the Component
    GetDimensionsFromSymbolArray oPropertyDescription.CAO
    'Now Construct the holes Array
    Set objEqpFoundationPort = oPropertyDescription.Object
    Call objEqpFoundationPort.GetHoles(holes())
    holes(0, 1) = (m_parSupportLength - m_parYboltHole) / 2
    holes(0, 2) = m_parXboltHole
    holes(1, 1) = (m_parSupportLength + m_parYboltHole) / 2
    holes(1, 2) = m_parXboltHole
    holes(2, 1) = (m_parSupportLength + m_parYboltHole) / 2
    holes(2, 2) = -m_parXboltHole
    holes(3, 1) = (m_parSupportLength - m_parYboltHole) / 2
    holes(3, 2) = -m_parXboltHole
    'Set the calculated holes on foundation port
    Call objEqpFoundationPort.SetHoles(holes)
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleEqpFoundationPort19(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleEqpFoundationPort19"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
    m_oEquipCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 9
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleEqpFoundationPort19(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleEqpFoundationPort19"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleEqpFoundationPort19(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleEqpFoundationPort19"
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
    
    m_dVesselDiameter = m_avSymbolArrayOfInputs(2)
    m_dVesselTantoTan = m_avSymbolArrayOfInputs(3)
    m_dFirstSupportLocation = m_avSymbolArrayOfInputs(4)
    m_dSecondSupportLocation = m_avSymbolArrayOfInputs(5)
    m_dInsulationThickness = m_avSymbolArrayOfInputs(6)
    m_parSupportLength = m_avSymbolArrayOfInputs(7)
    m_parXboltHole = m_avSymbolArrayOfInputs(16)
    m_parYboltHole = m_avSymbolArrayOfInputs(17)

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
