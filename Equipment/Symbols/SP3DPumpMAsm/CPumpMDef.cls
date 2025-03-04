VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPumpMDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"CAD"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/2/2005-7:28:18 PM"
Attribute VB_Ext_KEY = "SP3DEqpCADTemplateVersion" ,"1.2"
'******************************************************************************
' Copyright (C) 2004, Intergraph Corporation. All rights reserved.
'
'File
'   CPumpMDef.cls
'
'Author
'   BBK
'
'Description
'   Defines the behavior of members of an Equipment that make a custom assembly
'   When a symbol is used with the smart occurrence that contain rigid geometry,
'   This class usually, but is not limited to, manages the nozzles of the equipment.
'Notes
'   From Template Version:1.2
'
'History:
'   02.Jan.2005    V6UpgradeSO     Initial Creation
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'   04.Jun.2005    MS  DI78930  Modified Eqp/EqpComp Symbols to use new methods of Equipment CAD helper
'   16.June.2005   MS TR79730 Some issues with the deletion of members is resolved;made datumshapes/foundationports can be deleted etc
'   24.June.2005   MS TR80576
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SP3DPumpMAsm:CPumpMDef"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const IID_IJDGEOMETRY = "{A1732CBF-5136-11D1-9770-080036754203}"

Private m_oEquipCADHelper As IJEquipCADHelper
Private m_oEditErrors As IJEditErrors

' Implement User Symbol Services(USS)
Implements IJDUserSymbolServices
Implements IJEquipUserAttrMgmt
Private m_avSymbolArrayOfInputs()   As Variant
Private PI                          As Double

Private m_dimpellerWidth As Double
Private m_dimpellerDia As Double
Private m_dshaftWidth As Double
Private m_dshaftDia As Double
Private m_dmotorWidth As Double
Private m_dmotorDia As Double
Private m_dsupportLength As Double
Private m_dsupportHeight As Double
Private m_dsupportThk As Double
Private m_dbasePlateThk As Double

Private m_oNorth                    As IJDVector
Private m_oEast                     As IJDVector
Private m_oElevation                As IJDVector

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
    Set m_oEquipCADHelper = New CADServices
    Set m_oEditErrors = New JServerErrors
    
    m_oEquipCADHelper.ProjectName = "SP3DPumpMAsm"
    m_oEquipCADHelper.ClassName = "CPumpMDef"
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
    'DP2
    'Add new member DP2 to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("DP2", 2, "CMConstructDP2", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsDP2"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructDP2"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalDP2"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseDP2"
    
    'Add properties for DP2
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "DP2Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateDP2", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "DP2GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryDP2", imsCOOKIE_ID_USS_LIB
 
    'Add new member(NozzleSuction1) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleSuction1", 3, "CMConstructNozzleSuction1", imsCOOKIE_ID_USS_LIB)
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
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleDischarge2", 4, "CMConstructNozzleDischarge2", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleDischarge2"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleDischarge2"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleDischarge2"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleDischarge2"
    'Add properties for (NozzleDischarge2)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleDischarge2Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleDischarge2", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleDischarge2GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleDischarge2", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzlePumpFoundationPort3) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzlePumpFoundationPort3", 5, "CMConstructNozzlePumpFoundationPort3", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzlePumpFoundationPort3"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzlePumpFoundationPort3"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzlePumpFoundationPort3"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzlePumpFoundationPort3"
    'Add properties for (NozzlePumpFoundationPort3)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzlePumpFoundationPort3Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzlePumpFoundationPort3", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzlePumpFoundationPort3GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzlePumpFoundationPort3", imsCOOKIE_ID_USS_LIB
    
        
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
        Case "DP1", "DP2", "NozzleSuction1", "NozzleDischarge2", "NozzlePumpFoundationPort3"
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
        Case "DP1", "DP2"
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
        Case "NozzlePumpFoundationPort3"
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
    
    Set oDatumShape = oPropertyDescription.object
    
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

Public Sub CMConstructDP2(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructDP2"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    Dim oDatumShape As IJShape
    Dim oDesignEquipment As IJDesignEquipment
        
    'Create Datum Shape DP2
    Set oDatumShape = m_oEquipCADHelper.CreateShape(pMemberDescription, pResourceManager, "DatumShape 001", "DP2")
    If Not oDatumShape Is Nothing Then
        Set pObject = oDatumShape
        oDatumShape.RepresentationId = ReferenceGeometry

        Set oDesignEquipment = pMemberDescription.CAO
        oDesignEquipment.AddShape oDatumShape

        GetDimensionsFromSymbolArray oDesignEquipment
        PositionAndOrientDP2 oDesignEquipment, oDatumShape
    End If
    
    Set oDesignEquipment = Nothing
    Set oDatumShape = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructDP2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructDP2"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsDP2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsDP2"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateDP2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateDP2"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryDP2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryDP2"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    Dim oEquipment As IJEquipment
'    Dim dEqpX As Double, dEqpY As Double, dEqpZ As Double
    Dim oDatumShape As IJShape
    
    Set oDatumShape = oPropertyDescription.object
    
    'Test if this DatumShape follows the symbol inputs
    If m_oEquipCADHelper.IsShapeFreeToTransform(oDatumShape) = False Then
        Set oEquipment = oPropertyDescription.CAO
        GetDimensionsFromSymbolArray oEquipment
        PositionAndOrientDP2 oEquipment, oDatumShape
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

Public Sub CMConditionalDP2(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalDP2"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = False
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseDP2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseDP2"
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
    oNozzle.Length = 0.175

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Axial
    oOrientation.N1 = -(oNozzle.Length * 0.9)
    oOrientation.N2 = 0
    oOrientation.OR1 = 0

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
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.object)
    
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
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleSuction1(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleSuction1"
    LogCalls METHOD
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
    Dim oPipePort As IJDPipePort
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 2, pResourceManager, DistribPortType_PIPE, pObject, False
    Dim dNozzleLength As Double
    dNozzleLength = m_dimpellerDia * 0.5 + 0.2
    Set oNozzle = pObject
    oNozzle.Length = dNozzleLength
    
    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Tangential
    Set oPipePort = oNozzle
    oOrientation.N1 = m_dimpellerWidth * 0.5
    oOrientation.N2 = m_dimpellerDia * 0.5 + 0.15
    oOrientation.N3 = -(m_dimpellerDia * 0.5 - oPipePort.PipingOutsideDiameter * 0.5)
    oOrientation.OR1 = 3 * PI / 2

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
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(oShape, pMemberDesc.object)
    
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
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleDischarge2(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleDischarge2"
    LogCalls METHOD
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
' Custom Methods for NozzlePumpFoundationPort3
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzlePumpFoundationPort3(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzlePumpFoundationPort3"
    On Error GoTo ErrorHandler
        
    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleFromPH pMemberDescription, pResourceManager, pObject, 3
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzlePumpFoundationPort3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzlePumpFoundationPort3"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzlePumpFoundationPort3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzlePumpFoundationPort3"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzlePumpFoundationPort3(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzlePumpFoundationPort3"
    On Error GoTo ErrorHandler
    
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzlePumpFoundationPort3(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzlePumpFoundationPort3"
    On Error GoTo ErrorHandler
    
    'Transform the nozzle so that it behaves like a rigid body inside the equipment
    m_oEquipCADHelper.TransformNozzleWrtPH oPropertyDescription, pObject, 3
        
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzlePumpFoundationPort3(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzlePumpFoundationPort3"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzlePumpFoundationPort3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzlePumpFoundationPort3"
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
    m_dimpellerWidth = m_avSymbolArrayOfInputs(2)
    m_dimpellerDia = m_avSymbolArrayOfInputs(3)
    m_dshaftWidth = m_avSymbolArrayOfInputs(4)
    m_dshaftDia = m_avSymbolArrayOfInputs(5)
    m_dmotorWidth = m_avSymbolArrayOfInputs(6)
    m_dmotorDia = m_avSymbolArrayOfInputs(7)
    m_dsupportLength = m_avSymbolArrayOfInputs(8)
    m_dsupportHeight = m_avSymbolArrayOfInputs(9)
    m_dsupportThk = m_avSymbolArrayOfInputs(10)
    m_dbasePlateThk = m_avSymbolArrayOfInputs(11)

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
      
    Dim Offset As Double
    Offset = 0.05
    Dim Ymax As Double
    Ymax = Offset + m_dsupportLength + Offset
    
    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    oPosition.Set Offset, -Ymax + Offset + m_dsupportLength * 0.5, m_dbasePlateThk + m_dsupportHeight + m_dimpellerDia * 0.5

    'We want:
    'the X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'the Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oEast, m_oNorth
    
    Set oPosition = Nothing

End Sub
Private Sub PositionAndOrientDP2(Equipment As IJEquipment, Shape As IJShape)
    Dim oPosition As IJDPosition
    Set oPosition = New DPosition
    oPosition.Set 0, 0, 0

    'We want:
    'the X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'the Y (secondary) of the shape on the Y or North of the equipment (ECS)
    m_oEquipCADHelper.PositionAndOrientShape Equipment, Shape, oPosition, m_oElevation, m_oNorth
    
    Set oPosition = Nothing


End Sub

