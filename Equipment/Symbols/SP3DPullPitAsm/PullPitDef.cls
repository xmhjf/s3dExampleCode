VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PullPitDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"CAD"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/2/2005-7:23:31 PM"
Attribute VB_Ext_KEY = "SP3DEqpCADTemplateVersion" ,"1.2"
'******************************************************************************
' Copyright (C) 2008, Intergraph Corporation. All rights reserved.
'
'   File:PullPitDef.cls
'   Author:RH
'
'Description
'   Defines the behavior of members of an Equipment that make a custom assembly
'   When a symbol is used with the smart occurrence that contain rigid geometry,
'   This class usually, but is not limited to, manages the ports of the equipment.
'   This Assembly has four Cable Tray Ports
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   13.June.2008     VRK     CR-134560:Provide pull-pit/manhole equipment symbol for use with duct banks
'   20.Mar.2009      Haneef  DI-CP-156623  Equipment: Resolve Compiler Warnings for VC++ and .net Projects
'   14.May.2009      JMS     TR#164799 - Previous changes introduced a bug where the variable oReOrderedShapesColl
'                               was being accessed outside the scope of its declaration
'
'+++-++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SP3DPullPitAsm:PullPitDef"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const IID_IJDGEOMETRY = "{A1732CBF-5136-11D1-9770-080036754203}"
Private Const IJControlPoint = "{F54DCDA8-2C24-47B9-A4ED-ACA4F3DF15D7}"

Private m_oEquipCADHelper As IJEquipCADHelper
Private m_oEditErrors As IJEditErrors

Private m_avSymbolArrayOfInputs() As Variant
Private m_PullBoxHeight As Double
Private m_PullBoxWidth As Double
Private m_PullBoxLength As Double
Private m_PullBoxThickness As Double

Private m_PullBoxHole1Height As Double
Private m_PullBoxHole1Width As Double
Private m_PullBoxHole1Offset As Double

Private m_PullBoxHole2Height As Double
Private m_PullBoxHole2Width As Double
Private m_PullBoxHole2Offset As Double

Private m_bPullBoxHole1ThruStart As Boolean
Private m_bPullBoxHole1ThruEnd As Boolean
Private m_bPullBoxHole2ThruStart As Boolean
Private m_bPullBoxHole2ThruEnd As Boolean

Private PI As Double

' Implement User Symbol Services(USS)
Implements IJDUserSymbolServices
Implements IJEquipUserAttrMgmt

Private Sub Class_Initialize()
    Set m_oEquipCADHelper = New CADServices
    Set m_oEditErrors = New JServerErrors

    PI = 4 * Atn(1)
    
    m_oEquipCADHelper.ProjectName = "SP3DPullPitAsm"
    m_oEquipCADHelper.CLASSNAME = "PullPitDef"
    m_oEquipCADHelper.OccurrenceRootClass = orcEquipment
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
    IJDUserSymbolServices_GetDefinitionName = m_oEquipCADHelper.ProjectName & "." & m_oEquipCADHelper.CLASSNAME
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

    'PullPit
    Set oMemberDescriptions = oSymbolDefinition
    oMemberDescriptions.RemoveAll
    Set oMemberDescription = oMemberDescriptions.AddMember("PullBoxSolid", 1, "CMConstructPullBoxSolid", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsPullBoxSolid"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructPullBoxSolid"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalPullBoxSolid"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleasePullBoxSolid"
    Set oPropertyDescriptions = oMemberDescription
    'output watching IJDAttributes
    oPropertyDescriptions.AddProperty "PullBoxGeometry", 1, IID_IJDGEOMETRY, "CMEvaluatePullBoxGeometry", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "PullBoxProperties", 2, IID_IJDATTRIBUTES, "CMEvaluatePullBox", imsCOOKIE_ID_USS_LIB

    'Port1
    'Add new member(NozzleCableTrayPort1) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleCableTrayPort1", 2, "CMConstructNozzleCableTrayPort1", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleCableTrayPort1"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleCableTrayPort1"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleCableTrayPort1"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleCableTrayPort1"
    'Add properties for (NozzleCableTrayPort1)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleCableTrayPort1Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleCableTrayPort1", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleCableTrayPort1GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleCableTrayPort1", imsCOOKIE_ID_USS_LIB

    'Port 2
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleCableTrayPort2", 3, "CMConstructNozzleCableTrayPort2", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleCableTrayPort2"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleCableTrayPort2"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleCableTrayPort2"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleCableTrayPort2"
    'Add properties for (NozzleCableTrayPort2)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleCableTrayPort2Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleCableTrayPort2", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleCableTrayPort2GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleCableTrayPort2", imsCOOKIE_ID_USS_LIB

    'Port 3
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleCableTrayPort3", 4, "CMConstructNozzleCableTrayPort3", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleCableTrayPort3"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleCableTrayPort3"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleCableTrayPort3"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleCableTrayPort3"
    'Add properties for (NozzleCableTrayPort3)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleCableTrayPort3Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleCableTrayPort3", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleCableTrayPort3GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleCableTrayPort3", imsCOOKIE_ID_USS_LIB

    'Port 4
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleCableTrayPort4", 5, "CMConstructNozzleCableTrayPort4", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleCableTrayPort4"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleCableTrayPort4"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleCableTrayPort4"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleCableTrayPort4"
    'Add properties for (NozzleCableTrayPort4)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleCableTrayPort4Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleCableTrayPort4", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleCableTrayPort4GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleCableTrayPort4", imsCOOKIE_ID_USS_LIB
    '
    'Add new member(ControlPoint) to the definition

    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("ControlPoint", 6, "CMConstructControlPoint", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputControlPoint"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructControlPoint"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalControlPoint"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseControlPoint"

    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "ControlPointProperties", 1, IID_IJDATTRIBUTES, "CMComputeControlPoints", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "ControlPointGeometry", 2, IID_IJDGEOMETRY, "CMEvaluateControlPointGeometry", imsCOOKIE_ID_USS_LIB


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
    Select Case oMemberDescription.name
        Case "NozzleCableTrayPort1", "NozzleCableTrayPort2", "NozzleCableTrayPort3", "NozzleCableTrayPort4"
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
    Select Case oMemberDescription.name
    Case "NozzleCableTrayPort1", "NozzleCableTrayPort2", "NozzleCableTrayPort3", "NozzleCableTrayPort4"

        For m = 1 To oAttrCollection.Count
            Set oAttributeDescriptor = oAttrCollection.Item(m)
            Select Case UCase(oAttributeDescriptor.InterfaceName)
            Case "IJDELETABLEMEMBER"
                If UCase(oAttributeDescriptor.AttrName) = "CANBEDELETED" Then
                    oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsChanged
                End If
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
' Custom Methods for NozzleCableTrayPort1
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleCableTrayPort1(ByVal pMemberDescription As IJDMemberDescription, _
                                           ByVal pResourceManager As IUnknown, _
                                           ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleCableTrayPort1"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJCableTrayPortOcc
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    Dim oRadialOrient As New DVector
    oVec.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1

    GetDimensionsFromSymbolArray pMemberDescription.CAO

    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 1, pResourceManager, DistribPortType_CABLETRAY, pObject, False
    Set oNozzle = pObject

    Dim oMat As IJDT4x4
    Dim oNzlMatrixAccess As IJDMatrixAccess
    Set oNzlMatrixAccess = oNozzle
    Set oMat = oNzlMatrixAccess.Matrix
    oMat.LoadIdentity
    oMat.IndexValue(12) = m_PullBoxThickness
    oMat.IndexValue(13) = 0
    oMat.IndexValue(14) = -(m_PullBoxHeight / 2 - m_PullBoxHole2Height / 2) + m_PullBoxHole2Offset
    oNzlMatrixAccess.Matrix = oMat
    Set oNzlMatrixAccess = Nothing
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    oDistribPort.SetRadialOrient oRadialOrient

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Set oDistribPort = Nothing
    Set oVec = Nothing
    Set oMat = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleCableTrayPort1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleCableTrayPort1"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleCableTrayPort1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleCableTrayPort1"
    On Error GoTo ErrorHandler
    
    Dim oOrientation As IJNozzleOrientation
    'Establish the relationship with the equipment origin if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations( _
            pMemberDesc.CAO, pMemberDesc.Object)
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleCableTrayPort1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleCableTrayPort1"
    On Error GoTo ErrorHandler

    Dim oNozzle As IJCableTrayPortOcc
    Set oNozzle = pObject

    Dim oEquipment As IJEquipment
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment

    oNozzle.ActualDepth = m_PullBoxHole2Height
    oNozzle.ActualWidth = m_PullBoxHole2Width
    oNozzle.NominalDepth = m_PullBoxHole2Height
    oNozzle.NominalWidth = m_PullBoxHole2Width

    Set oNozzle = Nothing
    Set oEquipment = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleCableTrayPort1(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleCableTrayPort1"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJCableTrayPortOcc
    Dim oEquipment As IJEquipment
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    Dim oRadialOrient As New DVector
    oVec.Set -1, 0, 0
    oRadialOrient.Set 0, 0, 1

    GetDimensionsFromSymbolArray oPropertyDescription.CAO
    Set oNozzle = oPropertyDescription.Object
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    oDistribPort.SetRadialOrient oRadialOrient
    Set oEquipment = oPropertyDescription.CAO

    TransformAndRotateCTNozzle oEquipment, pObject, PI / 2, 0, 0, m_PullBoxThickness, _
                0, -(m_PullBoxHeight / 2 - m_PullBoxHole2Height / 2) + m_PullBoxHole2Offset

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Set oDistribPort = Nothing
    Set oVec = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleCableTrayPort1(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleCableTrayPort1"
    On Error GoTo ErrorHandler

    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleCableTrayPort1(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleCableTrayPort1"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleCableTrayPort2
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleCableTrayPort2(ByVal pMemberDescription As IJDMemberDescription, _
                                           ByVal pResourceManager As IUnknown, _
                                           ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleCableTrayPort2"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJCableTrayPortOcc
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    Dim oRadialOrient As New DVector
    oVec.Set 0, -1, 0
    oRadialOrient.Set 0, 0, 1
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO
    
    'Create nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 2, pResourceManager, DistribPortType_CABLETRAY, pObject, False
    Set oNozzle = pObject
    
    Dim oMat As IJDT4x4
    Dim oNzlMatrixAccess As IJDMatrixAccess
    Set oNzlMatrixAccess = oNozzle
    Set oMat = oNzlMatrixAccess.Matrix
    oMat.LoadIdentity
    oMat.IndexValue(12) = m_PullBoxLength / 2
    oMat.IndexValue(13) = -m_PullBoxWidth / 2 + m_PullBoxThickness
    oMat.IndexValue(14) = -(m_PullBoxHeight / 2 - m_PullBoxHole1Height / 2) + m_PullBoxHole1Offset
    oNzlMatrixAccess.Matrix = oMat
    Set oNzlMatrixAccess = Nothing
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    oDistribPort.SetRadialOrient oRadialOrient
    
    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Set oDistribPort = Nothing
    Set oVec = Nothing
    Set oMat = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleCableTrayPort2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleCableTrayPort2"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleCableTrayPort2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleCableTrayPort2"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    'Establish the relationship with the equipment origin if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations( _
            pMemberDesc.CAO, pMemberDesc.Object)
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleCableTrayPort2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleCableTrayPort2"
    On Error GoTo ErrorHandler

    Dim oNozzle As IJCableTrayPortOcc
    Set oNozzle = pObject

    Dim oEquipment As IJEquipment
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment

    oNozzle.ActualDepth = m_PullBoxHole1Height
    oNozzle.ActualWidth = m_PullBoxHole1Width
    oNozzle.NominalDepth = m_PullBoxHole1Height
    oNozzle.NominalWidth = m_PullBoxHole1Width
    Set oNozzle = Nothing
    Set oEquipment = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleCableTrayPort2(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleCableTrayPort2"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oEquipment As IJEquipment
    Dim oNozzle As IJCableTrayPortOcc
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    Dim oRadialOrient As New DVector
    oVec.Set 0, -1, 0
    oRadialOrient.Set 0, 0, 1
    
    GetDimensionsFromSymbolArray oPropertyDescription.CAO
    Set oNozzle = oPropertyDescription.Object
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    oDistribPort.SetRadialOrient oRadialOrient
    Set oEquipment = oPropertyDescription.CAO
    
    TransformAndRotateCTNozzle oEquipment, pObject, 0, PI / 2, PI / 2, m_PullBoxLength / 2, _
                     -m_PullBoxWidth / 2 + m_PullBoxThickness, _
                      -(m_PullBoxHeight / 2 - m_PullBoxHole1Height / 2) + m_PullBoxHole1Offset
    
    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Set oDistribPort = Nothing
    Set oVec = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleCableTrayPort2(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleCableTrayPort2"
    On Error GoTo ErrorHandler

    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleCableTrayPort2(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleCableTrayPort2"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'port 3
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleCableTrayPort3
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleCableTrayPort3(ByVal pMemberDescription As IJDMemberDescription, _
                                           ByVal pResourceManager As IUnknown, _
                                           ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleCableTrayPort3"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJCableTrayPortOcc
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    Dim oRadialOrient As New DVector
    oVec.Set 1, 0, 0
    oRadialOrient.Set 0, 0, 1

    GetDimensionsFromSymbolArray pMemberDescription.CAO

    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 3, pResourceManager, DistribPortType_CABLETRAY, pObject, False
    Set oNozzle = pObject

    Dim oMat As IJDT4x4
    Dim oNzlMatrixAccess As IJDMatrixAccess
    Set oNzlMatrixAccess = oNozzle
    Set oMat = oNzlMatrixAccess.Matrix
    oMat.LoadIdentity
    oMat.IndexValue(12) = m_PullBoxLength - m_PullBoxThickness
    oMat.IndexValue(13) = 0
    oMat.IndexValue(14) = -(m_PullBoxHeight / 2 - m_PullBoxHole2Height / 2) + m_PullBoxHole2Offset
    oNzlMatrixAccess.Matrix = oMat
    Set oNzlMatrixAccess = Nothing
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    oDistribPort.SetRadialOrient oRadialOrient

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Set oDistribPort = Nothing
    Set oVec = Nothing
    Set oMat = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleCableTrayPort3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleCableTrayPort3"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleCableTrayPort3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleCableTrayPort3"
    On Error GoTo ErrorHandler
    
    Dim oOrientation As IJNozzleOrientation
    'Establish the relationship with the equipment origin if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations( _
         pMemberDesc.CAO, pMemberDesc.Object)
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleCableTrayPort3(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleCableTrayPort3"
    On Error GoTo ErrorHandler

    Dim oNozzle As IJCableTrayPortOcc
    Set oNozzle = pObject

    Dim oEquipment As IJEquipment
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment

    oNozzle.ActualDepth = m_PullBoxHole2Height
    oNozzle.ActualWidth = m_PullBoxHole2Width
    oNozzle.NominalDepth = m_PullBoxHole2Height
    oNozzle.NominalWidth = m_PullBoxHole2Width

    Set oNozzle = Nothing
    Set oEquipment = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleCableTrayPort3(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleCableTrayPort3"
    On Error GoTo ErrorHandler
    
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJCableTrayPortOcc
    Dim oEquipment As IJEquipment
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    Dim oRadialOrient As New DVector
    oVec.Set 1, 0, 0
    oRadialOrient.Set 0, 0, 1

    GetDimensionsFromSymbolArray oPropertyDescription.CAO
    Set oNozzle = oPropertyDescription.Object
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    oDistribPort.SetRadialOrient oRadialOrient
    Set oEquipment = oPropertyDescription.CAO

    TransformAndRotateCTNozzle oEquipment, pObject, PI / 2, PI, 0, m_PullBoxLength - m_PullBoxThickness, _
         0, -(m_PullBoxHeight / 2 - m_PullBoxHole2Height / 2) + m_PullBoxHole2Offset

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Set oDistribPort = Nothing
    Set oVec = Nothing
     
     Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleCableTrayPort3(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleCableTrayPort3"
    On Error GoTo ErrorHandler

    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleCableTrayPort3(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleCableTrayPort3"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'port 4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleCableTrayPort4
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleCableTrayPort4(ByVal pMemberDescription As IJDMemberDescription, _
                                           ByVal pResourceManager As IUnknown, _
                                           ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleCableTrayPort4"
    On Error GoTo ErrorHandler

     Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJCableTrayPortOcc
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    Dim oRadialOrient As New DVector
    oVec.Set 0, 1, 0
    oRadialOrient.Set 0, 0, 1
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO

    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 4, pResourceManager, DistribPortType_CABLETRAY, pObject, False
    Set oNozzle = pObject
    
    Dim oMat As IJDT4x4
    Dim oNzlMatrixAccess As IJDMatrixAccess
    Set oNzlMatrixAccess = oNozzle
    Set oMat = oNzlMatrixAccess.Matrix
    oMat.LoadIdentity
    oMat.IndexValue(12) = m_PullBoxLength / 2
    oMat.IndexValue(13) = m_PullBoxWidth / 2 - m_PullBoxThickness
    oMat.IndexValue(14) = -(m_PullBoxHeight / 2 - m_PullBoxHole1Height / 2) + m_PullBoxHole1Offset
    oNzlMatrixAccess.Matrix = oMat
    Set oNzlMatrixAccess = Nothing
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    oDistribPort.SetRadialOrient oRadialOrient
    
    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Set oDistribPort = Nothing
    Set oVec = Nothing
    Set oMat = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleCableTrayPort4(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleCableTrayPort4"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleCableTrayPort4(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleCableTrayPort4"
    On Error GoTo ErrorHandler
    
    Dim oOrientation As IJNozzleOrientation
    'Establish the relationship with the equipment origin if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations( _
            pMemberDesc.CAO, pMemberDesc.Object)
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleCableTrayPort4(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleCableTrayPort4"
    On Error GoTo ErrorHandler

    Dim oNozzle As IJCableTrayPortOcc
    Set oNozzle = pObject

    Dim oEquipment As IJEquipment
    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment

    oNozzle.ActualDepth = m_PullBoxHole1Height
    oNozzle.ActualWidth = m_PullBoxHole1Width
    oNozzle.NominalDepth = m_PullBoxHole1Height
    oNozzle.NominalWidth = m_PullBoxHole1Width

    Set oNozzle = Nothing
    Set oEquipment = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleCableTrayPort4(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleCableTrayPort4"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJCableTrayPortOcc
    Dim oEquipment As IJEquipment
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    Dim oRadialOrient As New DVector
    oVec.Set 0, 1, 0
    oRadialOrient.Set 0, 0, 1
  
    GetDimensionsFromSymbolArray oPropertyDescription.CAO
    Set oNozzle = oPropertyDescription.Object
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    oDistribPort.SetRadialOrient oRadialOrient
    Set oEquipment = oPropertyDescription.CAO
    
    TransformAndRotateCTNozzle oEquipment, pObject, 0, -PI / 2, -PI / 2, m_PullBoxLength / 2, _
                        m_PullBoxWidth / 2 - m_PullBoxThickness, _
               -(m_PullBoxHeight / 2 - m_PullBoxHole1Height / 2) + m_PullBoxHole1Offset

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Set oDistribPort = Nothing
    Set oVec = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleCableTrayPort4(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleCableTrayPort4"
    On Error GoTo ErrorHandler

    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleCableTrayPort4(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleCableTrayPort4"
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for PullBoxSolid
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructPullBoxSolid(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    Const METHOD = "CMConstructPullBoxSolid"
    On Error GoTo ErrorHandler

    Dim oDesignEquipment As IJDesignEquipment
    Dim oDesignSolid As IJDesignSolid
    Dim oDesignSolidFactory As New GSCADEqpEntities.DesignSolidFactory
    Dim oNamedItem As IJNamedItem

    ' Get the parent Designed Equipment
    Set oDesignEquipment = pMemberDescription.CAO

    'Create Solid
    Set oDesignSolid = oDesignSolidFactory.CreateDesignSolid(pResourceManager, oDesignEquipment)

    ' Add the solid as a shape of the designed equipment (i.e., solids behave a shapes so they
    '   are added to the equipment as shpaes)
    oDesignEquipment.AddShape oDesignSolid

    ' Provide the solid with a name

    Set oNamedItem = oDesignSolid
    oNamedItem.name = "CablePit - Solid"

    ' Return the solid as a custom assembly member
    Set pObject = oDesignSolid

    Set oDesignEquipment = Nothing
    Set oDesignSolid = Nothing
    Set oNamedItem = Nothing
    Set oDesignSolidFactory = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsPullBoxSolid(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsPullBoxSolid"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalPullBoxSolid(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalPullBoxSolid"
    On Error GoTo ErrorHandler

    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructPullBoxSolid(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructPullBoxSolid"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleasePullBoxSolid(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleasePullBoxSolid"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluatePullBox(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluatePullBox"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluatePullBoxGeometry(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluatePullBoxGeometry"
    On Error GoTo ErrorHandler

    ' These parameters will be computed from the attributes of the pull box
    Dim dPullBoxHeight As Double
    Dim dPullBoxWidth As Double
    Dim dPullBoxLength As Double
    Dim dPullBoxThickness As Double
    Dim dPullBoxHole1Height As Double
    Dim dPullBoxHole1Width As Double
    Dim dPullBoxHole1VerticalOffset As Double
    Dim dPullBoxHole2Height As Double
    Dim dPullBoxHole2Width As Double
    Dim dPullBoxHole2VerticalOffset As Double
    '
    '              Side - 3
    '           ------------
    '           |          |
    '  Side - 2 |          | Side - 4
    '           |          |
    '           ------------
    '              Side - 1
    '
    Dim bPullBoxHoleSide1 As Boolean
    Dim bPullBoxHoleSide2 As Boolean
    Dim bPullBoxHoleSide3 As Boolean
    Dim bPullBoxHoleSide4 As Boolean
'
'    ' Default the parameters (which later will be overridden by the equipment
'    '   attributes)
'
'    dPullBoxHeight = 10#
'    dPullBoxWidth = 12#
'    dPullBoxLength = 18#
'    dPullBoxThickness = 1#
'    dPullBoxHole1Height = 2#
'    dPullBoxHole1Width = 3#
'    dPullBoxHole1VerticalOffset = 2#
'    dPullBoxHole2Height = 3#
'    dPullBoxHole2Width = 4#
'    dPullBoxHole2VerticalOffset = 2#
'    bPullBoxHoleSide1 = True
'    bPullBoxHoleSide2 = False
'    bPullBoxHoleSide3 = True
'    bPullBoxHoleSide4 = True

    Dim numShapes As Long
    Dim oDesignSolid As IJDesignSolid
    Dim oExistingShapesColl As IJDObjectCollection
    Dim oShapes(1 To 4) As Object
    Dim oSmartOccurrence As IJSmartOccurrence
    
    Set oExistingShapesColl = New JObjectCollection
    ' Get the smart occurrence
    Set oSmartOccurrence = oPropertyDescription.CAO

    ' Get the solid
    Set oDesignSolid = oPropertyDescription.Object

    ' Get the operators (i.e. shapes) associated to the solid. If the shapes already have been
    '   added to this solid then this will contain the array of shapes
    oDesignSolid.GetOrderedOperators oExistingShapesColl

    numShapes = 0
    If Not IsEmpty(oExistingShapesColl) Then
        numShapes = oExistingShapesColl.Count
    End If

    Dim oShapePart As IJDPart
    Dim oShape As IJShape
    Dim oShapeFactory As New GSCADEqpEntities.ShapeFactory
    Dim oCatalogResourceMgr As Object
    Dim oIJDObject As IJDObject
    Dim oResourceManager As Object

    ' Get the catalog manager from the smart occurrence
    Set oCatalogResourceMgr = oSmartOccurrence.CatalogResourceMgr

    ' Get the resource manager from the design solid
    Set oIJDObject = oDesignSolid
    Set oResourceManager = oIJDObject.ResourceManager

    ' Get the Outside box shape and if it does not exist create it.
    '   (The GetOperatorByName will access the shape assigned to the solid by an
    '   internal name, which you can assign to the shape operator)
    oDesignSolid.GetOperatorByName "Box", oShapes(1)
    If oShapes(1) Is Nothing Then   ' Shape does not exist yet
        ' Get a rectangular shape block from the catalog
        Set oShapePart = GetPartFromPartNumber(oCatalogResourceMgr, "RectangularSolid 001", oResourceManager)
        ' Create the rectangular shape underneath the solid
        '   (This shape factory method is not new but it does more than just add the shape
        '   to the solid. It also, adds the shape as the last shape operator because order
        '   is important and it defaults the operator as an operator that adds material)
        Set oShape = oShapeFactory.CreateShape2(oShapePart, oResourceManager, oDesignSolid)
        ' oShape.RepresentationId = SimplePhysical
        Set oShapes(1) = oShape
        ' We will name this added shape so we can retrieve it later. This is not
        '   a required step but will help us keep track of the shapes we require to exist
        '   on the solid.)
        oDesignSolid.SetOperatorName oShape, "Box"  ' Give the shape an internal (we will access it by this name later)
        '"Outside Box" ' Give the shape a name
        SetObjNameRule oShape, "CPShape"

    End If

    ' Get the Inside box shape and if it does not exist create it
    '   (The GetOperatorByName is accessing the shape by it's internal name called "InsideBox",
    '   which was assigned to the shape operator below on the call to SetOperatorName)
    oDesignSolid.GetOperatorByName "InsideBox", oShapes(2)
    If oShapes(2) Is Nothing Then   ' Shape does not exist yet
        ' Get a rectangular shape block from the catalog
        Set oShapePart = GetPartFromPartNumber(oCatalogResourceMgr, "RectangularSolid 001", oResourceManager)
        ' Create another rectangular shape underneath the solid
        '   (This shape factory method adds the shape after the previously added outside box as
        '   an operator that adds material. This is not what we want so we will change it below)
        Set oShape = oShapeFactory.CreateShape2(oShapePart, oResourceManager, oDesignSolid)
        Set oShapes(2) = oShape
        ' Because the default operator type is "add" the inside box needs to subtract material;
        '   therefore, we must change this operator to a subtract operator
        oDesignSolid.SetOperatorType oShape, BooleanOperatorSubtract
        ' Again we will name this added shape so we can retrieve it later. This is not
        '   a required step but will help us keep track of the shapes we require to exist
        '   on the solid.)
        oDesignSolid.SetOperatorName oShape, "InsideBox"
        '"Inside Box"
        SetObjNameRule oShape, "CPShape"

    End If

    ' Get the first hole shape and if it does not exist create it
    '   (The GetOperatorByName is accessing the shape by it's internal name called "Hole1",
    '   which was assigned to the shape operator below)
    oDesignSolid.GetOperatorByName "Hole1", oShapes(3)
    If oShapes(3) Is Nothing Then   ' Shape does not exist yet
        ' Get a rectangular shape block from the catalog
        Set oShapePart = GetPartFromPartNumber(oCatalogResourceMgr, "RectangularSolid 001", oResourceManager)
        ' Create another rectangular shape underneath the solid
        '   (This shape factory method adds the shape after the previously added inside box
        '   operator. Similarly, this shape adds material so we will change it to subtract below).
        '   Also note that this shape operator could have been added before the Inside of the box operator
        '   (i.e., the order of these two operators does not matter).
        Set oShape = oShapeFactory.CreateShape2(oShapePart, oResourceManager, oDesignSolid)
        Set oShapes(3) = oShape
        ' Because the default operator type is "add" the hole that penetrates through two sides
        '   of the box needs to subtract material (i.e., change to a Subtract Operator)
        oDesignSolid.SetOperatorType oShape, BooleanOperatorSubtract
        ' Again we will name this added shape so we can retrieve it later. This is not
        '   a required step but will help us keep track of which shape is the first hole.)
        oDesignSolid.SetOperatorName oShape, "Hole1"
        '"Duct Hole #1"
        SetObjNameRule oShape, "CPShape"

    End If

    ' Get the second hole shape and if it does not exist create it
    '   (The GetOperatorByName is accessing the shape by it's internal name called "Hole2",
    '   which was assigned to the shape operator below)
    oDesignSolid.GetOperatorByName "Hole2", oShapes(4)
    If oShapes(4) Is Nothing Then   ' Shape does not exist yet
        Set oShapePart = GetPartFromPartNumber(oCatalogResourceMgr, "RectangularSolid 001", oResourceManager)
        ' Create another rectangular shape underneath the solid
        '   (This shape factory method adds the shape after the previously added first hole
        '   operator. Similarly, this shape adds material so we will change it to subtract below).
        '   Also note that this shape operator could have been added before the Inside of the box operator
        '   (i.e., the order of these last three operators does not matter).
        Set oShape = oShapeFactory.CreateShape2(oShapePart, oResourceManager, oDesignSolid)
        Set oShapes(4) = oShape
        ' Because the default operator type is "add" the hole that penetrates through two sides
        '   of the box needs to subtract material (i.e., change to a Subtract Operator)
        oDesignSolid.SetOperatorType oShape, BooleanOperatorSubtract
        ' Again we will name this added shape so we can retrieve it later. This is not
        '   a required step but will help us keep track of which shape is the 2nd hole.)
        oDesignSolid.SetOperatorName oShape, "Hole2"
        '"Duct Hole #2"
        SetObjNameRule oShape, "CPShape"

    End If

    ' Earlier we accessed the solid to get the set of assigned shape operators via
    '   the call to GetOrderedOperators. If this list did not exist yet (i.e., the
    '   shapes were not added yet), then we don't need to do anything. However, if
    '   the operators did exist, we can make sure the user did not mess with ordering of
    '   these shapes.
    If numShapes > 0 Then
        ' If the first shape is not the outside box shape then we need to re-order the shapes.
        '  oShapes(1) should be the first shape when we called GetOrderedOperators and it will be the
        '  first shape that we constructed above (which we assigned to oShapes(1)).
        Set oShape = oExistingShapesColl.Item(CStr(1))
        If Not oShapes(1) Is oShape Then
            Dim oReOrderedShapesColl As IJDObjectCollection
            Dim oShapeUnknown As Object
            Dim i As Long, j As Long
            
            Set oReOrderedShapesColl = New JObjectCollection
            ' The following simply shifts the other operators in the array such
            ' that they follow the outside box operator.
            oReOrderedShapesColl.Add oShape(1), CStr(1)
            j = 2
            For i = 1 To numShapes
                ' Shape is not the outside box shape
                Set oShapeUnknown = oExistingShapesColl.Item(CStr(i))
                If Not oShapeUnknown Is oShapes(1) Then
                    oReOrderedShapesColl.Add oShapeUnknown, CStr(j)
                    j = j + 1
                End If
                Set oShapeUnknown = Nothing
            Next i
            ' The following call on the solid will reset the operators so they are order the way
            '   we want (i.e., the outside box operator 1st).
            oDesignSolid.SetOrderedOperators oReOrderedShapesColl
            
            oReOrderedShapesColl.Clear
            Set oReOrderedShapesColl = Nothing
        End If
        Set oShape = Nothing
    End If

    ' Above completes the solid portion of the problem. If you followed the code carefully
    '   you will note that the evaluate is called upon a compute of the equipment and the shapes
    '   that you've named will get automatically re-created if the user deletes it. This probably
    '   is not what is wanted but there currently is nothing to prevent the shapes from being
    '   removed at this time. Also, the evaluate does not get called when a user deletes a shape.
    '   It is not until an attribute of the equipment/solid is modified before this evaluate occurs
    '   which regenerates the shape again. One could simply move the construction of the shape code
    '   to the Construct method and if the user deletes one of the shapes then they will simply lose
    '   that portion of their pull-pit.

    Dim dHole1Length As Double
    Dim dHole1Offset As Double
    Dim dHole2Length As Double
    Dim dHole2Offset As Double
    Dim xPos As Double, yPos As Double, zPos As Double
    Dim oLocalCoor As IJLocalCoordinateSystem
    Dim oAttributes As IJDAttributes
    Dim oAttributesCol As IJDAttributesCol

    ' Retrieve the attributes from the equipment
    '   You do not have to do this this way. Some of these values could have been controlled by the shape
    '   and then this evaluate only concerns itself with positioning of the shape. I just chose
    '   to do it this way top see if it would work.

    Set oAttributes = oSmartOccurrence
    Set oAttributesCol = oAttributes.CollectionOfAttributes("IJUAPullBoxDimensions")
    If Not oAttributesCol Is Nothing Then
        dPullBoxHeight = oAttributesCol.Item("PullBoxHeight").Value
        dPullBoxWidth = oAttributesCol.Item("PullBoxWidth").Value
        dPullBoxLength = oAttributesCol.Item("PullBoxLength").Value
        dPullBoxThickness = oAttributesCol.Item("PullBoxThickness").Value
        dPullBoxHole1Height = oAttributesCol.Item("PullBoxHole1Height").Value
        dPullBoxHole1Width = oAttributesCol.Item("PullBoxHole1Width").Value
        dPullBoxHole1VerticalOffset = oAttributesCol.Item("PullBoxHole1Offset").Value

        bPullBoxHoleSide1 = oAttributesCol.Item("PullBoxHole1ThruStart").Value
        bPullBoxHoleSide3 = oAttributesCol.Item("PullBoxHole1ThruEnd").Value
        dPullBoxHole2Height = oAttributesCol.Item("PullBoxHole2Height").Value
        dPullBoxHole2Width = oAttributesCol.Item("PullBoxHole2Width").Value
        dPullBoxHole2VerticalOffset = oAttributesCol.Item("PullBoxHole2Offset").Value
        bPullBoxHoleSide2 = oAttributesCol.Item("PullBoxHole2ThruStart").Value
        bPullBoxHoleSide4 = oAttributesCol.Item("PullBoxHole2ThruEnd").Value
    End If

    '
    ' This next block of code is simply code to size and position the shapes relative
    '   to one another. It actually has nothing to do with the solid other than the
    '   result of the size and position will be the final solid
    '

    ' Get the position of design solid because the shapes will be placed
    ' relative to its origin
    Set oLocalCoor = oDesignSolid
    xPos = oLocalCoor.Position.x
    yPos = oLocalCoor.Position.y
    zPos = oLocalCoor.Position.z

    ' Set the position of the first shape
    Set oShape = oShapes(1)
    oShape.SetOrigin xPos, yPos, zPos    ' Position the shape at the origin of the solid
   
    ' Change attributes of the first shape
    oShape.GetPart oShapePart
    Set oAttributes = oShape
    Set oAttributesCol = oAttributes.CollectionOfAttributes(oShapePart.GetPartOccIID)
    If Not oAttributesCol Is Nothing Then
        oAttributesCol.Item("A").Value = dPullBoxLength
        oAttributesCol.Item("B").Value = dPullBoxWidth
        oAttributesCol.Item("C").Value = dPullBoxHeight
    End If

    ' Set the position of the second shape
    Set oShape = oShapes(2)
    oShape.SetOrigin xPos + dPullBoxThickness, yPos, zPos + dPullBoxThickness / 2

    ' Change attributes of the second shape (inside of the box)
    oShape.GetPart oShapePart
    Set oAttributes = oShape
    Set oAttributesCol = oAttributes.CollectionOfAttributes(oShapePart.GetPartOccIID)
    If Not oAttributesCol Is Nothing Then
        oAttributesCol.Item("A").Value = dPullBoxLength - 2 * dPullBoxThickness
        oAttributesCol.Item("B").Value = dPullBoxWidth - 2 * dPullBoxThickness
        oAttributesCol.Item("C").Value = dPullBoxHeight - dPullBoxThickness

    End If
    
    Dim oEquipment As IJEquipment
    Set oEquipment = oPropertyDescription.CAO

    TransformAndRotate oEquipment, oShape, 0, 0, 0, _
                                dPullBoxThickness, 0, dPullBoxThickness / 2
                                
                                
    ' Set the position of the third shape (Hole 1)
    Set oShape = oShapes(3)
    dHole1Offset = 0
    dHole1Length = dPullBoxWidth
    ' if both bPullBoxHoleSide1 and bPullBoxHoleSide3 are false then we could suppress the shape
    ' via oDesignSolid.SetOperatorType oShape, BooleanOperatorSuppress
    If Not bPullBoxHoleSide1 Then
        dHole1Offset = -dPullBoxThickness / 2
        dHole1Length = dHole1Length - dPullBoxThickness
    End If
    If Not bPullBoxHoleSide3 Then
        dHole1Offset = dPullBoxThickness / 2
        dHole1Length = dHole1Length - dPullBoxThickness
    End If

    If (Not bPullBoxHoleSide3) And (Not bPullBoxHoleSide1) Then
        oShape.SetOrigin xPos + (dPullBoxLength - dPullBoxHole1Width) / 2, yPos, zPos - (dPullBoxHeight / 2# - dPullBoxHole1Height / 2) + dPullBoxHole1VerticalOffset
    Else
        oShape.SetOrigin xPos + (dPullBoxLength - dPullBoxHole1Width) / 2, yPos + dHole1Offset, zPos - (dPullBoxHeight / 2# - dPullBoxHole1Height / 2) + dPullBoxHole1VerticalOffset
    End If
    
    If (Not bPullBoxHoleSide3) And (Not bPullBoxHoleSide1) Then
        TransformAndRotate oEquipment, oShape, 0, 0, 0, _
                                (dPullBoxLength - dPullBoxHole1Width) / 2, 0, -(dPullBoxHeight / 2# - dPullBoxHole1Height / 2) + dPullBoxHole1VerticalOffset
    Else
        TransformAndRotate oEquipment, oShape, 0, 0, 0, _
                                (dPullBoxLength - dPullBoxHole1Width) / 2, dHole1Offset, -(dPullBoxHeight / 2# - dPullBoxHole1Height / 2) + dPullBoxHole1VerticalOffset
    End If
    
    ' Change attributes of the third shape (Hole 2)
    oShape.GetPart oShapePart
    Set oAttributes = oShape
    Set oAttributesCol = oAttributes.CollectionOfAttributes(oShapePart.GetPartOccIID)
    If Not oAttributesCol Is Nothing Then
        oAttributesCol.Item("A").Value = dPullBoxHole1Width
        oAttributesCol.Item("B").Value = dHole1Length
        oAttributesCol.Item("C").Value = dPullBoxHole1Height
    End If

    m_PullBoxHole1Height = dPullBoxHole1Height
    m_PullBoxHole1Width = dPullBoxHole1Width

    ' Set the position of the fourth shape (Hole 2)
    Set oShape = oShapes(4)
    dHole2Offset = 0#
    dHole2Length = dPullBoxLength
    ' if both bPullBoxHoleSide2 and bPullBoxHoleSide4 are false then we could suppress the shape
    '   via oDesignSolid.SetOperatorType oShape, BooleanOperatorSuppress
    If Not bPullBoxHoleSide2 Then
        dHole2Offset = dPullBoxThickness
        dHole2Length = dHole2Length - dPullBoxThickness
    End If
    If Not bPullBoxHoleSide4 Then
        dHole2Length = dHole2Length - dPullBoxThickness
    End If

    oShape.SetOrigin xPos + dHole2Offset, yPos, zPos - (dPullBoxHeight / 2# - dPullBoxHole2Height / 2) + dPullBoxHole2VerticalOffset

    ' Change attributes of the fourth shape (Hole 2)
    oShape.GetPart oShapePart
    Set oAttributes = oShape
    Set oAttributesCol = oAttributes.CollectionOfAttributes(oShapePart.GetPartOccIID)
    If Not oAttributesCol Is Nothing Then
        oAttributesCol.Item("A").Value = dHole2Length
        oAttributesCol.Item("B").Value = dPullBoxHole2Width
        oAttributesCol.Item("C").Value = dPullBoxHole2Height
    End If

    m_PullBoxHole2Height = dPullBoxHole2Height
    m_PullBoxHole2Width = dPullBoxHole2Width

    TransformAndRotate oEquipment, oShape, 0, 0, 0, _
                                dHole2Offset, 0, -(dPullBoxHeight / 2# - dPullBoxHole2Height / 2) + dPullBoxHole2VerticalOffset

    Set oAttributes = Nothing
    Set oAttributesCol = Nothing
    Set oShape = Nothing
    Set oLocalCoor = Nothing
    Set oDesignSolid = Nothing
    Set oSmartOccurrence = Nothing
    Set oShapePart = Nothing
    Set oShapeFactory = Nothing
    Set oCatalogResourceMgr = Nothing
    Set oIJDObject = Nothing
    Set oResourceManager = Nothing
    oExistingShapesColl.Clear
    Set oExistingShapesColl = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function GetPartFromPartNumber(ByVal CatalogResourceMgr As Object, ByVal PartNumber As String, ByVal ModelResourceMgr As Object) As Object
    Const METHOD = "GetPartFromPartNumber"
    On Error GoTo ErrHandler

    Dim oPart As IJDPart
    Dim oModelPOM As IJDPOM
    Dim oNamingCntxObject As IJDNamingContextObject

    Set oNamingCntxObject = New NamingContextObject

    If Trim(PartNumber) <> vbNullString Then
        Set oPart = oNamingCntxObject.ObjectMoniker(CatalogResourceMgr, PartNumber)

        ' Get the proxy of the part in model
        Set oModelPOM = ModelResourceMgr
        Set GetPartFromPartNumber = oModelPOM.GetProxy(oPart, True)
    End If

    Set oPart = Nothing
    Set oModelPOM = Nothing
    Set oNamingCntxObject = Nothing
    Exit Function

ErrHandler:
    Set oPart = Nothing
    Set oModelPOM = Nothing
    Set oNamingCntxObject = Nothing
    HandleError MODULE, METHOD
End Function

Private Sub GetDimensionsFromSymbolArray(SmartOccurrence As IJSmartOccurrence)
    Const METHOD = "GetDimensionsFromSymbolArray"
    On Error GoTo ErrorHandler

    m_avSymbolArrayOfInputs = m_oEquipCADHelper.GetSymbolArrayOfInputs(SmartOccurrence)

    'Inputs,  from equipment symbol code
    'Set m_oPartFclt = m_avSymbolArrayOfInputs(1)
    m_PullBoxHeight = m_avSymbolArrayOfInputs(2)
    m_PullBoxWidth = m_avSymbolArrayOfInputs(3)
    m_PullBoxLength = m_avSymbolArrayOfInputs(4)
    m_PullBoxThickness = m_avSymbolArrayOfInputs(5)
    m_PullBoxHole1Height = m_avSymbolArrayOfInputs(6)
    m_PullBoxHole1Width = m_avSymbolArrayOfInputs(7)
    m_PullBoxHole1Offset = m_avSymbolArrayOfInputs(8)
    m_bPullBoxHole1ThruStart = m_avSymbolArrayOfInputs(9)
    m_bPullBoxHole1ThruEnd = m_avSymbolArrayOfInputs(10)
    m_PullBoxHole2Height = m_avSymbolArrayOfInputs(11)
    m_PullBoxHole2Width = m_avSymbolArrayOfInputs(12)
    m_PullBoxHole2Offset = m_avSymbolArrayOfInputs(13)
    m_bPullBoxHole2ThruStart = m_avSymbolArrayOfInputs(14)
    m_bPullBoxHole2ThruEnd = m_avSymbolArrayOfInputs(15)
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub SetObjNameRule(ByRef obj As Object, ByRef CLASSNAME As String)
' Apply the namerule using the IJDNamingRulesHelper helper interface
    Const METHOD = "SetNameRule"
    On Error GoTo ErrorHandler

    Dim NameRule As String
    Dim NamingRules As IJElements

    Dim oNameRuleHlpr As GSCADNameRuleSemantics.IJDNamingRulesHelper

    'Returns a collection of the naming rules available in the catalog database
    'for the given object
    Set oNameRuleHlpr = New GSCADNameRuleHlpr.NamingRulesHelper
    Call oNameRuleHlpr.GetEntityNamingRulesGivenName(CLASSNAME, NamingRules)
    'get the first namerule from the collection
    Dim oNameRuleHolder As GSCADGenericNamingRulesFacelets.IJDNameRuleHolder
    
    If CLASSNAME = "CPShape" Then
        Set oNameRuleHolder = NamingRules.Item(2) 'Set Name rule as "Default Name Rule" for cable pits
    Else
        Set oNameRuleHolder = NamingRules.Item(1) 'Set name rule as Control Point name rule for Control Point
    End If
    'Create relations "NamedEntity" and "EntityNamingRule" and obj
    Dim oNameRuleAE As GSCADGenNameRuleAE.IJNameRuleAE
    Call oNameRuleHlpr.AddNamingRelations(obj, oNameRuleHolder, oNameRuleAE)

    GoTo CleanObjects
ErrorHandler:
    HandleError MODULE, METHOD

CleanObjects:
    Set oNameRuleHlpr = Nothing
    Set oNameRuleHolder = Nothing
    Set oNameRuleAE = Nothing
    Set NamingRules = Nothing
End Sub

' Custom Methods for ControlPoint

Public Sub CMConstructControlPoint(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
    Const METHOD = "CMConstructControlPoint"
    On Error GoTo ErrorHandler

    Dim oEquipment As IJEquipment
    Set oEquipment = pMemberDescription.CAO

    Dim oControlPoint As IJControlPoint
    Dim oGBSFactory As IJGeneralBusinessObjectsFactory
    Set oGBSFactory = New GeneralBusinessObjectsFactory

    Set oControlPoint = oGBSFactory.CreateControlPoint(pResourceManager, 0#, 0#, 0#, 0.1, , False, True)

    oControlPoint.SubType = cpProcessEquipment
    oControlPoint.Type = cpControlPoint
    oControlPoint.IsAssociative = True
    oControlPoint.Diameter = 0.1

    Set pObj = oControlPoint
    SetObjNameRule oControlPoint, "CPControlPoint"

    Set oControlPoint = Nothing
    Set oEquipment = Nothing
    Set oGBSFactory = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateControlPointGeometry(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateControlPointGeometry"
    On Error GoTo ErrorHandler

    Call CMComputeControlPoints(pPropertyDescriptions, pObject)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMComputeControlPoints(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMComputeControlPoints"
    On Error GoTo ErrorHandler

    Dim oControlPoint As IJControlPoint
    Set oControlPoint = pPropertyDescriptions.Object

    oControlPoint.SubType = cpProcessEquipment
    oControlPoint.Type = cpControlPoint
    oControlPoint.IsAssociative = True
    oControlPoint.Diameter = 0.1

    Dim oEquipment As IJEquipment
    Set oEquipment = pPropertyDescriptions.CAO

    Dim x As Double, y As Double, z As Double

    oEquipment.GetPosition x, y, z

    Dim ControlPt As IJPoint
    Set ControlPt = New Point3d
    Set ControlPt = oControlPoint

    ControlPt.SetPoint x + m_PullBoxLength / 2, y, z + m_PullBoxHeight / 2

    Set oControlPoint = Nothing
    Set ControlPt = Nothing
    Set oEquipment = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalControlPoint(ByVal pMemberDescription As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalControlPoint"
    On Error GoTo ErrorHandler


    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructControlPoint(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructControlPoint"
    On Error GoTo ErrorHandler

    Call CreateControlPointRelation(pMemberDesc)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseControlPoint(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseControlPoint"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputControlPoint(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputControlPoint"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CreateControlPointRelation(pMemberDesc As IJDMemberDescription)
    Const METHOD = "CreateControlPointRelation"
    On Error GoTo ErrorHandler

    Dim oParent As IJDObject
    Dim oChild As IJDObject
    Dim lConditionId As Long
    Dim CollHlpr As IMSRelation.DCollectionHelper
    Dim TOColl As IJDTargetObjectCol
    Dim AssocRel As IJDAssocRelation
    Dim oRevision As IJRevision
    Dim oRelationship As IJDRelationship

    lConditionId = -1
    Set oParent = pMemberDesc.CAO
    Set oChild = pMemberDesc.Object
    If Not oParent Is Nothing Then
        lConditionId = oParent.PermissionGroup
        If Not lConditionId = -1 Then
            oChild.PermissionGroup = lConditionId
        End If
    End If
    Set oParent = Nothing
    Set oChild = Nothing

    Set AssocRel = pMemberDesc.Object
    Set CollHlpr = AssocRel.CollectionRelations(IJControlPoint, "Parent")

    If Not CollHlpr Is Nothing Then
        Set TOColl = CollHlpr
        On Error Resume Next
        If TOColl.Count > 0 Then
            TOColl.Remove 1
        End If
        On Error GoTo ErrorHandler
        TOColl.Add pMemberDesc.CAO, vbNullString, oRelationship
        Set oRevision = New JRevision
        oRevision.AddRelationship oRelationship

        Set CollHlpr = Nothing
        Set TOColl = Nothing
        Set AssocRel = Nothing
    End If

    Set oRevision = Nothing
    Set oRelationship = Nothing
    Exit Sub
ErrorHandler:     HandleError MODULE, METHOD
End Sub
Private Sub TransformAndRotateCTNozzle(oEquipment As IJEquipment, oNozzle As IJCableTrayPortOcc, _
                                    dAngleOfRotAboutX As Double, _
                                    dAngleOfRotAboutY As Double, dAngleOfRotAboutZ As Double, _
                                    dXOffset As Double, dYOffset As Double, dZOffset As Double)
    Const METHOD = "TransformAndRotateCTNozzle"
    On Error GoTo ErrorHandler
    
    If Not oNozzle Is Nothing Then
        Dim oNzlMatrixAccess As IJDMatrixAccess
        Set oNzlMatrixAccess = oNozzle

        Dim oEqpMatrix As IJDT4x4
        Dim oDirVector As New AutoMath.DVector
        Dim oEquipPrimary As New AutoMath.DVector
        Dim oEquipSecondary As New AutoMath.DVector
        Dim oEquipNormal As New AutoMath.DVector

        oEquipment.GetMatrix oEqpMatrix
        oEquipPrimary.Set oEqpMatrix.IndexValue(0), oEqpMatrix.IndexValue(1), oEqpMatrix.IndexValue(2)
        oEquipSecondary.Set oEqpMatrix.IndexValue(4), oEqpMatrix.IndexValue(5), oEqpMatrix.IndexValue(6)
        oEquipNormal.Set oEqpMatrix.IndexValue(8), oEqpMatrix.IndexValue(9), oEqpMatrix.IndexValue(10)

        If Not CmpDblEqual(dAngleOfRotAboutX, 0) Then
            oDirVector.Set 1, 0, 0
            oEqpMatrix.Rotate dAngleOfRotAboutX, oDirVector
        End If
        If Not CmpDblEqual(dAngleOfRotAboutY, 0) Then
            oDirVector.Set 0, 1, 0
            oEqpMatrix.Rotate dAngleOfRotAboutY, oDirVector
        End If
        If Not CmpDblEqual(dAngleOfRotAboutZ, 0) Then
            oDirVector.Set 0, 0, 1
            oEqpMatrix.Rotate dAngleOfRotAboutZ, oDirVector
        End If

        Dim oCompOrigin As IJDPosition
        Dim x As Double, y As Double, z As Double

        Set oCompOrigin = New AutoMath.DPosition
        oEquipment.GetPosition x, y, z
        oCompOrigin.Set x, y, z

        If Not CmpDblEqual(dXOffset, 0) Then
            oEquipPrimary.Length = dXOffset
            Set oCompOrigin = oCompOrigin.Offset(oEquipPrimary)
        End If

        If Not CmpDblEqual(dYOffset, 0) Then
            oEquipSecondary.Length = dYOffset
            Set oCompOrigin = oCompOrigin.Offset(oEquipSecondary)
        End If

        If Not CmpDblEqual(dZOffset, 0) Then
            oEquipNormal.Length = dZOffset
            Set oCompOrigin = oCompOrigin.Offset(oEquipNormal)
        End If
        oEqpMatrix.IndexValue(12) = oCompOrigin.x
        oEqpMatrix.IndexValue(13) = oCompOrigin.y
        oEqpMatrix.IndexValue(14) = oCompOrigin.z
        oNzlMatrixAccess.Matrix = oEqpMatrix

        Set oNzlMatrixAccess = Nothing
        Set oEqpMatrix = Nothing
        Set oDirVector = Nothing
    End If

    Exit Sub
ErrorHandler:
    Set oEqpMatrix = Nothing
    Set oDirVector = Nothing
    HandleError MODULE, METHOD
End Sub

Private Sub TransformAndRotate(Equipment As IJEquipment, oShape As IJShape, _
                                    dAngleOfRotAboutX As Double, _
                                    dAngleOfRotAboutY As Double, dAngleOfRotAboutZ As Double, _
                                    dXOffset As Double, dYOffset As Double, dZOffset As Double)
    Const METHOD = "TransformAndRotate"
    On Error GoTo ErrorHandler
    
    If Not oShape Is Nothing Then

            Dim oEqpMatrix As IJDT4x4
            Dim oTransform As IJShape
            Dim oDirVector As New AutoMath.DVector
            Dim oEquipPrimary As New AutoMath.DVector
            Dim oEquipSecondary As New AutoMath.DVector
            Dim oEquipNormal As New AutoMath.DVector
            
            Equipment.GetMatrix oEqpMatrix
            oEquipPrimary.Set oEqpMatrix.IndexValue(0), oEqpMatrix.IndexValue(1), oEqpMatrix.IndexValue(2)
            oEquipSecondary.Set oEqpMatrix.IndexValue(4), oEqpMatrix.IndexValue(5), oEqpMatrix.IndexValue(6)
            oEquipNormal.Set oEqpMatrix.IndexValue(8), oEqpMatrix.IndexValue(9), oEqpMatrix.IndexValue(10)
            Set oTransform = oShape

            If Not CmpDblEqual(dAngleOfRotAboutX, 0) Then
                oDirVector.Set 1, 0, 0
                oEqpMatrix.Rotate dAngleOfRotAboutX, oDirVector
            End If
            If Not CmpDblEqual(dAngleOfRotAboutY, 0) Then
                oDirVector.Set 0, 1, 0
                oEqpMatrix.Rotate dAngleOfRotAboutY, oDirVector
            End If
            If Not CmpDblEqual(dAngleOfRotAboutZ, 0) Then
                oDirVector.Set 0, 0, 1
                oEqpMatrix.Rotate dAngleOfRotAboutZ, oDirVector
            End If
            'oTransform.DTransform oEqpMatrix
            oTransform.SetMatrix oEqpMatrix
            
            Dim oCompOrigin As IJDPosition
            Dim x As Double, y As Double, z As Double
            
            Set oCompOrigin = New AutoMath.DPosition
            Equipment.GetPosition x, y, z
            oCompOrigin.Set x, y, z
            
            If Not CmpDblEqual(dXOffset, 0) Then
                oEquipPrimary.Length = dXOffset
                Set oCompOrigin = oCompOrigin.Offset(oEquipPrimary)
            End If
            
           If Not CmpDblEqual(dYOffset, 0) Then
                oEquipSecondary.Length = dYOffset
                Set oCompOrigin = oCompOrigin.Offset(oEquipSecondary)
            End If
            
            If Not CmpDblEqual(dZOffset, 0) Then
                oEquipNormal.Length = dZOffset
                Set oCompOrigin = oCompOrigin.Offset(oEquipNormal)
            End If
            
            oTransform.SetOrigin oCompOrigin.x, oCompOrigin.y, oCompOrigin.z
            oTransform.GetMatrix oEqpMatrix
            
            Set oTransform = Nothing
            Set oEqpMatrix = Nothing
            Set oDirVector = Nothing
        End If
     
    Exit Sub
ErrorHandler:
    Set oTransform = Nothing
    Set oEqpMatrix = Nothing
    Set oDirVector = Nothing
    HandleError MODULE, METHOD
End Sub
