VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFD3ACoolBDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"CAD"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/2/2005-7:23:42 PM"
Attribute VB_Ext_KEY = "SP3DEqpCADTemplateVersion" ,"1.2"
'******************************************************************************
' Copyright (C) 2004, Intergraph Corporation. All rights reserved.
'
'File
'   CSFD3AirCoolBDef.cls
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
'   16.June.2005   MS TR79730 Some issues with the deletion of members is resolved;made datumshapes/foundationports can be deleted etc
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  06.SEP.2007    PK  TR109377 Made changes so that deleted nozzles do not reappear after modification of equipment properties and also nozzles can be moved.
'                                For this,the nozzles have been created in the def file
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SP3DFD3AirCoolerBayAsm:CFD3ACoolBDef"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const IID_IJDGEOMETRY = "{A1732CBF-5136-11D1-9770-080036754203}"

Private m_oEquipCADHelper As IJEquipCADHelper
Private m_oEditErrors As IJEditErrors
Private m_avSymbolArrayOfInputs() As Variant

Private m_dCool1toCool2CentoCen As Double
Private m_dCool2toCool3CentoCen As Double
Private m_dCoolerWidth As Double
Private m_dInletBlockHeight As Double
Private m_dInletNozzLength As Double
Private m_dOutletNozzLength As Double
Private m_dCoolerLength As Double
Private m_dCoolLeftfromPP As Double
 
' Implement User Symbol Services(USS)
Implements IJDUserSymbolServices
Implements IJEquipUserAttrMgmt

Private Sub Class_Initialize()
    Set m_oEquipCADHelper = New CADServices
    Set m_oEditErrors = New JServerErrors
    
    m_oEquipCADHelper.ProjectName = "SP3DFD3AirCoolerBayAsm"
    m_oEquipCADHelper.ClassName = "CFD3ACoolBDef"
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
    'Add new member(NozzleAirCoolingNoz11) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleAirCoolingNoz11", 1, "CMConstructNozzleAirCoolingNoz11", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleAirCoolingNoz11"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleAirCoolingNoz11"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleAirCoolingNoz11"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleAirCoolingNoz11"
    'Add properties for (NozzleAirCoolingNoz11)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleAirCoolingNoz11Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleAirCoolingNoz11", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleAirCoolingNoz11GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleAirCoolingNoz11", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleAirCoolingNoz22) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleAirCoolingNoz22", 2, "CMConstructNozzleAirCoolingNoz22", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleAirCoolingNoz22"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleAirCoolingNoz22"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleAirCoolingNoz22"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleAirCoolingNoz22"
    'Add properties for (NozzleAirCoolingNoz22)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleAirCoolingNoz22Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleAirCoolingNoz22", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleAirCoolingNoz22GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleAirCoolingNoz22", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleAirCoolingNoz33) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleAirCoolingNoz33", 3, "CMConstructNozzleAirCoolingNoz33", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleAirCoolingNoz33"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleAirCoolingNoz33"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleAirCoolingNoz33"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleAirCoolingNoz33"
    'Add properties for (NozzleAirCoolingNoz33)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleAirCoolingNoz33Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleAirCoolingNoz33", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleAirCoolingNoz33GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleAirCoolingNoz33", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleAirCoolingNoz44) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleAirCoolingNoz44", 4, "CMConstructNozzleAirCoolingNoz44", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleAirCoolingNoz44"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleAirCoolingNoz44"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleAirCoolingNoz44"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleAirCoolingNoz44"
    'Add properties for (NozzleAirCoolingNoz44)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleAirCoolingNoz44Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleAirCoolingNoz44", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleAirCoolingNoz44GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleAirCoolingNoz44", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleAirCoolingNoz55) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleAirCoolingNoz55", 5, "CMConstructNozzleAirCoolingNoz55", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleAirCoolingNoz55"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleAirCoolingNoz55"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleAirCoolingNoz55"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleAirCoolingNoz55"
    'Add properties for (NozzleAirCoolingNoz55)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleAirCoolingNoz55Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleAirCoolingNoz55", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleAirCoolingNoz55GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleAirCoolingNoz55", imsCOOKIE_ID_USS_LIB
    
    'Add new member(NozzleAirCoolingNoz66) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("NozzleAirCoolingNoz66", 6, "CMConstructNozzleAirCoolingNoz66", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsNozzleAirCoolingNoz66"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructNozzleAirCoolingNoz66"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalNozzleAirCoolingNoz66"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseNozzleAirCoolingNoz66"
    'Add properties for (NozzleAirCoolingNoz66)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "NozzleAirCoolingNoz66Properties", 1, IID_IJDATTRIBUTES, "CMEvaluateNozzleAirCoolingNoz66", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "NozzleAirCoolingNoz66GeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryNozzleAirCoolingNoz66", imsCOOKIE_ID_USS_LIB
    
        
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
        Case "NozzleAirCoolingNoz11", "NozzleAirCoolingNoz22", "NozzleAirCoolingNoz33", "NozzleAirCoolingNoz44", "NozzleAirCoolingNoz55", "NozzleAirCoolingNoz66"
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
        Case "NozzleAirCoolingNoz11", "NozzleAirCoolingNoz22", "NozzleAirCoolingNoz33", "NozzleAirCoolingNoz44", "NozzleAirCoolingNoz55", "NozzleAirCoolingNoz66"
            For m = 1 To oAttrCollection.Count
                Set oAttributeDescriptor = oAttrCollection.Item(m)
                Select Case UCase(oAttributeDescriptor.InterfaceName)
                Case "IJDELETABLEMEMBER"
                    If UCase(oAttributeDescriptor.AttrName) = "CANBEDELETED" Then
                        oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsChanged
                    End If
                Case "IJNOZZLEORIENTATION"
                    'The nozzle is placed by Point
                    oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsChanged
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
' Custom Methods for NozzleAirCoolingNoz11
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleAirCoolingNoz11(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleAirCoolingNoz11"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    oVec.Set 0, 0, 1
    GetDimensionsFromSymbolArray pMemberDescription.CAO

    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 1, pResourceManager, DistribPortType_PIPE, pObject, False

    Set oNozzle = pObject
    
    oNozzle.Length = 1.25 * m_dInletNozzLength
          
    Dim matObj As IJDT4x4
    Dim oNzlMatrixAccess As IJDMatrixAccess
    Set oNzlMatrixAccess = oNozzle
    Set matObj = oNzlMatrixAccess.Matrix
    matObj.IndexValue(12) = 0
    matObj.IndexValue(13) = m_dCoolerWidth / 3
    matObj.IndexValue(14) = m_dInletBlockHeight / 2 + m_dInletNozzLength
    oNzlMatrixAccess.Matrix = matObj
    Set oNzlMatrixAccess = Nothing
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    Set oDistribPort = Nothing
    Set oVec = Nothing

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)

    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleAirCoolingNoz11(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleAirCoolingNoz11"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleAirCoolingNoz11(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleAirCoolingNoz11"
    On Error GoTo ErrorHandler

    Dim oSmartOcc As IJSmartOccurrence
    Dim oOrientation As IJNozzleOrientation

    Set oSmartOcc = pMemberDesc.CAO

    'Establish the relationship with the equipment origin if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations( _
    pMemberDesc.CAO, pMemberDesc.Object)

    Set oSmartOcc = Nothing
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleAirCoolingNoz11(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleAirCoolingNoz11"
    On Error GoTo ErrorHandler


    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleAirCoolingNoz11(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleAirCoolingNoz11"
    On Error GoTo ErrorHandler

    Dim oEquipment As IJEquipment
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment

    Set pObject = oPropertyDescription.Object
    Set oNozzle = pObject
    oNozzle.Length = 1.25 * m_dInletNozzLength

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
     oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleAirCoolingNoz11(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleAirCoolingNoz11"

    On Error GoTo ErrorHandler

    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleAirCoolingNoz11(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleAirCoolingNoz11"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleAirCoolingNoz22
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleAirCoolingNoz22(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleAirCoolingNoz22"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    oVec.Set 0, 0, -1
    
    GetDimensionsFromSymbolArray pMemberDescription.CAO

    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 2, pResourceManager, DistribPortType_PIPE, pObject, False

    Set oNozzle = pObject
    oNozzle.Length = 1.25 * m_dOutletNozzLength
    
    Dim matObj As IJDT4x4
    Dim oNzlMatrixAccess As IJDMatrixAccess
    Set oNzlMatrixAccess = oNozzle
    Set matObj = oNzlMatrixAccess.Matrix
    matObj.IndexValue(12) = 2 * m_dCoolLeftfromPP + m_dCoolerLength
    matObj.IndexValue(13) = -m_dCoolerWidth / 3
    matObj.IndexValue(14) = -(m_dInletBlockHeight / 2 + m_dOutletNozzLength)
    oNzlMatrixAccess.Matrix = matObj
    Set oNzlMatrixAccess = Nothing
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    Set oDistribPort = Nothing
    Set oVec = Nothing

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)

    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleAirCoolingNoz22(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleAirCoolingNoz22"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleAirCoolingNoz22(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleAirCoolingNoz22"
    On Error GoTo ErrorHandler

    Dim oSmartOcc As IJSmartOccurrence
    Dim oOrientation As IJNozzleOrientation

    Set oSmartOcc = pMemberDesc.CAO

    'Establish the relationship with the equipment origin if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations( _
    pMemberDesc.CAO, pMemberDesc.Object)

    Set oSmartOcc = Nothing
    Set oOrientation = Nothing

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleAirCoolingNoz22(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleAirCoolingNoz22"
    On Error GoTo ErrorHandler


    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleAirCoolingNoz22(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleAirCoolingNoz22"
    On Error GoTo ErrorHandler

    Dim oEquipment As IJEquipment
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment

    Set pObject = oPropertyDescription.Object
    Set oNozzle = pObject
    oNozzle.Length = 1.25 * m_dOutletNozzLength

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)

    oOrientation.PlacementType = Position_By_Point

    Set oNozzle = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleAirCoolingNoz22(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleAirCoolingNoz22"

    On Error GoTo ErrorHandler

    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleAirCoolingNoz22(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleAirCoolingNoz22"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleAirCoolingNoz33
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleAirCoolingNoz33(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleAirCoolingNoz33"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    oVec.Set 0, 0, 1
    GetDimensionsFromSymbolArray pMemberDescription.CAO

    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 3, pResourceManager, DistribPortType_PIPE, pObject, False

    Set oNozzle = pObject
    oNozzle.Length = 1.25 * m_dInletNozzLength
    
    Dim matObj As IJDT4x4
    Dim oNzlMatrixAccess As IJDMatrixAccess
    Set oNzlMatrixAccess = oNozzle
    Set matObj = oNzlMatrixAccess.Matrix
    matObj.IndexValue(12) = 0
    matObj.IndexValue(13) = -(m_dCool1toCool2CentoCen - m_dCoolerWidth / 3)
    matObj.IndexValue(14) = m_dInletBlockHeight / 2 + m_dInletNozzLength
    oNzlMatrixAccess.Matrix = matObj
    Set oNzlMatrixAccess = Nothing
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    Set oDistribPort = Nothing
    Set oVec = Nothing
    
    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)

    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleAirCoolingNoz33(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleAirCoolingNoz33"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleAirCoolingNoz33(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleAirCoolingNoz33"
    On Error GoTo ErrorHandler

   Dim oSmartOcc As IJSmartOccurrence
    Dim oOrientation As IJNozzleOrientation

    Set oSmartOcc = pMemberDesc.CAO

    'Establish the relationship with the equipment origin if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations( _
    pMemberDesc.CAO, pMemberDesc.Object)

    Set oSmartOcc = Nothing
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleAirCoolingNoz33(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleAirCoolingNoz33"
    On Error GoTo ErrorHandler


    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleAirCoolingNoz33(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleAirCoolingNoz33"
    On Error GoTo ErrorHandler

    Dim oEquipment As IJEquipment
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment

    Set pObject = oPropertyDescription.Object
    Set oNozzle = pObject
    oNozzle.Length = 1.25 * m_dInletNozzLength
    
    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)

    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing

   Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleAirCoolingNoz33(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleAirCoolingNoz33"

    On Error GoTo ErrorHandler

    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleAirCoolingNoz33(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleAirCoolingNoz33"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleAirCoolingNoz44
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleAirCoolingNoz44(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleAirCoolingNoz44"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    oVec.Set 0, 0, -1

    GetDimensionsFromSymbolArray pMemberDescription.CAO

    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 4, pResourceManager, DistribPortType_PIPE, pObject, False

    Set oNozzle = pObject
    oNozzle.Length = 1.25 * m_dOutletNozzLength
    
    Dim matObj As IJDT4x4
    Dim oNzlMatrixAccess As IJDMatrixAccess
    Set oNzlMatrixAccess = oNozzle
    Set matObj = oNzlMatrixAccess.Matrix
    matObj.IndexValue(12) = 2 * m_dCoolLeftfromPP + m_dCoolerLength
    matObj.IndexValue(13) = -(m_dCool1toCool2CentoCen + m_dCoolerWidth / 3)
    matObj.IndexValue(14) = -(m_dInletBlockHeight / 2 + m_dOutletNozzLength)
    oNzlMatrixAccess.Matrix = matObj
    Set oNzlMatrixAccess = Nothing
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    Set oDistribPort = Nothing
    Set oVec = Nothing

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)
    
    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    
    Exit Sub
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleAirCoolingNoz44(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleAirCoolingNoz44"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleAirCoolingNoz44(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleAirCoolingNoz44"
    On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oOrientation As IJNozzleOrientation

    Set oSmartOcc = pMemberDesc.CAO

    'Establish the relationship with the equipment origin if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations( _
    pMemberDesc.CAO, pMemberDesc.Object)

    Set oSmartOcc = Nothing
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleAirCoolingNoz44(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleAirCoolingNoz44"
    On Error GoTo ErrorHandler


    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleAirCoolingNoz44(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleAirCoolingNoz44"
    On Error GoTo ErrorHandler

    Dim oEquipment As IJEquipment
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment

    Set pObject = oPropertyDescription.Object
    Set oNozzle = pObject
    oNozzle.Length = 1.25 * m_dOutletNozzLength

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)

    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleAirCoolingNoz44(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleAirCoolingNoz44"

    On Error GoTo ErrorHandler

    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleAirCoolingNoz44(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleAirCoolingNoz44"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleAirCoolingNoz55
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleAirCoolingNoz55(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleAirCoolingNoz55"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    oVec.Set 0, 0, 1
    GetDimensionsFromSymbolArray pMemberDescription.CAO

    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 5, pResourceManager, DistribPortType_PIPE, pObject, False

    Set oNozzle = pObject
    oNozzle.Length = 1.25 * m_dInletNozzLength
    
    Dim matObj As IJDT4x4
    Dim oNzlMatrixAccess As IJDMatrixAccess
    Set oNzlMatrixAccess = oNozzle
    Set matObj = oNzlMatrixAccess.Matrix
    matObj.IndexValue(12) = 0
    matObj.IndexValue(13) = -(m_dCool1toCool2CentoCen + m_dCool2toCool3CentoCen - m_dCoolerWidth / 3)
    matObj.IndexValue(14) = m_dInletBlockHeight / 2 + m_dInletNozzLength
    oNzlMatrixAccess.Matrix = matObj
    Set oNzlMatrixAccess = Nothing
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    Set oDistribPort = Nothing
    Set oVec = Nothing

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)

    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleAirCoolingNoz55(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleAirCoolingNoz55"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleAirCoolingNoz55(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleAirCoolingNoz55"
    On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oOrientation As IJNozzleOrientation

    Set oSmartOcc = pMemberDesc.CAO

    'Establish the relationship with the equipment origin if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations( _
    pMemberDesc.CAO, pMemberDesc.Object)

    Set oSmartOcc = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleAirCoolingNoz55(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleAirCoolingNoz55"
    On Error GoTo ErrorHandler


    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleAirCoolingNoz55(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleAirCoolingNoz55"
    On Error GoTo ErrorHandler

     Dim oEquipment As IJEquipment
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment

    Set pObject = oPropertyDescription.Object
    Set oNozzle = pObject
    oNozzle.Length = 1.25 * m_dInletNozzLength
 
    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)

    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing

   Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleAirCoolingNoz55(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleAirCoolingNoz55"

    On Error GoTo ErrorHandler

    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleAirCoolingNoz55(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleAirCoolingNoz55"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Custom Methods for NozzleAirCoolingNoz66
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub CMConstructNozzleAirCoolingNoz66(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructNozzleAirCoolingNoz66"
    On Error GoTo ErrorHandler

    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle
    Dim oDistribPort As IJDistribPort
    Dim oVec As New DVector
    oVec.Set 0, 0, -1

    GetDimensionsFromSymbolArray pMemberDescription.CAO

    'Create Nozzle
    m_oEquipCADHelper.CreateNozzleGivenIndex pMemberDescription, 6, pResourceManager, DistribPortType_PIPE, pObject, False

    Set oNozzle = pObject
    oNozzle.Length = 1.25 * m_dOutletNozzLength
    
    Dim matObj As IJDT4x4
    Dim oNzlMatrixAccess As IJDMatrixAccess
    Set oNzlMatrixAccess = oNozzle
    Set matObj = oNzlMatrixAccess.Matrix
    matObj.IndexValue(12) = 2 * m_dCoolLeftfromPP + m_dCoolerLength
    matObj.IndexValue(13) = -(m_dCool1toCool2CentoCen + m_dCool2toCool3CentoCen + m_dCoolerWidth / 3)
    matObj.IndexValue(14) = -(m_dInletBlockHeight / 2 + m_dOutletNozzLength)
    oNzlMatrixAccess.Matrix = matObj
    Set oNzlMatrixAccess = Nothing
    Set oDistribPort = oNozzle
    oDistribPort.SetDirectionVector oVec
    Set oDistribPort = Nothing
    Set oVec = Nothing

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)

    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructNozzleAirCoolingNoz66(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructNozzleAirCoolingNoz66"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsNozzleAirCoolingNoz66(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsNozzleAirCoolingNoz66"
    On Error GoTo ErrorHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oOrientation As IJNozzleOrientation

    Set oSmartOcc = pMemberDesc.CAO

    'Establish the relationship with the equipment origin if it doesn't already exist.
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations( _
    pMemberDesc.CAO, pMemberDesc.Object)

    Set oSmartOcc = Nothing
    Set oOrientation = Nothing
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateNozzleAirCoolingNoz66(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateNozzleAirCoolingNoz66"
    On Error GoTo ErrorHandler


    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryNozzleAirCoolingNoz66(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryNozzleAirCoolingNoz66"
    On Error GoTo ErrorHandler

    Dim oEquipment As IJEquipment
    Dim oOrientation As IJNozzleOrientation
    Dim oNozzle As IJDNozzle

    Set oEquipment = oPropertyDescription.CAO
    GetDimensionsFromSymbolArray oEquipment

    Set pObject = oPropertyDescription.Object
    Set oNozzle = pObject
    oNozzle.Length = 1.25 * m_dOutletNozzLength

    'Create the nozzle orientation and set it on the nozzle
    Set oOrientation = m_oEquipCADHelper.CreateOrientationAndSetRelations(Nothing, oNozzle)

    'Set the default values
    oOrientation.PlacementType = Position_By_Point
    Set oNozzle = Nothing
    Set oOrientation = Nothing
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalNozzleAirCoolingNoz66(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalNozzleAirCoolingNoz66"

    On Error GoTo ErrorHandler

    IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseNozzleAirCoolingNoz66(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseNozzleAirCoolingNoz66"
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Private Sub HandleError(sModule As String, sMethod As String)
    
    If Not m_oEditErrors Is Nothing Then
        m_oEditErrors.AddFromErr Err, "", sMethod, sModule
    End If
    
    Err.Raise Err.Number, Err.Source & " " & sMethod, Err.description, _
              Err.HelpFile, Err.HelpContext
End Sub

'The code of this method is specific to each Equipment Component as it converts the array of inputs in a set of global variables
Private Sub GetDimensionsFromSymbolArray(SmartOccurrence As IJSmartOccurrence)
Const METHOD = "GetDimensionsFromSymbolArray"
On Error GoTo ErrorHandler
    m_avSymbolArrayOfInputs = m_oEquipCADHelper.GetSymbolArrayOfInputs(SmartOccurrence)
    'Inputs,  from equipment symbol code
    m_dCool1toCool2CentoCen = m_avSymbolArrayOfInputs(2)
    m_dCool2toCool3CentoCen = m_avSymbolArrayOfInputs(3)
    m_dCoolerWidth = m_avSymbolArrayOfInputs(4)
    m_dCoolerLength = m_avSymbolArrayOfInputs(5)
    m_dInletBlockHeight = m_avSymbolArrayOfInputs(11)
    m_dCoolLeftfromPP = m_avSymbolArrayOfInputs(12)
    m_dInletNozzLength = m_avSymbolArrayOfInputs(18)
    m_dOutletNozzLength = m_avSymbolArrayOfInputs(19)
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

