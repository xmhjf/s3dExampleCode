VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CKXNDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"CAD"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 1/2/2005-7:27:51 PM"
Attribute VB_Ext_KEY = "SP3DEqpCADTemplateVersion" ,"1.2"
'******************************************************************************
' Copyright (C) 2004, Intergraph Corporation. All rights reserved.
'
'File
'   KettleExchNesteDef.cls
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
'   18.Mar.2005    MS  Added Datum Shapes
'   03.Jun.2005    MS  Added the three equipment components as members
'   16.June.2005   MS TR79730 Some issues with the deletion of members is resolved;made datumshapes/foundationports can be deleted etc
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "SP3DKettleXNestedAsm:CKXNDef"

Private Const IID_IJDATTRIBUTES = "{B25FD387-CFEB-11D1-850B-080036DE8E03}"
Private Const IID_IJDGEOMETRY = "{A1732CBF-5136-11D1-9770-080036754203}"

Private m_oEquipCADHelper As IJEquipCADHelper
Private m_oEditErrors As IJEditErrors

Private m_avSymbolArrayOfInputs()   As Variant

Private m_dDatumShapeSize           As Double

Private m_oNorth                    As IJDVector
Private m_oEast                     As IJDVector
Private m_oElevation                As IJDVector

Private m_dExchangerLength As Double
Private m_dExchangerNeckLength As Double
Private m_dExchangerNeckDiameter As Double
Private m_dExchangerTaperLength As Double
Private m_dExchangerDiameter As Double
Private m_dBundleFlangeTk As Double
Private m_dBundleFlangeDia As Double
Private m_dExchangerFlangeTk1 As Double
Private m_dBundlePullingLength As Double
Private m_dBotSupportCenFromPP As Double
Private m_dBottomSupportCentoCen As Double
Private m_dSupport1Thickness As Double
Private m_dSupport2Thickness As Double
Private m_dBottomSupportHeight As Double
Private m_dSupportLength As Double
Private m_dFrontEndFlangeDia As Double
Private m_dFrontEndFlangeTk1 As Double
Private m_dFrontEndLength1 As Double
Private m_dFrontEndLength2 As Double
Private m_dFrontEndFlangeTk2 As Double
Private m_dFrontEndFlangeTk3 As Double
Private m_dFrontEndDiameter As Double
Private m_dFrontEndType As Long
Private m_dInsulationThickness As Double

'   Define front end types
Private Const FrontEndType_A = 1
Private Const FrontEndType_B = 2
Private Const FrontEndType_C = 3
Private Const FrontEndType_N = 5
Private Const FrontEndType_D = 4
Private Const FrontEndType_Q = 6

Private m_bComputeEqpComp            As Boolean

' Implement User Symbol Services(USS)
Implements IJDUserSymbolServices
Implements IJEquipUserAttrMgmt

Private Sub Class_Initialize()
    Set m_oEquipCADHelper = New CADServices
    Set m_oEditErrors = New JServerErrors
    
    m_oEquipCADHelper.ProjectName = "SP3DKettleXNestedAsm"
    m_oEquipCADHelper.ClassName = "CKXNDef"
    m_oEquipCADHelper.OccurrenceRootClass = orcEquipment
        LogCalls "Class_Initialize"
    
    m_dDatumShapeSize = 9 * 25.4
        
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
    
    'Remove all the previous member descriptions
    Set oMemberDescriptions = oSymbolDefinition
    oMemberDescriptions.RemoveAll
    oPropertyDescriptions.AddProperty "EquipmentProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateEquipment", imsCOOKIE_ID_USS_LIB
    
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


    'Add new member(TyACDN) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("TyACDN", 2, "CMConstructTyACDN", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsTyACDN"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTyACDN"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalTyACDN"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseTyACDN"
    'Add properties for (TyACDN)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "TyACDNProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateTyACDN", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "TyACDNGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryTyACDN", imsCOOKIE_ID_USS_LIB


    'Add new member(TyB) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("TyB", 3, "CMConstructTyB", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsTyB"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTyB"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalTyB"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseTyB"
    'Add properties for (TyB)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "TyBProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateTyB", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "TyBGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryTyB", imsCOOKIE_ID_USS_LIB

    'Add new member(TyQ) to the definition
    Set oMemberDescription = Nothing
    Set oMemberDescription = oMemberDescriptions.AddMember("TyQ", 4, "CMConstructTyQ", imsCOOKIE_ID_USS_LIB)
    oMemberDescription.SetCMSetInputs imsCOOKIE_ID_USS_LIB, "CMSetInputsTyQ"
    oMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTyQ"
    oMemberDescription.SetCMConditional imsCOOKIE_ID_USS_LIB, "CMConditionalTyQ"
    oMemberDescription.SetCMRelease imsCOOKIE_ID_USS_LIB, "CMReleaseTyQ"
    'Add properties for (TyQ)
    Set oPropertyDescriptions = Nothing
    Set oPropertyDescriptions = oMemberDescription
    oPropertyDescriptions.AddProperty "TyQProperties", 1, IID_IJDATTRIBUTES, "CMEvaluateTyQ", imsCOOKIE_ID_USS_LIB
    oPropertyDescriptions.AddProperty "TyQGeometryProperties", 2, IID_IJDGEOMETRY, "CMEvaluateGeometryTyQ", imsCOOKIE_ID_USS_LIB


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
    LogCalls METHOD
    On Error GoTo ErrorHandler
  
    Dim oMemberDescription As IJDMemberDescription

    Set oMemberDescription = m_oEquipCADHelper.GetMemberDescriptionFromChild(pIJDAttrs)
        
    Select Case oMemberDescription.name
        Case "DP1","TyACDN","TyB","TyQ"
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
    LogCalls METHOD
    On Error GoTo ErrorHandler
    Dim oMemberDescription As IJDMemberDescription
    
    IJEquipUserAttrMgmt_OnPreLoad = "ERROR"
    Set oMemberDescription = m_oEquipCADHelper.GetMemberDescriptionFromChild(pIJDAttrs)
    Dim oAttrCollection As Collection
    Dim oAttributeDescriptor As IJEquipAttrDescriptor
    Dim m As Long
    Set oAttrCollection = CollAllDisplayedValues
    Select Case oMemberDescription.name
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
            Set oAttrCollection = Nothing
        Case "TyACDN", "TyB", "TyQ"
                'Set oAttrCollection = CollAllDisplayedValues
                For m = 1 To oAttrCollection.Count
                    Set oAttributeDescriptor = oAttrCollection.Item(m)
                    Select Case UCase(oAttributeDescriptor.InterfaceName)
                        Case "IJUAHEATEXCHANGERFRONTEND", "IJUAEXCHANGERCHANNEL"
                            oAttributeDescriptor.AttrState = oAttributeDescriptor.AttrState Or adsReadOnly
                        Case Else
                            '
                    End Select
                Next
                Set oAttrCollection = Nothing
        Case Else
        
    End Select
    
 Set oMemberDescription = Nothing
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
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMFinalConstructEquipment"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMSetInputsEquipment"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    m_oEquipCADHelper.SetSmartItemAsInputToSymbol pAggregatorDescription
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMRemoveInputsEquipment(ByVal pAggregatorDescription As IJDAggregatorDescription)
    Const METHOD = "CMRemoveInputsEquipment"
    LogCalls METHOD
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


Private Sub HandleError(sModule As String, sMethod As String)
    
    If Not m_oEditErrors Is Nothing Then
        m_oEditErrors.AddFromErr Err, "", sMethod, sModule
    End If
    
    Err.Raise Err.Number, Err.Source & " " & sMethod, Err.description, _
              Err.HelpFile, Err.HelpContext
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
    If IsFree(oDatumShape) = False Then
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

    m_dExchangerLength = m_avSymbolArrayOfInputs(2)           'P1
    m_dExchangerNeckLength = m_avSymbolArrayOfInputs(3)       'P2
    m_dExchangerNeckDiameter = m_avSymbolArrayOfInputs(4)     'P3
    m_dExchangerTaperLength = m_avSymbolArrayOfInputs(5)      'P4
    m_dExchangerDiameter = m_avSymbolArrayOfInputs(6)         'P5
    m_dBundleFlangeTk = m_avSymbolArrayOfInputs(7)            'P6
    m_dBundleFlangeDia = m_avSymbolArrayOfInputs(8)           'P7
    m_dExchangerFlangeTk1 = m_avSymbolArrayOfInputs(9)        'P8
    m_dBundlePullingLength = m_avSymbolArrayOfInputs(10)      'P9
    m_dBotSupportCenFromPP = m_avSymbolArrayOfInputs(11)      'P10
    m_dBottomSupportCentoCen = m_avSymbolArrayOfInputs(12)    'P11
    m_dSupport1Thickness = m_avSymbolArrayOfInputs(13)        'P12
    m_dSupport2Thickness = m_avSymbolArrayOfInputs(14)        'P13
    m_dBottomSupportHeight = m_avSymbolArrayOfInputs(15)      'P14
    m_dSupportLength = m_avSymbolArrayOfInputs(16)            'P15
    m_dFrontEndFlangeDia = m_avSymbolArrayOfInputs(17)        'P30
    m_dFrontEndFlangeTk1 = m_avSymbolArrayOfInputs(18)        'P31
    m_dFrontEndLength1 = m_avSymbolArrayOfInputs(19)          'P32
    m_dFrontEndLength2 = m_avSymbolArrayOfInputs(20)          'P33 of type A/C/D/N or B
    m_dFrontEndFlangeTk2 = m_avSymbolArrayOfInputs(21)        'P34
    m_dFrontEndFlangeTk3 = m_avSymbolArrayOfInputs(22)        'P35
    m_dFrontEndDiameter = m_avSymbolArrayOfInputs(23)         'P33 of type Q End
    m_dFrontEndType = m_avSymbolArrayOfInputs(24)
    m_dInsulationThickness = m_avSymbolArrayOfInputs(25)

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
Private Sub LogCalls(sMethod As String)
   m_oEquipCADHelper.LogError "Entering " & sMethod
End Sub

Private Sub TransformFromECStoGCS(Equipment As IJEquipment, Object As Object)
    Const METHOD = "TransformFromECStoGCS"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    Dim oEqpMatrix As IJDT4x4
    Dim oShapeMatrix As IJDT4x4
    Dim oTransform As IJDGeometry
    Dim oShape As IJShape
    
    If Not Object Is Nothing Then
        If TypeOf Object Is IJDGeometry Then
            Equipment.GetMatrix oEqpMatrix
            Set oTransform = Object
            oTransform.DTransform oEqpMatrix
            Set oTransform = Nothing
            Set oEqpMatrix = Nothing
        End If
    End If
    
    Set oTransform = Nothing
    Set oEqpMatrix = Nothing
    Set oShape = Nothing
    Set oShapeMatrix = Nothing
    
    Exit Sub
ErrorHandler:
    Set oTransform = Nothing
    Set oEqpMatrix = Nothing
    Set oShape = Nothing
    Set oShapeMatrix = Nothing
    HandleError MODULE, METHOD
End Sub

Private Function IsFree(Shape As IJShape) As Boolean
    Const METHOD = "IsFree"
    On Error GoTo ErrorHandler
    Dim oAttributes As IJDAttributes
    Dim oAttribute As IJDAttribute
    
    IsFree = False
    Set oAttributes = Shape
    Set oAttribute = oAttributes.CollectionOfAttributes("IJTransformableMember").Item("IsAssociative")
    IsFree = CBool(oAttribute.Value)
    
    Set oAttribute = Nothing
    Set oAttributes = Nothing
    
    If IsFree = True Then
        LogCalls METHOD & " = True"
    Else
        LogCalls METHOD & " = False"
    End If
    
    Exit Function
ErrorHandler:
    Set oAttribute = Nothing
    Set oAttributes = Nothing
    HandleError MODULE, METHOD
End Function

Private Sub PositionAndOrientDP1(Equipment As IJEquipment, Shape As IJShape)
    Dim oDatumShapeLCS As IJLocalCoordinateSystem
    
    'We want:
    'the X (primary) of the shape on the Z or Elevation of the equipment (ECS)
    'the Y (secondary) of the shape on the Y or North of the equipment (ECS)
    Set oDatumShapeLCS = Shape
    oDatumShapeLCS.XAxis = m_oEast
    oDatumShapeLCS.YAxis = m_oNorth
    Shape.SetOrigin 0, 0, 0
    
    'The shape is persisted in GCS
    TransformFromECStoGCS Equipment, Shape
    
    Set oDatumShapeLCS = Nothing

End Sub

Public Sub CMConstructTyACDN(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructTyACDN"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO
        
    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "FrontEndTypeACDN 01-EC", "TyACDN")
    TransformFromECStoGCS oDesignEquipment, pObject
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructTyACDN(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructTyACDN"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsTyACDN(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsTyACDN"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    GetDimensionsFromSymbolArray pMemberDesc.CAO
'    If m_bComputeEqpComp = True Then
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeDia").Value = m_dFrontEndFlangeDia
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeTk1").Value = m_dFrontEndFlangeTk1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndLength1").Value = m_dFrontEndLength1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndLength2").Value = m_dFrontEndLength2
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeTk2").Value = m_dFrontEndFlangeTk2
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeTk3").Value = m_dFrontEndFlangeTk3
    oAttribs.CollectionOfAttributes("IJUAExchangerChannel").Item("ChannelDiameter").Value = m_dExchangerNeckDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    m_bComputeEqpComp = False
    
'    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateTyACDN(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateTyACDN"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryTyACDN(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryTyACDN"
    On Error GoTo ErrorHandler
    LogCalls METHOD
   
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalTyACDN(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalTyACDN"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dFrontEndType = FrontEndType_A Or m_dFrontEndType = FrontEndType_C Or _
        m_dFrontEndType = FrontEndType_D Or m_dFrontEndType = FrontEndType_N Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMReleaseTyACDN(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseTyACDN"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConstructTyB(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructTyB"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO
     
    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "FrontEndTypeB 01-EC", "TyB")
    TransformFromECStoGCS oDesignEquipment, pObject
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructTyB(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructTyB"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsTyB(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsTyB"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    GetDimensionsFromSymbolArray pMemberDesc.CAO
'        If m_bComputeEqpComp = True Then
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject

    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeDia").Value = m_dFrontEndFlangeDia
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeTk1").Value = m_dFrontEndFlangeTk1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndLength1").Value = m_dFrontEndLength1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndLength2").Value = m_dFrontEndLength2
    oAttribs.CollectionOfAttributes("IJUAExchangerChannel").Item("ChannelDiameter").Value = m_dExchangerNeckDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness

    m_bComputeEqpComp = False
    '    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateTyB(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateTyB"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryTyB(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryTyB"
    On Error GoTo ErrorHandler
    LogCalls METHOD
   
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalTyB(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalTyB"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dFrontEndType = FrontEndType_B Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub


Public Sub CMReleaseTyB(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseTyB"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
Public Sub CMConstructTyQ(ByVal pMemberDescription As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef pObject As Object)
    Const METHOD = "CMConstructTyQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    Dim oEqpComponent As IJEquipmentComponent
    Dim oDesignEquipment As IJDesignEquipment
    Set oDesignEquipment = pMemberDescription.CAO
    
    'Create Equipment Component
    Set pObject = m_oEquipCADHelper.CreateEquipmentComponent(pMemberDescription, pResourceManager, "FrontEndTypeQ 01-EC", "TyQ")
    TransformFromECStoGCS oDesignEquipment, pObject
    Set oDesignEquipment = Nothing
    m_bComputeEqpComp = True
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructTyQ(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMFinalConstructTyQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler

    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMSetInputsTyQ(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMSetInputsTyQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
'        If m_bComputeEqpComp = True Then
    Dim oAttribs As IJDAttributes
    Dim oSmartOcc As IJSmartOccurrence
    
    Set oSmartOcc = pMemberDesc.Object
    Set oAttribs = oSmartOcc '.ItemObject
    
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeDia").Value = m_dFrontEndFlangeDia
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndFlangeTk1").Value = m_dFrontEndFlangeTk1
    oAttribs.CollectionOfAttributes("IJUAHeatExchangerFrontEnd").Item("FrontEndLength1").Value = m_dFrontEndLength1
    oAttribs.CollectionOfAttributes("IJUAExchangerChannel").Item("FrontEndDiameter").Value = m_dFrontEndDiameter
    oAttribs.CollectionOfAttributes("IJUAExchangerChannel").Item("ChannelDiameter").Value = m_dExchangerNeckDiameter
    oAttribs.CollectionOfAttributes("IJInsulationThickness").Item("InsulationThickness").Value = m_dInsulationThickness
    
    m_bComputeEqpComp = False
    
'    End If
 
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateTyQ(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateTyQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateGeometryTyQ(ByVal oPropertyDescription As IJDPropertyDescription, pObject As Object)
    Const METHOD = "CMEvaluateGeometryTyQ"
    On Error GoTo ErrorHandler
    LogCalls METHOD
   
    Exit Sub
    
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub CMConditionalTyQ(ByVal pMemberDesc As IJDMemberDescription, ByRef IsNeeded As Boolean)
    Const METHOD = "CMConditionalTyQ"
    LogCalls METHOD
    On Error GoTo ErrorHandler
    
    GetDimensionsFromSymbolArray pMemberDesc.CAO
    
    If m_dFrontEndType = FrontEndType_Q Then
        IsNeeded = True
        IsNeeded = m_oEquipCADHelper.CheckMemberConditional(pMemberDesc)
    Else
        IsNeeded = False
    End If
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
Public Sub CMReleaseTyQ(ByVal pMemberDesc As IJDMemberDescription)
    Const METHOD = "CMReleaseTyQ"
    On Error GoTo ErrorHandler
    
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub





