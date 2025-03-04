VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BUFlat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2008, Intergraph Corporation. All rights reserved.
'
'File
'    BUFlat.cls
'
'Author
'       Will Rosa 28 Feb 2008
'
'Description
'       Built Up Flat Definitions
'
'Notes
'
'History:
'
'   22-Sept-2009 GG TR#167167 - DesignedMember does not set PG of plates to its own PG
'   06-Aug-2010 GG DI-169828  SCHEMA: Interface to indicate member symmetry and open/closed cross sections
'*******************************************************************
'                   <-- Flat
'                       Thickness ->
'                  (Z)
'                   |   __           .
'                   |  |  |         / \
'                   |  |  |          |
'                   |  |  |          D
'                   |  |  |          e
'                   |  |  |          p
'                   |  |  |          t
'                   |  |  |          h
'                   |  |  |          |
'                   |  |  |          |
'                   |  |__|         \ /
' (Y) --------------+----------------. (-Y)
'                 / |
'               /   |
'             /     |
'           /       |
'         /         |
'       /           |
'     (-X)          |
'                   |
'                  (-Z)
'
'  (X) is into the paper, direction of the extrusion
'  (0,0) is Cardinal Point = 1
'
'   NOTE: Cardinal Pts are based on initial thickness, if user changes plate web
'   thickness, then cardinal points will not be correct!!!
  
Option Explicit

Private Const MODULE = "BUFlat"
Private Const strSourceFile = "BUFlat.def"
Private Const CONST_ItemProgId As String = "SM3DBUFlat.BUFlat"

Private m_strErrorDescr As String
Private m_oLocalizer As IJLocalizer
Private m_oDesignMemberHelper As BUHelperUtils
Private m_oCalcXProps As BUCalcSectionProperties

Private Enum BUFMembers
  Web = 3
End Enum

Implements ICustomSectionShapeService
Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements ISPSDesignedMemberHelper

Private Function IJDUserSymbolServices_EditOccurence(pSymbolOccurrence As Object, ByVal pTransactionMgr As Object) As Boolean
     IJDUserSymbolServices_EditOccurence = False
End Function

Private Function IJDUserSymbolServices_GetDefinitionName(ByVal definitionParameters As Variant) As String
     IJDUserSymbolServices_GetDefinitionName = CONST_ItemProgId
End Function

Private Sub IJDUserSymbolServices_InitializeSymbolDefinition(pDefinition As IMSSymbolEntities.IJDSymbolDefinition)
Const METHOD = "IJDUserSymbolServices_InitializeSymbolDefinition"
On Error GoTo ErrorHandler
     pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
     pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA

     ' Define the inputs -
     Dim pIH As IJDInputsHelper
     Set pIH = New InputHelper
     pIH.definition = pDefinition
     pIH.SetInput "DefiningCurve"
     
     ' Aggregator Type
     Dim pAD As IJDAggregatorDescription
     Set pAD = pDefinition
     pAD.AggregatorClsid = "{F4CDE773-A760-4561-A43A-D44A9C8340A7}" 'CSPSDesignedMember
     pAD.UserTypeClsid = "{308AF706-D414-4F71-968F-3C762119D2B4}"   'BUFlat (Generated here)
     pAD.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructAsm"
     pAD.SetCMConstruct imsCOOKIE_ID_USS_LIB, "CMConstructAsm"
     pAD.SetCMSetInputs -1, -1
     pAD.SetCMRemoveInputs -1, -1
     Set pAD = Nothing
           
     ' Aggregator property
     Dim pAPDs As IJDPropertyDescriptions
     Set pAPDs = pDefinition
     pAPDs.RemoveAll ' Remove all the previous property descriptions
     pAPDs.AddProperty "IStructCrossSection", 1, IStructCrossSection, , imsCOOKIE_ID_USS_LIB
     pAPDs.AddProperty "IStructCrossSectionDimensions", 2, IStructCrossSectionDimensions, "CMEvaluateCAO", imsCOOKIE_ID_USS_LIB
     pAPDs.AddProperty "IStructCrossSectionDesignProperties", 3, IStructCrossSectionDesignProperties, "CMEvaluateCAO", imsCOOKIE_ID_USS_LIB
     pAPDs.AddProperty "ISPSDesignedMemberDesignNotifyInput", 4, "ISPSDesignedMemberDesignNotifyInput", "CMEvaluateCAO1", imsCOOKIE_ID_USS_LIB
     pAPDs.AddProperty "IUABuiltUpLengthExt", 5, IID_IUABuiltUpLengthExt, , imsCOOKIE_ID_USS_LIB
     pAPDs.AddProperty "IUABuiltUpWeb", 6, IID_IUABuiltUpWeb, , imsCOOKIE_ID_USS_LIB
     pAPDs.AddProperty "IUABuiltUpCompute", 7, IID_IUABuiltUpCompute, , imsCOOKIE_ID_USS_LIB

     Set pAPDs = Nothing
               
     ' Define the members
     Dim pMemberDescriptions As IJDMemberDescriptions
     Dim pMemberDescription As IJDMemberDescription
     Dim pPropertyDescriptions As IJDPropertyDescriptions
     Set pMemberDescriptions = pDefinition
     
     ' Remove all the previous member descriptions
     pMemberDescriptions.RemoveAll
     
'''''
'' Web Members 1,2,3
'''''
     Set pMemberDescription = pMemberDescriptions.AddMember("WebCurveToExtrude", 1, "CMConstructWebCurveToExtrude", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "TopWebToExtrudeProperties", 1, IJCurve, "CMComputeWebCurveToExtrude", imsCOOKIE_ID_USS_LIB

     Set pMemberDescription = pMemberDescriptions.AddMember("WebExtrusionPath", 2, "CMConstructWebExtrusionPath", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "WebExtrusionPathProperties", 1, IJCurve, "CMComputeWebExtrusionPath", imsCOOKIE_ID_USS_LIB

    ' Web
     Set pMemberDescription = pMemberDescriptions.AddMember("Web", 3, "CMConstructWebPlateSystem", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructWebPlateSystem"
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateWebPlateSystem"
          
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "WebProperties", 1, IJCurve, "CMComputeWebPlateSystem", imsCOOKIE_ID_USS_LIB
        
     Set pMemberDescriptions = Nothing
     Set pMemberDescription = Nothing
     Set pPropertyDescriptions = Nothing
     
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Private Function IJDUserSymbolServices_InstanciateDefinition(ByVal CodeBase As String, ByVal defParams As Variant, ByVal pResourceMgr As Object) As Object
' This method is in charge of the creation of the symbol definition object
Const METHOD = "IJDUserSymbolServices_InstanciateDefinition"
On Error GoTo ErrorHandler
     Dim pDefinition As IJDSymbolDefinition
     Dim pFact As IJCAFactory
     Set pFact = New CAFactory
     Set pDefinition = pFact.CreateCAD(pResourceMgr)
     
     ' Set definition progId and codebase
     pDefinition.ProgId = CONST_ItemProgId
     pDefinition.CodeBase = CodeBase
     
     ' Initialize the definition
     IJDUserSymbolServices_InitializeSymbolDefinition pDefinition
     pDefinition.Name = IJDUserSymbolServices_GetDefinitionName(defParams)
     
     ' Persistence behavior
     pDefinition.SupportOnlyOption = igSYMBOL_NOT_SUPPORT_ONLY
     pDefinition.MetaDataOption = igSYMBOL_DYNAMIC_METADATA
     
     'returned symbol definition
     Set IJDUserSymbolServices_InstanciateDefinition = pDefinition
  
Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

'-------------------------------------------------------
'                       Web
'-------------------------------------------------------
Public Sub CMConstructWebCurveToExtrude(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructWebCurveToExtrude"
On Error GoTo ErrorHandler
    
    'Create a Line/Curve to use
    Dim pCurve As IJCurve
   
    m_oDesignMemberHelper.CreateCurveBy2Points pResourceManager, 0, 0, 0, 0, 0, 1, pCurve
    CopyPermissionGroup pCurve, pMemberDescription.CAO

    Set pObj = pCurve
    
    Dim oControlFlags As IJControlFlags
    Set oControlFlags = pObj
    oControlFlags.ControlFlags(&H4) = &H4
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMComputeWebCurveToExtrude(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeWebCurveToExtrude"
On Error GoTo ErrorHandler

    Dim dDepth As Double
    Dim dThickness As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If

    GetFlatExtrusionParameters oSmartOcc, dDepth, dThickness, dLength

    Dim pIJLine As IJLine
    Set pIJLine = pObject
  
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDepth, dThickness, oMatrix
        
    Dim dWebExtension As Double
    m_oDesignMemberHelper.GetWebDepthExtension oSmartOcc, dWebExtension
    
    pIJLine.DefineBy2Points 0, -dThickness / 2, 0 - dWebExtension, _
                            0, -dThickness / 2, dDepth + dWebExtension

    pIJLine.Transform oMatrix
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
Err.Raise E_FAIL
End Sub

Public Sub CMConstructWebExtrusionPath(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructWebExtrusionPath"
On Error GoTo ErrorHandler
 
    'Create a Line/Curve to use
    Dim pCurve As IJCurve
    
    m_oDesignMemberHelper.CreateCurveBy2Points pResourceManager, 0, 0, 0, 1, 0, 0, pCurve
    CopyPermissionGroup pCurve, pMemberDescription.CAO
    
    Set pObj = pCurve
    
    Dim oControlFlags As IJControlFlags
    Set oControlFlags = pObj
    oControlFlags.ControlFlags(&H4) = &H4
        
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMComputeWebExtrusionPath(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeWebExtrusionPath"
On Error GoTo ErrorHandler

    Dim dDepth As Double
    Dim dThickness As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
        
    GetFlatExtrusionParameters oSmartOcc, dDepth, dThickness, dLength
    
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDepth, dThickness, oMatrix

    Dim pIJLine As IJLine
    Set pIJLine = pObject
    
    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt

    pIJLine.DefineBy2Points 0 - dLengthExt, -dThickness / 2, dDepth, _
                           dLength + dLengthExt, -dThickness / 2, dDepth
  
    pIJLine.Transform oMatrix
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
Err.Raise E_FAIL
End Sub

' Construction Method for the Web
Public Sub CMConstructWebPlateSystem(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructWebPlateSystem"
On Error GoTo ErrorHandler
    
    ' Setup the Smart Item
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Set oSmartOcc = pMemberDescription.CAO
    Set oSmartItem = oSmartOcc.ItemObject
  
    ' Get the default Attributes
    Dim dThickness As Double
    Dim strMaterial As String
    Dim strGrade As String
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
   
    Set oAttr = oSmartItem
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpWeb)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("WebThickness").Value
        strMaterial = oAttrCol.Item("WebMaterial").Value
        strGrade = oAttrCol.Item("WebGrade").Value
    Else
        GoTo ErrorHandler
    End If
        
    Dim pIJDMemberObject As IJDMemberObjects
    Set pIJDMemberObject = oSmartOcc

    ' create the curves to extrude
    Dim pCurveToExtrude As IJCurve
    Dim pExtrusionPath As IJCurve

    Set pCurveToExtrude = pIJDMemberObject.ItemByDispid(1)
    Set pExtrusionPath = pIJDMemberObject.ItemByDispid(2)
    
    Set pObj = m_oDesignMemberHelper.CreateABuitUpPlate(pResourceManager, pCurveToExtrude, _
                                                        pExtrusionPath, dThickness, _
                                                        strMaterial, strGrade, oSmartOcc, WebPlate, _
                                                        , , , Centered, MoldedFormPlateNameCat.NameCatWeb)

    m_oDesignMemberHelper.SetPlateBoundaries oSmartOcc, pObj
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructWebPlateSystem(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructWebPlateSystem"
On Error GoTo ErrorHandler
 
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMComputeWebPlateSystem(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeWebPlateSystem"
On Error GoTo ErrorHandler
    ' get the web plate
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    Dim pIJDMemberObject As IJDMemberObjects
    Set pIJDMemberObject = oSmartOcc
    Dim oWebPlate As IJPlate
    Set oWebPlate = pIJDMemberObject.ItemByDispid(Web)
    
    ' get the parameters
    Dim oSmartItem As IJSmartItem
    Set oSmartItem = oSmartOcc.ItemObject
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
   
    Dim dThickness As Double
    Dim strMaterial As String
    Dim strGrade As String
    Set oAttr = oSmartItem
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpWeb)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("WebThickness").Value
        strMaterial = oAttrCol.Item("WebMaterial").Value
        strGrade = oAttrCol.Item("WebGrade").Value
    Else
        GoTo ErrorHandler
    End If
    
    ' apply the new parameters
    With m_oDesignMemberHelper
        .SetMaterialAndGrade oWebPlate, strMaterial, strGrade
        .SetPlateDimensions oWebPlate, dThickness
    End With

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMMigrateWebPlateSystem(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateWebPlateSystem"
On Error GoTo ErrorHandler

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructAsm(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMFinalConstructAsm"
On Error GoTo ErrorHandler
  
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

Public Sub CMConstructAsm(pAggregatorDescription As IJDAggregatorDescription)
Const METHOD = "CMConstructAsm"
On Error GoTo ErrorHandler

  
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateCAO(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAO"
On Error GoTo ErrHandler

    CalcuteOutputParameters pPropertyDescriptions.CAO
    
Exit Sub
ErrHandler: HandleError MODULE, METHOD
End Sub

Public Sub CMEvaluateCAO1(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMEvaluateCAO1"
On Error GoTo ErrHandler
    
    Dim oSmartOcc As IJSmartOccurrence
    Dim oDesMem As ISPSDesignedMember
    
    Set oSmartOcc = pPropertyDescriptions.CAO
    Set oDesMem = oSmartOcc
    
    oDesMem.NotifyDesignChange
    
Exit Sub
ErrHandler: HandleError MODULE, METHOD
End Sub

Public Sub ISPSDesignedMemberHelper_GetNominalSectionSize(ByVal pIDesignedMember As ISPSDesignedMember, _
                                                          ByVal pPosAlong As IJDPosition, _
                                                          ByRef pdWidth As Double, _
                                                          ByRef pdDepth As Double)

Const METHOD = "ISPSDesignedMemberHelper_GetNominalSectionSize"
On Error GoTo ErrHandler
    
    Dim dLength As Double
    Dim dDepth As Double
    Dim dThickness As Double
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pIDesignedMember
    
    GetFlatExtrusionParameters oSmartOcc, dDepth, dThickness, dLength
   
    pdWidth = dThickness
    pdDepth = dDepth
   
Exit Sub
ErrHandler: HandleError MODULE, METHOD
End Sub

Private Function IJUserAttributeMgmt_OnAttributeChange(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object, ByVal pAttrToChange As IJAttributeDescriptor, ByVal varNewAttrValue As Variant) As String
Const METHOD = "IJUserAttributeMgmt_OnAttributeChange"
On Error GoTo ErrHandler

    ' Validate the attribute new value first before any further processing
    Dim ErrStr As String

    Dim oAttrCol As IJDInfosCol
    Set oAttrCol = Nothing
    Set oAttrCol = m_oDesignMemberHelper.GetInfosCollection(pIJDAttrs, pAttrToChange.InterfaceName, pAttrToChange.AttrName)

    If oAttrCol Is Nothing Then
        GoTo ErrHandler
    End If
    
    Dim oAttrObj As IJDAttributeInfo
    Dim oInterfaceInfo As IJDInterfaceInfo
    Dim AttrCount As Long
    Dim AttrType As Long
    
    ' loop on the attributes on the interface to match the supplied attribute type
    For AttrCount = 1 To oAttrCol.Count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        If oAttrObj.Name = pAttrToChange.AttrName Then
            Select Case oAttrObj.Type
                Case m_oDesignMemberHelper.DoubleValue
                    ErrStr = BuiltUpDefValidate(pAttrToChange.AttrName, varNewAttrValue, 0#)
                    If Len(ErrStr) > 0 Then
                        IJUserAttributeMgmt_OnAttributeChange = ErrStr
                        Exit Function
                    End If
                    
                    Dim oDesMem As ISPSDesignedMember
                    Set oDesMem = pIJDAttrs
                    If Not oDesMem Is Nothing And IsDesignParameter(pAttrToChange) Then
                        oDesMem.NotifyDesignChange
                    End If
                End Select
        End If
    Next
    
    IJUserAttributeMgmt_OnAttributeChange = ""
   
Exit Function
ErrHandler:
    IJUserAttributeMgmt_OnAttributeChange = m_oLocalizer.GetString(IDS_BUILTUP_ERROR, "ERROR")
    HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreCommit(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreCommit"
On Error GoTo ErrHandler

    IJUserAttributeMgmt_OnPreCommit = ""
    
Exit Function
ErrHandler: HandleError MODULE, METHOD
End Function

Private Function IJUserAttributeMgmt_OnPreLoad(ByVal pIJDAttrs As IJDAttributes, ByVal CollAllDisplayedValues As Object) As String
Const METHOD = "IJUserAttributeMgmt_OnPreLoad"
On Error GoTo ErrHandler
    
    Dim oAttrCol As IJDAttributesCol
    
    'set the common read only attributes.  The flat does not have additional attributes that are
    'dependent on the IsModifiable flag
    m_oDesignMemberHelper.SetCrossSectionReadOnlyAttributes pIJDAttrs, CollAllDisplayedValues
    
    Dim bSectionProperties As Boolean
    bSectionProperties = m_oDesignMemberHelper.AreSectionPropertiesModifiable(pIJDAttrs)
    
    Dim i As Integer
     
    If bSectionProperties = False Then
        Dim pAttrColl As Collection
        Dim pAttrDescr As IJAttributeDescriptor
        Set pAttrColl = CollAllDisplayedValues
        For i = 1 To pAttrColl.Count
            Set pAttrDescr = pAttrColl.Item(i)
            If pAttrDescr.InterfaceName = "IStructCrossSectionDesignProperties" _
                Or pAttrDescr.InterfaceName = "IStructCrossSectionUnitWeight" Then
                pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
            End If
        Next
    End If
    
    IJUserAttributeMgmt_OnPreLoad = ""
    
Exit Function
ErrHandler: HandleError MODULE, METHOD
End Function
Private Sub Class_Initialize()
    Set m_oLocalizer = New IMSLocalizer.Localizer
    m_oLocalizer.Initialize App.Path & "\" & "SPSDesignedMemberDefs"

    Set m_oCalcXProps = New BUCalcSectionProperties
    Set m_oDesignMemberHelper = New BUHelperUtils
End Sub

Private Sub Class_Terminate()
    Set m_oLocalizer = Nothing
    Set m_oCalcXProps = Nothing
    Set m_oDesignMemberHelper = Nothing
End Sub

Private Sub GetFlatExtrusionParameters(ByVal oSmartOcc As IJSmartOccurrence, _
                                           ByRef dDepth As Double, _
                                           ByRef dThickness As Double, _
                                           ByRef dLength As Double)
                                           
Const METHOD = "GetFlatExtrusionParameters"
On Error GoTo ErrorHandler
    
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    Dim pIJDMemberObject As IJDMemberObjects
    Dim oSmartItem As IJSmartItem
    
    ' Get Parameters of BU
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    If Not oAttrCol Is Nothing Then
        dDepth = oAttrCol.Item("Depth").Value
    Else
        GoTo ErrorHandler
    End If
    
    Set oAttr = Nothing
    Set oAttrCol = Nothing
    Set oSmartItem = oSmartOcc.ItemObject
    Set oAttr = oSmartItem
    ' check to see if plate exist yet
    If Not oAttr Is Nothing Then
        Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpWeb)
        If Not oAttrCol Is Nothing Then
            dThickness = oAttrCol.Item("WebThickness").Value
        Else
            GoTo ErrorHandler
        End If
    End If

    
    ' Get Length of input curve
    Dim pISPSMemberPartCommon As ISPSMemberPartCommon
    Set pISPSMemberPartCommon = oSmartOcc
    dLength = pISPSMemberPartCommon.Axis.Length

Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

Public Function AreOccurrencePropertiesValid(ByVal oSmartOcc As IJSmartOccurrence) As Boolean
Const METHOD = "AreOccurrencePropertiesValid"
On Error GoTo ErrorHandler
    
    AreOccurrencePropertiesValid = False
    
    Dim dDepth As Double
    Dim dThickness As Double
    Dim dLength As Double
    Dim sError As String
        
    GetFlatExtrusionParameters oSmartOcc, dDepth, dThickness, dLength
    
    sError = BuiltUpDefValidate("Depth", dDepth, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("Thickness", dThickness, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("Length", dLength, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt
    
    sError = BuiltUpDefValidate("LengthExt", dLengthExt, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If

    Dim dWebExtension As Double
    m_oDesignMemberHelper.GetWebDepthExtension oSmartOcc, dWebExtension
    
    sError = BuiltUpDefValidate("DepthExt", dWebExtension, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
  
    AreOccurrencePropertiesValid = True

Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Sub CalcuteOutputParameters(ByVal oSmartOcc As IJSmartOccurrence)
Const METHOD = "CalcuteOutputParameters"
On Error GoTo ErrorHandler
        
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    Dim oSmartItem As IJSmartItem
    Dim dArea As Double
    Dim dPerimeter As Double
    Dim dDepth As Double
    Dim dWebThickness As Double
           
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    If Not oAttrCol Is Nothing Then
        dDepth = oAttrCol.Item("Depth").Value
    Else
        GoTo ErrorHandler
    End If
    
    Set oAttr = Nothing
    Set oAttrCol = Nothing
    Set oSmartItem = oSmartOcc.ItemObject
    Set oAttr = oSmartItem
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpWeb)
    If Not oAttrCol Is Nothing Then
        dWebThickness = oAttrCol.Item("WebThickness").Value
    Else
        GoTo ErrorHandler
    End If
            
    Set oAttr = Nothing
    Set oAttrCol = Nothing
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    If Not oAttrCol Is Nothing Then
        dArea = dDepth * dWebThickness
        dPerimeter = dWebThickness * 2 + dDepth * 2
        
        oAttrCol.Item("Depth").Value = dDepth
        oAttrCol.Item("Width").Value = dWebThickness
        oAttrCol.Item("Area").Value = dArea
        oAttrCol.Item("Perimeter").Value = dPerimeter
    Else
        GoTo ErrorHandler
    End If
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDesignProperties)
    If Not oAttrCol Is Nothing Then
        On Error Resume Next
        oAttrCol.Item("IsHollow").Value = False
        oAttrCol.Item("IsSymmetricAboutX").Value = True
        oAttrCol.Item("IsSymmetricAboutY").Value = True
        On Error GoTo ErrorHandler
    End If
    
      
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

Private Function BuiltUpDefValidate(sAttributeName As String, _
                                    varAttributeValue As Variant, _
                                    dLowRange As Double, _
                                    Optional dUpRange As Double = -1908) As String
                                    '-1908 is an arbritraty value any negative number would do
                                    '1908 is the last time the Chicago Cubs won a world series
                                    
Const METHOD = "BuiltUpDefValidate"
On Error GoTo ErrorHandler

    If dUpRange > 0 Then
         If (varAttributeValue < dLowRange) Or ((varAttributeValue - dUpRange) > m_oDesignMemberHelper.distTol) Then
            BuiltUpDefValidate = sAttributeName & ": " & m_oLocalizer.GetString(IDS_BUILTUP_VALUE_MUSTBE_IN_RANGE, _
                                 "Value must be within range of " & dLowRange & " and " & dUpRange) _
                                 & " [" & dLowRange & ", " & dUpRange & "]"
            Exit Function
        End If
    End If
    
    Select Case sAttributeName
        Case "Depth", "Length", "Thickness"
            If varAttributeValue <= 0# Then
                BuiltUpDefValidate = sAttributeName & ": " & m_oLocalizer.GetString(IDS_BUILTUP_VALUE_MUSTBE_POSITIVE, "Value must be > 0")
                Exit Function
            End If
        Case Else
            If (varAttributeValue < dLowRange) Then
                BuiltUpDefValidate = sAttributeName & ": " & m_oLocalizer.GetString(IDS_BUILTUP_VALUE_MUSTBE_GREATERTHAN_OR_EQUAL_TO_ZERO, "Value must be >= 0")
                Exit Function
            End If
    End Select

Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function

Public Sub ISPSDesignedMemberHelper_LoadEmulatedFacePorts(ByVal oDesignedMember As SPSMembers.ISPSDesignedMember)
Const METHOD = "ISPSDesignedMemberHelper_LoadEmulatedFacePorts"
On Error GoTo ErrorHandler

    Dim oAssyMembers     As IJDMemberObjects
    Dim oPlateSystem     As IJPlateSystem
    Dim eSectionAlias As EnumSectionTypeAlias
    
    eSectionAlias = Section_Flat
    'get assy members from SO
    Set oAssyMembers = oDesignedMember
    
    'handle the web
    Set oPlateSystem = oAssyMembers.Item(BUFMembers.Web)
    If Not oPlateSystem Is Nothing Then
        m_oDesignMemberHelper.EvaluateWebPlateSystemXIDs oPlateSystem, StructBUPlateType.Web, eSectionAlias, oDesignedMember
    End If
    
    Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

Public Sub ISPSDesignedMemberHelper_ResolveAmbiguity(ByVal pDesignedMember As SPSMembers.ISPSDesignedMember)
Const METHOD = "ISPSDesignedMemberHelper_ResolveAmbiguity"
On Error GoTo ErrorHandler
    'Nothing required for this particular Built-Up Definition
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

' check lets the caller decide to take a certain action for changes in deaign parameters, like notification
' to the listeners
Private Function IsDesignParameter(ByVal oAttrDesc As IJAttributeDescriptor) As Boolean
Const METHOD = "IsDesignParameter"
On Error GoTo ErrorHandler

    Dim sAttrName As String
    Dim sIntfName As String
    
    If Not oAttrDesc Is Nothing Then
    
        IsDesignParameter = False
        sAttrName = oAttrDesc.AttrName
        sIntfName = oAttrDesc.InterfaceName
        
        Select Case sIntfName
            Case "IStructCrossSectionDimensions"
                If sAttrName = "Depth" Then
                    IsDesignParameter = True
                End If
            Case Else
        End Select ' I/F name
    End If

Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub ICustomSectionShapeService_GetCrossSectionData(ByVal pExtrusionObject As Object, ByVal distFromStart As Double, ByVal eRepresentationType As SP3DStructInterfaces.structShapeRepresentationEnum, ByVal bBreakIntoComponents As Boolean, sectionPrev As SP3DStructInterfaces.IStructSectionShape, sectionNext As SP3DStructInterfaces.IStructSectionShape)
Const METHOD = "IStructSectionShapeService_GetCrossSectionData"
On Error GoTo ErrorHandler
    
    ' we're only creating a section_profile_i_type for now.  There is no use case for creating a compound shape
    ' or a graphic shape yet and no way for the user to ask for something different.
    
    ' in the future, utility methods will be implemented to convert to a graphic shape and to decompose the shape
    ' if there is a need to do this

    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pExtrusionObject
    
    Dim depth As Double
    Dim width As Double
    ISPSDesignedMemberHelper_GetNominalSectionSize pExtrusionObject, Nothing, width, depth

    Dim sectionMaterial As IStructSectionMaterial
    Dim Mat As String
    Dim Grade As String
    Dim Thickness As Double
   
    GetComponentMaterial pExtrusionObject, IID_IUABuiltUpWeb, Mat, Grade, Thickness
    
    ' always create a parametric shape.  It may be used to create the other shapes as needed
    Dim parametricShape As IStructSectionRectangleSolidShape
    Set parametricShape = New TransientClassesLib.StructSectionRectangleSolidShape
    parametricShape.SetProperties depth, width, 0#
    Set sectionMaterial = parametricShape
    sectionMaterial.SetMaterial Mat, Grade
   

    Set sectionNext = parametricShape
            
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub ICustomSectionShapeService_GetTransitionLocations(ByVal pProfileObject As Object, distFromStart() As Double)
    'the section size for this type do not change along the length
    'so just add the value of 0, representing the start of the member, to the array.
    ReDim distFromStart(0 To 0)
    
    distFromStart(0) = 0#
End Sub
    
