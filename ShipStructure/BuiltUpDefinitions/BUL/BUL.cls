VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BUL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2008, Intergraph Corporation. All rights reserved.
'
'File
'    BUL.cls
'
'Author
'       10 March 2008 Will Rosa
'
'Description
'       Built Up L definition
'
'Notes
'
'History:
'
'   22-Sept-2009 GG TR#167167 - DesignedMember does not set PG of plates to its own PG
'   06-Aug-2010 GG DI-169828  SCHEMA: Interface to indicate member symmetry and open/closed cross sections
'*******************************************************************
'
'                  (Z)   __                      .
'                   |   |  |                    / \
'                   |   |  |                     |
'                   |   |  |                     |
'                   |   |  |                     D
'                   |   |  |                     e
'                   |   |  |                     p
'                   |   |  |                     t
'                   |   |  |                     h
'                   | -----------------------    |
'                   ||BO                     |   |
'                   | -----------------------   \ /
' (Y) --------------+--------------------------- .(-Y)
'                 / |
'               /   |  <-- Botom Flge W ----->
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
'   BO - Bottom Web Offset
'
'   (-Y) Computations
'
'   Bottom Flange
'   Y = 0
'
'   Web
'   If(BO > 0) Y = -BO else Y = 0
'
'   NOTE1:  BO is based on moving the bottom flange according to the offset.
'
'
'   NOTE2: Cardinal Pts are based on initial thickness, if user changes plate web
'   thickness, then cardinal points will not be correct!!!
  
Option Explicit

Private Const MODULE = "BUL"
Private Const strSourceFile = "BUL.def"
Private Const CONST_ItemProgId As String = "SM3DBUL.BUL"

Private m_strErrorDescr As String
Private m_oLocalizer As IJLocalizer
Private m_oDesignMemberHelper As BUHelperUtils
Private m_oCalcXProps As BUCalcSectionProperties

Private Enum BULMembers
  BotFlange = 3
  Web = 6
End Enum

Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements ISPSDesignedMemberHelper
Implements ICustomSectionShapeService



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
     pAD.UserTypeClsid = "{26552791-07D7-42BA-ABD5-8576E4B82EF2}"   'CUBuiltUpL (Generated here)
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
     pAPDs.AddProperty "IUABuiltUpBottomFlange", 7, IID_IUABuiltUpBottomFlange, , imsCOOKIE_ID_USS_LIB
     pAPDs.AddProperty "IUABuiltUpL", 8, IID_IUABuiltUpL, , imsCOOKIE_ID_USS_LIB
     pAPDs.AddProperty "IUABuiltUpCompute", 9, IID_IUABuiltUpCompute, , imsCOOKIE_ID_USS_LIB
     
     Set pAPDs = Nothing

     ' Define the members
     Dim pMemberDescriptions As IJDMemberDescriptions
     Dim pMemberDescription As IJDMemberDescription
     Dim pPropertyDescriptions As IJDPropertyDescriptions
     Set pMemberDescriptions = pDefinition

     ' Remove all the previous member descriptions
     pMemberDescriptions.RemoveAll

'''''
'' Bottom Flange Members 1,2,3
'''''
     Set pMemberDescription = pMemberDescriptions.AddMember("BottomFlangeCurveToExtrude", 1, "CMConstructBotFlangeCurveToExtrude", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "BottomFlangeCurveToExtrudeProperties", 1, IJCurve, "CMComputeBotFlangeCurveToExtrude", imsCOOKIE_ID_USS_LIB

     Set pMemberDescription = pMemberDescriptions.AddMember("BottomFlangeExtrusionPath", 2, "CMConstructBotFlangeExtrusionPath", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "BottomFlangeExtrusionPathProperties", 1, IJCurve, "CMComputeBotFlangeExtrusionPath", imsCOOKIE_ID_USS_LIB

     ' Bottom flange
     Set pMemberDescription = pMemberDescriptions.AddMember("BottomFlange", 3, "CMConstructBottomFlangePlateSystem", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructBottomFlangePlateSystem"
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateBottomFlangePlateSystem"

     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "BottomFlangeProperties", 1, IJCurve, "CMComputeBottomFlangePlateSystem", imsCOOKIE_ID_USS_LIB

'''''
'' Web Members 4,5,6
'''''
     Set pMemberDescription = pMemberDescriptions.AddMember("WebCurveToExtrude", 4, "CMConstructWebCurveToExtrude", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "WebToExtrudeProperties", 1, IJCurve, "CMComputeWebCurveToExtrude", imsCOOKIE_ID_USS_LIB

     Set pMemberDescription = pMemberDescriptions.AddMember("WebExtrusionPath", 5, "CMConstructWebExtrusionPath", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "WebExtrusionPathProperties", 1, IJCurve, "CMComputeWebExtrusionPath", imsCOOKIE_ID_USS_LIB

     ' Web
     Set pMemberDescription = pMemberDescriptions.AddMember("Web", 6, "CMConstructWebPlateSystem", imsCOOKIE_ID_USS_LIB)
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

     ' Returned symbol definition
     Set IJDUserSymbolServices_InstanciateDefinition = pDefinition

Exit Function
ErrorHandler:  HandleError MODULE, METHOD
End Function

Private Sub IJDUserSymbolServices_InvokeRepresentation(ByVal pSymbolOccurrence As Object, ByVal pRepName As String, ByVal pOutputColl As Object, arrayOfInputs() As Variant)

End Sub

'-------------------------------------------------------
'                       Bottom Flange
'-------------------------------------------------------

Public Sub CMConstructBotFlangeCurveToExtrude(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructBotFlangeCurveToExtrude"
On Error GoTo ErrorHandler

    'Create a Line/Curve to use
    Dim pCurve As IJCurve

    m_oDesignMemberHelper.CreateCurveBy2Points pResourceManager, 0, 0, 0, 0, -1, 0, pCurve
    CopyPermissionGroup pCurve, pMemberDescription.CAO

    Set pObj = pCurve
    
    Dim oControlFlags As IJControlFlags
    Set oControlFlags = pObj
    oControlFlags.ControlFlags(&H4) = &H4

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMComputeBotFlangeCurveToExtrude(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeBotFlangeCurveToExtrude"
On Error GoTo ErrorHandler

    Dim dDepth As Double
    Dim dBottomWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If

    GetLExtrusionParameters oSmartOcc, dDepth, dBottomWidth, dLength

    Dim dOffsetBot As Double
    dOffsetBot = GetLOffsetWeb(oSmartOcc)

    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDepth, dBottomWidth, oMatrix

    Dim dBtmExt As Double
    m_oDesignMemberHelper.GetBottomWidthExtension oSmartOcc, dBtmExt

    Dim pIJLine As IJLine
    Set pIJLine = pObject

    pIJLine.DefineBy2Points 0, dBtmExt, 0, 0, -dBottomWidth - dBtmExt, 0

    pIJLine.Transform oMatrix

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
Err.Raise E_FAIL
End Sub

Public Sub CMConstructBotFlangeExtrusionPath(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructBotFlangeExtrusionPath"
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

Public Sub CMComputeBotFlangeExtrusionPath(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeBotFlangeExtrusionPath"
On Error GoTo ErrorHandler

    Dim dDepth As Double
    Dim dBottomWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If

    GetLExtrusionParameters oSmartOcc, dDepth, dBottomWidth, dLength

    Dim pIJLine As IJLine
    Set pIJLine = pObject

    Dim dOffsetBot As Double
    dOffsetBot = GetLOffsetWeb(oSmartOcc)

    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt

    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDepth, dBottomWidth, oMatrix

    pIJLine.DefineBy2Points -dLengthExt, -dBottomWidth, 0, _
                            dLength + dLengthExt, -dBottomWidth, 0

    pIJLine.Transform oMatrix

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
Err.Raise E_FAIL
End Sub

Public Sub CMConstructBottomFlangePlateSystem(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructBottomFlangePlateSystem"
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
    Dim dDepth As Double
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes

    Set oAttr = oSmartItem
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpBottomFlange)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("BottomFlangeThickness").Value
        strMaterial = oAttrCol.Item("BottomFlangeMaterial").Value
        strGrade = oAttrCol.Item("BottomFlangeGrade").Value
    Else
        GoTo ErrorHandler
    End If

    ' Create the curves to extrude
    Dim pCurveToExtrude As IJCurve
    Dim pExtrusionCurve As IJCurve

    ' Get the outputs
    Dim pIJDMemberObject As IJDMemberObjects
    Set pIJDMemberObject = oSmartOcc

    Set pCurveToExtrude = pIJDMemberObject.ItemByDispid(1)
    Set pExtrusionCurve = pIJDMemberObject.ItemByDispid(2)

    Set pObj = m_oDesignMemberHelper.CreateABuitUpPlate(pResourceManager, pCurveToExtrude, _
                                                        pExtrusionCurve, dThickness, _
                                                        strMaterial, strGrade, oSmartOcc, FlangePlate, _
                                                        , , , InDir, MoldedFormPlateNameCat.NameCatBFlange)

    m_oDesignMemberHelper.SetPlateBoundaries oSmartOcc, pObj

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructBottomFlangePlateSystem(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructBottomFlangePlateSystem"
On Error GoTo ErrorHandler

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub
Public Sub CMComputeBottomFlangePlateSystem(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeBottomFlangePlateSystem"
On Error GoTo ErrorHandler
    ' get the plate
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    Dim pIJDMemberObject As IJDMemberObjects
    Set pIJDMemberObject = oSmartOcc
    Dim oPlate As IJPlate
    Set oPlate = pIJDMemberObject.ItemByDispid(BotFlange)
    
    ' get the parameters
    Dim oSmartItem As IJSmartItem
    Set oSmartItem = oSmartOcc.ItemObject
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
   
    Dim dThickness As Double
    Dim strMaterial As String
    Dim strGrade As String
    Set oAttr = oSmartItem
     Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpBottomFlange)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("BottomFlangeThickness").Value
        strMaterial = oAttrCol.Item("BottomFlangeMaterial").Value
        strGrade = oAttrCol.Item("BottomFlangeGrade").Value
    Else
        GoTo ErrorHandler
    End If
    
    ' apply the new parameters
    With m_oDesignMemberHelper
        .SetMaterialAndGrade oPlate, strMaterial, strGrade
        .SetPlateDimensions oPlate, dThickness
    End With

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMMigrateBottomFlangePlateSystem(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateBottomFlangePlateSystem"
On Error GoTo ErrorHandler

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
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
    Dim dBottomWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If

    GetLExtrusionParameters oSmartOcc, dDepth, dBottomWidth, dLength

    Dim pIJLine As IJLine
    Set pIJLine = pObject

    Dim dOffsetBot As Double
    Dim y As Double
    dOffsetBot = GetLOffsetWeb(oSmartOcc)
    y = GetWebPlacement(dOffsetBot)

    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDepth, dBottomWidth, oMatrix

    Dim dWebExtension As Double
    m_oDesignMemberHelper.GetWebDepthExtension oSmartOcc, dWebExtension
    
    Dim trimExt As Double
    trimExt = dWebExtension
    If trimExt < 0.000001 Then
        trimExt = m_oDesignMemberHelper.BoundingExt
    End If
    
    pIJLine.DefineBy2Points 0, y, 0 - trimExt, 0, y, dDepth + dWebExtension

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
    Dim dBottomWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If

    GetLExtrusionParameters oSmartOcc, dDepth, dBottomWidth, dLength

    Dim pIJLine As IJLine
    Set pIJLine = pObject

    Dim dOffsetBot As Double
    Dim y As Double
    dOffsetBot = GetLOffsetWeb(oSmartOcc)
    y = GetWebPlacement(dOffsetBot)

    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt

    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDepth, dBottomWidth, oMatrix

    pIJLine.DefineBy2Points -dLengthExt, y, dDepth, dLength + dLengthExt, y, dDepth

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
    Dim dDepth As Double
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

    ' Create the curves to extrude
    Dim pCurveToExtrude As IJCurve
    Dim pExtrusionCurve As IJCurve

    Set pCurveToExtrude = pIJDMemberObject.ItemByDispid(4)
    Set pExtrusionCurve = pIJDMemberObject.ItemByDispid(5)

    Set pObj = m_oDesignMemberHelper.CreateABuitUpPlate(pResourceManager, pCurveToExtrude, _
                                                        pExtrusionCurve, dThickness, _
                                                        strMaterial, strGrade, oSmartOcc, WebPlate, _
                                                        , , , InDir, MoldedFormPlateNameCat.NameCatWeb)

    ' Add the boundaries
    Dim oBottomFlange As Object
    Set oBottomFlange = pIJDMemberObject.ItemByDispid(3)

    m_oDesignMemberHelper.SetPlateBoundaries oSmartOcc, pObj, oBottomFlange

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
Const METHOD = "CMComputeWebPlateSystem "
On Error GoTo ErrorHandler
    ' get the plate
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    Dim pIJDMemberObject As IJDMemberObjects
    Set pIJDMemberObject = oSmartOcc
    Dim oPlate As IJPlate
    Set oPlate = pIJDMemberObject.ItemByDispid(Web)
    
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
        .SetMaterialAndGrade oPlate, strMaterial, strGrade
        .SetPlateDimensions oPlate, dThickness
    End With
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMMigrateWebPlateSystem(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateWebPlateSystem "
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
    Dim dBottomWidth As Double
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pIDesignedMember
    
    GetLExtrusionParameters oSmartOcc, dDepth, dBottomWidth, dLength
    
    pdWidth = dBottomWidth
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
    Dim dBotWidth As Double
       
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pIJDAttrs
    
    m_oDesignMemberHelper.GetBottomWidth oSmartOcc, dBotWidth
       
    ' loop on the attributes on the interface to match the supplied attribute type
    For AttrCount = 1 To oAttrCol.Count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        If oAttrObj.Name = pAttrToChange.AttrName Then
            Select Case oAttrObj.Type
                Case m_oDesignMemberHelper.DoubleValue
                    If pAttrToChange.AttrName = "OffsetWeb" Then
                        ErrStr = BuiltUpDefValidate(pAttrToChange.AttrName, varNewAttrValue, 0#, dBotWidth)
                    Else
                        ErrStr = BuiltUpDefValidate(pAttrToChange.AttrName, varNewAttrValue, 0#)
                    End If
                    
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
    Dim bIsModifiable As Boolean
    Dim i As Integer
    Dim pAttrColl As Collection
    Dim pAttrDescr As IJAttributeDescriptor
    
    ' Setup the Smart Item
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Set oSmartOcc = pIJDAttrs
    Set oSmartItem = oSmartOcc.ItemObject
    
    Dim oAttr As IJDAttributes
    Set oAttr = oSmartItem
    
    bIsModifiable = m_oDesignMemberHelper.IsAttributeModifiable(oAttr)
    
    'set the common read only attributes
    m_oDesignMemberHelper.SetCrossSectionReadOnlyAttributes pIJDAttrs, CollAllDisplayedValues
    
    If bIsModifiable = False Then
        Set pAttrColl = CollAllDisplayedValues
        For i = 1 To pAttrColl.Count
            Set pAttrDescr = pAttrColl.Item(i)
            If pAttrDescr.InterfaceName = IID_IUABuiltUpBottomFlange Then
                If pAttrDescr.AttrName = "BottomFlangeWidth" Then
                    pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
                    Exit For 'get out of the loop since there is only one attribute to set
                End If
            End If
        Next
    End If
    
    Dim bSectionProperties As Boolean
    bSectionProperties = m_oDesignMemberHelper.AreSectionPropertiesModifiable(oSmartOcc)
    
    If bSectionProperties = False Then
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

Private Sub GetLExtrusionParameters(ByVal oSmartOcc As IJSmartOccurrence, _
                                           ByRef dDepth As Double, _
                                           ByRef dBottomWidth As Double, _
                                           ByRef dLength As Double)

Const METHOD = "GetLExtrusionParameters"
On Error GoTo ErrorHandler

    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes

    ' Get Parameters of BU
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    If Not oAttrCol Is Nothing Then
        dDepth = oAttrCol.Item("Depth").Value
    Else
        GoTo ErrorHandler
    End If

    Set oAttrCol = Nothing
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpBottomFlange)
    If Not oAttrCol Is Nothing Then
        dBottomWidth = oAttrCol.Item("BottomFlangeWidth").Value
    Else
        GoTo ErrorHandler
    End If

    ' Get Length of input curve
    Dim pISPSMemberPartCommon As ISPSMemberPartCommon
    Set pISPSMemberPartCommon = oSmartOcc
    dLength = pISPSMemberPartCommon.Axis.Length

Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

Private Function GetLOffsetWeb(ByVal oSmartOcc As IJSmartOccurrence) As Double

Const METHOD = "GetLOffsetWeb"
On Error GoTo ErrorHandler

    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    Dim dOffset As Double

    ' Get Parameters of BU
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpL)
    If Not oAttrCol Is Nothing Then
        dOffset = oAttrCol.Item("OffsetWeb").Value
    Else
        GoTo ErrorHandler
    End If

    GetLOffsetWeb = dOffset

Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function

Private Function GetWebPlacement(ByVal dWebOffset As Double) As Double

Const METHOD = "GetWebPlacement"
On Error GoTo ErrorHandler

    Dim y As Double

    If dWebOffset > 0 Then
       y = -dWebOffset
    Else
       y = 0
    End If

    GetWebPlacement = y

Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function

Public Function AreOccurrencePropertiesValid(ByVal oSmartOcc As IJSmartOccurrence) As Boolean
Const METHOD = "AreOccurrencePropertiesValid"

    On Error GoTo ErrorHandler
    AreOccurrencePropertiesValid = False
     
    Dim dDepth As Double
    Dim dBottomWidth As Double
    Dim dLength As Double
    Dim sError As String
        
    GetLExtrusionParameters oSmartOcc, dDepth, dBottomWidth, dLength
    
    sError = BuiltUpDefValidate("Depth", dDepth, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("BottomFlangeWidth", dBottomWidth, 0#)
    
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

    Dim dBtmExt As Double
    m_oDesignMemberHelper.GetBottomWidthExtension oSmartOcc, dBtmExt
    
    sError = BuiltUpDefValidate("BottomFlangeWidthExt", dBtmExt, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    Dim dBotOffset As Double
    dBotOffset = GetLOffsetWeb(oSmartOcc)
    
    sError = BuiltUpDefValidate("OffsetWeb", dBotOffset, 0#, dBottomWidth)
    
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
   
    Dim dBottomWidth As Double
    Dim dBottomThickness As Double
    Dim dDepth As Double
    Dim dWebThickness As Double
    Dim dBotOffset As Double
    
    Dim dArea As Double
    Dim dPerimeter As Double
    Dim dCentroidX As Double
    Dim dCentroidY As Double
    Dim dXp As Double
    Dim dYp As Double
    Dim dIxx As Double
    Dim dIyy As Double
    Dim dRo As Double
    Dim dRxx As Double
    Dim dRxy As Double
    Dim dRyy As Double
    Dim dSw As Double
    Dim dSxx As Double
    Dim dSyy As Double
    Dim dZxx As Double
    Dim dZyy As Double
    Dim dJ As Double
    Dim dCw As Double
    Dim dH As Double
    Dim dUnitWt As Double
    
    'Initialize all parameters to Zero (0)
    dArea = dPerimeter = dCentroidX = dCentroidY = dXp = dYp = dIxx = dIyy = _
    dRo = dRxx = dRxy = dRyy = dSw = dSxx = dSyy = dZxx = dZyy = dJ = dCw = dH = dUnitWt = 0#
                 
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    If Not oAttrCol Is Nothing Then
        dDepth = oAttrCol.Item("Depth").Value
    Else
        GoTo ErrorHandler
    End If
    
    Set oAttrCol = Nothing
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpBottomFlange)
    If Not oAttrCol Is Nothing Then
        dBottomWidth = oAttrCol.Item("BottomFlangeWidth").Value
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
    
    Set oAttrCol = Nothing
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpBottomFlange)
    If Not oAttrCol Is Nothing Then
        dBottomThickness = oAttrCol.Item("BottomFlangeThickness").Value
    Else
        GoTo ErrorHandler
    End If
    
    Dim ErrorStatus As ErrorSectionStatus
    ErrorStatus = GetSectionPropertiesFromCalculator(dArea, dPerimeter, dCentroidX, dCentroidY, _
                                                     dXp, dYp, dIxx, dIyy, dRo, dRxx, dRxy, dRyy, _
                                                     dSw, dSxx, dSyy, dZxx, dZyy, dJ, dCw, dH, dUnitWt, _
                                                     dDepth, dBottomWidth, dWebThickness, dBottomThickness, dBotOffset)
                                                     
    If ErrorStatus <> SectionProperties_OK Then
        GoTo ErrorHandler
    End If
            
    Set oAttr = Nothing
    Set oAttrCol = Nothing
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    If Not oAttrCol Is Nothing Then
        oAttrCol.Item("Depth").Value = dDepth
        oAttrCol.Item("Width").Value = dBottomWidth
        oAttrCol.Item("Area").Value = dArea
        oAttrCol.Item("Perimeter").Value = dPerimeter
    Else
        GoTo ErrorHandler
    End If
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDesignProperties)
    If Not oAttrCol Is Nothing Then
        On Error Resume Next
        oAttrCol.Item("IsHollow").Value = False
        oAttrCol.Item("IsSymmetricAboutX").Value = False
        oAttrCol.Item("IsSymmetricAboutY").Value = SymmetricAboutY()
        On Error GoTo ErrorHandler
    End If
    
    Dim bIsSetOk As Boolean
    bIsSetOk = True
    
    ' set the cross section design properties if they are not user defined
    If (m_oDesignMemberHelper.AreSectionPropertiesModifiable(oAttr) = False) Then
        bIsSetOk = m_oDesignMemberHelper.SetXSectionDesignProperties(oSmartOcc, dCentroidX, dCentroidY, _
                                                                     dXp, dYp, dIxx, dIyy, dRo, dRxx, dRxy, dRyy, _
                                                                     dSw, dSxx, dSyy, dZxx, dZyy, dJ, dCw, dH, dUnitWt)
    End If
                                           
    ' if there was an error while setting the design properties go the the error handler
    If bIsSetOk = False Then
        GoTo ErrorHandler
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
        Case "Depth", "Length", "BottomFlangeWidth"
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

Private Function GetSectionPropertiesFromCalculator(ByRef dArea As Double, _
                                                    ByRef dPerimeter As Double, _
                                                    ByRef dCentroidX As Double, _
                                                    ByRef dCentroidY As Double, _
                                                    ByRef dXp As Double, ByRef dYp As Double, _
                                                    ByRef dIxx As Double, ByRef dIyy As Double, _
                                                    ByRef dRo As Double, ByRef dRxx As Double, _
                                                    ByRef dRxy As Double, ByRef dRyy As Double, _
                                                    ByRef dSw As Double, ByRef dSxx As Double, _
                                                    ByRef dSyy As Double, ByRef dZxx As Double, _
                                                    ByRef dZyy As Double, ByRef dJ As Double, _
                                                    ByRef dCw As Double, dH As Double, _
                                                    ByRef dUnitWt As Double, _
                                                    ByVal dDepth As Double, _
                                                    ByVal dBotFlangeWidth As Double, _
                                                    ByVal dWebThickness As Double, _
                                                    ByVal dBotThickness As Double, _
                                                    ByVal dBotOffset As Double) As ErrorSectionStatus
                                               
    
Const METHOD = "GetSectionPropertiesFromCalculator"
On Error GoTo ErrorHandler
    
    Dim ErrorStatus As ErrorSectionStatus
    Dim bIsSetOk As Boolean
      
    m_oCalcXProps.SectionTypeAlias = EnumSectionTypeAlias.Section_LG
    m_oCalcXProps.Width = dBotFlangeWidth
    m_oCalcXProps.Depth = dDepth
    m_oCalcXProps.BottomFlangeWidth = dBotFlangeWidth
    m_oCalcXProps.BottomFlangeThickness = dBotThickness
    m_oCalcXProps.WebThickness = dWebThickness
    m_oCalcXProps.OffsetWebBottom = dBotOffset

    ErrorStatus = m_oCalcXProps.SectionProperties(dArea, dPerimeter, dCentroidX, dCentroidY, _
                                                     dXp, dYp, dIxx, dIyy, dRo, dRxx, dRxy, dRyy, _
                                                     dSw, dSxx, dSyy, dZxx, dZyy, dJ, dCw, dH, dUnitWt)
                                                                         
Exit Function
ErrorHandler: HandleError MODULE, METHOD
ErrorStatus = SectionProperties_UnExpectedError
End Function

Public Sub ISPSDesignedMemberHelper_LoadEmulatedFacePorts(ByVal oDesignedMember As SPSMembers.ISPSDesignedMember)
Const METHOD = "ISPSDesignedMemberHelper_LoadEmulatedFacePorts"
On Error GoTo ErrorHandler

    Dim oAssyMembers     As IJDMemberObjects
    Dim oPlateSystem     As IJPlateSystem
    Dim eSectionAlias As EnumSectionTypeAlias
    
    eSectionAlias = Section_LG
    'get assy members from SO
    Set oAssyMembers = oDesignedMember
    
    'First handle the bot flange
    Set oPlateSystem = oAssyMembers.Item(BULMembers.BotFlange)
    If Not oPlateSystem Is Nothing Then
        m_oDesignMemberHelper.EvaluateBotFlangePlateSystemXIDs oPlateSystem, eSectionAlias, oDesignedMember
    End If
    'handle the web
    Set oPlateSystem = oAssyMembers.Item(BULMembers.Web)
    If Not oPlateSystem Is Nothing Then
        m_oDesignMemberHelper.EvaluateWebPlateSystemXIDs oPlateSystem, StructBUPlateType.Web, eSectionAlias, oDesignedMember
    End If
    
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

Private Sub ISPSDesignedMemberHelper_ResolveAmbiguity(ByVal pDesignedMember As SPSMembers.ISPSDesignedMember)
Const METHOD = "ISPSDesignedMemberHelper_ResolveAmbiguity"
On Error GoTo ErrorHandler
 
    Dim pCurveToExtrude As IJCurve
    Dim pCreatedPlate As IJPlate
    Dim pPointOnSurface As IJDPosition
    Set pPointOnSurface = New DPosition
    Dim strPlate As String
    Dim res As VbTriState
    Dim dMidX As Double
    Dim dMidY As Double
    Dim dMidZ As Double
    Dim sX As Double
    Dim sY As Double
    Dim sZ As Double
    Dim eX As Double
    Dim eY As Double
    Dim eZ As Double
    Dim tX As Double
    Dim tY As Double
    Dim tZ As Double
    Dim t2X As Double
    Dim t2Y As Double
    Dim t2Z As Double
    Dim curveDir As IJDVector
    Dim oSmartOcc As IJSmartOccurrence
    Dim pIJDMemberObjects As IJDMemberObjects
    
    Set oSmartOcc = pDesignedMember
    Set pIJDMemberObjects = oSmartOcc
    Dim i As Long
    'The ambiguity of the BUL differs a bit from the BUT because the BUL can be bounded by
    'the flange or bounded by the web.  The for loop ensures abiguity call for both cases.
    '
    ' Flange bounded by WEB
    ' __
    '|  |
    '|  |
    '|  |
    '|  |
    '|  |____________
    '|  |____________|
    '|__|
    '
    ' Web bounded by Flange
    '   __
    '  |  |
    '  |  |
    '  |  |
    '  |  |
    ' _|__|__________
    '|_______________|
    '
    ' indices are 1 and 3 for the flange and
    ' 4 and 6 for the web
    For i = 1 To 4 Step 3
        'get plate
        Set pCreatedPlate = pIJDMemberObjects.Item(i + 2)
        strPlate = pIJDMemberObjects.MemberDescriptions.Item(i + 2).Name
        ' get the mother curve used to create the plate
        Set pCurveToExtrude = pIJDMemberObjects.Item(i)
        ' get a point on the mother curve
        Dim dCurveExtension As Double
        If i = 1 Then
            m_oDesignMemberHelper.GetBottomWidthExtension oSmartOcc, dCurveExtension
        Else
             m_oDesignMemberHelper.GetWebDepthExtension oSmartOcc, dCurveExtension
        End If
        If dCurveExtension < 0.000001 Then
            pCurveToExtrude.PositionFRatio 0.5, dMidX, dMidY, dMidZ
        Else
            Dim startParam As Double
            Set curveDir = New DVector
            pCurveToExtrude.EndPoints sX, sY, sZ, eX, eY, eZ
            pCurveToExtrude.Parameter sX, sY, sZ, startParam
            
            pCurveToExtrude.Evaluate startParam, dMidX, dMidY, dMidZ, tX, tY, tZ, t2X, t2Y, t2Z
            curveDir.Set tX, tY, tZ
            curveDir.Length = dCurveExtension + (((pCurveToExtrude.Length()) - (2 * dCurveExtension)) / 2)
            curveDir.Get tX, tY, tZ
            dMidX = dMidX + tX
            dMidY = dMidY + tY
            dMidZ = dMidZ + tZ
        End If
        
        'this code was added to basically place the point towards the middle of the
        'plate surface instead of having the point at the edge of the plate surface
        Dim oMP As ISPSMemberPartCommon
        Set oMP = pDesignedMember
        Dim pTF As IJDT4x4
        Dim pKTF As IJDT4x4
        oMP.Rotation.GetTransformAtPosition dMidX, dMidY, dMidZ, pTF, pKTF
        Dim xV As IJDVector
        Set xV = New DVector
        xV.Set pTF.IndexValue(0), pTF.IndexValue(1), pTF.IndexValue(2)
        'xV.Length = (pIJDMemberObjects.Item(2).Length) / 2
        xV.Length = oMP.Axis.Length / 2
          pPointOnSurface.Set dMidX, dMidY, dMidZ
        Set pPointOnSurface = pPointOnSurface.Offset(xV)
        pPointOnSurface.Get dMidX, dMidY, dMidZ
        
        ' call the helper to rersolve any ambiguities
        res = m_oDesignMemberHelper.ResolveAmbiguity(pCreatedPlate, pPointOnSurface)
        If res = vbUseDefault Then
            Err.Raise E_FAIL, MODULE & "." & METHOD, "Failed to resolve ambiguity of " & strPlate
        ElseIf res = vbFalse Then
            Debug.Print strPlate & " is not ambiguous"
        End If
    Next
Exit Sub
ErrorHandler:
Dim errMsg As String
Dim errSrc As String
With Err
    errMsg = .Description
    errSrc = .Source
End With

HandleError MODULE, METHOD
SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
Err.Raise E_FAIL, errSrc, errMsg
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
            Case "IUABuiltUpBotFlange"
                If sAttrName = "BottomFlangeWidth" Then
                    IsDesignParameter = True
                End If
            Case "IUABuiltUpL"
                If sAttrName = "OffsetWeb" Then
                    IsDesignParameter = True
                End If
            
            Case Else
        End Select ' I/F name
    End If

Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Function SymmetricAboutY() As Boolean
    SymmetricAboutY = False
    If (Abs(m_oCalcXProps.BottomFlangeWidth - m_oCalcXProps.WebThickness - 2 * m_oCalcXProps.OffsetWebBottom) < dTol) Then
        SymmetricAboutY = True
    End If

End Function


Private Sub GetComponentMaterials(ByVal pProfileObject As Object, BFlangeMat As String, BFlangeGrade As String, WebMat As String, WebGrade As String)
    
    Dim Thickness As Double
    
    GetComponentMaterial pProfileObject, IID_IUABuiltUpBottomFlange, BFlangeMat, BFlangeGrade, Thickness
    GetComponentMaterial pProfileObject, IID_IUABuiltUpWeb, WebMat, WebGrade, Thickness
    
End Sub

Private Sub GetComponentParams(ByVal pProfileObject As Object, dBottomFlangeWidth As Double, dBottomFlangeThickness As Double, dWebThickness As Double, dWebHeight As Double)
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    Dim dDepth#
   
    Set oSmartOcc = pProfileObject
    
    Set oSmartItem = oSmartOcc.ItemObject
    
    Set oAttr = oSmartItem
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpBottomFlange)
    dBottomFlangeThickness = oAttrCol.Item("BottomFlangeThickness").Value
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpWeb)
    dWebThickness = oAttrCol.Item("WebThickness").Value
    
    Set oAttr = oSmartOcc
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpBottomFlange)
    dBottomFlangeWidth = oAttrCol.Item("BottomFlangeWidth").Value
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    dDepth = oAttrCol.Item("Depth").Value
    dWebHeight = dDepth - dBottomFlangeThickness
    
    

End Sub

Private Sub GetComponentCP5Delta(ByVal pProfileObject As Object, dBFlangeDX As Double, dBFlangeDY As Double, _
dWebDX As Double, dWebDY As Double)
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    Dim dDepth#
    Dim dOffsetWeb As Double
    Dim iCrossSection As ISPSCrossSection
    Dim locX#, locY#
    Dim dBottomFlangeThickness#, dBottomFlangeWidth#, dWebThickness#, dWebHeight#
    
    'computes the offset from the current CP of the member to the CP5 of the component
    
    Set oSmartOcc = pProfileObject
    Set iCrossSection = pProfileObject
    Set oSmartItem = oSmartOcc.ItemObject
    
    Set oAttr = oSmartItem
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpBottomFlange)
    dBottomFlangeThickness = oAttrCol.Item("BottomFlangeThickness").Value
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpWeb)
    dWebThickness = oAttrCol.Item("WebThickness").Value
    
    Set oAttr = oSmartOcc
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpBottomFlange)
    dBottomFlangeWidth = oAttrCol.Item("BottomFlangeWidth").Value
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    dDepth = oAttrCol.Item("Depth").Value
    dWebHeight = dDepth - dBottomFlangeThickness
    
    
   
    'get distances from cp1  position to current cp
    iCrossSection.GetCardinalPointDelta Nothing, 1, 0, locX, locY
    'the origin is at cp1 position of the member
    
    dBFlangeDY = 0.5 * dBottomFlangeThickness - locY
    'the minus sign  below is due to the conversion from 2d coordinates to 3d
    dBFlangeDX = -(0.5 * dBottomFlangeWidth - locX)
    
    dOffsetWeb = GetLOffsetWeb(oSmartOcc)

    dWebDX = -((dOffsetWeb + 0.5 * dWebThickness) - locX)
  
    dWebDY = (0.5 * (dDepth - dBottomFlangeThickness) + dBottomFlangeThickness) - locY


End Sub


Private Sub GetComplexBoundary(ByVal pProfileObject As Object, boundaryOuter As IJComplexString, boundaryInner As IJComplexString)
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    Dim dDepth#
    Dim dOffsetWeb As Double
    Dim iCrossSection As ISPSCrossSection
    Dim dBottomFlangeThickness#, dBottomFlangeWidth#, dWebThickness#, dWebHeight#
    
    
    Set oSmartOcc = pProfileObject
    Set iCrossSection = pProfileObject
    Set oSmartItem = oSmartOcc.ItemObject
    
    Set oAttr = oSmartItem
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpBottomFlange)
    dBottomFlangeThickness = oAttrCol.Item("BottomFlangeThickness").Value
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpWeb)
    dWebThickness = oAttrCol.Item("WebThickness").Value
    
    Set oAttr = oSmartOcc
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpBottomFlange)
    dBottomFlangeWidth = oAttrCol.Item("BottomFlangeWidth").Value
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    dDepth = oAttrCol.Item("Depth").Value
    dWebHeight = dDepth - dBottomFlangeThickness
    
   
    dOffsetWeb = GetLOffsetWeb(oSmartOcc)
  
    'mark external corners starting from bottom left
    Dim oPos(1 To 8) As IJDPosition
    
    Set oPos(1) = New DPosition
    oPos(1).Set 0, 0, 0
    
    Set oPos(2) = New DPosition
    oPos(2).Set 0, -dBottomFlangeWidth, 0
    
    Set oPos(3) = New DPosition
    oPos(3).Set 0, -dBottomFlangeWidth, dBottomFlangeThickness
    
    Set oPos(4) = New DPosition
    oPos(4).Set 0, -(dOffsetWeb + dWebThickness), dBottomFlangeThickness


    Set oPos(5) = New DPosition
    oPos(5).Set 0, -(dOffsetWeb + dWebThickness), dDepth

    Set oPos(6) = New DPosition
    oPos(6).Set 0, -dOffsetWeb, dDepth

    Set oPos(7) = New DPosition
    oPos(7).Set 0, -dOffsetWeb, dBottomFlangeThickness
    
    Set oPos(8) = New DPosition
    oPos(8).Set 0, 0, dBottomFlangeThickness

  
    Set boundaryOuter = CreateComplexStringFromPositions(oPos)
  
    Erase oPos 'free the memory
    Dim locX#, locY#
    Dim oMat As New DT4x4
    Dim oVec As New DVector

      
    'get distances to cp1  from current cp
    iCrossSection.GetCardinalPointDelta Nothing, 0, 1, locX, locY
   

    oVec.Set 0, -locX, locY '-sign due to section to member coord sys difference
  
    oMat.LoadIdentity
    oMat.Translate oVec
    'make the boundaries relative to the cp of the member
    boundaryOuter.Transform oMat
    

End Sub



Private Sub ICustomSectionShapeService_GetCrossSectionData(ByVal pProfileObject As Object, ByVal distFromStart As Double, ByVal eRepresentationType As SP3DStructInterfaces.structShapeRepresentationEnum, ByVal bBreakIntoComponents As Boolean, sectionPrev As SP3DStructInterfaces.IStructSectionShape, sectionNext As SP3DStructInterfaces.IStructSectionShape)


    Dim BFlangeMat As String, BFlangeGrade As String, WebMat As String, WebGrade As String
    Dim iSectionMaterial As IStructSectionMaterial
    Dim pISPSMemberPartCommon As ISPSMemberPartCommon
    Dim dMembLength As Double
    
    'get the materials
    
    GetComponentMaterials pProfileObject, BFlangeMat, BFlangeGrade, WebMat, WebGrade
    'if representation type is size_by_parameter
    If eRepresentationType = shape_by_parametric Then

        Dim oCompoundSection As StructSectionCompoundShape
        Dim oBottomFlange As IStructSectionRectangleSolidShape
        Dim oWeb As IStructSectionRectangleSolidShape
        Dim dBFlangeW#, dBFlangeT#
        Dim dWebT#, dWebH#
        Dim iSectionChild As IStructSectionShapeChild
        Dim matChild As New DT4x4
        Dim xVec As New DVector
        Dim dBFlangeDZ#, dBFlangeDY#, dWebDZ#, dWebDY#
        Dim deltaVec As New DVector
        
        
        'get sizes of component plates
        GetComponentParams pProfileObject, dBFlangeW, dBFlangeT, dWebT, dWebH
        'get cp5 offsets of component plates
        GetComponentCP5Delta pProfileObject, dBFlangeDY, dBFlangeDZ, dWebDY, dWebDZ
        
        xVec.Set 1, 0, 0 'member axis direction
        
        Set oCompoundSection = New StructSectionCompoundShape
        'add bottom flange
        Set oBottomFlange = oCompoundSection.CreateComponentSection(sectionType_Rectangle)
        
        oBottomFlange.SetProperties dBFlangeW, dBFlangeT, 0#
        
        Set iSectionChild = oBottomFlange
        matChild.LoadIdentity

       
        deltaVec.Set 0, dBFlangeDY, dBFlangeDZ
        matChild.Translate deltaVec
        
        'rotate matrix 90 degree to correspond to  flange
        matChild.Rotate 0.5 * PI, xVec
        
        
        iSectionChild.SetPositionInformation 5, False, matChild
        Set iSectionMaterial = oBottomFlange
        iSectionMaterial.SetMaterial BFlangeMat, BFlangeGrade
        
        'add  web
        Set oWeb = oCompoundSection.CreateComponentSection(sectionType_Rectangle)
        
        oWeb.SetProperties dWebH, dWebT, 0#
        
        Set iSectionChild = oWeb
        matChild.LoadIdentity
        
        
        deltaVec.Set 0, dWebDY, dWebDZ
        matChild.Translate deltaVec
        
        
        iSectionChild.SetPositionInformation 5, False, matChild
        Set iSectionMaterial = oWeb
        iSectionMaterial.SetMaterial WebMat, WebGrade
        
        If distFromStart > distTol Then
            Set sectionPrev = oCompoundSection
        End If
        

        Set pISPSMemberPartCommon = pProfileObject
        dMembLength = pISPSMemberPartCommon.Axis.Length
        If (dMembLength - distFromStart) > distTol Then
            Set sectionNext = oCompoundSection
        End If
            
            
        
    Else ' size by graphic

        'return complex string boundary of the section
        Dim oGraphicSection As IStructSectionCurveBoundedShape
        Dim outerCurve As IJComplexString
        
        
        Set oGraphicSection = New StructSectionCurveBoundedShape
        GetComplexBoundary pProfileObject, outerCurve, Nothing

        oGraphicSection.SetBoundary outerCurve, Nothing
        Set iSectionMaterial = oGraphicSection
        iSectionMaterial.SetMaterial BFlangeMat, BFlangeGrade 'need to pick the mat of one of the component to represent the whole section

        If distFromStart > distTol Then
            Set sectionPrev = oGraphicSection
        End If
        

        Set pISPSMemberPartCommon = pProfileObject
        dMembLength = pISPSMemberPartCommon.Axis.Length
        If (dMembLength - distFromStart) > distTol Then
            Set sectionNext = oGraphicSection
        End If

    End If

End Sub

Private Sub ICustomSectionShapeService_GetTransitionLocations(ByVal pProfileObject As Object, distFromStart() As Double)
    'the section size for this type do not change along the length
    'so just add the value of 0, representing the start of the member, to the array.
    ReDim distFromStart(0 To 0)
    
    distFromStart(0) = 0#
End Sub

