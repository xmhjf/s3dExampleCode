VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2008, Intergraph Corporation. All rights reserved.
'
'File
'    BUT.cls
'
'Author
'       Will Rosa 22 Feb 2008
'
'Description
'       Built Up T definition
'
'Notes
'
'History:
'
'   22-Sept-2009 GG TR#167167 - DesignedMember does not set PG of plates to its own PG
'   06-Aug-2010 GG DI-169828  SCHEMA: Interface to indicate member symmetry and open/closed cross sections
'*******************************************************************
'                   <-- Top Flange W ->
'
'                  (Z)
'                   | ----------------    .
'                   ||                |  / \ |
'                   | ----------------    |
'                   |       |  |          H
'                   |       |  |          e
'                   |       |  |          i
'                   |       |  |          g
'                   |       |  |          h
'                   |       |  |          t
'                   |       |  |          |
'                   |       |__|         \ /
' (Y) --------------+-------------------- . (-Y)
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

Private Const MODULE = "BUT"
Private Const strSourceFile = "BUT.def"
Private Const CONST_ItemProgId As String = "SM3DBUT.BUT"

Private m_strErrorDescr As String
Private m_oLocalizer As IJLocalizer
Private m_oDesignMemberHelper As BUHelperUtils
Private m_oCalcXProps As BUCalcSectionProperties

Private Enum BUTMembers
  TopFlangeMotherCurve = 1
  TopFlange = 3
  WebMotherCurve = 4
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
     pAD.UserTypeClsid = "{F7CB178D-9A87-4039-AC44-C2112028EC91}"   'CUBuiltUpT (Generated here)
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
     pAPDs.AddProperty "IUABuiltUpTopFlange", 7, IID_IUABuiltUpTopFlange, , imsCOOKIE_ID_USS_LIB
     pAPDs.AddProperty "IUABuiltUpCompute", 8, IID_IUABuiltUpCompute, , imsCOOKIE_ID_USS_LIB
     
     Set pAPDs = Nothing
               
     ' Define the members
     Dim pMemberDescriptions As IJDMemberDescriptions
     Dim pMemberDescription As IJDMemberDescription
     Dim pPropertyDescriptions As IJDPropertyDescriptions
     Set pMemberDescriptions = pDefinition
     
     ' Remove all the previous member descriptions
     pMemberDescriptions.RemoveAll
     
'''''
'' Top Flange Members 1,2,3
'''''

     Set pMemberDescription = pMemberDescriptions.AddMember("TopFlangeCurveToExtrude", TopFlangeMotherCurve, "CMConstructTopFlangeCurveToExtrude", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "TopFlangeCurveToExtrudeProperties", 1, IJCurve, "CMComputeTopFlangeCurveToExtrude", imsCOOKIE_ID_USS_LIB

     Set pMemberDescription = pMemberDescriptions.AddMember("TopFlangeExtrusionPath", 2, "CMConstructTopFlangeExtrusionPath", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "TopFlangeExtrusionPathProperties", 1, IJCurve, "CMComputeTopFlangeExtrusionPath", imsCOOKIE_ID_USS_LIB
                   
     ' Top Flange
     Set pMemberDescription = pMemberDescriptions.AddMember("TopFlange", TopFlange, "CMConstructTopFlangePlateSystem", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTopFlangePlateSystem"
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateTopFlangePlateSystem"
     
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "TopFlangePlateSystemProperties", TopFlangeMotherCurve, IJCurve, "CMComputeTopFlangePlateSystem", imsCOOKIE_ID_USS_LIB

'''''
'' Web Members 4,5,6
'''''

     Set pMemberDescription = pMemberDescriptions.AddMember("WebCurveToExtrude", WebMotherCurve, "CMConstructWebCurveToExtrude", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "TopWebToExtrudeProperties", 1, IJCurve, "CMComputeWebCurveToExtrude", imsCOOKIE_ID_USS_LIB

     Set pMemberDescription = pMemberDescriptions.AddMember("WebExtrusionPath", 5, "CMConstructWebExtrusionPath", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "WebExtrusionPathProperties", 1, IJCurve, "CMComputeWebExtrusionPath", imsCOOKIE_ID_USS_LIB

     ' Web
     Set pMemberDescription = pMemberDescriptions.AddMember("Web", Web, "CMConstructWebPlateSystem", imsCOOKIE_ID_USS_LIB)
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
'                       Top Flange
'-------------------------------------------------------
Public Sub CMConstructTopFlangeCurveToExtrude(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructTopFlangeCurveToExtrude"
On Error GoTo ErrorHandler
   
    'Create a Line/Curve to use
    Dim pCurveToExtrude As IJCurve
    
    m_oDesignMemberHelper.CreateCurveBy2Points pResourceManager, 0, 0, 1, 0, -1, 1, pCurveToExtrude
    CopyPermissionGroup pCurveToExtrude, pMemberDescription.CAO
    
    Set pObj = pCurveToExtrude
    
    Dim oControlFlags As IJControlFlags
    Set oControlFlags = pObj
    oControlFlags.ControlFlags(&H4) = &H4

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMComputeTopFlangeCurveToExtrude(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeTopFlangeCurveToExtrude"
On Error GoTo ErrorHandler

    Dim dDepth As Double
    Dim dTopWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
    
    GetTExtrusionParameters oSmartOcc, dDepth, dTopWidth, dLength

    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDepth, dTopWidth, oMatrix
    
    ' Modify the input object/line
    Dim pIJLine As IJLine
    Set pIJLine = pObject
    
    Dim dTopExt As Double
    m_oDesignMemberHelper.GetTopWidthExtension oSmartOcc, dTopExt
    
    pIJLine.DefineBy2Points 0, 0 + dTopExt, dDepth, _
                            0, -dTopWidth - dTopExt, dDepth

    pIJLine.Transform oMatrix

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
Err.Raise E_FAIL
End Sub

Public Sub CMConstructTopFlangeExtrusionPath(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructTopFlangeExtrusionPath"
On Error GoTo ErrorHandler
     
    'Create a Line/Curve to use
    Dim pCurve As IJCurve
    
    m_oDesignMemberHelper.CreateCurveBy2Points pResourceManager, 0, 0, 1, 1, 0, 1, pCurve
    CopyPermissionGroup pCurve, pMemberDescription.CAO
    
    Set pObj = pCurve
    
    Dim oControlFlags As IJControlFlags
    Set oControlFlags = pObj
    oControlFlags.ControlFlags(&H4) = &H4
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMComputeTopFlangeExtrusionPath(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeTopFlangeExtrusionPath"
On Error GoTo ErrorHandler

    Dim dDepth As Double
    Dim dTopWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
        
    GetTExtrusionParameters oSmartOcc, dDepth, dTopWidth, dLength

    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDepth, dTopWidth, oMatrix

    Dim pIJLine As IJLine
    Set pIJLine = pObject
        
    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt
    
    pIJLine.DefineBy2Points 0 - dLengthExt, 0, dDepth, _
                            dLength + dLengthExt, 0, dDepth
    
    pIJLine.Transform oMatrix

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
Err.Raise E_FAIL
End Sub

Public Sub CMConstructTopFlangePlateSystem(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructTopFlangePlateSystem"
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
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTopFlange)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("TopFlangeThickness").Value
        strMaterial = oAttrCol.Item("TopFlangeMaterial").Value
        strGrade = oAttrCol.Item("TopFlangeGrade").Value
    Else
        GoTo ErrorHandler
    End If
    
    ' Create the curves to extrude
    Dim pCurveToExtrude As IJCurve
    Dim pExtrusionPath As IJCurve
  
    ' Get the outputs
    Dim pIJDMemberObject As IJDMemberObjects
    Set pIJDMemberObject = oSmartOcc
    
    Set pCurveToExtrude = pIJDMemberObject.ItemByDispid(1)
    Set pExtrusionPath = pIJDMemberObject.ItemByDispid(2)
    
    Set pObj = m_oDesignMemberHelper.CreateABuitUpPlate(pResourceManager, pCurveToExtrude, _
                                                       pExtrusionPath, dThickness, _
                                                       strMaterial, strGrade, oSmartOcc, FlangePlate, _
                                                       , , , InDir, MoldedFormPlateNameCat.NameCatTFlange)
   
     
    m_oDesignMemberHelper.SetPlateBoundaries oSmartOcc, pObj
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructTopFlangePlateSystem(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructTopFlangePlateSystem"
On Error GoTo ErrorHandler
 
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub
Public Sub CMComputeTopFlangePlateSystem(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeTopFlangePlateSystem"
On Error GoTo ErrorHandler
    ' get the plate
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    Dim pIJDMemberObject As IJDMemberObjects
    Set pIJDMemberObject = oSmartOcc
    Dim oPlate As IJPlate
    Set oPlate = pIJDMemberObject.ItemByDispid(TopFlange)
    
    ' get the parameters
    Dim oSmartItem As IJSmartItem
    Set oSmartItem = oSmartOcc.ItemObject
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
   
    Dim dThickness As Double
    Dim strMaterial As String
    Dim strGrade As String
    Set oAttr = oSmartItem
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTopFlange)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("TopFlangeThickness").Value
        strMaterial = oAttrCol.Item("TopFlangeMaterial").Value
        strGrade = oAttrCol.Item("TopFlangeGrade").Value
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
Public Sub CMMigrateTopFlangePlateSystem(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateTopFlangePlateSystem"
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
    Dim dTopWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If

    GetTExtrusionParameters oSmartOcc, dDepth, dTopWidth, dLength

    Dim pIJLine As IJLine
    Set pIJLine = pObject
  
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDepth, dTopWidth, oMatrix
    
    Dim dWebExtension As Double
    m_oDesignMemberHelper.GetWebDepthExtension oSmartOcc, dWebExtension
    Dim trimExt As Double
    trimExt = dWebExtension
    If trimExt < 0.000001 Then
        trimExt = m_oDesignMemberHelper.BoundingExt
    End If
    
    pIJLine.DefineBy2Points 0, -dTopWidth / 2, 0, _
                            0, -dTopWidth / 2, dDepth + dWebExtension + trimExt

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
    Dim dTopWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If

    GetTExtrusionParameters oSmartOcc, dDepth, dTopWidth, dLength
    
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDepth, dTopWidth, oMatrix

    Dim pIJLine As IJLine
    Set pIJLine = pObject
    
    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt
    
    pIJLine.DefineBy2Points 0 - dLengthExt, -dTopWidth / 2, dDepth, _
                            dLength + dLengthExt, -dTopWidth / 2, dDepth
  
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
    Dim pExtrusionPath As IJCurve

    Set pCurveToExtrude = pIJDMemberObject.ItemByDispid(4)
    Set pExtrusionPath = pIJDMemberObject.ItemByDispid(5)
       
    Set pObj = m_oDesignMemberHelper.CreateABuitUpPlateWithProfileOrientation(pResourceManager, pCurveToExtrude, _
                                                        pExtrusionPath, dThickness, strMaterial, strGrade, oSmartOcc, WebPlate, _
                                                        RightDir, , , , Centered, MoldedFormPlateNameCat.NameCatWeb)
                                            
    ' Add the boundaries
    Dim oTopFlange As Object
    
    Set oTopFlange = pIJDMemberObject.ItemByDispid(3)
    
    m_oDesignMemberHelper.SetPlateBoundaries oSmartOcc, pObj, oTopFlange, Nothing, Nothing, Nothing
                                            
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
    Dim dWidth As Double
    Dim dDepth As Double
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pIDesignedMember
    
    GetTExtrusionParameters oSmartOcc, dDepth, dWidth, dLength
    
    pdDepth = dDepth
    pdWidth = dWidth
   
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
    
    If Not bIsModifiable Then
        Set pAttrColl = CollAllDisplayedValues
        For i = 1 To pAttrColl.Count
            Set pAttrDescr = pAttrColl.Item(i)
            If pAttrDescr.InterfaceName = IID_IUABuiltUpTopFlange Then
                If pAttrDescr.AttrName = "TopFlangeWidth" Then
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

Private Sub GetTExtrusionParameters(ByVal oSmartOcc As IJSmartOccurrence, _
                                           ByRef dDepth As Double, _
                                           ByRef dTopWidth As Double, _
                                           ByRef dLength As Double)
Const METHOD = "GetTExtrusionParameters"
On Error GoTo ErrorHandler
    
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes

    ' Get Parameters of BU
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTopFlange)
    If Not oAttrCol Is Nothing Then
        dTopWidth = oAttrCol.Item("TopFlangeWidth").Value
    Else
        GoTo ErrorHandler
    End If
    
    Set oAttrCol = Nothing
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    If Not oAttrCol Is Nothing Then
        dDepth = oAttrCol.Item("Depth").Value
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

Public Function AreOccurrencePropertiesValid(ByVal oSmartOcc As IJSmartOccurrence) As Boolean
Const METHOD = "AreOccurrencePropertiesValid"
On Error GoTo ErrorHandler
    
    AreOccurrencePropertiesValid = False
    
    Dim dDepth As Double
    Dim dTopWidth As Double
    Dim dLength As Double
    Dim sError As String
        
    GetTExtrusionParameters oSmartOcc, dDepth, dTopWidth, dLength
    
    sError = BuiltUpDefValidate("Depth", dDepth, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("TopFlangeWidth", dTopWidth, 0#)
    
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
    
    Dim dTopExt As Double
    m_oDesignMemberHelper.GetTopWidthExtension oSmartOcc, dTopExt
    
    sError = BuiltUpDefValidate("TopFlangeWidthExt", dTopExt, 0#)
    
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
         
    Dim dTopWidth As Double
    Dim dTopThickness As Double
    Dim dDepth As Double
    Dim dWebThickness As Double
    
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
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTopFlange)
    If Not oAttrCol Is Nothing Then
        dTopWidth = oAttrCol.Item("TopFlangeWidth").Value
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
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTopFlange)
    If Not oAttrCol Is Nothing Then
        dTopThickness = oAttrCol.Item("TopFlangeThickness").Value
    Else
        GoTo ErrorHandler
    End If
            
    Dim ErrorStatus As ErrorSectionStatus
    ErrorStatus = GetSectionPropertiesFromCalculator(dArea, dPerimeter, dCentroidX, dCentroidY, _
                                                     dXp, dYp, dIxx, dIyy, dRo, dRxx, dRxy, dRyy, _
                                                     dSw, dSxx, dSyy, dZxx, dZyy, dJ, dCw, dH, dUnitWt, _
                                                     dDepth, dTopWidth, dWebThickness, dTopThickness)
                                                     
    If ErrorStatus <> SectionProperties_OK Then
        GoTo ErrorHandler
    End If
                        
    Set oAttr = Nothing
    Set oAttrCol = Nothing
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    If Not oAttrCol Is Nothing Then
        oAttrCol.Item("Depth").Value = dDepth
        oAttrCol.Item("Width").Value = dTopWidth
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
        oAttrCol.Item("IsSymmetricAboutY").Value = True
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
        Case "Depth", "Length", "TopFlangeWidth"
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
                                                    ByVal dTopFlangeWidth As Double, _
                                                    ByVal dWebThickness As Double, _
                                                    ByVal dTopThickness As Double) As ErrorSectionStatus
                                               
    
Const METHOD = "GetSectionPropertiesFromCalculator"
On Error GoTo ErrorHandler
    
    Dim ErrorStatus As ErrorSectionStatus
    Dim bIsSetOk As Boolean
      
    m_oCalcXProps.SectionTypeAlias = EnumSectionTypeAlias.Section_TG
    m_oCalcXProps.Width = dTopFlangeWidth
    m_oCalcXProps.Depth = dDepth
    m_oCalcXProps.BottomFlangeWidth = dTopFlangeWidth
    m_oCalcXProps.WebThickness = dWebThickness
    m_oCalcXProps.TopFlangeWidth = dTopFlangeWidth
    m_oCalcXProps.TopFlangeThickness = dTopThickness

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
    
    eSectionAlias = Section_TG
    'get assy members from SO
    Set oAssyMembers = oDesignedMember
    
    'First handle the top flange
    Set oPlateSystem = oAssyMembers.Item(BUTMembers.TopFlange)
    If Not oPlateSystem Is Nothing Then
        m_oDesignMemberHelper.EvaluateTopFlangePlateSystemXIDs oPlateSystem, eSectionAlias, oDesignedMember
    End If
    'handle the web
    Set oPlateSystem = oAssyMembers.Item(BUTMembers.Web)
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
    Dim pPointOnWeb As IJDPosition
    Set pPointOnWeb = New DPosition
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
    
    'get the web
    Set pCreatedPlate = pIJDMemberObjects.Item(Web)
    strPlate = pIJDMemberObjects.MemberDescriptions.Item(Web).Name
    ' get the mother curve used to create the top flange
    Set pCurveToExtrude = pIJDMemberObjects.Item(WebMotherCurve)
    ' get a point on the mother curve
    Dim dWebExtension As Double
    m_oDesignMemberHelper.GetWebDepthExtension oSmartOcc, dWebExtension
    If dWebExtension < 0.000001 Then
        pCurveToExtrude.PositionFRatio 0.5, dMidX, dMidY, dMidZ
    Else
        Dim startParam As Double
        Set curveDir = New DVector
        pCurveToExtrude.EndPoints sX, sY, sZ, eX, eY, eZ
        pCurveToExtrude.Parameter sX, sY, sZ, startParam
        
        pCurveToExtrude.Evaluate startParam, dMidX, dMidY, dMidZ, tX, tY, tZ, t2X, t2Y, t2Z
        curveDir.Set tX, tY, tZ
        curveDir.Length = dWebExtension + (((pCurveToExtrude.Length()) - (2 * dWebExtension)) / 2)
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
    xV.Length = oMP.Axis.Length / 2
    'xV.Length = (pIJDMemberObjects.Item(2).Length) / 2
    pPointOnWeb.Set dMidX, dMidY, dMidZ
    Set pPointOnWeb = pPointOnWeb.Offset(xV)
    pPointOnWeb.Get dMidX, dMidY, dMidZ
    
    ' call the helper to rersolve any ambiguities
    res = m_oDesignMemberHelper.ResolveAmbiguity(pCreatedPlate, pPointOnWeb)
    If res = vbUseDefault Then
        Err.Raise E_FAIL, MODULE & "." & METHOD, "Failed to resolve ambiguity of " & strPlate
    ElseIf res = vbFalse Then
        Debug.Print strPlate & " is not ambiguous"
    End If

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
            Case "IUABuiltUpTopFlange"
                If sAttrName = "TopFlangeWidth" Then
                    IsDesignParameter = True
                End If
            
            Case Else
        End Select ' I/F name
    End If

Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function



Private Sub GetComponentParams(ByVal pProfileObject As Object, dTopFlangeWidth As Double, dTopFlangeThickness As Double, dWebThickness As Double, dWebHeight As Double)
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    Dim dDepth#
    
    Set oSmartOcc = pProfileObject
    
    Set oSmartItem = oSmartOcc.ItemObject
    
    Set oAttr = oSmartItem
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTopFlange)
    dTopFlangeThickness = oAttrCol.Item("TopFlangeThickness").Value
    
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpWeb)
    dWebThickness = oAttrCol.Item("WebThickness").Value
    
    
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTopFlange)
    dTopFlangeWidth = oAttrCol.Item("TopFlangeWidth").Value
    
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    dDepth = oAttrCol.Item("Depth").Value
    dWebHeight = dDepth - dTopFlangeThickness
    

End Sub





Private Sub GetComponentMaterials(ByVal pProfileObject As Object, TFlangeMat As String, TFlangeGrade As String, WebMat As String, WebGrade As String)
    
    
    Dim Thickness As Double
        
    GetComponentMaterial pProfileObject, IID_IUABuiltUpTopFlange, TFlangeMat, TFlangeGrade, Thickness
    GetComponentMaterial pProfileObject, IID_IUABuiltUpWeb, WebMat, WebGrade, Thickness
End Sub
    


Private Sub ICustomSectionShapeService_GetCrossSectionData(ByVal pProfileObject As Object, ByVal distFromStart As Double, ByVal eRepresentationType As SP3DStructInterfaces.structShapeRepresentationEnum, ByVal bBreakIntoComponents As Boolean, sectionPrev As SP3DStructInterfaces.IStructSectionShape, sectionNext As SP3DStructInterfaces.IStructSectionShape)
    Dim TFlangeMat As String, TFlangeGrade As String, BFlangeMat As String, BFlangeGrade As String, WebMat As String, WebGrade As String
    Dim iSectionMaterial As IStructSectionMaterial
    Dim oTsection As StructSectionTShape
    Dim pISPSMemberPartCommon As ISPSMemberPartCommon
    Dim dMembLength As Double

    Dim dTFlangeW#, dTFlangeT#
    Dim dWebT#, dWebH#
    
    'get sizes of component plates
    GetComponentParams pProfileObject, dTFlangeW, dTFlangeT, dWebT, dWebH
    
    
    GetComponentMaterials pProfileObject, TFlangeMat, TFlangeGrade, WebMat, WebGrade
  


    Set oTsection = New StructSectionTShape
    oTsection.SetProperties dWebH + dTFlangeT, dTFlangeW, dWebT, dTFlangeT, 0, 0, 0, 0
    
    Set iSectionMaterial = oTsection
    iSectionMaterial.SetMaterial TFlangeMat, TFlangeGrade 'need to pick the mat of one of the component to represent the whole section
    
    If distFromStart > distTol Then
        Set sectionPrev = oTsection
    End If
    

    Set pISPSMemberPartCommon = pProfileObject
    dMembLength = pISPSMemberPartCommon.Axis.Length
    If (dMembLength - distFromStart) > distTol Then
        Set sectionNext = oTsection
    End If
    


End Sub

Private Sub ICustomSectionShapeService_GetTransitionLocations(ByVal pProfileObject As Object, distFromStart() As Double)
    'the section size for this type do not change along the length
    'so just add the value of 0, representing the start of the member, to the array.
    ReDim distFromStart(0 To 0)
    
    distFromStart(0) = 0#
End Sub

