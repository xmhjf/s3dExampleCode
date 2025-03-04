VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BUITaperWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2008, Intergraph Corporation. All rights reserved.
'
'File
'    BUITaperWeb.cls
'
'Author
'       6 March 2008 Will Rosa
'
'Description
'       Built Up Tapered I Web
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
'                   ||                |  / \
'                   | ----------------    |
'                   |       |  |          D
'                   |       |  |          e
'                   |       |  |          p
'                   |       |  |          t
'                   |       |  |          h
'                   |   -------------     |
'                   |  |             |    |
'                   |   -------------    \ /
' (Y) --------------+-------------------- . (-Y)
'                 / |
'               /   |  <-Btm Flge W->
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
'
'*******************************************************************
'
'  View rotated counter clock wise about the z axis from above view
'  This view shows the side view of the Tapered Web
'
'                  (Z)D
'                   | e -_
'                   | p | -_
'                   | t |   -_
'                   | h |     -_
'                   |   |       -_  <--Top Flange Length
'                   | S |         -_
'                   | t |           -_
'                   | a |             -_
'                   | r | Tapered Web   -_
'                   | t |                 | <-- Depth End
'                   | ------------------------
' (-X) -------------+-------------------------- (X)
'                 / |
'               /   | <----Btm Flge Length---->
'             /     |
'           /       |
'         /         |
'       /           |
'     (-y)          |
'                   |
'                  (-Z)
  
Option Explicit

Private Const MODULE = "BUITaperWeb"
Private Const strSourceFile = "BUITaperWeb.def"
Private Const CONST_ItemProgId As String = "SM3DBUITaperWeb.BUITaperWeb"

Private m_strErrorDescr As String
Private m_oLocalizer As IJLocalizer
Private m_oDesignMemberHelper As BUHelperUtils
Private m_oCalcXProps As BUCalcSectionProperties

Private Enum BUIMembers
  TopFlange = 3
  BotFlange = 6
  Web = 9
End Enum

Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements ISPSDesignedMemberHelper
Implements ICustomSectionShapeService
Private Enum Tapered
    Top
    Center
    Bottom
End Enum



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
     pAD.UserTypeClsid = "{BDE5E9A1-608F-438C-90FB-2322945B8420}"   'BuiltUpTaperedIWDef (Generated here)
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
     pAPDs.AddProperty "IUABuiltUpBottomFlange", 8, IID_IUABuiltUpBottomFlange, , imsCOOKIE_ID_USS_LIB
     pAPDs.AddProperty "IUABuiltUpITaperWeb", 9, IID_IUABuiltUpITaperWeb, , imsCOOKIE_ID_USS_LIB
     pAPDs.AddProperty "IUABuiltUpCompute", 10, IID_IUABuiltUpCompute, , imsCOOKIE_ID_USS_LIB

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
     Set pMemberDescription = pMemberDescriptions.AddMember("TopFlangeCurveToExtrude", 1, "CMConstructTopFlangeCurveToExtrude", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "TopFlangeCurveToExtrudeProperties", 1, IJCurve, "CMComputeTopFlangeCurveToExtrude", imsCOOKIE_ID_USS_LIB

     Set pMemberDescription = pMemberDescriptions.AddMember("TopFlangeExtrusionPath", 2, "CMConstructTopFlangeExtrusionPath", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "TopFlangeExtrusionPathProperties", 1, IJCurve, "CMComputeTopFlangeExtrusionPath", imsCOOKIE_ID_USS_LIB
                   
     ' Top Flange
     Set pMemberDescription = pMemberDescriptions.AddMember("TopFlange", 3, "CMConstructTopFlangePlateSystem", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTopFlangePlateSystem"
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateTopFlangePlateSystem"
          
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "TopFlangePlateSystemProperties", 1, IJCurve, "CMComputeTopFlangePlateSystem", imsCOOKIE_ID_USS_LIB
     
'''''
'' Bottom Flange Members 4,5,6
'''''
     Set pMemberDescription = pMemberDescriptions.AddMember("BottomFlangeCurveToExtrude", 4, "CMConstructBotFlangeCurveToExtrude", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "BottomFlangeCurveToExtrudeProperties", 1, IJCurve, "CMComputeBotFlangeCurveToExtrude", imsCOOKIE_ID_USS_LIB

     Set pMemberDescription = pMemberDescriptions.AddMember("BottomFlangeExtrusionPath", 5, "CMConstructBotFlangeExtrusionPath", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "BottomFlangeExtrusionPathProperties", 1, IJCurve, "CMComputeBotFlangeExtrusionPath", imsCOOKIE_ID_USS_LIB

     ' Bottom flange
     Set pMemberDescription = pMemberDescriptions.AddMember("BottomFlange", 6, "CMConstructBottomFlangePlateSystem", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructBottomFlangePlateSystem"
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateBottomFlangePlateSystem"
          
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "BottomFlangeProperties", 1, IJCurve, "CMComputeBottomFlangePlateSystem", imsCOOKIE_ID_USS_LIB

'''''
'' Web Members 7,8,9
'''''
     Set pMemberDescription = pMemberDescriptions.AddMember("WebCurveToExtrude", 7, "CMConstructWebCurveToExtrude", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "TopWebToExtrudeProperties", 1, IJCurve, "CMComputeWebCurveToExtrude", imsCOOKIE_ID_USS_LIB

     Set pMemberDescription = pMemberDescriptions.AddMember("WebExtrusionCurve", 8, "CMConstructWebExtrusionPath", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "WebExtrusionPathProperties", 1, IJCurve, "CMComputeWebExtrusionPath", imsCOOKIE_ID_USS_LIB

     ' Web
     Set pMemberDescription = pMemberDescriptions.AddMember("Web", 9, "CMConstructWebPlateSystem", imsCOOKIE_ID_USS_LIB)
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

    ' Create a Line/Curve to use
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

    Dim dDepthStart As Double
    Dim dDepthEnd As Double
    Dim dBottomWidth As Double
    Dim dTopWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
    
    GetTaperedIWExtrusionParameters oSmartOcc, dDepthStart, dDepthEnd, dBottomWidth, dTopWidth, dLength
    
    Dim dMaxWidth As Double
    dMaxWidth = GetMaxWidth(dBottomWidth, dTopWidth)

    Dim dMaxDepth As Double
    dMaxDepth = GetMaxDepth(dDepthStart, dDepthEnd)
    
    Dim taperLocation As Tapered
    taperLocation = GetTaperLocation(oSmartOcc)

    m_oDesignMemberHelper.CreateTransform oSmartOcc, dMaxDepth, dMaxWidth, oMatrix
        
    Dim dTopExt As Double
    m_oDesignMemberHelper.GetTopWidthExtension oSmartOcc, dTopExt
        
    ' Modify the input object/line
    Dim pIJLine As IJLine
    Set pIJLine = pObject
    
    Dim z1 As Double
    Dim z2 As Double
    
    z1 = (dMaxDepth / 2) + (dDepthStart / 2)
    z2 = (dMaxDepth / 2) + (dDepthEnd / 2)
    
    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt

    Dim dDepthDelta As Double
    Dim dYOffset As Double
    
    dDepthDelta = Abs(dDepthStart - dDepthEnd)
     
    If taperLocation = Center Then
        dYOffset = (dDepthDelta / 2 * dLengthExt) / dLength
    Else
        dYOffset = (dDepthDelta * dLengthExt) / dLength
    End If
    
    If dDepthStart < dMaxDepth Then
        dYOffset = -dYOffset
    End If

    If taperLocation = Bottom Then
        pIJLine.DefineBy2Points 0 - dLengthExt, -(0.5 * dTopWidth) - (0.5 * dMaxWidth) - dTopExt, dMaxDepth, _
                                0 - dLengthExt, -(dMaxWidth - dTopWidth) * 0.5 + dTopExt, dMaxDepth
    ElseIf taperLocation = Top Then
        pIJLine.DefineBy2Points 0 - dLengthExt, -(0.5 * dTopWidth) - (0.5 * dMaxWidth) - dTopExt, dDepthStart + dYOffset, _
                                0 - dLengthExt, -(dMaxWidth - dTopWidth) * 0.5 + dTopExt, dDepthStart + dYOffset
    Else ' center
        pIJLine.DefineBy2Points 0 - dLengthExt, -(0.5 * dTopWidth) - (0.5 * dMaxWidth) - dTopExt, z1 + dYOffset, _
                                0 - dLengthExt, -(dMaxWidth - dTopWidth) * 0.5 + dTopExt, z1 + dYOffset
    End If
                            
    pIJLine.Transform oMatrix
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
Err.Raise E_FAIL
End Sub

Public Sub CMConstructTopFlangeExtrusionPath(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructTopFlangeExtrusionPath"
On Error GoTo ErrorHandler
     
    ' Create a Line/Curve to use
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

    Dim dDepthStart As Double
    Dim dDepthEnd As Double
    Dim dBottomWidth As Double
    Dim dTopWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
        
    GetTaperedIWExtrusionParameters oSmartOcc, dDepthStart, dDepthEnd, dBottomWidth, dTopWidth, dLength

    Dim dMaxWidth As Double
    dMaxWidth = GetMaxWidth(dBottomWidth, dTopWidth)
    
    Dim dMaxDepth As Double
    dMaxDepth = GetMaxDepth(dDepthStart, dDepthEnd)
    
    Dim taperLocation As Tapered
    taperLocation = GetTaperLocation(oSmartOcc)

    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt
        
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dMaxDepth, dMaxWidth, oMatrix
    
    Dim pIJLine As IJLine
    Set pIJLine = pObject
    
    Dim z1 As Double
    Dim z2 As Double
    
    z1 = (dMaxDepth / 2) + (dDepthStart / 2)
    z2 = (dMaxDepth / 2) + (dDepthEnd / 2)
    
    Dim dDepthDelta As Double
    Dim dYOffset As Double
    
    dDepthDelta = Abs(dDepthStart - dDepthEnd)
    
    If taperLocation = Center Then
        dYOffset = (dDepthDelta / 2 * dLengthExt) / dLength
    Else
        dYOffset = (dDepthDelta * dLengthExt) / dLength
    End If
    
    If dDepthStart < dMaxDepth Then
        dYOffset = -dYOffset
    End If
                               
    If (taperLocation = Top) Then
        pIJLine.DefineBy2Points 0 - dLengthExt, -(0.5 * dTopWidth) - (0.5 * dMaxWidth), dDepthStart + dYOffset, _
                                dLength + dLengthExt, -(0.5 * dTopWidth) - (0.5 * dMaxWidth), dDepthEnd - dYOffset
    ElseIf (taperLocation = Bottom) Then
        pIJLine.DefineBy2Points 0 - dLengthExt, -(0.5 * dTopWidth) - (0.5 * dMaxWidth), dMaxDepth, _
                                dLength + dLengthExt, -(0.5 * dTopWidth) - (0.5 * dMaxWidth), dMaxDepth
    Else 'center
        pIJLine.DefineBy2Points 0 - dLengthExt, -(0.5 * dTopWidth) - (0.5 * dMaxWidth), z1 + dYOffset, _
                                dLength + dLengthExt, -(0.5 * dTopWidth) - (0.5 * dMaxWidth), z2 - dYOffset
    End If
                            
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
    Dim pExtrusionCurve As IJCurve
  
     ' Get the outputs
    Dim pIJDMemberObject As IJDMemberObjects
    Set pIJDMemberObject = oSmartOcc
    
    Set pCurveToExtrude = pIJDMemberObject.ItemByDispid(1)
    Set pExtrusionCurve = pIJDMemberObject.ItemByDispid(2)
    
    Set pObj = m_oDesignMemberHelper.CreateABuitUpPlate(pResourceManager, pCurveToExtrude, _
                                                        pExtrusionCurve, dThickness, _
                                                        strMaterial, strGrade, oSmartOcc, FlangePlate, _
                                                        , , , InDir, MoldedFormPlateNameCat.NameCatTFlange)
     
    m_oDesignMemberHelper.SetPlateBoundaries oSmartOcc, pObj
   
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructTopFlangePlateSystem(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructTopFlangePlateSystem "
On Error GoTo ErrorHandler
 
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub
Public Sub CMComputeTopFlangePlateSystem(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeTopFlangePlateSystem "
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
Const METHOD = "CMMigrateTopFlangePlateSystem "
On Error GoTo ErrorHandler
 
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
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

    Dim dDepthStart As Double
    Dim dDepthEnd As Double
    Dim dBottomWidth As Double
    Dim dTopWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
        
    GetTaperedIWExtrusionParameters oSmartOcc, dDepthStart, dDepthEnd, dBottomWidth, dTopWidth, dLength

    Dim dMaxWidth As Double
    dMaxWidth = GetMaxWidth(dBottomWidth, dTopWidth)
    
    Dim dMaxDepth As Double
    dMaxDepth = GetMaxDepth(dDepthStart, dDepthEnd)
    
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dMaxDepth, dMaxWidth, oMatrix
    
    Dim dBtmExt As Double
    m_oDesignMemberHelper.GetBottomWidthExtension oSmartOcc, dBtmExt
    
    Dim pIJLine As IJLine
    Set pIJLine = pObject
    
    Dim z1 As Double
    Dim z2 As Double
        
    z1 = (dMaxDepth / 2) - (dDepthStart / 2)
    z2 = (dMaxDepth / 2) - (dDepthEnd / 2)
        
    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt

    Dim dDepthDelta As Double
    Dim dYOffset As Double
    dDepthDelta = Abs(dDepthStart - dDepthEnd)
    
    Dim taperLocation As Tapered
    taperLocation = GetTaperLocation(oSmartOcc)
    
    If taperLocation = Center Then
        dYOffset = (dDepthDelta / 2 * dLengthExt) / dLength
    Else
        dYOffset = (dDepthDelta * dLengthExt) / dLength
    End If
    
    If dDepthStart < dMaxDepth Then
        dYOffset = -dYOffset
    End If
        
    If (taperLocation = Bottom) Then
        pIJLine.DefineBy2Points 0 - dLengthExt, -(dBottomWidth * 0.5) - (dMaxWidth * 0.5) - dBtmExt, dMaxDepth - dDepthStart - dYOffset, _
                                0 - dLengthExt, -(dMaxWidth - dBottomWidth) * 0.5 + dBtmExt, dMaxDepth - dDepthStart - dYOffset
    ElseIf (taperLocation = Top) Then
        pIJLine.DefineBy2Points 0 - dLengthExt, -(dBottomWidth * 0.5) - (dMaxWidth * 0.5) - dBtmExt, 0, _
                                0 - dLengthExt, -(dMaxWidth - dBottomWidth) * 0.5 + dBtmExt, 0
    Else ' centered
        pIJLine.DefineBy2Points 0 - dLengthExt, -(dBottomWidth * 0.5) - (dMaxWidth * 0.5) - dBtmExt, z1 - dYOffset, _
                                0 - dLengthExt, -(dMaxWidth - dBottomWidth) * 0.5 + dBtmExt, z1 - dYOffset
    End If
    
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

    Dim dDepthStart As Double
    Dim dDepthEnd As Double
    Dim dBottomWidth As Double
    Dim dTopWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
        
    GetTaperedIWExtrusionParameters oSmartOcc, dDepthStart, dDepthEnd, dBottomWidth, dTopWidth, dLength

    Dim dMaxWidth As Double
    dMaxWidth = GetMaxWidth(dBottomWidth, dTopWidth)
    
    Dim dMaxDepth As Double
    dMaxDepth = GetMaxDepth(dDepthStart, dDepthEnd)
    
    Dim taperLocation As Tapered
    taperLocation = GetTaperLocation(oSmartOcc)
    
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dMaxDepth, dMaxWidth, oMatrix
    
    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt
        
    Dim pIJLine As IJLine
    Set pIJLine = pObject
    
    Dim z1 As Double
    Dim z2 As Double
    
    z1 = (dMaxDepth / 2) - (dDepthStart / 2)
    z2 = (dMaxDepth / 2) - (dDepthEnd / 2)
            
    Dim dDepthDelta As Double
    Dim dYOffset As Double
    
    dDepthDelta = Abs(dDepthStart - dDepthEnd)
    
    If taperLocation = Center Then
        dYOffset = (dDepthDelta / 2 * dLengthExt) / dLength
    Else
        dYOffset = (dDepthDelta * dLengthExt) / dLength
    End If
    
    If dDepthStart < dMaxDepth Then
        dYOffset = -dYOffset
    End If
    
    If (taperLocation = Bottom) Then
        pIJLine.DefineBy2Points 0 - dLengthExt, -(dBottomWidth * 0.5) - (dMaxWidth * 0.5), dMaxDepth - dDepthStart - dYOffset, _
                            dLength + dLengthExt, -(dBottomWidth * 0.5) - (dMaxWidth * 0.5), dMaxDepth - dDepthEnd + dYOffset
    ElseIf (taperLocation = Top) Then
        pIJLine.DefineBy2Points 0 - dLengthExt, -(dBottomWidth * 0.5) - (dMaxWidth * 0.5), 0, _
                            dLength + dLengthExt, -(dBottomWidth * 0.5) - (dMaxWidth * 0.5), 0
    Else ' center
        pIJLine.DefineBy2Points 0 - dLengthExt, -(dBottomWidth * 0.5) - (dMaxWidth * 0.5), z1 - dYOffset, _
                            dLength + dLengthExt, -(dBottomWidth * 0.5) - (dMaxWidth * 0.5), z2 + dYOffset
    End If
    
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
    
    Set pCurveToExtrude = pIJDMemberObject.ItemByDispid(4)
    Set pExtrusionCurve = pIJDMemberObject.ItemByDispid(5)
                            
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

    Dim dDepthStart As Double
    Dim dDepthEnd As Double
    Dim dBottomWidth As Double
    Dim dTopWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
    
    GetTaperedIWExtrusionParameters oSmartOcc, dDepthStart, dDepthEnd, dBottomWidth, dTopWidth, dLength

    Dim pIJLine As IJLine
    Set pIJLine = pObject
    
    Dim dMaxWidth As Double
    dMaxWidth = GetMaxWidth(dBottomWidth, dTopWidth)

    Dim dMaxDepth As Double
    dMaxDepth = GetMaxDepth(dDepthStart, dDepthEnd)
    
    Dim dWebExtension As Double
    m_oDesignMemberHelper.GetWebDepthExtension oSmartOcc, dWebExtension
    
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dMaxDepth, dMaxWidth, oMatrix
    
    ' Web depth has to be expanded when extension exists
    Dim taperLocation As Tapered
    taperLocation = GetTaperLocation(oSmartOcc)
    
    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt
    
    Dim dDepthDelta As Double
    Dim dYOffset As Double
    
    dDepthDelta = Abs(dDepthStart - dDepthEnd)
    
    If taperLocation = Center Then
        dYOffset = (dDepthDelta / 2 * dLengthExt) / dLength
    Else
        dYOffset = (dDepthDelta * dLengthExt) / dLength
    End If

    If dWebExtension > 0 Then
        pIJLine.DefineBy2Points 0, -dMaxWidth / 2, 0 - dWebExtension - dYOffset, _
                                0, -dMaxWidth / 2, dMaxDepth + dWebExtension + dYOffset
    Else
        Dim BOUNDING_EXT As Double
        BOUNDING_EXT = m_oDesignMemberHelper.BoundingExt
        
        pIJLine.DefineBy2Points 0, -dMaxWidth / 2, 0 - BOUNDING_EXT - dYOffset, _
                                0, -dMaxWidth / 2, dMaxDepth + BOUNDING_EXT + dYOffset
    End If
    
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

    Dim dDepthStart As Double
    Dim dDepthEnd As Double
    Dim dBottomWidth As Double
    Dim dTopWidth As Double
    Dim dLength As Double
    Dim oMatrix As IJDT4x4
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
        
    GetTaperedIWExtrusionParameters oSmartOcc, dDepthStart, dDepthEnd, dBottomWidth, dTopWidth, dLength
    
    Dim dMaxWidth As Double
    dMaxWidth = GetMaxWidth(dBottomWidth, dTopWidth)

    Dim dMaxDepth As Double
    dMaxDepth = GetMaxDepth(dDepthStart, dDepthEnd)
    
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dMaxDepth, dMaxWidth, oMatrix
    
    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt
   
    Dim pIJLine As IJLine
    Set pIJLine = pObject
    
    pIJLine.DefineBy2Points 0 - dLengthExt, -dMaxWidth / 2, 0, _
                            dLength + dLengthExt, -dMaxWidth / 2, 0
       
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

    ' Create the curves to extrude
    Dim pCurveToExtrude As IJCurve
    Dim pExtrusionCurve As IJCurve

    Set pCurveToExtrude = pIJDMemberObject.ItemByDispid(7)
    Set pExtrusionCurve = pIJDMemberObject.ItemByDispid(8)

    Set pObj = m_oDesignMemberHelper.CreateABuitUpPlateWithProfileOrientation(pResourceManager, pCurveToExtrude, _
                                                                              pExtrusionCurve, dThickness, _
                                                                              strMaterial, strGrade, oSmartOcc, WebPlate, _
                                                                              RightDir, , , , Centered, MoldedFormPlateNameCat.NameCatWeb)
                                            
    ' Add the boundaries
    Dim oTopFlange As Object
    Dim oBottomFlange As Object
    
    Set oTopFlange = pIJDMemberObject.ItemByDispid(3)
    Set oBottomFlange = pIJDMemberObject.ItemByDispid(6)
    
    m_oDesignMemberHelper.SetPlateBoundaries oSmartOcc, pObj, oTopFlange, oBottomFlange, Nothing, Nothing
                                            
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
    Dim dDepthStart As Double
    Dim dDepthEnd As Double
    Dim dBottomWidth As Double
    Dim dTopWidth As Double
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pIDesignedMember
    
    GetTaperedIWExtrusionParameters oSmartOcc, dDepthStart, dDepthEnd, dBottomWidth, dTopWidth, dLength
    
  
    If Not pPosAlong Is Nothing Then
    
        Dim pISPSMemberPartCommon As ISPSMemberPartCommon
        Set pISPSMemberPartCommon = pIDesignedMember
        Dim axisCurve As IJCurve
        Set axisCurve = pISPSMemberPartCommon.Axis
        'BUITaperedWeb is a linear transition
        ' get the difference from start to end
        Dim dDepthDiff As Double
        dDepthDiff = dDepthEnd - dDepthStart
        ' get the parameter of the input position
        Dim posParam As Double
        axisCurve.Parameter pPosAlong.x, pPosAlong.y, pPosAlong.z, posParam
        ' get the parameter range
        Dim startParam As Double
        Dim endParam As Double
        
        axisCurve.ParamRange startParam, endParam
        
        If posParam < startParam Then
            posParam = startParam
        ElseIf posParam > endParam Then
            posParam = endParam
        End If
        
        
        'get the ratio of the input point along the curve
        Dim dPositionRatio As Double
        dPositionRatio = posParam / (endParam - startParam)
        
        
        pdDepth = dDepthStart + (dDepthDiff * dPositionRatio)
    Else
        pdDepth = GetMaxDepth(dDepthStart, dDepthEnd)
    End If
    pdWidth = GetMaxWidth(dBottomWidth, dTopWidth)
   
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
    m_oDesignMemberHelper.SetCrossSectionReadOnlyAttributesSpecCase pIJDAttrs, CollAllDisplayedValues
    
    If bIsModifiable = False Then
        Set pAttrColl = CollAllDisplayedValues
        For i = 1 To pAttrColl.Count
            Set pAttrDescr = pAttrColl.Item(i)
            If pAttrDescr.InterfaceName = IID_IUABuiltUpTopFlange Then
                If pAttrDescr.AttrName = "TopFlangeWidth" Then
                    pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
                End If
            ElseIf pAttrDescr.InterfaceName = IID_IUABuiltUpBottomFlange Then
                If pAttrDescr.AttrName = "BottomFlangeWidth" Then
                    pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
                End If
            ElseIf pAttrDescr.InterfaceName = IID_IUABuiltUpITaperWeb Then
                If pAttrDescr.AttrName = "DepthStart" Or pAttrDescr.AttrName = "DepthEnd" Then
                    pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
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

Private Sub GetTaperedIWExtrusionParameters(ByVal oSmartOcc As IJSmartOccurrence, _
                                           ByRef dDepthStart As Double, _
                                           ByRef dDepthEnd As Double, _
                                           ByRef dBottomWidth As Double, _
                                           ByRef dTopWidth As Double, _
                                           ByRef dLength As Double)
                                           
Const METHOD = "GetTaperedIWExtrusionParameters"

On Error GoTo ErrorHandler
    
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes

    ' Get Parameters of BU
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpITaperWeb)
    If Not oAttrCol Is Nothing Then
        dDepthStart = oAttrCol.Item("DepthStart").Value
    Else
        GoTo ErrorHandler
    End If
    
    Set oAttrCol = Nothing
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpITaperWeb)
    If Not oAttrCol Is Nothing Then
        dDepthEnd = oAttrCol.Item("DepthEnd").Value
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
    
    Set oAttrCol = Nothing
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTopFlange)
    If Not oAttrCol Is Nothing Then
        dTopWidth = oAttrCol.Item("TopFlangeWidth").Value
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

Private Function GetTaperLocation(ByVal oSmartOcc As IJSmartOccurrence) As Tapered
Const METHOD = "GetTaperLocation"
On Error GoTo ErrorHandler
    
    Dim dTaperLocation As Double
    Dim lCardinalPoint As Integer
    Dim pISPSCrossSection As ISPSCrossSection
    
    Set pISPSCrossSection = oSmartOcc
    lCardinalPoint = pISPSCrossSection.CardinalPoint
    
    Select Case lCardinalPoint
        Case 1, 2, 3, 11
            GetTaperLocation = Top
        Case 4, 5, 6, 12, 13, 15
            GetTaperLocation = Center
        Case 7, 8, 9, 14
            GetTaperLocation = Bottom
        Case Else
            GoTo ErrorHandler
    End Select

Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function

Private Function GetMaxDepth(ByVal dDepthStart As Double, ByVal dDepthEnd As Double) As Double
Const METHOD = "GetMaxDepth"
On Error GoTo ErrorHandler
    
    If dDepthStart > dDepthEnd Then
        GetMaxDepth = dDepthStart
    Else
        GetMaxDepth = dDepthEnd
    End If

Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function

Private Function GetMinDepth(ByVal dDepthStart As Double, ByVal dDepthEnd As Double) As Double
Const METHOD = "GetMinDepth"
On Error GoTo ErrorHandler
    
    If dDepthStart > dDepthEnd Then
        GetMinDepth = dDepthEnd
    Else
        GetMinDepth = dDepthStart
    End If

Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function

Private Function GetMaxWidth(ByVal dBottomWidth As Double, ByVal dTopWidth As Double) As Double
Const METHOD = "GetMaxWidth"
On Error GoTo ErrorHandler
    
    If dBottomWidth > dTopWidth Then
        GetMaxWidth = dBottomWidth
    Else
        GetMaxWidth = dTopWidth
    End If

Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function

Public Function AreOccurrencePropertiesValid(ByVal oSmartOcc As IJSmartOccurrence) As Boolean
Const METHOD = "AreOccurrencePropertiesValid"
On Error GoTo ErrorHandler
    
    AreOccurrencePropertiesValid = False
    
    Dim dDepthStart As Double
    Dim dDepthEnd As Double
    Dim dBottomWidth As Double
    Dim dTopWidth As Double
    Dim dLength As Double
    Dim sError As String
    
    GetTaperedIWExtrusionParameters oSmartOcc, dDepthStart, dDepthEnd, dBottomWidth, dTopWidth, dLength
    
    sError = BuiltUpDefValidate("DepthStart", dDepthStart, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("DepthEnd", dDepthEnd, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
        
    sError = BuiltUpDefValidate("BottomFlangeWidth", dBottomWidth, 0#)
    
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

    Dim dBtmExt As Double
    m_oDesignMemberHelper.GetBottomWidthExtension oSmartOcc, dBtmExt
    
    sError = BuiltUpDefValidate("BottomFlangeWidthExt", dBtmExt, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
  
    AreOccurrencePropertiesValid = True

Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function

Private Sub CalcuteOutputParameters(ByVal oSmartOcc As IJSmartOccurrence)
Const METHOD = "CalcuteOutputParameters"
On Error GoTo ErrorHandler
        
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    Dim oSmartItem As IJSmartItem
    
    Dim dTopWidth As Double
    Dim dTopThickness As Double
    Dim dBottomWidth As Double
    Dim dBottomThickness As Double
    Dim dDepth As Double
    Dim dWebThickness As Double
    Dim dDepthStart As Double
    Dim dDepthEnd As Double
    
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
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTopFlange)
    If Not oAttrCol Is Nothing Then
        dTopWidth = oAttrCol.Item("TopFlangeWidth").Value
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
    
    Set oAttrCol = Nothing
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpITaperWeb)
    If Not oAttrCol Is Nothing Then
        dDepthStart = oAttrCol.Item("DepthStart").Value
    Else
        GoTo ErrorHandler
    End If
    
    Set oAttrCol = Nothing
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpITaperWeb)
    If Not oAttrCol Is Nothing Then
        dDepthEnd = oAttrCol.Item("DepthEnd").Value
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
    
    Set oAttrCol = Nothing
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpBottomFlange)
    If Not oAttrCol Is Nothing Then
        dBottomThickness = oAttrCol.Item("BottomFlangeThickness").Value
    Else
        GoTo ErrorHandler
    End If
            
    Dim dMaxDepth As Double
    dMaxDepth = GetMaxDepth(dDepthStart, dDepthEnd)
    
    Dim dMaxWidth As Double
    dMaxWidth = GetMaxWidth(dBottomWidth, dTopWidth)
    
    Dim ErrorStatus As ErrorSectionStatus
    ErrorStatus = GetSectionPropertiesFromCalculator(dArea, dPerimeter, dCentroidX, dCentroidY, _
                                                     dXp, dYp, dIxx, dIyy, dRo, dRxx, dRxy, dRyy, _
                                                     dSw, dSxx, dSyy, dZxx, dZyy, dJ, dCw, dH, dUnitWt, _
                                                     dMaxDepth, dMaxWidth, dTopWidth, dBottomWidth, _
                                                     dWebThickness, dBottomThickness, dTopThickness)
                                                     
    If ErrorStatus <> SectionProperties_OK Then
        GoTo ErrorHandler
    End If
    
    Set oAttr = Nothing
    Set oAttrCol = Nothing
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    If Not oAttrCol Is Nothing Then
        oAttrCol.Item("Width").Value = dMaxWidth
        oAttrCol.Item("Depth").Value = dMaxDepth
        oAttrCol.Item("Area").Value = dArea
        oAttrCol.Item("Perimeter").Value = dPerimeter
    Else
        GoTo ErrorHandler
    End If
    
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDesignProperties)
    If Not oAttrCol Is Nothing Then
        On Error Resume Next
        oAttrCol.Item("IsHollow").Value = False
        oAttrCol.Item("IsSymmetricAboutX").Value = SymmetricAboutX()
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
        Case "DepthStart", "DepthEnd", "Length", "TopFlangeWidth", "BottomFlangeWidth"
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
                                                    ByVal dMaxDepth As Double, _
                                                    ByVal dMaxWidth As Double, _
                                                    ByVal dTopFlangeWidth As Double, _
                                                    ByVal dBotFlangeWidth As Double, _
                                                    ByVal dWebThickness As Double, _
                                                    ByVal dBotThickness As Double, _
                                                    ByVal dTopThickness As Double) As ErrorSectionStatus
                                               
    
Const METHOD = "GetSectionPropertiesFromCalculator"
On Error GoTo ErrorHandler
    
    Dim ErrorStatus As ErrorSectionStatus
    Dim bIsSetOk As Boolean
      
    m_oCalcXProps.SectionTypeAlias = EnumSectionTypeAlias.Section_IGS
    m_oCalcXProps.Width = dMaxWidth
    m_oCalcXProps.depth = dMaxDepth
    m_oCalcXProps.BottomFlangeWidth = dBotFlangeWidth
    m_oCalcXProps.bottomFlangeThickness = dBotThickness
    m_oCalcXProps.webThickness = dWebThickness
    m_oCalcXProps.TopFlangeWidth = dTopFlangeWidth
    m_oCalcXProps.topFlangeThickness = dTopThickness

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
    
    eSectionAlias = Section_IG
    'get assy members from SO
    Set oAssyMembers = oDesignedMember
    
    'First handle the top flange
    Set oPlateSystem = oAssyMembers.Item(BUIMembers.TopFlange)
    If Not oPlateSystem Is Nothing Then
        m_oDesignMemberHelper.EvaluateTopFlangePlateSystemXIDs oPlateSystem, eSectionAlias, oDesignedMember
    End If
    'handle the bottom flange
    Set oPlateSystem = oAssyMembers.Item(BUIMembers.BotFlange)
    If Not oPlateSystem Is Nothing Then
        m_oDesignMemberHelper.EvaluateBotFlangePlateSystemXIDs oPlateSystem, eSectionAlias, oDesignedMember
    End If
    'handle the web
    Set oPlateSystem = oAssyMembers.Item(BUIMembers.Web)
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
            Case "IUABuiltUpBotFlange"
                If sAttrName = "BottomFlangeWidth" Then
                    IsDesignParameter = True
                End If
            Case "IUABuiltUpTopFlange"
                If sAttrName = "TopFlangeWidth" Then
                    IsDesignParameter = True
                End If
            Case "IUABuiltUpITaperWeb"
            'actualy don't neeed this logic below. All attributes on this I/F are design attrs but
            ' its an extra check in case user decides to add a non-design attribute on the same I/F
                Select Case sAttrName
                    Case "DepthStart"
                        IsDesignParameter = True
                    Case "DepthEnd"
                        IsDesignParameter = True
                    Case Else
                End Select ' attr name
            
            Case Else
        End Select ' I/F name
    End If

Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function
    
Private Function SymmetricAboutX() As Boolean
    SymmetricAboutX = False
    If ((Abs(m_oCalcXProps.TopFlangeWidth - m_oCalcXProps.BottomFlangeWidth) < dTol) And _
        (Abs(m_oCalcXProps.topFlangeThickness - m_oCalcXProps.bottomFlangeThickness) < dTol)) Then
        SymmetricAboutX = True
    End If

End Function
    
Private Sub ICustomSectionShapeService_GetCrossSectionData(ByVal pProfileObject As Object, ByVal distFromStart As Double, ByVal eRepresentationType As SP3DStructInterfaces.structShapeRepresentationEnum, ByVal bBreakIntoComponents As Boolean, sectionPrev As SP3DStructInterfaces.IStructSectionShape, sectionNext As SP3DStructInterfaces.IStructSectionShape)
Const METHOD = "IStructSectionShapeService_GetCrossSectionData"
On Error GoTo ErrorHandler
    
    ' we're only creating a section_profile_i_type for now.  There is no use case for creating a compound shape
    ' or a graphic shape yet and no way for the user to ask for something different.
    
    ' in the future, utility methods will be implemented to convert to a graphic shape and to decompose the shape
    ' if there is a need to do this

    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pProfileObject
    Dim dDepthStart As Double
    Dim dDepthEnd As Double
    Dim dBottomWidth As Double
    Dim dTopWidth As Double
    Dim dLength As Double
    
    GetTaperedIWExtrusionParameters oSmartOcc, dDepthStart, dDepthEnd, dBottomWidth, dTopWidth, dLength
    
    ' the only parameter that should vary along the length of this member is its depth
    Dim depth As Double
    If distFromStart < dTol Then
        ' user wants the definition at the start
        depth = dDepthStart
    ElseIf dLength - distFromStart < dTol Then
        ' also handles the case where distFromStart is greater than the length
        depth = dDepthEnd
    Else
        ' somewhere along the length of the member
        'BUITaperedWeb is a linear transition
        ' get the difference from start to end
        Dim dDepthDiff As Double
        dDepthDiff = dDepthEnd - dDepthStart
        depth = dDepthStart + ((distFromStart / dLength) * dDepthDiff)
    End If
    

    Dim sectionMaterial As IStructSectionMaterial
    Dim webThickness As Double
    Dim topFlangeThickness As Double
    Dim bottomFlangeThickness As Double
    Dim internalDepth As Double
    Dim flangeSlope As Double
    Dim rootRadius As Double
    Dim edgeRadius As Double
    Dim topFlangeMaterial As String
    Dim topFlangeGrade As String
    
    GetComponentMaterial pProfileObject, IID_IUABuiltUpTopFlange, topFlangeMaterial, topFlangeGrade, topFlangeThickness
    
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes

    Set oAttr = oSmartOcc.ItemObject
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpWeb)
    If Not oAttrCol Is Nothing Then
        webThickness = oAttrCol.Item("WebThickness").Value
    End If
    
    Set oAttrCol = Nothing
   
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpBottomFlange)
    If Not oAttrCol Is Nothing Then
        bottomFlangeThickness = oAttrCol.Item("BottomFlangeThickness").Value
    End If


    internalDepth = depth - (topFlangeThickness + bottomFlangeThickness)
    flangeSlope = 0#
    rootRadius = 0#
    edgeRadius = 0#
    
    
    ' always create a parametric shape.  It may be used to create the other shapes as needed
    Dim parametricShape As IStructSectionIAsymmetricShape
    Set parametricShape = New TransientClassesLib.StructSectionIAsymmetricShape
    
    parametricShape.SetProperties depth, dTopWidth, dBottomWidth, webThickness, topFlangeThickness, bottomFlangeThickness, internalDepth, flangeSlope, flangeSlope, rootRadius, rootRadius, edgeRadius, edgeRadius
    Set sectionMaterial = parametricShape
    sectionMaterial.SetMaterial topFlangeMaterial, topFlangeGrade
        If distFromStart < dTol Then
        ' user wants the definition at the start
        Set sectionNext = parametricShape
        Set sectionPrev = Nothing
    ElseIf dLength - distFromStart < dTol Then
        ' also handles the case where distFromStart is greater than the length
        Set sectionNext = Nothing
        Set sectionPrev = parametricShape

    Else
        ' somewhere along the length of the member
        Set sectionPrev = parametricShape
        Set sectionNext = sectionPrev
    End If
         
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Sub ICustomSectionShapeService_GetTransitionLocations(ByVal pProfileObject As Object, distFromStart() As Double)
Const METHOD = "IStructSectionShapeService_GetTransitionLocations"
On Error GoTo ErrorHandler

    ReDim distFromStart(0 To 1)
    distFromStart(0) = 0#
    ' Get Length of the member curve
    Dim pISPSMemberPartCommon As ISPSMemberPartCommon
    Set pISPSMemberPartCommon = pProfileObject
    distFromStart(1) = pISPSMemberPartCommon.Axis.Length

Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

