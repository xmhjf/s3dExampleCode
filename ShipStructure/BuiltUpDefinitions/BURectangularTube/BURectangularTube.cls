VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BURectTube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2011, Intergraph Corporation. All rights reserved.
'
'File
'    BUSquareTube.cls
'
'Author
'       9/2011  cre
'
'Description
'       BU Square Tube definition. Derived from the BUTube definition
'
'Notes
'
'History:
'
'   22-Sept-2009 GG TR#167167 - DesignedMember does not set PG of plates to its own PG
'   06-Aug-2010 GG DI-169828  SCHEMA: Interface to indicate member symmetry and open/closed cross sections
'*******************************************************************
'                   <Radius>
'
'                  (Z)
'                   |
'                   | /---------\
'                   ||           |
'                   ||           |
'                   ||     +     |
'                   ||           |
'                   ||           |
'                   ||           |
'                   | \---------/
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
'  Shape is built in local Member CS and transformed into location for extrusion as Linear-Extruses PlateSystem
'  Rounded corners are created automatically by the plate system where the curve breaks as bent knuckles
'  Knuckle radius is cointrolled by rule
'  (+X) is into the paper, direction of the extrusion
'  (0,0,0) is Cardinal Point = 1
'
'   NOTE: Cardinal Pts are based on initial thickness, if user changes plate web
'   thickness, then cardinal points will not be correct!!!

Option Explicit

Private Const MODULE = "BURectTube"
'Private Const strSourceFile = "BUSquareTube.def"
Private Const CONST_ItemProgId As String = "SM3DBURectangularTube.BURectTube"

Private m_strErrorDescr As String
Private m_oLocalizer As IJLocalizer
Private m_oDesignMemberHelper As BUHelperUtils
Private m_oCalcXProps As BUCalcSectionProperties

Private Enum BUTubeMembers
  Tube = 3
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
     pAD.UserTypeClsid = "{573C9A29-CAD1-4b48-B6DE-FA4562B0B09A}"   'CBUSquareTube (Generated here)
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
     pAPDs.AddProperty "IUABuiltUpRectTube", 6, IID_IUABuiltUpRectTube, , imsCOOKIE_ID_USS_LIB
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
'' Tube Shape (curve to extrude)
'''''

     Set pMemberDescription = pMemberDescriptions.AddMember("TubeCurveToExtrude", 1, "CMConstructCurveToExtrude", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "TubeCurveToExtrudeProperties", 1, IJCurve, "CMComputeCurveToExtrude", imsCOOKIE_ID_USS_LIB

'''''
'' Tube Extrusion curve
'''''

     Set pMemberDescription = pMemberDescriptions.AddMember("TubeExtrusionPath", 2, "CMConstructExtrusionPath", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "TubeExtrusionPathProperties", 1, IJCurve, "CMComputeExtrusionPath", imsCOOKIE_ID_USS_LIB

'''''
'' Tube
'''''
     Set pMemberDescription = pMemberDescriptions.AddMember("Tube", 3, "CMConstructTubePlateSystem", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructTubePlateSystem"
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateTubePlateSystem"
          
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "TubeProperties", 1, IJCurve, "CMComputeTubePlateSystem", imsCOOKIE_ID_USS_LIB
             
    
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
'                       TUBE
'-------------------------------------------------------

Public Sub CMConstructCurveToExtrude(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructCurveToExtrude "
On Error GoTo ErrorHandler

    'Create a Line/Curve to use
    Dim pCurveToExtrude As IJCurve

    'Create the square tube shape by linestring points in CCW order from lower left corner in Member local coordinate system
    Dim pGeomFact As New GeometryFactory
    Dim dPoints(1 To 15) As Double
    dPoints(1) = 0
    dPoints(2) = 0
    dPoints(3) = 0
    
    dPoints(4) = 0
    dPoints(5) = -1
    dPoints(6) = 0

    dPoints(7) = 0
    dPoints(8) = -1
    dPoints(9) = 1
    
    dPoints(10) = 0
    dPoints(11) = 0
    dPoints(12) = 1
    
    dPoints(13) = 0
    dPoints(14) = 0
    dPoints(15) = 0
    
    Set pCurveToExtrude = pGeomFact.LineStrings3d.CreateByPoints(pResourceManager, 5, dPoints)
    
    CopyPermissionGroup pCurveToExtrude, pMemberDescription.CAO
    
    Set pObj = pCurveToExtrude
    
    'Make the construction curve geometry invisible
    Dim oControlFlags As IJControlFlags
    Set oControlFlags = pObj
    oControlFlags.ControlFlags(&H4) = &H4

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMComputeCurveToExtrude(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeCurveToExtrude"
On Error GoTo ErrorHandler

    Dim oMatrix As IJDT4x4
    Dim dLength As Double
    Dim dWidth As Double
    Dim dDepth As Double
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
    
    GetTubeExtrusionParameters oSmartOcc, dWidth, dDepth, dLength
    
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDepth, dWidth, oMatrix
    
    ' Modify the input object/line
    Dim pILineString As IJLineString
    Set pILineString = pObject
    
    Dim dPoints(1 To 15) As Double
    dPoints(1) = 0
    dPoints(2) = 0
    dPoints(3) = 0
    
    dPoints(4) = 0
    dPoints(5) = -dWidth
    dPoints(6) = 0
    
    dPoints(7) = 0
    dPoints(8) = -dWidth
    dPoints(9) = dDepth
    
    dPoints(10) = 0
    dPoints(11) = 0
    dPoints(12) = dDepth
    
    dPoints(13) = 0
    dPoints(14) = 0
    dPoints(15) = 0
    
    pILineString.SetPoints 5, dPoints

    pILineString.Transform oMatrix
   
Exit Sub
Set oSmartOcc = Nothing
ErrorHandler:  HandleError MODULE, METHOD
SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
Err.Raise E_FAIL
End Sub

Public Sub CMConstructExtrusionPath(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructExtrusionPath"
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

Public Sub CMComputeExtrusionPath(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeExtrusionPath"
On Error GoTo ErrorHandler

    Dim oMatrix As IJDT4x4
    Dim dDepth As Double
    Dim dWidth As Double
    Dim dLength As Double
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
    
    GetTubeExtrusionParameters oSmartOcc, dWidth, dDepth, dLength

    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDepth, dWidth, oMatrix

    Dim pIJLine As IJLine
    Set pIJLine = pObject
    
    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt
    
    pIJLine.DefineBy2Points 0 - dLengthExt, 0, 0, _
                            dLength + dLengthExt, 0, 0
    
    pIJLine.Transform oMatrix
   
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
Err.Raise E_FAIL
End Sub

Public Sub CMConstructTubePlateSystem(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructTubePlateSystem"
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
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpRectTube)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("TubeThickness").Value
        strMaterial = oAttrCol.Item("TubeMaterial").Value
        strGrade = oAttrCol.Item("TubeGrade").Value
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
                                                        strMaterial, strGrade, oSmartOcc, TubePlate, _
                                                        , , , InDir, MoldedFormPlateNameCat.NameCatTube)
                                  
    m_oDesignMemberHelper.SetPlateBoundaries oSmartOcc, pObj
     


Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructTubePlateSystem(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructTubePlateSystem"
On Error GoTo ErrorHandler

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub
Public Sub CMComputeTubePlateSystem(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeTubePlateSystem"
On Error GoTo ErrorHandler

    ' get the  plate
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    Dim pIJDMemberObject As IJDMemberObjects
    Set pIJDMemberObject = oSmartOcc
    Dim oTubePlate As IJPlate
    Set oTubePlate = pIJDMemberObject.ItemByDispid(Tube)
    
    ' get the parameters
    Dim oSmartItem As IJSmartItem
    Set oSmartItem = oSmartOcc.ItemObject
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
   
    Dim dThickness As Double
    Dim strMaterial As String
    Dim strGrade As String
    Set oAttr = oSmartItem
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpRectTube)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("TubeThickness").Value
        strMaterial = oAttrCol.Item("TubeMaterial").Value
        strGrade = oAttrCol.Item("TubeGrade").Value
    Else
        GoTo ErrorHandler
    End If
    
    ' apply the new parameters
    With m_oDesignMemberHelper
        .SetMaterialAndGrade oTubePlate, strMaterial, strGrade
        .SetPlateDimensions oTubePlate, dThickness
    End With

    ProcessKnuckles oTubePlate
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub
Public Sub CMMigrateTubePlateSystem(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateTubePlateSystem "
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

    CalculateOutputParameters pPropertyDescriptions.CAO
    
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
    
    Dim dDepth As Double
    Dim dWidth As Double
    Dim dLength As Double
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pIDesignedMember

    GetTubeExtrusionParameters oSmartOcc, dWidth, dDepth, dLength

    pdWidth = dWidth
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
                    Select Case pAttrToChange.AttrName
                        Case "TubeDepth", "TubeWidth", "BendRadius"
                            Dim dWidth#, dDepth#, dLength#, dBendRadius#
                            'Get other values
                            GetTubeExtrusionParameters pIJDAttrs, dWidth, dDepth, dLength
                            GetRectangularTubeProperties pIJDAttrs, dBendRadius
                            'Now update the new value
                            Select Case pAttrToChange.AttrName
                                Case "TubeDepth"
                                    dDepth = varNewAttrValue
                                Case "TubeWidth"
                                    dWidth = varNewAttrValue
                                Case "BendRadius"
                                    dBendRadius = varNewAttrValue
                                End Select
                                IJUserAttributeMgmt_OnAttributeChange = IsBendRadiusValid(pIJDAttrs, dWidth, dDepth, dBendRadius)
                                If Len(IJUserAttributeMgmt_OnAttributeChange) > 0 Then
                                    Exit Function
                                End If
                        End Select
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
    
    If Not bIsModifiable Then
        Set pAttrColl = CollAllDisplayedValues
        For i = 1 To pAttrColl.Count
            Set pAttrDescr = pAttrColl.Item(i)
            If pAttrDescr.InterfaceName = IID_IUABuiltUpRectTube Then
                If pAttrDescr.AttrName = "TubeDepth" Or pAttrDescr.AttrName = "TubeWidth" Then
                    pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
                End If
            End If
        Next
    End If
    
    'set the common read only attributes
    m_oDesignMemberHelper.SetCrossSectionReadOnlyAttributesSpecCase pIJDAttrs, CollAllDisplayedValues
    
    Dim bSectionProperties As Boolean
    bSectionProperties = m_oDesignMemberHelper.AreSectionPropertiesModifiable(pIJDAttrs)
     
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

Private Sub GetTubeExtrusionParameters(ByVal oSmartOcc As IJSmartOccurrence, _
                                       ByRef dWidth As Double, _
                                       ByRef dDepth As Double, _
                                       ByRef dLength As Double)

Const METHOD = "GetTubeExtrusionParameters"
On Error GoTo ErrorHandler
    
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes

    ' Get Parameters of BU
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes("IUABuiltUpRectTube")
    If Not oAttrCol Is Nothing Then
        dWidth = oAttrCol.Item("TubeWidth").Value
        dDepth = oAttrCol.Item("TubeDepth").Value
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

'Get properties specific to this (rectangular) built-up
Private Sub GetRectangularTubeProperties(ByVal oSmartOcc As IJSmartOccurrence, _
                                       ByRef dBendRadius As Double)

Const METHOD = "GetRectangularTubeProperties"
On Error GoTo ErrorHandler
    
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes

    ' Get Parameters of BU
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes("IUABuiltUpRectTube")
    If Not oAttrCol Is Nothing Then
        dBendRadius = oAttrCol.Item("BendRadius").Value
    Else
        GoTo ErrorHandler
    End If
    
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub


Public Function AreOccurrencePropertiesValid(ByVal oSmartOcc As IJSmartOccurrence) As Boolean
Const METHOD = "AreOccurrencePropertiesValid"

    On Error GoTo ErrorHandler
    AreOccurrencePropertiesValid = False
    
    Dim dWidth As Double
    Dim dDepth As Double
    Dim dLength As Double
    Dim dBendRadius As Double
    Dim sError As String
    
    GetTubeExtrusionParameters oSmartOcc, dWidth, dDepth, dLength
    GetRectangularTubeProperties oSmartOcc, dBendRadius
    
    sError = BuiltUpDefValidate("TubeDepth", dDepth, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("TubeWidth", dWidth, 0#)
    
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
    
    'Bend radius should be between 0 and half of depth and width (the property which can be modifed by user)
    'Set oSmartItem = oSmartOcc.ItemObject
    Dim dTubeThickness As Double
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    Set oAttr = oSmartOcc.ItemObject
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpRectTube)
    If Not oAttrCol Is Nothing Then
        dTubeThickness = oAttrCol.Item("TubeThickness").Value
    Else
        GoTo ErrorHandler
    End If
    
    sError = BuiltUpDefValidate("BendRadius", dBendRadius, 0#, dDepth / 2 - dTubeThickness)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("BendRadius", dBendRadius, 0#, dWidth / 2 - dTubeThickness)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
  
    AreOccurrencePropertiesValid = True


Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function


Private Sub CalculateOutputParameters(ByVal oSmartOcc As IJSmartOccurrence)
Const METHOD = "CalcuteOutputParameters"
On Error GoTo ErrorHandler
        
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    
    Dim oSmartItem As IJSmartItem

    Dim dWidth As Double
    Dim dDepth As Double
    Dim dLength As Double
   Dim dTubeThickness As Double

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

    GetTubeExtrusionParameters oSmartOcc, dWidth, dDepth, dLength

    Set oSmartItem = oSmartOcc.ItemObject
    Set oAttr = oSmartItem
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTube)
    If Not oAttrCol Is Nothing Then
        dTubeThickness = oAttrCol.Item("TubeThickness").Value
    Else
        GoTo ErrorHandler
    End If

    Dim ErrorStatus As ErrorSectionStatus
    ErrorStatus = GetSectionPropertiesFromCalculator(dArea, dPerimeter, dCentroidX, dCentroidY, _
                                                     dXp, dYp, dIxx, dIyy, dRo, dRxx, dRxy, dRyy, _
                                                     dSw, dSxx, dSyy, dZxx, dZyy, dJ, dCw, dH, dUnitWt, _
                                                     dWidth, dDepth, dTubeThickness)

    If ErrorStatus <> SectionProperties_OK Then
        GoTo ErrorHandler
    End If

    Set oAttr = Nothing
    Set oAttrCol = Nothing
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    If Not oAttrCol Is Nothing Then
        oAttrCol.Item("Width").Value = dWidth
        oAttrCol.Item("Depth").Value = dDepth
        oAttrCol.Item("Area").Value = dArea
        oAttrCol.Item("Perimeter").Value = dPerimeter
    Else
        GoTo ErrorHandler
    End If

    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDesignProperties)
    If Not oAttrCol Is Nothing Then
        On Error Resume Next
        oAttrCol.Item("IsHollow").Value = True
        oAttrCol.Item("IsSymmetricAboutX").Value = True
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
                                    Optional dUpRange As Double = -1) As String
                                    
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
        Case "Width", "Depth", "Length", "BendRadius"
            If varAttributeValue <= m_oDesignMemberHelper.distTol Then
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
                                                    ByVal dWidth As Double, dDepth As Double, _
                                                    ByVal dTubeThickness As Double) As ErrorSectionStatus
                                               
    
Const METHOD = "GetSectionPropertiesFromCalculator"
On Error GoTo ErrorHandler

'TODO: Need to update the Section calculator project to add property calculation logic.

    Dim ErrorStatus As ErrorSectionStatus
'    Dim bIsSetOk As Boolean
'
'    m_oCalcXProps.SectionTypeAlias = EnumSectionTypeAlias.Section_Circular
'    m_oCalcXProps.Depth = dDiameter
'    m_oCalcXProps.WebThickness = dTubeThickness
'
'    ErrorStatus = m_oCalcXProps.SectionProperties(dArea, dPerimeter, dCentroidX, dCentroidY, _
'                                                  dXp, dYp, dIxx, dIyy, dRo, dRxx, dRxy, dRyy, _
'                                                  dSw, dSxx, dSyy, dZxx, dZyy, dJ, dCw, dH, dUnitWt)
                                                                         
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
    
    eSectionAlias = Section_Circular
    'get assy members from SO
    Set oAssyMembers = oDesignedMember
    
   'handle the web
    Set oPlateSystem = oAssyMembers.Item(BUTubeMembers.Tube)
    If Not oPlateSystem Is Nothing Then
        m_oDesignMemberHelper.EvaluateWebPlateSystemXIDs oPlateSystem, StructBUPlateType.Tube, eSectionAlias, oDesignedMember
    End If
Exit Sub
ErrorHandler: HandleError MODULE, METHOD
End Sub

Private Sub ISPSDesignedMemberHelper_ResolveAmbiguity(ByVal pDesignedMember As SPSMembers.ISPSDesignedMember)
Const METHOD = "ISPSDesignedMemberHelper_ResolveAmbiguity"
On Error GoTo ErrorHandler
    'Nothing required for this particular Built-Up Definition
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub


' check lets the caller decide to take a certain action for changes in design parameters, like notification
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
            Case "IUABuiltUpRectTube"
                If sAttrName = "TubeThickness" Or sAttrName = "BendRadius" Or sAttrName = "TubeWidth" Or sAttrName = "TubeDepth" Then
                    IsDesignParameter = True
                End If
            Case Else
        End Select ' I/F name
    End If


Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub ProcessKnuckles(oPlate As IJPlateSystem)
Const METHOD = "ProcessKnuckles"
On Error GoTo ErrorHandler

    Dim oPlateUtils As IJPlateAttributes
    Set oPlateUtils = New PlateUtils
    Dim oKnuckleColl As IJElements
    
    Set oKnuckleColl = oPlateUtils.GetKnucklesOnPlate(oPlate)

    Dim oKnuckle As IJKnuckle
    Dim oKnuckleRule As IJKnuckleAttributes
    Dim oBUCAD As Object
    Set oBUCAD = oPlate.ParentBuiltup
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = oBUCAD

    If oKnuckleColl.Count > 0 Then
        ' get the bend radius
        Dim dBendRadius As Double

        Dim oAttrCol As IJDAttributesCol
        Dim oAttr As IJDAttributes
   
        Set oAttr = oSmartOcc
        Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpRectTube)
        If Not oAttrCol Is Nothing Then
            dBendRadius = oAttrCol.Item("BendRadius").Value
        Else
            GoTo ErrorHandler
        End If
    
        
         Dim nIndex As Long
        For nIndex = 1 To oKnuckleColl.Count
            Set oKnuckle = oKnuckleColl.Item(nIndex)
            Set oKnuckleRule = oKnuckle

            oKnuckleRule.PutOverrideRules True, True ' override both the knuckletype and radius
            oKnuckle.KnuckleType = eKT_Bent
            oKnuckle.InnerRadius = dBendRadius
        Next
    Else
    ' Need to do something to trigger a recompute to correc tthe knuckles

    End If
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD

End Sub

Private Function IsBendRadiusValid(oSmartOcc As IJSmartOccurrence, dWidth#, dDepth#, dBendRadius#) As String
Const METHOD = "IsBendRadiusValid"
On Error GoTo ErrorHandler
    'Bend radius should be between 0 and half of depth and width (the property which can be modifed by user)
    'Set oSmartItem = oSmartOcc.ItemObject
    Dim dTubeThickness As Double
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    Set oAttr = oSmartOcc.ItemObject
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpRectTube)
    If Not oAttrCol Is Nothing Then
        dTubeThickness = oAttrCol.Item("TubeThickness").Value
    Else
        GoTo ErrorHandler
    End If
    
    IsBendRadiusValid = BuiltUpDefValidate("BendRadius", dBendRadius, 0#, dDepth / 2 - dTubeThickness)
    
    If Len(IsBendRadiusValid) > 0 Then
        Exit Function
    End If
    
    IsBendRadiusValid = BuiltUpDefValidate("BendRadius", dBendRadius, 0#, dWidth / 2 - dTubeThickness)
    
Exit Function
ErrorHandler:
    HandleError MODULE, METHOD
End Function

Private Sub ICustomSectionShapeService_GetCrossSectionData(ByVal pProfileObject As Object, ByVal distFromStart As Double, ByVal eRepresentationType As SP3DStructInterfaces.structShapeRepresentationEnum, ByVal bBreakIntoComponents As Boolean, sectionPrev As SP3DStructInterfaces.IStructSectionShape, sectionNext As SP3DStructInterfaces.IStructSectionShape)
    Dim Mat As String, Grade As String
    Dim iSectionMaterial As IStructSectionMaterial
    Dim dWidth#, dDepth#, dThickness#, dBentRadius#
    Dim oRectSection As IStructSectionRectangleHollowShape
    Dim pISPSMemberPartCommon As ISPSMemberPartCommon
    Dim dMembLength As Double

    
    'get sizes of component plates
    GetComponentParams pProfileObject, dWidth, dDepth, dThickness, dBentRadius
    GetComponentMaterials pProfileObject, Mat, Grade
    
    Set oRectSection = New StructSectionRectangleHollowShape
    oRectSection.SetProperties dDepth, dWidth, dThickness, (dBentRadius + dThickness), dBentRadius
    
    Set iSectionMaterial = oRectSection
    iSectionMaterial.SetMaterial Mat, Grade
    
    If distFromStart > distTol Then
        Set sectionPrev = oRectSection
    End If
    

    Set pISPSMemberPartCommon = pProfileObject
    dMembLength = pISPSMemberPartCommon.Axis.Length
    If (dMembLength - distFromStart) > distTol Then
        Set sectionNext = oRectSection
    End If

    

End Sub

Private Sub ICustomSectionShapeService_GetTransitionLocations(ByVal pProfileObject As Object, distFromStart() As Double)
    'the section size for this type do not change along the length
    'so just add the value of 0, representing the start of the member, to the array.
    ReDim distFromStart(0 To 0)
    
    distFromStart(0) = 0#
End Sub


Private Sub GetComponentParams(ByVal pProfileObject As Object, dWidth As Double, dDepth As Double, dThickness As Double, dBentRadius As Double)
    Dim oSmartOcc As IJSmartOccurrence
    Dim oSmartItem As IJSmartItem
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    
    Set oSmartOcc = pProfileObject
    
    Set oSmartItem = oSmartOcc.ItemObject
    
    Set oAttr = oSmartItem
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpRectTube)
    dThickness = oAttrCol.Item("TubeThickness").Value
    
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpRectTube)
    dWidth = oAttrCol.Item("TubeWidth").Value
    
    dDepth = oAttrCol.Item("TubeDepth").Value
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpRectTube)
    dBentRadius = oAttrCol.Item("BendRadius").Value

End Sub


Private Sub GetComponentMaterials(ByVal pProfileObject As Object, TTubeMat As String, TTubeGrade As String)
    
    Dim Thickness As Double
    GetComponentMaterial pProfileObject, IID_IUABuiltUpRectTube, TTubeMat, TTubeGrade, Thickness

End Sub


