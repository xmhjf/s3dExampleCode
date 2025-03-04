VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BUCone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2006, Intergraph Corporation. All rights reserved.
'
'File
'    BUCone.cls
'
'Author
'       27 February 2008  ras
'
'Description
'       Cone definition
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
'                   |     ----
'                   |   /      \
'                   | /     -    \
'                   ||    /   \   |
'                   ||   |  +  |  |
'                   ||    \   /   |
'                   | \     -    /
'                   |   \      /
'                   |     ----
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


Option Explicit

Private Const MODULE = "BUCone"
Private Const strSourceFile = "BUCone.def"
Private Const CONST_ItemProgId As String = "SM3DBUCone.BUCone"

Private m_strErrorDescr As String
Private m_oLocalizer As IJLocalizer
Private m_oDesignMemberHelper As BUHelperUtils
Private m_oCalcXProps As BUCalcSectionProperties

Private Enum BUTubeMembers
  Tube = 3
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
     pAD.UserTypeClsid = "{CDB46E65-2844-431e-B2D3-76ADFFE1B3CB}"   'BUCone (Generated here)
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
     pAPDs.AddProperty "IUABuiltUpCone", 6, IID_IUABuiltUpCone, , imsCOOKIE_ID_USS_LIB
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
'' Cone Curve To Revolve
'''''
      
     Set pMemberDescription = pMemberDescriptions.AddMember("ConeCurveToRevolve", 1, "CMConstructCurveToRevolve", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "ConeCurveToRevolveProperties", 1, IJCurve, "CMComputeCurveToRevolve", imsCOOKIE_ID_USS_LIB

'''''
'' cone's Axis To Revolve Around
'''''

     Set pMemberDescription = pMemberDescriptions.AddMember("ConeAxisOfRevolution", 2, "CMConstructAxisOfRevolution", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "ConeAxisOfRevolutionProperties", 1, IJCurve, "CMComputeAxisOfRevolution", imsCOOKIE_ID_USS_LIB

'''''
'' Cone Plate
'''''
     Set pMemberDescription = pMemberDescriptions.AddMember("Cone", 3, "CMConstructConePlateSystem", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructConePlateSystem"
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateConePlateSystem"
          
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "ConeProperties", 1, IJCurve, "CMComputeConePlateSystem", imsCOOKIE_ID_USS_LIB
             
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
'                       Cone
'-------------------------------------------------------

Public Sub CMConstructCurveToRevolve(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructCurveToRevolve "
On Error GoTo ErrorHandler

    ' Just Create a initial dummy Curve
    Dim pCurve As IJCurve
    
    m_oDesignMemberHelper.CreatePositionLineStringCurveBy2PointsWithNPoints pResourceManager, 0, -1, 0, _
                                                                                              1, -1, 0, _
                                                                                              3, pCurve
    CopyPermissionGroup pCurve, pMemberDescription.CAO
                                                                                              
    Dim oLineString3d As LineString3d
    Set oLineString3d = pCurve
    oLineString3d.SetPoint 1, 0, -1, 0
    oLineString3d.SetPoint 2, 1, -1, 0
    oLineString3d.SetPoint 3, 1.5, -1, 0
    
    Set pObj = pCurve
    
    Dim oControlFlags As IJControlFlags
    Set oControlFlags = pObj
    oControlFlags.ControlFlags(&H4) = &H4

Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMComputeCurveToRevolve(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeCurveToRevolve"
On Error GoTo ErrorHandler

    Dim oMatrix As IJDT4x4
    Dim dDiameterStart As Double
    Dim dDiameterEnd As Double
    Dim dLength As Double
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
    
    GetConeExtrusionParameters oSmartOcc, dDiameterStart, dDiameterEnd, dLength

    Dim oLineString3d As LineString3d
    Set oLineString3d = pObject
    
    Dim dDiameterMax As Double
    dDiameterMax = GetMaxDiameter(dDiameterStart, dDiameterEnd)
  
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDiameterMax, dDiameterMax, oMatrix
    
    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt
           
    Dim oPt1 As IJDPosition
    Set oPt1 = New DPosition
    
    Dim oPt2 As IJDPosition
    Set oPt2 = New DPosition
    
    Dim oPt3 As IJDPosition
    Set oPt3 = New DPosition
    
    Dim oConeDir As IJDVector
    Dim oAxisDir As IJDVector

    oPt1.Set 0, -(dDiameterMax - dDiameterStart) / 2, dDiameterMax / 2
    oPt2.Set dLength, -(dDiameterMax - dDiameterEnd) / 2, dDiameterMax / 2
    
    If dDiameterStart < dDiameterEnd Then
        Set oConeDir = oPt1.Subtract(oPt2)
        oConeDir.Length = 0.005
        Set oPt1 = oPt1.Offset(oConeDir)
        
        'add lenght extension to other side
        oConeDir.Length = -dLengthExt
        Set oPt2 = oPt2.Offset(oConeDir)
        
        oPt3.Set -dLengthExt, oPt1.y, oPt1.z
        
        With oLineString3d
            .SetPoint 1, oPt3.x, oPt3.y, oPt3.z
            .SetPoint 2, oPt1.x, oPt1.y, oPt1.z
            .SetPoint 3, oPt2.x, oPt2.y, oPt2.z
        End With
    Else
        Set oConeDir = oPt2.Subtract(oPt1)
        oConeDir.Length = 0.005
        Set oPt2 = oPt2.Offset(oConeDir)
       
        oPt3.Set oPt2.x + dLengthExt, oPt2.y, oPt2.z
        
        'add lenght extension to other side
        oConeDir.Length = -dLengthExt
        Set oPt1 = oPt1.Offset(oConeDir)
        
        With oLineString3d
            .SetPoint 1, oPt1.x, oPt1.y, oPt1.z
            .SetPoint 2, oPt2.x, oPt2.y, oPt2.z
            .SetPoint 3, oPt3.x, oPt3.y, oPt3.z
        End With
    End If
        
    oLineString3d.Transform oMatrix
  
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
Err.Raise E_FAIL
End Sub

Public Sub CMConstructAxisOfRevolution(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructAxisOfRevolution"
On Error GoTo ErrorHandler
    
    ' Just Create a initial dummy Curve
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

Public Sub CMComputeAxisOfRevolution(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeAxisOfRevolution"
On Error GoTo ErrorHandler

    Dim oMatrix As IJDT4x4
    Dim dDiameterStart As Double
    Dim dDiameterEnd As Double
    Dim dLength As Double
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
    
    GetConeExtrusionParameters oSmartOcc, dDiameterStart, dDiameterEnd, dLength

    Dim dDiameterMax As Double
    dDiameterMax = GetMaxDiameter(dDiameterStart, dDiameterEnd)
    
    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDiameterMax, dDiameterMax, oMatrix

    Dim pIJLine As IJLine
    Set pIJLine = pObject
    
    pIJLine.DefineBy2Points 0, -dDiameterMax / 2, dDiameterMax / 2, dLength, -dDiameterMax / 2, dDiameterMax / 2
    pIJLine.Transform oMatrix
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
SPSToDoErrorNotify "StructBUToDoMessages", 14, oSmartOcc, Nothing
Err.Raise E_FAIL
End Sub

Public Sub CMConstructConePlateSystem(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructConePlateSystem"
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
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpCone)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("ConeThickness").Value
        strMaterial = oAttrCol.Item("ConeMaterial").Value
        strGrade = oAttrCol.Item("ConeGrade").Value
    Else
        GoTo ErrorHandler
    End If
    
   ' create the curves to extrude
    Dim pCurveToRevolve As IJCurve
    Dim pAxisCurve As IJCurve
  
     ' Get the outputs
    Dim pIJDMemberObject As IJDMemberObjects
    Set pIJDMemberObject = oSmartOcc
    
    Set pCurveToRevolve = pIJDMemberObject.ItemByDispid(1)
    Set pAxisCurve = pIJDMemberObject.ItemByDispid(2)
    
    Set pObj = m_oDesignMemberHelper.CreateARevolvedBuiltUpPlate(pResourceManager, pCurveToRevolve, _
                                                                pAxisCurve, dThickness, _
                                                                strMaterial, strGrade, oSmartOcc, TubePlate, _
                                                                , , , InDir)
    m_oDesignMemberHelper.SetPlateBoundaries oSmartOcc, pObj
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructConePlateSystem(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructConePlateSystem"
On Error GoTo ErrorHandler
 
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub
Public Sub CMComputeConePlateSystem(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeConePlateSystem"
On Error GoTo ErrorHandler
    ' get the  plate
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    Dim pIJDMemberObject As IJDMemberObjects
    Set pIJDMemberObject = oSmartOcc
    Dim oPlate As IJPlate
    Set oPlate = pIJDMemberObject.ItemByDispid(Tube)
    
    ' get the parameters
    Dim oSmartItem As IJSmartItem
    Set oSmartItem = oSmartOcc.ItemObject
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
   
    Dim dThickness As Double
    Dim strMaterial As String
    Dim strGrade As String
    Set oAttr = oSmartItem
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpCone)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("ConeThickness").Value
        strMaterial = oAttrCol.Item("ConeMaterial").Value
        strGrade = oAttrCol.Item("ConeGrade").Value
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
Public Sub CMMigrateConePlateSystem(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateConePlateSystem "
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
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pIDesignedMember

    Dim dDiameterStart As Double
    Dim dDiameterEnd As Double
    Dim dLength As Double

    GetConeExtrusionParameters oSmartOcc, dDiameterStart, dDiameterEnd, dLength

    If Not pPosAlong Is Nothing Then
        Dim pISPSMemberPartCommon As ISPSMemberPartCommon
        Set pISPSMemberPartCommon = pIDesignedMember
        Dim axisCurve As IJCurve
        Set axisCurve = pISPSMemberPartCommon.Axis
        'BUCone is a linear transition
        ' get the difference from start to end
        Dim dDiameterDiff As Double
        dDiameterDiff = dDiameterEnd - dDiameterStart
       
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
        
        Dim dPositionRatio As Double
        dPositionRatio = posParam / (endParam - startParam)
        
        pdWidth = dDiameterStart + (dDiameterDiff * dPositionRatio)
    Else
        pdWidth = GetMaxDiameter(dDiameterStart, dDiameterEnd)
    End If
    pdDepth = pdWidth
   
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
    
    If bIsModifiable = False Then
        Set pAttrColl = CollAllDisplayedValues
        For i = 1 To pAttrColl.Count
            Set pAttrDescr = pAttrColl.Item(i)
            If pAttrDescr.InterfaceName = IID_IUABuiltUpCone Then
                If pAttrDescr.AttrName = "DiameterStart" Or pAttrDescr.AttrName = "DiameterEnd" Then
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

Private Sub GetConeExtrusionParameters(ByVal oSmartOcc As IJSmartOccurrence, _
                                       ByRef dDiameterStart As Double, _
                                       ByRef dDiameterEnd As Double, _
                                       ByRef dLength As Double)

Const METHOD = "GetConeExtrusionParameters"

On Error GoTo ErrorHandler
    
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    
    ' Get Parameters of BU
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpCone)
    If Not oAttrCol Is Nothing Then
        dDiameterStart = oAttrCol.Item("DiameterStart").Value
        dDiameterEnd = oAttrCol.Item("DiameterEnd").Value
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
Private Function GetMaxDiameter(ByVal dDiameterStart As Double, _
                                ByVal dDiameterEnd As Double) As Double
Const METHOD = "GetMaxDiameter"
On Error GoTo ErrorHandler

    Dim dDiameterMax As Double
    dDiameterMax = 0#
    
    If dDiameterStart > dDiameterEnd Then
        dDiameterMax = dDiameterStart
    Else
        dDiameterMax = dDiameterEnd
    End If
    
    GetMaxDiameter = dDiameterMax
        
Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function
Public Function AreOccurrencePropertiesValid(ByVal oSmartOcc As IJSmartOccurrence) As Boolean
Const METHOD = "AreOccurrencePropertiesValid"

    On Error GoTo ErrorHandler
    AreOccurrencePropertiesValid = False
    
    Dim dDiameterStart As Double
    Dim dDiameterEnd As Double
    Dim dLength As Double
    Dim sError As String
    
    GetConeExtrusionParameters oSmartOcc, dDiameterStart, dDiameterEnd, dLength
    
    sError = BuiltUpDefValidate("DiameterStart", dDiameterStart, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("DiameterEnd", dDiameterEnd, 0#)
    
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
            
    Dim dDiameterStart As Double
    Dim dDiameterEnd As Double
    Dim dLength As Double
    Dim dConeThickness As Double
    
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
       
    GetConeExtrusionParameters oSmartOcc, dDiameterStart, dDiameterEnd, dLength

    Dim dDiameterMax As Double
    dDiameterMax = GetMaxDiameter(dDiameterStart, dDiameterEnd)
    
    Set oSmartItem = oSmartOcc.ItemObject
    Set oAttr = oSmartItem
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpCone)
    If Not oAttrCol Is Nothing Then
        dConeThickness = oAttrCol.Item("ConeThickness").Value
    Else
        GoTo ErrorHandler
    End If
    
    Dim ErrorStatus As ErrorSectionStatus
    ErrorStatus = GetSectionPropertiesFromCalculator(dArea, dPerimeter, dCentroidX, dCentroidY, _
                                                     dXp, dYp, dIxx, dIyy, dRo, dRxx, dRxy, dRyy, _
                                                     dSw, dSxx, dSyy, dZxx, dZyy, dJ, dCw, dH, dUnitWt, _
                                                     dDiameterMax, dConeThickness)
                                                     
    If ErrorStatus <> SectionProperties_OK Then
        GoTo ErrorHandler
    End If

    Set oAttr = Nothing
    Set oAttrCol = Nothing
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IStructCrossSectionDimensions)
    If Not oAttrCol Is Nothing Then
        oAttrCol.Item("Width").Value = dDiameterMax
        oAttrCol.Item("Depth").Value = dDiameterMax
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
        Case "DiameterStart", "DiameterEnd", "Length"
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
                                                    ByVal dDiameterMax As Double, _
                                                    ByVal dConeThickness As Double) As ErrorSectionStatus
                                               
    
Const METHOD = "GetSectionPropertiesFromCalculator"
On Error GoTo ErrorHandler
    
    Dim ErrorStatus As ErrorSectionStatus
    Dim bIsSetOk As Boolean
      
    m_oCalcXProps.SectionTypeAlias = EnumSectionTypeAlias.Section_Circular
    m_oCalcXProps.depth = dDiameterMax
    m_oCalcXProps.WebThickness = dConeThickness

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
            Case "IUABuiltUpCone"
                Select Case sAttrName
                    Case "DiameterStart"
                        IsDesignParameter = True
                    Case "DiameterEnd"
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

Private Sub ICustomSectionShapeService_GetCrossSectionData(ByVal pProfileObject As Object, ByVal distFromStart As Double, ByVal eRepresentationType As SP3DStructInterfaces.structShapeRepresentationEnum, ByVal bBreakIntoComponents As Boolean, sectionPrev As SP3DStructInterfaces.IStructSectionShape, sectionNext As SP3DStructInterfaces.IStructSectionShape)
Const METHOD = "IStructSectionShapeService_GetCrossSectionData"
On Error GoTo ErrorHandler
    
    ' create a sectionProfileRoundHollow shape. If a graphic shape is needed, then call the convert utility method

    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pProfileObject
    Dim dDiameterStart As Double
    Dim dDiameterEnd As Double
    Dim dLength As Double
    Dim sError As String
    
    GetConeExtrusionParameters oSmartOcc, dDiameterStart, dDiameterEnd, dLength
    
    ' the only parameter that should vary along the length of this member is its depth
    Dim depth As Double
    If distFromStart < dTol Then
        ' user wants the definition at the start
        depth = dDiameterStart
    ElseIf dLength - distFromStart < dTol Then
        ' also handles the case where distFromStart is greater than the length
        depth = dDiameterEnd
    Else
        ' somewhere along the length of the member
        'BUITaperedWeb is a linear transition
        ' get the difference from start to end
        Dim dDepthDiff As Double
        dDepthDiff = dDiameterEnd - dDiameterStart
        depth = dDiameterStart + ((distFromStart / dLength) * dDepthDiff)
    End If
    

    Dim sectionMaterial As IStructSectionMaterial
    Dim tubeThickness As Double
    Dim tubeMaterial As String
    Dim tubeGrade As String
    
    GetComponentMaterial pProfileObject, IID_IUABuiltUpCone, tubeMaterial, tubeGrade, tubeThickness
   
    ' always create a parametric shape.  It may be used to create the other shapes as needed
    'Dim parametricShape As IStructSectionRoundHollowShape
    Dim parametricShape As IStructSectionShape
    Set parametricShape = New TransientClassesLib.StructSectionRoundHollowShape
    
    Dim roundShape As IStructSectionRoundHollowShape
    Set roundShape = parametricShape
    roundShape.SetProperties depth, tubeThickness
    Set sectionMaterial = roundShape
    sectionMaterial.SetMaterial tubeMaterial, tubeGrade
    
    ' A cone section cannot be broken down into components
    If eRepresentationType = shape_by_graphic Then

        Dim outerBoundary As IJComplexString
        Dim innerBoundaries As IJElements
        parametricShape.GetShapeAsComplexString outerBoundary, innerBoundaries
        
        ' the conversion routine creates a 2D complex string in the member's LCS.
        ' we need to transform it to the member.
        Dim oMatrix As IJDT4x4
        Dim dDiameterMax As Double
        dDiameterMax = GetMaxDiameter(dDiameterStart, dDiameterEnd)
  
        m_oDesignMemberHelper.CreateTransform oSmartOcc, dDiameterMax, dDiameterMax, oMatrix
        outerBoundary.Transform oMatrix
        
        Dim innerBoundary As IJComplexString
        For Each innerBoundary In innerBoundaries
            innerBoundary.Transform oMatrix
        Next
        
        ' now create a graphic shape to return to the caller
        Set parametricShape = Nothing
        Set parametricShape = New TransientClassesLib.StructSectionCurveBoundedShape
        Dim curvedShape As IStructSectionCurveBoundedShape
        Set curvedShape = parametricShape
        
        curvedShape.SetBoundary outerBoundary, innerBoundaries
        ' need to set the material too since we've replaced the original shape
        Set sectionMaterial = curvedShape
        sectionMaterial.SetMaterial tubeMaterial, tubeGrade
    End If

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

