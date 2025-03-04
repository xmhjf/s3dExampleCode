VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BUEndCan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************
' Copyright (C) 2006, Intergraph Corporation. All rights reserved.
'
'File
'    BUEndCan.cls
'
'Author
'       4 March 2008  ras
'
'Description
'       End Can definition
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

Private Const MODULE = "BUEndCan"
Private Const strSourceFile = "BUEndCan.def"
Private Const CONST_ItemProgId As String = "SM3DBUEndCan.BUEndCan"

Private m_strErrorDescr As String
Private m_oLocalizer As IJLocalizer
Private m_oDesignMemberHelper As BUHelperUtils
Private m_oCalcXProps As BUCalcSectionProperties

Implements IJDUserSymbolServices
Implements IJUserAttributeMgmt
Implements ISPSDesignedMemberHelper
Implements ICustomSectionShapeService

Private Enum BUTubeMembers
  Tube = 3
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
     pAD.UserTypeClsid = "{6C8F4748-79AA-4c82-B7C0-6A48965C11A5}"   '**CUEndCan (Generated here for each new DM)
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
     pAPDs.AddProperty "IUABuiltUpTube", 6, IID_IUABuiltUpTube, , imsCOOKIE_ID_USS_LIB
     pAPDs.AddProperty "IUABuiltUpEndCan", 7, IID_IUABuiltUpEndCan, , imsCOOKIE_ID_USS_LIB
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
'' EndCan Curve To Revolve
'''''
      
     Set pMemberDescription = pMemberDescriptions.AddMember("EndCanCurveToRevolve", 1, "CMConstructCurveToRevolve", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "EndCanCurveToRevolveProperties", 1, IJCurve, "CMComputeCurveToRevolve", imsCOOKIE_ID_USS_LIB

'''''
'' EndCan's Axis To Revolve Around
'''''

     Set pMemberDescription = pMemberDescriptions.AddMember("EndCanAxisOfRevolution", 2, "CMConstructAxisOfRevolution", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "EndCanAxisOfRevolutionProperties", 1, IJCurve, "CMComputeAxisOfRevolution", imsCOOKIE_ID_USS_LIB

'''''
'' EndCan Plate
'''''
     Set pMemberDescription = pMemberDescriptions.AddMember("EndCan", 3, "CMConstructEndCanPlateSystem", imsCOOKIE_ID_USS_LIB)
     pMemberDescription.SetCMFinalConstruct imsCOOKIE_ID_USS_LIB, "CMFinalConstructEndCanPlateSystem"
     pMemberDescription.RelationshipClsid = AssemblyMembers1RelationshipCLSID
     pMemberDescription.SetCMMigrate imsCOOKIE_ID_USS_LIB, "CMMigrateEndCanPlateSystem"
          
     Set pPropertyDescriptions = pMemberDescription
     pPropertyDescriptions.AddProperty "EndCanProperties", 1, IJCurve, "CMComputeEndCanPlateSystem", imsCOOKIE_ID_USS_LIB
             
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
'                       EndCan
'-------------------------------------------------------

Public Sub CMConstructCurveToRevolve(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructCurveToRevolve "
On Error GoTo ErrorHandler

    ' Just Create a initial dummy Curve with points
    Dim pCurve As IJCurve
    m_oDesignMemberHelper.CreatePositionLineStringCurveBy2PointsWithNPoints pResourceManager, 0, -1, 0, _
                                                                            1, -1, 0, _
                                                                            4, pCurve
    CopyPermissionGroup pCurve, pMemberDescription.CAO
    
    Dim oLineString3d As LineString3d
    Set oLineString3d = pCurve
    oLineString3d.SetPoint 1, 0, -1, 0
    oLineString3d.SetPoint 2, 0.5, -1, 0
    oLineString3d.SetPoint 3, 1, -0.5, 0
    oLineString3d.SetPoint 4, 1.5, -0.5, 0
    
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
    Dim dDiameterEnd As Double
    Dim dLength As Double
    Dim dConeLength As Double
    Dim IsTransitionAtStart As Long
    Dim dTubeDiameter As Double

    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO

    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If

    GetEndCanExtrusionParameters oSmartOcc, dDiameterEnd, dConeLength, _
                                 IsTransitionAtStart, dTubeDiameter, dLength

    Dim dDiameterMax As Double
    dDiameterMax = GetMaxDiameter(dTubeDiameter, dDiameterEnd)

    m_oDesignMemberHelper.CreateTransform oSmartOcc, dDiameterMax, dDiameterMax, oMatrix

    Dim oLineString3d As LineString3d
    Set oLineString3d = pObject

    Dim dLengthExt As Double
    m_oDesignMemberHelper.GetLengthExtension oSmartOcc, dLengthExt

    Dim oPt1 As IJDPosition
    Set oPt1 = New DPosition
    
    Dim oPt2 As IJDPosition
    Set oPt2 = New DPosition
    
    Dim oPt3 As IJDPosition
    Set oPt3 = New DPosition
    
    Dim oUnext3 As IJDPosition
    Set oUnext3 = New DPosition
    
    Dim oPt4 As IJDPosition
    Set oPt4 = New DPosition
    
    Dim oConeDir As IJDVector
    Dim oAxisDir As IJDVector

    If IsTransitionAtStart = 0 Then ' tube at start
        oPt1.Set 0 - dLengthExt, -(dDiameterMax - dTubeDiameter) / 2, dDiameterMax / 2
        oPt2.Set dLength - dConeLength, -(dDiameterMax - dTubeDiameter) / 2, dDiameterMax / 2
        oUnext3.Set dLength, -(dDiameterMax - dDiameterEnd) / 2, dDiameterMax / 2
        
        Set oConeDir = oUnext3.Subtract(oPt2)
        oConeDir.Length = 0.005
        Set oPt3 = oUnext3.Offset(oConeDir)

        Set oAxisDir = oPt2.Subtract(oPt1)
        oAxisDir.Length = dLengthExt
        Set oPt4 = oPt3.Offset(oAxisDir)
        
        With oLineString3d
            .SetPoint 1, oPt1.x, oPt1.y, oPt1.z
            .SetPoint 2, oPt2.x, oPt2.y, oPt2.z
            .SetPoint 3, oPt3.x, oPt3.y, oPt3.z
            .SetPoint 4, oPt4.x, oPt4.y, oPt4.z
        End With

    Else ' transition is the start
        oPt1.Set 0, -(dDiameterMax - dDiameterEnd) / 2, dDiameterMax / 2
        oPt2.Set dConeLength, -(dDiameterMax - dTubeDiameter) / 2, dDiameterMax / 2
        oPt3.Set dLength + dLengthExt, -(dDiameterMax - dTubeDiameter) / 2, dDiameterMax / 2
        
        Set oConeDir = oPt1.Subtract(oPt2)
        oConeDir.Length = 0.005
        Set oPt1 = oPt1.Offset(oConeDir)

        Set oAxisDir = oPt2.Subtract(oPt3)
        oAxisDir.Length = dLengthExt
        Set oPt4 = oPt1.Offset(oAxisDir)
        
        With oLineString3d
            .SetPoint 1, oPt4.x, oPt4.y, oPt1.z
            .SetPoint 2, oPt1.x, oPt1.y, oPt2.z
            .SetPoint 3, oPt2.x, oPt2.y, oPt3.z
            .SetPoint 4, oPt3.x, oPt3.y, oPt4.z
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
    Dim dDiameterEnd As Double
    Dim dLength As Double
    Dim dConeLength As Double
    Dim IsTransitionAtStart As Long
    Dim dTubeDiameter As Double
    
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pPropertyDescriptions.CAO
    
    'validate attributes
    If Not AreOccurrencePropertiesValid(oSmartOcc) Then
        GoTo ErrorHandler
    End If
    
    GetEndCanExtrusionParameters oSmartOcc, dDiameterEnd, dConeLength, _
                                 IsTransitionAtStart, dTubeDiameter, dLength

    Dim dDiameterMax As Double
    dDiameterMax = GetMaxDiameter(dTubeDiameter, dDiameterEnd)
        
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

Public Sub CMConstructEndCanPlateSystem(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObj As Object)
Const METHOD = "CMConstructEndCanPlateSystem"
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
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTube)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("TubeThickness").Value
        strMaterial = oAttrCol.Item("TubeMaterial").Value
        strGrade = oAttrCol.Item("TubeGrade").Value
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
                                                                , , , InDir, MoldedFormPlateNameCat.NameCatCan)
     
    m_oDesignMemberHelper.SetPlateBoundaries oSmartOcc, pObj
    
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub

Public Sub CMFinalConstructEndCanPlateSystem(pMemberDesc As IJDMemberDescription)
Const METHOD = "CMFinalConstructEndCanPlateSystem"
On Error GoTo ErrorHandler
 
Exit Sub
ErrorHandler:  HandleError MODULE, METHOD
End Sub
Public Sub CMComputeEndCanPlateSystem(pPropertyDescriptions As IJDPropertyDescription, pObject As Object)
Const METHOD = "CMComputeEndCanPlateSystem"
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
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTube)
    If Not oAttrCol Is Nothing Then
        dThickness = oAttrCol.Item("TubeThickness").Value
        strMaterial = oAttrCol.Item("TubeMaterial").Value
        strGrade = oAttrCol.Item("TubeGrade").Value
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
Public Sub CMMigrateEndCanPlateSystem(pMemberDesc As IJDMemberDescription, pMigrateHelper As IJMigrateHelper)
Const METHOD = "CMMigrateEndCanPlateSystem "
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
    
    Dim dDiameterEnd As Double
    Dim dAxisLength As Double
    Dim dConeLength As Double

    Dim IsTransitionAtStart As Long
    Dim dDiameterTube As Double
    Dim dLengthToPosition As Double

   
    Dim axisCurve As IJCurve
    

    Dim oSmartOcc As IJSmartOccurrence
    
    Set oSmartOcc = pIDesignedMember
    
    GetEndCanExtrusionParameters oSmartOcc, dDiameterEnd, dConeLength, _
                                 IsTransitionAtStart, dDiameterTube, dAxisLength

    'measure diameter at the position
    If Not pPosAlong Is Nothing Then
    Dim pISPSMemberPartCommon As ISPSMemberPartCommon
        Set pISPSMemberPartCommon = pIDesignedMember
        Set axisCurve = pISPSMemberPartCommon.Axis
        Dim posParam As Double
        axisCurve.Parameter pPosAlong.x, pPosAlong.y, pPosAlong.z, posParam
 '
'         ____________________________
'        /
'       /
'      /
'     /
'
'    A-a-B----------------------------C
'
'     \
'      \
'       \
'        \____________________________
        
' A->B =  Cone Length
' B->C = Tube Length
' A->C = Overall Length
' the ratio for point 'a' (on the cone) is A->a/A->B
     
        Dim startParam As Double
        Dim endParam As Double
        axisCurve.ParamRange startParam, endParam
        If posParam < startParam Then
            posParam = startParam
        ElseIf posParam > endParam Then
            posParam = endParam
        End If
        dLengthToPosition = posParam - startParam      'for linear axis, param = distance
    
        If Not IsTransitionAtStart Then   ' measure from the end to match where the cone is
            dLengthToPosition = dAxisLength - dLengthToPosition
        End If
        
        If dAxisLength > 0.001 And dConeLength > dAxisLength Then        ' there is no tube, adjust "tube" diameter
            dDiameterTube = dDiameterEnd + (dAxisLength / dConeLength) * (dDiameterTube - dDiameterEnd)
        End If
        
        If dConeLength > 0.001 Then
            If dLengthToPosition < dConeLength Then     ' on the cone
                pdWidth = dDiameterEnd + (dLengthToPosition / dConeLength) * (dDiameterTube - dDiameterEnd)
            Else
                pdWidth = dDiameterTube
            End If
        Else
            pdWidth = dDiameterTube
        End If

    Else    ' no position given.   return the max diameter
        pdWidth = GetMaxDiameter(dDiameterTube, dDiameterEnd)
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
      
    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pIJDAttrs
    Dim dLength As Double
    
    dLength = GetLength(oSmartOcc)
    
    ' loop on the attributes on the interface to match the supplied attribute type
    For AttrCount = 1 To oAttrCol.Count
        Set oAttrObj = oAttrCol.Item(AttrCount)
        If oAttrObj.Name = pAttrToChange.AttrName Then
            Select Case oAttrObj.Type
                Case m_oDesignMemberHelper.DoubleValue
                    If pAttrToChange.AttrName = "ConeLength" And varNewAttrValue > 0 Then 'conelength cannot be zero
                        ErrStr = BuiltUpDefValidate(pAttrToChange.AttrName, varNewAttrValue, 0#, dLength)
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
    
    If bIsModifiable = False Then
        Set pAttrColl = CollAllDisplayedValues
        For i = 1 To pAttrColl.Count
            Set pAttrDescr = pAttrColl.Item(i)
            If pAttrDescr.InterfaceName = IID_IUABuiltUpEndCan Then
                If pAttrDescr.AttrName = "DiameterCone" Then
                    pAttrDescr.AttrState = AttributeDescriptor_ReadOnly
                End If
            ElseIf pAttrDescr.InterfaceName = IID_IUABuiltUpTube Then
                If pAttrDescr.AttrName = "TubeDiameter" Then
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

Private Sub GetEndCanExtrusionParameters(ByVal oSmartOcc As IJSmartOccurrence, _
                                         ByRef dDiameterEnd As Double, _
                                         ByRef dConeLength As Double, _
                                         ByRef IsTransitionAtStart As Long, _
                                         ByRef dTubeDiameter As Double, _
                                         ByRef dLength As Double)

Const METHOD = "GetEndCanExtrusionParameters"

On Error GoTo ErrorHandler
    
    Dim oAttrCol As IJDAttributesCol
    Dim oAttr As IJDAttributes
    
    ' Get Parameters of Tube/can
    Set oAttr = oSmartOcc
    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpTube)
    If Not oAttrCol Is Nothing Then
        dTubeDiameter = oAttrCol.Item("TubeDiameter").Value
    Else
        GoTo ErrorHandler
    End If

    Set oAttrCol = oAttr.CollectionOfAttributes(IID_IUABuiltUpEndCan)
    If Not oAttrCol Is Nothing Then
        dDiameterEnd = oAttrCol.Item("DiameterCone").Value
        dConeLength = oAttrCol.Item("ConeLength").Value
        IsTransitionAtStart = oAttrCol.Item("IsTransitionAtStart").Value
    Else
        GoTo ErrorHandler
    End If
     
    Dim pISPSMemberPartCommon As ISPSMemberPartCommon
    Set pISPSMemberPartCommon = oSmartOcc
    dLength = pISPSMemberPartCommon.Axis.Length
     
    ' if we have curves then change length when given a negative number
    Dim pCurveToRevolve As IJCurve
    Dim pAxisCurve As IJCurve
    Dim pIJDMemberObject As IJDMemberObjects
    Set pIJDMemberObject = oSmartOcc
    Set pCurveToRevolve = pIJDMemberObject.ItemByDispid(1)
    Set pAxisCurve = pIJDMemberObject.ItemByDispid(2)
    If pCurveToRevolve Is Nothing Or pAxisCurve Is Nothing Then
        Exit Sub
    End If
 
    ' do we need to initialize StartLength, EndLength parameters ?
    ' the answer is yes if either is < 0, or if newly created and sum of Start + End > AxisLength

    Dim iFlags As Long
    Const RELATION_INSERTED = &H100000
    Dim structAssocCompute As IJStructAssocCompute
    Dim bResetLength As Boolean
    bResetLength = False

    If dConeLength < 0 Then                            ' initial placement
        bResetLength = True
    End If
    
    If Not bResetLength Then
        
        Set structAssocCompute = New StructAssocTools
        structAssocCompute.GetAssocFlags oSmartOcc, iFlags
        
        If (iFlags And RELATION_INSERTED) = RELATION_INSERTED Then          ' inserted this compute
    
            If dConeLength > dLength Then                     ' axis smaller than cone length
                bResetLength = True
            End If
        End If
    End If

    If bResetLength Then
    
        dConeLength = 0.5 * dLength
        oAttrCol.Item("ConeLength").Value = dConeLength
        
    End If

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Private Function GetLength(ByVal oSmartOcc As IJSmartOccurrence) As Double
Const METHOD = "GetLength"
On Error GoTo ErrorHandler
    
    Dim dLength As Double
    
    Dim pISPSMemberPartCommon As ISPSMemberPartCommon
    Set pISPSMemberPartCommon = oSmartOcc
    dLength = pISPSMemberPartCommon.Axis.Length
    
    GetLength = dLength

Exit Function
ErrorHandler: HandleError MODULE, METHOD
End Function
Private Function GetMaxDiameter(ByVal dTubeDiameter As Double, _
                                ByVal dDiameterEnd As Double) As Double
Const METHOD = "GetMaxDiameter"
On Error GoTo ErrorHandler

    Dim dDiameterMax As Double
    dDiameterMax = 0#
    
    If dTubeDiameter > dDiameterEnd Then
        dDiameterMax = dTubeDiameter
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
    
    Dim dDiameterEnd As Double
    Dim dLength As Double
    Dim dConeLength As Double
    Dim IsTransitionAtStart As Long
    Dim dTubeDiameter As Double
    Dim sError As String
    
    GetEndCanExtrusionParameters oSmartOcc, dDiameterEnd, dConeLength, _
                                 IsTransitionAtStart, dTubeDiameter, dLength
        
    sError = BuiltUpDefValidate("DiameterCone", dDiameterEnd, 0#)
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("ConeLength", dConeLength, 0#) 'cannot be zero
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("ConeLength", dConeLength, 0#, dLength) 'cannot be >= Length
    
    If Len(sError) > 0 Then
        AreOccurrencePropertiesValid = False
        Exit Function
    End If
    
    sError = BuiltUpDefValidate("TubeDiameter", dTubeDiameter, 0#)
    
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
            
    Dim dDiameterEnd As Double
    Dim dLength As Double
    Dim dConeLength As Double
    Dim IsTransitionAtStart As Long
    Dim dTubeDiameter As Double
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
    
    GetEndCanExtrusionParameters oSmartOcc, dDiameterEnd, dConeLength, _
                                 IsTransitionAtStart, dTubeDiameter, dLength

    Dim dDiameterMax As Double
    dDiameterMax = GetMaxDiameter(dTubeDiameter, dDiameterEnd)
    
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
                                                     dDiameterMax, dTubeThickness)
                                                     
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
        Case "DiameterCone", "ConeLength", "TubeDiameter", "Length"
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
                                                    ByVal dTubeThickness As Double) As ErrorSectionStatus
                                               
    
Const METHOD = "GetSectionPropertiesFromCalculator"
On Error GoTo ErrorHandler
    
    Dim ErrorStatus As ErrorSectionStatus
    Dim bIsSetOk As Boolean
      
    m_oCalcXProps.SectionTypeAlias = EnumSectionTypeAlias.Section_Circular
    m_oCalcXProps.depth = dDiameterMax
    m_oCalcXProps.WebThickness = dTubeThickness

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
            Case "IUABuiltUpTube"
                If sAttrName = "TubeDiameter" Then
                    IsDesignParameter = True
                End If
            Case "IUABuiltUpEndCan"
            'actualy don't neeed this logic below. All attributes on this I/F are design attrs but
            ' its an extra check in case user decides to add a non-design attribute on the same I/F
                Select Case sAttrName
                    Case "ConeLength"
                        IsDesignParameter = True
                    Case "DiameterCone"
                        IsDesignParameter = True
                    Case "IsTransitionAtStart"
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
    
    ' we're only creating a section_profile_i_type for now.  There is no use case for creating a compound shape
    ' or a graphic shape yet and no way for the user to ask for something different.
    
    ' in the future, utility methods will be implemented to convert to a graphic shape and to decompose the shape
    ' if there is a need to do this

   
    ' the only parameter that varys along the length of this member is its depth(diameter)
    ' thickness of the tube and cone (for this CAD) are always the same
    
    Dim tubeMaterial As String
    Dim tubeGrade As String
    Dim tubeThickness As Double
    GetComponentMaterial pProfileObject, IID_IUABuiltUpTube, tubeMaterial, tubeGrade, tubeThickness

    
    ' get an IJDPosition at the requested distance along the profile's curve
    Dim pISPSMemberPartCommon As ISPSMemberPartCommon
    Set pISPSMemberPartCommon = pProfileObject
    
    Dim myAxisCurve As IJCurve
    
    Set myAxisCurve = pISPSMemberPartCommon.Axis
    
    Dim pX As Double
    Dim pY As Double
    Dim pZ As Double
    Dim axisLength As Double
    axisLength = myAxisCurve.Length
    
    myAxisCurve.PositionFRatio distFromStart / axisLength, pX, pY, pZ
    
    Dim pointOnCurve As IJDPosition
    Set pointOnCurve = New DPosition
    pointOnCurve.Set pX, pY, pZ
    
    Dim meAsDM As ISPSDesignedMember
    Set meAsDM = pProfileObject
    
    Dim depth As Double
    Dim width As Double
    
    ISPSDesignedMemberHelper_GetNominalSectionSize meAsDM, pointOnCurve, width, depth
    
    
    ' always create a parametric shape.  It may be used to create the other shapes as needed
    Dim parametricShape As IStructSectionRoundHollowShape
    Set parametricShape = New TransientClassesLib.StructSectionRoundHollowShape

    parametricShape.SetProperties depth, tubeThickness
    
    
    
    Dim sectionMaterial As IStructSectionMaterial
    Set sectionMaterial = parametricShape
    
    sectionMaterial.SetMaterial tubeMaterial, tubeGrade
    
    If distFromStart < dTol Then
        ' user wants the definition at the start
        Set sectionNext = parametricShape
        Set sectionPrev = Nothing
    ElseIf axisLength - distFromStart < dTol Then
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
    ' even if the user has edited the parameters (setting the Cone diameter to be equal to the tube diameter)
    ' essentially turning a BUEndCan into a BUTube, we want to preserve the "design intent" so will always
    ' have three transition locations.

    Dim dDiameterEnd As Double
    Dim dLength As Double
    Dim dConeLength As Double
    Dim IsTransitionAtStart As Long
    Dim dTubeDiameter As Double

    Dim oSmartOcc As IJSmartOccurrence
    Set oSmartOcc = pProfileObject

    GetEndCanExtrusionParameters oSmartOcc, dDiameterEnd, dConeLength, _
                             IsTransitionAtStart, dTubeDiameter, dLength

    ReDim distFromStart(0 To 2)
    distFromStart(0) = 0#
    If IsTransitionAtStart = 0 Then ' tube at end
        distFromStart(1) = dLength - dConeLength
    Else
        distFromStart(1) = dConeLength
    End If
       
    distFromStart(2) = dLength
       
       
Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub
