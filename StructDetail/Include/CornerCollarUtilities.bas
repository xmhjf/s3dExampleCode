Attribute VB_Name = "CornerCollarUtilities"
'*************************************************************************************************************************************
'  Copyright (C) 2011-13, Intergraph Corporation.  All rights reserved.
'
'  File        : CornerCollarUtilities.bas
'
'  Description : Common Methods for Corner Collar
'
'  Author      :
'
'  History     :
'   31/May/2013  -  svsmylav/vbbheema    -    CR-232822 'CreatePhysicalConns' method is updated
'                                    to create PCs at Support1/2. Collar thickness calculated
'                                    based on web/flange and available plate thickness.
'
'   27/July/2015 - pkakula: TR-243401: updated some methods in order to reference middle-tier dll's instead of client-tier dll's
'
'*************************************************************************************************************************************
Option Explicit

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\SMCornerFeatRules\CornerCollarUtilities.bas"
Private Const E_INVALIDARG = &H80070057
Public Const INPUT_CORNERFEATURE = "CornerFeature"
Public Const INPUT_CORNERCUTOUTPORT = "CornerCoutoutPort"

'

'
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function CreateCollarAtCorner(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown) As Object
Const METHOD = "CreateCollarAtCorner"
    On Error GoTo ErrorHandler
    
    Dim iCnt As Long
    Dim nCnt As Long
    Dim sMsg As String
    Dim dThickness As Double
    Dim oFacePort As Object
    Dim oEdgePort1 As Object
    Dim oEdgePort2 As Object
    Dim oCornerCollar As Object
    Dim oSmartOccurrence As IJSmartOccurrence
    
    Dim oInputCollection As JCmnShp_CollectionAlias
    
    Dim oSDOCollar As StructDetailObjectsex.IJSDOCollar
    Dim oStructFeatUtils As IJSDFeatureAttributes
    Dim oDSmartPlateAttributes As IJSDSmartPlateAttributes
    Dim oSDSmartPlateDefinition As IJSDSmartPlateDefinition
    
    Set oSmartOccurrence = pMemberDescription.CAO
    
    '   Smart Class:    SmartGeneralCornerCollar
    '   Smart Item:     CollarAtLongScallop
    Set oSDOCollar = New StructDetailObjectsex.SDOCollar
    oSDOCollar.Create oSmartOccurrence, "RootCornerCollar"
    
    Set oCornerCollar = oSDOCollar.object
    Set CreateCollarAtCorner = oCornerCollar
    
    Dim oPartWithCorner As Object
    Set oPartWithCorner = oSDOCollar.Penetrated
    If Not TypeOf oPartWithCorner Is IJPlate Then
        ' If the Corner Feature is NOT on a Plate part
        ' Or if the Material, Grade, Thickness is to be different
        ' Then need to (Initialize) set the Collar Plate properties
        SetPlatePartProperties oCornerCollar, pResourceManager, _
                              "", "", _
                              NonTight, Standalone, _
                              "Steel - Carbon", "A", 0.025
    End If
    
If False Then
    ' Kludge ... current using Corner Feature USS (requires 3 Inputs)
    ' ... add the additional inputs
    Set oStructFeatUtils = New SDFeatureUtils
    Set oDSmartPlateAttributes = New SDSmartPlateUtils
    Set oSDSmartPlateDefinition = oDSmartPlateAttributes
    oDSmartPlateAttributes.GetInputs_SmartCollar oCornerCollar, oInputCollection
    iCnt = oInputCollection.Count
    
    oStructFeatUtils.get_CornerCutInputsEx oSmartOccurrence, oFacePort, oEdgePort1, oEdgePort2
    oInputCollection.Add oEdgePort1
    oInputCollection.Add oEdgePort2
    nCnt = oInputCollection.Count
    
    Set oCornerCollar = oSDSmartPlateDefinition.CreateSmartCollarPart(pResourceManager, "", oInputCollection, Nothing, oCornerCollar)
End If
    
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub GetCollarThicknessFromCornerFeature(oCornerObject As Object, dThickness As Double)
Const METHOD = "GetCollarThicknessFromCornerFeature"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim lXid1 As Long
    Dim lXid2 As Long
    Dim bOnWeb As Boolean
    
    Dim oPort As IJPort
    Dim oCornerPart As Object
    
    Dim oSDO_PlatePart As StructDetailObjects.PlatePart
    Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
    Dim oSDO_MemberPart As StructDetailObjectsex.SDOMemberPart
    Dim oSDO_CornerFeature As StructDetailObjects.CornerFeature
    
    If TypeOf oCornerObject Is IJPort Then
        Set oPort = oCornerObject
        Set oCornerPart = oPort.Connectable
        
    ElseIf TypeOf oCornerObject Is IJStructFeature Then
        Set oSDO_CornerFeature = New StructDetailObjects.CornerFeature
        
        Dim oCF As New StructDetailObjectsex.CornerFeature
        
        Set oCF.object = oCornerObject
        Set oCornerPart = oCF.GetPartObject
    End If
    
    bOnWeb = False
    
    Dim oStructFeatUtils As IJSDFeatureAttributes
    Set oStructFeatUtils = New SDFeatureUtils
    
    Dim oFacePort As Object
    Dim oEdgePort1 As Object
    Dim oEdgePort2 As Object
    
    oStructFeatUtils.get_CornerCutInputsEx oCornerObject, oFacePort, oEdgePort1, oEdgePort2
    
    Dim oStructPort As IJStructPort
    Set oStructPort = oFacePort
    
    If oStructPort.operatorID = JXSEC_WEB_LEFT Or oStructPort.operatorID = JXSEC_WEB_RIGHT Then
        bOnWeb = True
    End If
    
    dThickness = 0#
    If TypeOf oCornerPart Is IJPlate Then
        Set oSDO_PlatePart = New StructDetailObjects.PlatePart
        Set oSDO_PlatePart.object = oCornerPart
        dThickness = oSDO_PlatePart.PlateThickness
        
    ElseIf TypeOf oCornerPart Is IJProfile Then
        Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
        Set oSDO_ProfilePart.object = oCornerPart
        If bOnWeb Then
            dThickness = oSDO_ProfilePart.WebThickness
        Else
            dThickness = oSDO_ProfilePart.flangeThickness
        End If
    ElseIf TypeOf oCornerPart Is ISPSMemberPartCommon Then
        Set oSDO_MemberPart = New StructDetailObjectsex.SDOMemberPart
        Set oSDO_MemberPart.object = oCornerPart
        If bOnWeb Then
            dThickness = oSDO_MemberPart.WebThickness
        Else
            dThickness = oSDO_MemberPart.flangeThickness
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub SetCollarMatlAndGrade(oPlate As IJPlate, sMaterial As String, sGrade As String)
Const METHOD = "SetCollarMatlAndGrade"
On Error GoTo ErrorHandler
    
    Dim oMaterial As IJDMaterial    
    Dim oRefDataQuery As RefDataMiddleServices.RefdataSOMMiddleServices
    Dim oPlateMaterial As IJStructureMaterial
    
    If oPlate Is Nothing Then
       Err.Raise E_INVALIDARG, MODULE & ":" & METHOD, "Plate not found"
    End If
    
    ' set up defaults
    If Len(Trim(sMaterial)) < 1 Then
      sMaterial = "Steel - Carbon"
    End If
    
    If Len(Trim(sGrade)) < 1 Then
      sGrade = "A"
    End If
        
    
    Set oRefDataQuery = New RefDataMiddleServices.RefdataSOMMiddleServices
    Set oMaterial = oRefDataQuery.GetMaterialByGrade(sMaterial, sGrade)
    Set oRefDataQuery = Nothing
    
    Set oPlateMaterial = oPlate
    oPlateMaterial.Material = oMaterial
    
    Set oPlateMaterial = Nothing
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub SetCollarThickness(oPlate As IJPlate, dThickness As Double)
Const METHOD = "SetCollarThickness"
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim nCount As Long
    
    Dim matlThickCol As RefDataMiddleServices.IJDCollection
        
    Dim oMaterial As IJDMaterial
    Dim oPlateDims As IJDPlateDimensions
    Dim oPlateMaterial As IJStructureMaterial
    Dim oStructServices As IJDStructServices
    Dim oStandAlonePart As IJDStandAlonePlatePart
    Dim catalogResMgr As Object
    
    If oPlate Is Nothing Then
       Err.Raise E_INVALIDARG, MODULE & ":" & METHOD, "Plate not found"
    End If
    
    Set oPlateMaterial = oPlate
    Set oMaterial = oPlateMaterial.Material
    
    ' If a thickness is given, retrieve the corresponding PlateDimension from
    ' the catalog. Otherwise, use the first PlateDimension
    If dThickness < 0.001 Then
       dThickness = 0.004
    End If
    
    Dim catUtils As New CatalogUtil
    Set catalogResMgr = catUtils.GetCatalogResourceManager
    
    ' As part of TR-CP-245512, removing the client reference
    Set oStructServices = New StructServices
    Set matlThickCol = oStructServices.GetPlateDimensions(catalogResMgr, oMaterial.MaterialType, oMaterial.MaterialGrade)
    
    Dim iIndex As Long
    Dim lLowerBound As Long
    Dim lUpperBound As Long
    
    Dim dMaxThickness As Double
    Dim dBestThickness As Double
    Dim oStructQuery2 As IJDStructQuery2
    Dim oPlateSystem As IJPlateSystem
    Dim oPlateSpec As IJDMoldedFormSpec
    Dim sGrade As String
    Dim sMaterial As String
    Dim oPlateDimensions As IJDPlateDimensions
    Dim oSpecCollection As GSCADMoldedFormSpecMgr.IJDCollection
    Dim oStructDetailHelper As GSCADStructDetailUtil.StructDetailHelper
    Dim oMoldedFormSpecMgr As GSCADMoldedFormSpecMgr.JMoldedFormSpecMgr
    Set oStructDetailHelper = New GSCADStructDetailUtil.StructDetailHelper
    
    If TypeOf oPlate Is IJModelBody Then
        oStructDetailHelper.IsPartDerivedFromSystem oPlate, oPlateSystem
    End If
    
    Set oStructDetailHelper = Nothing
    If oPlateSystem Is Nothing Then
        Set oMoldedFormSpecMgr = New GSCADMoldedFormSpecMgr.JMoldedFormSpecMgr
        Set oSpecCollection = oMoldedFormSpecMgr.GetSpecs
        If oSpecCollection.Size > 0 Then
            Set oPlateSpec = oSpecCollection.Item(1)
        End If
        
        Set oSpecCollection = Nothing
        Set oMoldedFormSpecMgr = Nothing
        If oPlateSpec Is Nothing Then
            Exit Sub
        End If
    Else
        Set oPlate = oPlateSystem
        Set oPlateSpec = oPlateSystem.MoldedFormSpec
        If oPlateSpec Is Nothing Then
            Exit Sub
        End If
    End If
    
    dMaxThickness = -1#
    dBestThickness = -1#
    
    For iIndex = 1 To matlThickCol.Size
        Set oPlateDims = matlThickCol.Item(iIndex)
        If oPlateDims.thickness >= dMaxThickness Then
            dMaxThickness = oPlateDims.thickness
        End If
        
        If oPlateDims.thickness >= dThickness Then
            If dBestThickness < 0# Or dBestThickness > oPlateDims.thickness Then
                dBestThickness = oPlateDims.thickness
            End If
        End If
    Next iIndex
    
    sMaterial = oMaterial.MaterialType
    sGrade = oMaterial.MaterialGrade
    Set oStructQuery2 = oPlateSpec
    If dBestThickness > 0# Then
        oStructQuery2.GetPlateDimension sMaterial, sGrade, dBestThickness, _
                                        oPlateDimensions
        oPlate.Dimensions = oPlateDimensions
    End If
    
    Set oStandAlonePart = oPlate
    
    Set matlThickCol = Nothing
    
    Set oMaterial = Nothing
    Set oStructQuery2 = Nothing
    Set oPlateDims = Nothing
    Set oStructServices = Nothing
    Set oPlateMaterial = Nothing
   
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub CreatePhysicalConns(pMemberDescription As IJDMemberDescription, _
                                          pResourceManager As IUnknown, _
                                          bWebEdge As Boolean, _
                                          sSmartItem As String, _
                                          pObject As Object)
Const METHOD = "CreatePhysicalConns"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sErrMsg As String
    
    Dim oEdgePort1 As Object
    Dim oEdgePort2 As Object
    Dim oEdgeFeature As Object
    Dim oInsertPlate As Object
    Dim oInsertInputs As Collection
    
    Dim oSDO_PlatePart As StructDetailObjects.PlatePart
    Dim oSDO_EdgeFeature As StructDetailObjects.EdgeFeature
    Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
    Dim oSDO_PhysicalConn As IJSDOPhysicalConnection
    
    'Physical Connection between Profile Edge Feature Port and Insert Plate Lateral Edge
    sMsg = ">>> GetInputs_InsertPlate"
    Set oInsertPlate = pMemberDescription.CAO
     
    Dim oSDOCollar As StructDetailObjectsex.IJSDOCollar
    Set oSDOCollar = New StructDetailObjectsex.SDOCollar
    
    Set oSDOCollar.object = oInsertPlate
    
    Dim oSDSmartPlateAttributes As IJSDSmartPlateAttributes
    Set oSDSmartPlateAttributes = New SDSmartPlateUtils
    
    Dim oCF As Object
    Set oCF = oSDOCollar.FeatureForCollar

    Dim lDIspid As Double
    Dim oPenetratedObject As Object
    
    Set oPenetratedObject = oSDOCollar.Penetrated
    Set oSDO_PlatePart = New StructDetailObjects.PlatePart
    Set oSDO_PlatePart.object = oInsertPlate
    
    lDIspid = pMemberDescription.Dispid
    
    If TypeOf oPenetratedObject Is IJPlatePart Then
        Dim oPlatePart As New StructDetailObjects.PlatePart
        Set oPlatePart.object = oSDOCollar.Penetrated
    
        If lDIspid = 1 Then
           Set oEdgePort1 = oPlatePart.baseport(BPT_Base)
           oSDSmartPlateAttributes.GetLateBindPort oInsertPlate, JS_TOPOLOGY_FILTER_SOLID_BASE_LFACE, True, oEdgePort2
           sSmartItem = "LapWeld"
        ElseIf lDIspid = 2 Then
           Set oEdgePort1 = oPlatePart.baseport(BPT_Offset)
           oSDSmartPlateAttributes.GetLateBindPort oInsertPlate, JS_TOPOLOGY_FILTER_SOLID_BASE_LFACE, True, oEdgePort2
           sSmartItem = "LapWeld"
        ElseIf lDIspid = 3 Then
            Set oEdgePort1 = oPlatePart.CutoutPort(oCF)
            Set oEdgePort2 = oSDO_PlatePart.baseport(BPT_Lateral)
            sSmartItem = "ButtWeld"
        End If
    ElseIf TypeOf oPenetratedObject Is IJProfilePart Then
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = oSDOCollar.Penetrated
        If lDIspid = 1 Then
           Set oEdgePort1 = oProfilePart.baseport(BPT_Lateral)
           oSDSmartPlateAttributes.GetLateBindPort oInsertPlate, JS_TOPOLOGY_FILTER_SOLID_BASE_LFACE, True, oEdgePort2
           sSmartItem = "LapWeld"
        ElseIf lDIspid = 2 Then
           Set oEdgePort1 = oProfilePart.baseport(BPT_Lateral)
           oSDSmartPlateAttributes.GetLateBindPort oInsertPlate, JS_TOPOLOGY_FILTER_SOLID_BASE_LFACE, True, oEdgePort2
           sSmartItem = "LapWeld"
        ElseIf lDIspid = 3 Then
            Set oEdgePort1 = oProfilePart.CutoutPort(oCF)
            Set oEdgePort2 = oSDO_PlatePart.baseport(BPT_Lateral)
            sSmartItem = "ButtWeld"
        End If
    ElseIf TypeOf oPenetratedObject Is ISPSMemberPartCommon Then
        Dim oMemberPart As New StructDetailObjects.MemberPart
        Set oMemberPart.object = oSDOCollar.Penetrated
        If lDIspid = 1 Then
           Set oEdgePort1 = oMemberPart.baseport(BPT_Lateral)
           oSDSmartPlateAttributes.GetLateBindPort oInsertPlate, JS_TOPOLOGY_FILTER_SOLID_BASE_LFACE, True, oEdgePort2
           sSmartItem = "LapWeld"
        ElseIf lDIspid = 2 Then
           Set oEdgePort1 = oMemberPart.baseport(BPT_Lateral)
           oSDSmartPlateAttributes.GetLateBindPort oInsertPlate, JS_TOPOLOGY_FILTER_SOLID_BASE_LFACE, True, oEdgePort2
           sSmartItem = "LapWeld"
        ElseIf lDIspid = 3 Then
            Dim oMbrPart As New StructDetailObjectsex.SDOMemberPart
            Set oMbrPart.object = oSDOCollar.Penetrated
            Set oEdgePort1 = oMbrPart.CutoutPort(oCF)
            Set oEdgePort2 = oSDO_PlatePart.baseport(BPT_Lateral)
            sSmartItem = "ButtWeld"
        End If
    End If
    
    'Check and create support1/2 regular and split PCs
    If lDIspid = 4 Or lDIspid = 5 Or lDIspid = 6 Or lDIspid = 7 Then
        Dim oSupp1Coll As JCmnShp_CollectionAlias
        Dim oSupp2Coll As JCmnShp_CollectionAlias
        Dim oSmartPltUtils As GSCADSDCreateModifyUtilities.IJSDSmartPlateAttributes
        Dim oS1Port As IJPort
        Dim oS2Port As IJPort
        Dim bComputeSupp1 As Boolean
        Dim bComputeSupp2 As Boolean
        bComputeSupp1 = IIf(lDIspid = 4 Or lDIspid = 6, True, False)
        bComputeSupp2 = IIf(lDIspid = 5 Or lDIspid = 7, True, False)
        Set oSmartPltUtils = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
        oSmartPltUtils.GetCornerCollarOverlappingPorts oInsertPlate, _
                bComputeSupp1, bComputeSupp2, oS1Port, oS2Port, oSupp1Coll, oSupp2Coll

        If lDIspid = 4 Or lDIspid = 6 Then
            Set oEdgePort1 = oS1Port
        Else
            Set oEdgePort1 = oS2Port
        End If
        
        Select Case lDIspid
        Case 4
            Set oEdgePort2 = oSupp1Coll.Item(1)
        Case 5
            Set oEdgePort2 = oSupp2Coll.Item(1)
        Case 6
            Set oEdgePort2 = oSupp1Coll.Item(2)
        Case 7
            Set oEdgePort2 = oSupp2Coll.Item(2)
        End Select
        sSmartItem = "TeeWeld"
    End If
    
    sMsg = ">>> oSDO_PhysicalConn.Create"
    Set oSDO_PhysicalConn = New StructDetailObjectsex.SDOPhysicalConn
    Call oSDO_PhysicalConn.Create(pResourceManager, oEdgePort1, oEdgePort2, _
                                  sSmartItem, oInsertPlate, ConnectionStandard)
    
    Set pObject = oSDO_PhysicalConn.object
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sErrMsg).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub SetPlatePartProperties(oSmartPlate As Object, _
                                  oResourceManager As IUnknown, _
                                  strEntity As String, _
                                  strNamingCategoryTable As String, _
                                  eTightness As GSCADShipGeomOps.StructPlateTightness, _
                                  ePlateType As GSCADShipGeomOps.StructPlateType, _
                                  strMatl As String, _
                                  strGrade As String, _
                                  dThickness As Double)
Const METHOD = "SetPlatePartProperties"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim strLongNames() As String
    Dim strShortNames() As String

    Dim iIndex As Long
    Dim lPriority() As Long

    Dim oPlate As IJPlate
    Dim oMoldedConv As IJDPlateMoldedConventions
    
    Dim oRules As IJElements
    Dim oDummyAE As IJNameRuleAE
    Dim oQueryUtil As IJMetaDataCategoryQuery
    Dim oNamingObject As IJDNamingRulesHelper
    
    Set oPlate = oSmartPlate
    If Len(Trim(strEntity)) > 0 Then
        'Retrieve first default naming rule
        Set oNamingObject = New NamingRulesHelper
        oNamingObject.GetEntityNamingRulesGivenName strEntity, oRules
        If oRules.Count >= 1 Then
            oNamingObject.AddNamingRelations oSmartPlate, oRules.Item(1), oDummyAE
        End If
        Set oDummyAE = Nothing
        Set oNamingObject = Nothing
    End If
    
    If Len(Trim(strNamingCategoryTable)) > 0 Then
        ' Default naming category to first non-negative value
        If oResourceManager Is Nothing Then
            Dim oIJDObject As IJDObject
            Set oIJDObject = oSmartPlate
            Set oResourceManager = oIJDObject.ResourceManager
        End If
        
        Set oQueryUtil = New CMetaDataCategoryQuery
        oQueryUtil.GetCategoryInfo oResourceManager, _
                                   strNamingCategoryTable, _
                                   strLongNames, _
                                   strShortNames, _
                                   lPriority
        Set oQueryUtil = Nothing
    
        oPlate.NamingCategory = -1
        For iIndex = LBound(lPriority) To UBound(lPriority)
            If lPriority(iIndex) >= 0 Then
                oPlate.NamingCategory = lPriority(iIndex)
                Exit For
            End If
        Next iIndex
    
        Erase strLongNames
        Erase strShortNames
        Erase lPriority
    End If
    
    If dThickness > 0.0001 Then
        'Set Plate Type
        'Set Plate Tightness
        oPlate.plateType = ePlateType
        oPlate.Tightness = eTightness
                
        'Set Plate Material, Grade and Thickness
        SetCollarMatlAndGrade oSmartPlate, strMatl, strGrade
        SetCollarThickness oSmartPlate, dThickness
    End If
  
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

