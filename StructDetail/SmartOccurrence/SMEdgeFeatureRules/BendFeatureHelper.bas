Attribute VB_Name = "BendFeatureHelper"
Option Explicit

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\SharedVB\BendFeatureHelper.bas"
Private Const E_INVALIDARG = &H80070057
'

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub GetEdgeFeatureInputs(oEdgeFeature As Object, oEdgePort As Object, oFeatureLocation As Object)
Const METHOD = "GetEdgeFeatureInputs"
    On Error GoTo ErrorHandler

    Dim oIJStructFeatUtils As IJSDFeatureAttributes
    If TypeOf oEdgeFeature Is IJStructFeature Then
        Set oIJStructFeatUtils = New SDFeatureUtils
        oIJStructFeatUtils.Get_Inputs_EdgeCut oEdgeFeature, oEdgePort, oFeatureLocation
    End If

  Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function IsKnuckleBendFeature(oEdgeFeature As Object) As Boolean
Const METHOD = "IsKnuckleBendFeature"
    On Error GoTo ErrorHandler

    Dim oEdgePort As Object
    Dim oFeatureLocation As Object
    Dim oIJStructFeatUtils As IJSDFeatureAttributes
    
    IsKnuckleBendFeature = False
    If oEdgeFeature Is Nothing Then
        Exit Function
    ElseIf Not TypeOf oEdgeFeature Is IJStructFeature Then
        Exit Function
    End If
    
    Set oIJStructFeatUtils = New SDFeatureUtils
    oIJStructFeatUtils.Get_Inputs_EdgeCut oEdgeFeature, oEdgePort, oFeatureLocation

    IsKnuckleBendFeature = IsKnucklePlusFeature(oFeatureLocation)

 Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function IsKnucklePlusFeature(oFeatureLocation As Object) As Boolean
Const METHOD = "IsKnucklePlusFeature"
    On Error GoTo ErrorHandler

    Dim oEdgePort As Object
    Dim oProfileKnuckle As IJProfileKnuckle
    Dim oProfileKnuckleMfg As IJProfileKnuckleMfg
    
    IsKnucklePlusFeature = False
    If oFeatureLocation Is Nothing Then
        Exit Function
    ElseIf Not TypeOf oFeatureLocation Is IJProfileKnuckle Then
        Exit Function
    End If

    Set oProfileKnuckleMfg = oFeatureLocation
    If oProfileKnuckleMfg.ManufacturingMethod = pkmmBendPlusFeature Then
        IsKnucklePlusFeature = True
'''    ElseIf oProfileKnuckleMfg.ManufacturingMethod = pkmmBend Then
'''        IsKnucklePlusFeature = True
    End If
    
 Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub CreateInsertPlate(pMemberDescription As IJDMemberDescription, _
                             pResourceManager As IUnknown, _
                             bWebEdge As Boolean, _
                             sSmartItem As String, _
                             pObject As Object)
Const METHOD = "CreateInsertPlate"
    On Error GoTo ErrorHandler
    
    Dim dThickness As Double
    
    Dim oEdgeFeature As Object
    Dim oInsertPlate As Object
    Dim oPartWithFeature As Object
    Dim oFeatureLocation As Object
    Dim oPhysicalEdgePort As Object
    
    Dim oSmartPlate As IJSmartPlate
    Dim oStructFeature As IJStructFeature
    Dim eStructFeatureType As StructFeatureTypes
    Dim oParentSystem As Object
    Dim oInputCollection As JCmnShp_CollectionAlias
    
    Dim oSDO_EdgeFeature As StructDetailObjects.EdgeFeature
    Dim oSDSmartPlateUtils As GSCADSDCreateModifyUtilities.SDSmartPlateUtils
    Dim oSDFeatureAttributes As GSCADSDCreateModifyUtilities.IJSDFeatureAttributes
    
    ' Get edge feature
    Set oEdgeFeature = pMemberDescription.CAO
    If Not TypeOf oEdgeFeature Is IJStructFeature Then
        Exit Sub
    End If
    
    Set oStructFeature = oEdgeFeature
    eStructFeatureType = oStructFeature.get_StructFeatureType
    If (eStructFeatureType <> SF_EdgeFeature) Then
        Exit Sub
    End If
    
    ' Set up edge feature wrapper
    Set oSDO_EdgeFeature = New StructDetailObjects.EdgeFeature
    Set oSDO_EdgeFeature.object = oEdgeFeature
    
    ' Get the cut out part
    Set oPartWithFeature = oSDO_EdgeFeature.GetPartObject
        
    Set oSDFeatureAttributes = New GSCADSDCreateModifyUtilities.SDFeatureUtils
    oSDFeatureAttributes.Get_Inputs_EdgeCut oEdgeFeature, _
                                       oPhysicalEdgePort, _
                                       oFeatureLocation
    Set oSDFeatureAttributes = Nothing
                                       
    Set oParentSystem = oEdgeFeature
''Dim oSystemChild As IJSystemChild
''Set oSystemChild = oEdgeFeature
''Set oParentSystem = oSystemChild.GetParent
''
    
    Set oInputCollection = New Collection
    oInputCollection.Add oEdgeFeature

    Set oSDSmartPlateUtils = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
    Set oInsertPlate = oSDSmartPlateUtils.CreateInsertPlatePart(pResourceManager, _
                                                                sSmartItem, _
                                                                oInputCollection, _
                                                                oParentSystem)
    
    ' Need to initialize the Plate Material, Grade, Thickness
    ' What is the Inital Material, Grade, Thickness based on ...
    '   is it based on Material, Grade of Profile
    '   is it based on Material, Grade of Stiffened Plate
    '   other?
    '       for now, just hard code it
    ' ... for Naming Category: ? ... BearingPlateCategory ... PlatePartCategory
    GetInsertPlateThicknessFromEdgeFeature oEdgeFeature, dThickness
    SetPlatePartProperties oInsertPlate, pResourceManager, _
                           "CPlatePart", "InsertPlateCategory", _
                           NonTight, Standalone, _
                           "Steel - Carbon", "A", dThickness
    
    Set pObject = oInsertPlate
    Set oSDSmartPlateUtils = Nothing
        
    Set oSDO_EdgeFeature = Nothing
    Set oEdgeFeature = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Function IsEdgeFeatureOnWeb(oEdgeFeature As Object, lXid1 As Long, lXid2 As Long) As Boolean
Const METHOD = "IsEdgeFeatureonWeb"
    On Error GoTo ErrorHandler
    
    Dim lLen As Long
    Dim lXid As Long
    Dim sXid As String
    Dim oEdgePort As Object
    Dim oEdgeLocation As Object

    Dim oStructPort As IJStructPort

    lXid1 = 0
    lXid2 = 0
    IsEdgeFeatureOnWeb = False
    GetEdgeFeatureInputs oEdgeFeature, oEdgePort, oEdgeLocation
    
    If TypeOf oEdgePort Is IJStructPort Then
        Set oStructPort = oEdgePort
        lXid = oStructPort.SectionID
        sXid = "00000000" & Trim(Hex(lXid))
        lLen = Len(sXid)
        lXid1 = CLng("&H" & Mid(sXid, lLen - 3, 4))
        lXid2 = CLng("&H" & Mid(sXid, lLen - 7, 4))
        
        If lXid1 = 257 Then
            IsEdgeFeatureOnWeb = True
        ElseIf lXid1 = 258 Then
            IsEdgeFeatureOnWeb = True
        ElseIf lXid2 = 257 Then
            IsEdgeFeatureOnWeb = True
        ElseIf lXid2 = 258 Then
            IsEdgeFeatureOnWeb = True
        End If
        
'''Dim sText As String
'''sText = "oEdgePort..."
'''
'''If TypeOf oEdgePort Is IJStructPort Then
'''Set oStructPort = oEdgePort
'''sText = sText & vbCrLf & _
'''        "   Ctx: " & oStructPort.ContextID & _
'''        "   Opt: " & oStructPort.OperationID & _
'''        "   Opr: " & oStructPort.OperatorID & _
'''        "   Xid: " & oStructPort.SectionID & "  x:" & Hex(oStructPort.SectionID) & vbCrLf & _
'''        "   Xid1: " & lXid1 & "  x:" & Hex(lXid1) & _
'''        "   Xid2: " & lXid2 & "  x:" & Hex(lXid2)
'''End If
'''MsgBox sText
    
        
    End If

  Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub GetInsertPlateThicknessFromEdgeFeature(oEdgeFeature As Object, dThickness As Double)
Const METHOD = "GetInsertPlateThickness"
    On Error GoTo ErrorHandler
    
    Dim sMsg As String
    Dim lXid1 As Long
    Dim lXid2 As Long
    Dim bOnWeb As Boolean
    
    Dim oProfilePart As Object
    
    Dim oSDO_EdgeFeature As StructDetailObjects.EdgeFeature
    Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
    
    Set oSDO_EdgeFeature = New StructDetailObjects.EdgeFeature
    Set oSDO_EdgeFeature.object = oEdgeFeature
    Set oProfilePart = oSDO_EdgeFeature.GetPartObject
    
    bOnWeb = IsEdgeFeatureOnWeb(oEdgeFeature, lXid1, lXid2)
    
    dThickness = 0#
    Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
    Set oSDO_ProfilePart.object = oProfilePart
    If bOnWeb Then
        dThickness = oSDO_ProfilePart.WebThickness
    Else
        dThickness = oSDO_ProfilePart.flangeThickness
    End If
    
    If dThickness < 0.001 Then
        dThickness = 0.01
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub SetInsertPlateMatlAndGrade(oPlate As IJPlate, sMaterial As String, sGrade As String)
Const METHOD = "SetInsertPlateMatlAndGrade"
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
    
    Dim oResMgr As IJDPOM
    
    Dim oMatProxy As Object
    Dim oPlateObject As IJDObject
    Set oPlateObject = oPlate
    Set oResMgr = oPlateObject.ResourceManager
    
    Set oPlateMaterial = oPlate
    Set oMatProxy = oResMgr.GetProxy(oMaterial)
    oPlateMaterial.Material = oMatProxy
        
    Set oPlateMaterial = Nothing
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub SetInsertPlateThickness(oPlate As IJPlate, dThickness As Double)
Const METHOD = "SetInsertPlateThickness"
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim nCount As Long
    
    Dim matlThickCol As RefDataMiddleServices.IJDCollection 
    
    Dim oMaterial As IJDMaterial
    Dim oPlateDims As IJDPlateDimensions
    Dim oPlateMaterial As IJStructureMaterial   
    Dim oRefDataQuery As RefDataMiddleServices.RefdataSOMMiddleServices
    Dim oStandAlonePart As IJDStandAlonePlatePart
    
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
    
    
    Set oRefDataQuery = New RefDataMiddleServices.RefdataSOMMiddleServices    
    Set matlThickCol = oRefDataQuery.GetPlateDimensions(oMaterial.MaterialType, oMaterial.MaterialGrade)
    Dim oPlateDimProxy As Object
    Dim oResMgr As IJDPOM
    Dim oPlateObject As IJDObject
    Set oPlateObject = oPlate
    Set oResMgr = oPlateObject.ResourceManager
    If matlThickCol Is Nothing Then
    ElseIf matlThickCol.Size < 1 Then
    Else
        nCount = matlThickCol.Size
        
        For i = 1 To nCount
           Set oPlateDims = matlThickCol.Item(i)
           If Abs(oPlateDims.thickness - dThickness) < 0.000005 Then
              Exit For
           End If
           Set oPlateDims = Nothing
        Next
        
        If oPlateDims Is Nothing Then
            Set oPlateDims = matlThickCol.Item(1)
        End If
        Set oPlateDimProxy = oResMgr.GetProxy(oPlateDims)
        oPlate.Dimensions = oPlateDimProxy
    End If
    
    Set oStandAlonePart = oPlate
    oStandAlonePart.plateThicknessOffset = -oPlate.thickness

    Set matlThickCol = Nothing
    
    Set oMaterial = Nothing
    Set oPlateDims = Nothing
    Set oRefDataQuery = Nothing
    Set oPlateMaterial = Nothing
   
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub CreatePC_FeatureAndInsertPlate(pMemberDescription As IJDMemberDescription, _
                                          pResourceManager As IUnknown, _
                                          bWebEdge As Boolean, _
                                          sSmartItem As String, _
                                          pObject As Object)
Const METHOD = "CreatePC_FeatureAndInsertPlate"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    Dim sErrMsg As String
    
    Dim oEdgePort1 As Object
    Dim oEdgePort2 As Object
    Dim oEdgeFeature As Object
    Dim oInsertPlate As Object
    Dim oProfilePart As Object
    Dim oInsertInputs As Collection
    
    Dim oSDO_PlatePart As StructDetailObjects.PlatePart
    Dim oSDO_EdgeFeature As StructDetailObjects.EdgeFeature
    Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
    Dim oSDO_PhysicalConn As IJSDOPhysicalConnection
    
    Dim oSDSmartPlateAttributes As GSCADSDCreateModifyUtilities.IJSDSmartPlateAttributes
    
    ' Physical Connection between Profile Edge Feature Port and Insert Plate Lateral Edge
    sMsg = ">>> GetInputs_InsertPlate"
    Set oInsertPlate = pMemberDescription.CAO
    Set oSDSmartPlateAttributes = New GSCADSDCreateModifyUtilities.SDSmartPlateUtils
    oSDSmartPlateAttributes.GetInputs_InsertPlate oInsertPlate, oInsertInputs
    Set oEdgeFeature = oInsertInputs.Item(1)
    
    sMsg = ">>> oSDO_EdgeFeature.GetPartObject"
    Set oSDO_EdgeFeature = New StructDetailObjects.EdgeFeature
    Set oSDO_EdgeFeature.object = oEdgeFeature
    Set oProfilePart = oSDO_EdgeFeature.GetPartObject
    
    sMsg = ">>> oSDO_ProfilePart.BasePort(BPT_Lateral)"
    Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
    Set oSDO_ProfilePart.object = oProfilePart
    Set oEdgePort1 = oSDO_ProfilePart.CutoutPort(oEdgeFeature)

    sMsg = ">>> oSDO_PlatePart.BasePort(BPT_Lateral)"
    Set oSDO_PlatePart = New StructDetailObjects.PlatePart
    Set oSDO_PlatePart.object = oInsertPlate
    Set oEdgePort2 = oSDO_PlatePart.BasePort(BPT_Lateral)

    sMsg = ">>> oSDO_PhysicalConn.Create"
    Set oSDO_PhysicalConn = New StructDetailObjectsEx.SDOPhysicalConn
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
Const METHOD = "CreatePC_FeatureAndInsertPlate"
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
    
    'Retrieve first default naming rule
    Set oNamingObject = New NamingRulesHelper
    oNamingObject.GetEntityNamingRulesGivenName strEntity, oRules
    If oRules.Count >= 1 Then
        oNamingObject.AddNamingRelations oSmartPlate, oRules.Item(1), oDummyAE
    End If
    Set oDummyAE = Nothing
    Set oNamingObject = Nothing
    
    ' Default naming category to first non-negative value
    Set oQueryUtil = New CMetaDataCategoryQuery
    oQueryUtil.GetCategoryInfo oResourceManager, _
                               strNamingCategoryTable, _
                               strLongNames, _
                               strShortNames, _
                               lPriority
    Set oQueryUtil = Nothing

    Set oPlate = oSmartPlate
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
            
    'Set Plate Type
    'Set Plate Tightness
    oPlate.plateType = ePlateType
    oPlate.Tightness = eTightness
            
    'Set Plate Material, Grade and Thickness
    SetInsertPlateMatlAndGrade oSmartPlate, strMatl, strGrade
    SetInsertPlateThickness oSmartPlate, dThickness
  
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub GetKnuckleCutData(oEdgeFeature As Object, _
                             oProfileKnuckle As IJProfileKnuckle, _
                             dHieght As Double, _
                             dOffset As Double, _
                             dBendAngle As Double, _
                             dBendWidth As Double, _
                             dCutWidth As Double, _
                             dMinCutWidth As Double, _
                             dCutAngle As Double)
Const METHOD = "GetKnuckleCutData"
    On Error GoTo ErrorHandler
    Dim sMsg As String
    
    Dim dAngle As Double
    Dim dInner As Double
    Dim dOuter As Double
    Dim dTmpAngle As Double
    Dim dCutRadius As Double
    Dim dBendRadius As Double
    
    Dim oProfileBendKnuckle As IJProfileBendKnuckle
    
    Set oProfileBendKnuckle = oProfileKnuckle
    
    ' Expect 90.0 <= ProfileKnuckle.Angle <= 180.0
    dAngle = (Atn(1) * 4#) - Abs(oProfileKnuckle.Angle)
    dInner = oProfileBendKnuckle.InnerRadius
    dOuter = oProfileBendKnuckle.OuterRadius
    
    ' Base calulcations on the ProfileBendKnuckle.InnerRadius
    ' ... Expect this to be a valid Bend Radius value
    dBendRadius = dInner + dHieght
    dBendAngle = dAngle / 2#
    dBendWidth = Tan(dBendAngle) * dBendRadius
    
    dMinCutWidth = Tan(dBendAngle) * (dInner + dOffset)
    dCutWidth = Sin(dBendAngle) * dBendRadius
    dCutRadius = dCutWidth / Tan(dBendAngle)
    
    dTmpAngle = dCutWidth / (dCutRadius - dInner - dOffset)
    dCutAngle = Atn(dTmpAngle)
    
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'''sMsg = METHOD
'''sMsg = sMsg & vbCrLf & _
'''       "ProfileKnuckle.Angle: " & Format(oProfileKnuckle.Angle, "0.0000") & vbCrLf & _
'''       " ...(deg) " & Format(((45# / Atn(1)) * oProfileKnuckle.Angle), "0.0000") & vbCrLf & _
'''       "dHieght: " & Format(dHieght, "0.0000") & vbCrLf & _
'''       "dOffset: " & Format(dOffset, "0.0000")
'''
'''sMsg = sMsg & vbCrLf & _
'''       "dAngle: " & Format(dAngle, "0.0000") & vbCrLf & _
'''       " ...(deg) " & Format(((45# / Atn(1)) * dAngle), "0.0000") & vbCrLf & _
'''       "dInner: " & Format(dInner, "0.0000") & vbCrLf & _
'''       "dOuter: " & Format(dOuter, "0.0000")
'''
'''sMsg = sMsg & vbCrLf & _
'''       "dBendRadius: " & Format(dBendRadius, "0.0000") & vbCrLf & _
'''       "dBendAngle: " & Format(dBendAngle, "0.0000") & vbCrLf & _
'''       " ...(deg) " & Format(((45# / Atn(1)) * dBendAngle), "0.0000") & vbCrLf & _
'''       "dBendWidth: " & Format(dBendWidth, "0.0000")
'''
'''sMsg = sMsg & vbCrLf & _
'''       "dCutWidth: " & Format(dCutWidth, "0.0000") & vbCrLf & _
'''       "dMinCutWidth: " & Format(dMinCutWidth, "0.0000") & vbCrLf & _
'''       "dCutRadius: " & Format(dCutRadius, "0.0000") & vbCrLf & _
'''       "dCutAngle: " & Format(dCutAngle, "0.0000") & vbCrLf & _
'''       " ...(deg) " & Format(((45# / Atn(1)) * dCutAngle), "0.0000")
'''MsgBox sMsg
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sMsg).Number
End Sub


