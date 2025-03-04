Attribute VB_Name = "CanRuleCommon"
'   July 17, 2009       GG   166187 Attribute management code needs to validate user keyed in values and update related values
Option Explicit

' for inputs and outputs, index refers to the order that the inputs / outputs are listed.

Public Const DESIGNEDMEMBPROGID = "SPSMembers.SPSDesignedMember"

Public Const InterfaceName_IJUASMCanRuleDiameter = "IJUASMCanRuleDiameter"
Public Const InterfaceName_IJUASMCanRuleCanMaterial = "IJUASMCanRuleCanMaterial"
Public Const InterfaceName_IJUASMCanRuleCone1Material = "IJUASMCanRuleCone1Material"
Public Const InterfaceName_IJUASMCanRuleCone2Material = "IJUASMCanRuleCone2Material"
Public Const InterfaceName_IJUASMCanRuleConeMaterial = "IJUASMCanRuleConeMaterial"
Public Const InterfaceName_IJUASMCanRuleInLine = "IJUASMCanRuleInLine"
Public Const InterfaceName_IJUASMCanRuleStubEnd = "IJUASMCanRuleStubEnd"
Public Const InterfaceName_IJUASMCanRuleEnd = "IJUASMCanRuleEnd"
Public Const InterfaceName_IJUASMCanRuleResult = "IJUASMCanRuleResult"

Public Const InterfaceName_IUABuiltUpTube = "IUABuiltUpTube"
Public Const attrTubeThickness As String = "TubeThickness"
Public Const attrChamfer1Length As String = "Chamfer1Length"
Public Const attrChamfer2Length As String = "Chamfer2Length"

Public Const InterfaceName_IUABuiltUpCan = "IUABuiltUpCan"
Public Const attrDiameterStart As String = "DiameterStart"
Public Const attrDiameterEnd As String = "DiameterEnd"
Public Const attrLengthStartCone As String = "LengthStartCone"
Public Const attrLengthEndCone As String = "LengthEndCone"

Public Const crname_SplitConnection = "SplitBySurface-1"       ' see StructSplitConnections.xls
Public Const crname_BuiltUpCan = "BUCan_Custom"       ' part name of output builtup

Public Const E_FAIL = -2147467259
Public Const S_FALSE = 1 ' Used to log warnings to the error log

Public Const crName_BuiltUpTubeInterface = "IUABuiltUpTube"
Public Const crNameCanRuleXSectionInterface = "IJUASMCanRuleDiameter"

Public Const DiameterRule_ByOD = 1
Public Const DiameterRule_ByID = 2
Public Const DiameterRule_User = 3
Public Const TubeExtension_HullFactor = 1
Public Const TubeExtension_HullLength = 2
Public Const TubeExtension_CLFactor = 3
Public Const TubeExtension_CLLength = 4
Public Const ConeMethod_Slope = 1
Public Const ConeMethod_Angle = 2
Public Const ConeMethod_Length = 3

Public Const attrCanType As String = "CanType"
Public Const CanType_InLine = 1
Public Const CanType_StubEnd = 2
Public Const CanType_End = 3

Public Const attrDiameterRule As String = "DiameterRule"
Public Const attrTubeDiameter As String = "TubeDiameter"
Public Const attrTubePlateThickness As String = "TubePlateThickness"
Public Const attrTubeExtensionMinimum As String = "TubeExtensionMinimum"

Public Const attrTubeExtension As String = "TubeExtension"
Public Const attrTubeLength As String = "TubeLength"
Public Const attrTubeFactor As String = "TubeFactor"

Public Const attrConeMethod As String = "ConeMethod"
Public Const attrConeThickness As String = "ConeThickness"
Public Const attrConeLength As String = "ConeLength"
Public Const attrConeSlope As String = "ConeSlope"
Public Const attrConeAngle As String = "ConeAngle"
Public Const attrChamferSlope As String = "ChamferSlope"
Public Const attrRoundoffDistance As String = "RoundoffDistance"
Public Const attrMinExtensionDistance As String = "MinExtensionDistance"

Public Const attrTube1Extension As String = "Tube1Extension"
Public Const attrTube2Extension As String = "Tube2Extension"
Public Const attrTube1Length As String = "Tube1Length"
Public Const attrTube2Length As String = "Tube2Length"
Public Const attrTube1Factor As String = "Tube1Factor"
Public Const attrTube2Factor As String = "Tube2Factor"

'SymbolAttributes
Public Const attrInnerDiameter As String = "CanInsideDiameter"
Public Const attrOuterDiameter As String = "CanOutsideDiameter"
Public Const attrL2CompMethod As String = "L2ComputationMethod"
Public Const attrL3CompMethod As String = "L3ComputationMethod"
Public Const attrConeLengthMethod As String = "ConeLengthMethod"
Public Const attrCone1LengthMethod As String = "Cone1LengthMethod"
Public Const attrCone2LengthMethod As String = "Cone2LengthMethod"
Public Const attrMinimumExtensionDistance As String = "MinimumExtensionDistance"

'SymbolParameter
Public Const attrCone1Thickness As String = "Cone1Thickness"
Public Const attrCone2Thickness As String = "Cone2Thickness"
Public Const attrCone1Method As String = "Cone1Method"
Public Const attrCone2Method As String = "Cone2Method"
Public Const attrCone1Length As String = "Cone1Length"
Public Const attrCone2Length As String = "Cone2Length"
Public Const attrCone1Slope As String = "Cone1Slope"
Public Const attrCone2Slope As String = "Cone2Slope"
Public Const attrCone1Angle As String = "Cone1Angle"
Public Const attrCone2Angle As String = "Cone2Angle"

Public Const attrL2Method As String = "L2Method"
Public Const attrL3Method As String = "L3Method"
Public Const attrL2Length As String = "L2Length"
Public Const attrL3Length As String = "L3Length"
Public Const attrL2Factor As String = "L2Factor"
Public Const attrL3Factor As String = "L3Factor"

Public Const attrCanDiameter As String = "CanDiameter"
Public Const attrCanThickness As String = "CanThickness"
Public Const attrResultCanLength As String = "CanLength"
Public Const attrResultUncutLength As String = "UncutLength"
Public Const attrResultCanType As String = "CanType"
Public Const attrResultL2HullLength As String = "L2HullLength"
Public Const attrResultL2CenterlineLength As String = "L2CenterlineLength"
Public Const attrResultL3HullLength As String = "L3HullLength"
Public Const attrResultL3CenterlineLength As String = "L3CenterlineLength"

Public Const dFootInMeter As Double = 0.3048    'One foot in meter
Public Const dTol As Double = 0.000001
Public Const strCodeListTablename = "StructCanRuleToDoMessages"
Public Const PI As Double = 3.14159265

Const MODULE = "CanRuleCommon"
'See StructCustomCodeList.xls
Public Enum CodelistedErrors
    MISSING_MANDATORY_INPUT = 1
    SPLIT_ALREADY_EXISTS = 2
    MISSING_SPLIT_DEF = 3
    MISSING_CAN_DEF = 4
    MISSING_TUBE_DEF = 5
    INVALID_PARAMETERS = 6
    UNEXPECTED_ERROR = 7
End Enum


' conditionally create the splitting plane and the split connection.  make the plane the output of the GC

Public Function CreateSplitConnection(MyGCMacro As IJGeometricConstructionMacro, _
                    dSplitOffset As Double, outputKey As String) As SPSCanRuleStatus

    Const METHOD = "CreateSurfaceSplitConnection"
    On Error GoTo ErrorHandler

    Dim oGC As IJGeometricConstruction
    Dim oCanRule As ISPSCanRule
    Dim oPOM As Object
    Dim oMSPrimary As iSPSMemberSystem
    Dim oVecPrimary As IJDVector
    
    Dim oMemberFactory As SPSMemberFactory
    Dim oSplitConn As ISPSSplitMemberConnection
    Dim oILC As IJStructILCConnection
    Dim oSplitHelper As IJStructILCHelper
    Dim elesPlane As IJElements, elesGCPlanes As IJElements
    Dim elesSplitConnections As IJElements
    Dim splitStatus As StructSOCInputHelperStatus

    Dim oGeomFactory As GeometryFactory
    Dim oPoint As IJPoint
    Dim uvecX As Double, uvecY As Double, uvecZ As Double
    Dim PosX As Double, PosY As Double, PosZ As Double
    
    Dim iControlFlags As IJControlFlags
    Dim oPlane As IJPlane
    
    Dim lKey As Long
    Dim bNewSplit As Boolean, bNewPlane As Boolean

    ' get the PrimaryMemberSystem
    ' get the PrimaryAxis
    ' get POM
    ' create a plane by point normal
    ' set the plane to invisible and set its PG
    ' add the plane to GC outputs
    '
    ' create a collection for secondary objects
    ' create a split connection
    ' set the PG of the split connection
    ' set the splitConn DefinitionName
    ' set the inputs to the split connection

    Set oPOM = GetPOM(MyGCMacro)

    Set oMemberFactory = New SPSMemberFactory
    Set oGeomFactory = New GeometryFactory

    Set oGC = MyGCMacro
    Set oCanRule = MyGCMacro
    
    oCanRule.GetPrimaryMemberSystem oMSPrimary
    GetUnitVectorTangent oMSPrimary.PhysicalAxis, uvecX, uvecY, uvecZ
    
    ' get the CanRule position and offset it by the specified amount
    Set oPoint = MyGCMacro

    oPoint.GetPoint PosX, PosY, PosZ
    PosX = PosX + dSplitOffset * uvecX
    PosY = PosY + dSplitOffset * uvecY
    PosZ = PosZ + dSplitOffset * uvecZ

    Set elesGCPlanes = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_Planes)
    Set elesSplitConnections = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_SplitConnections)
    bNewPlane = False
    bNewSplit = False

    lKey = CLng(outputKey)
    
    If lKey <= elesGCPlanes.count Then
        Set oPlane = elesGCPlanes.Item(outputKey)
    End If
    If oPlane Is Nothing Then
        Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(oPOM, PosX, PosY, PosZ, uvecX, uvecY, uvecZ)
        Set iControlFlags = oPlane
        iControlFlags.ControlFlags(CTL_FLAG_NO_DISPLAY) = CTL_FLAG_NO_DISPLAY
        iControlFlags.ControlFlags(CTL_FLAG_NOT_IN_SPATIAL_INDEX) = CTL_FLAG_NOT_IN_SPATIAL_INDEX
        elesGCPlanes.Add oPlane, outputKey
        bNewPlane = True
    Else
        oPlane.DefineByPointNormal PosX, PosY, PosZ, uvecX, uvecY, uvecZ
    End If

    If lKey <= elesSplitConnections.count Then
        Set oSplitConn = elesSplitConnections.Item(outputKey)
    End If
    If oSplitConn Is Nothing Then
        Set oSplitConn = oMemberFactory.CreateSplitMemberConnection(oPOM)
        elesSplitConnections.Add oSplitConn, outputKey
        bNewSplit = True
    End If

    SetPG oMSPrimary, oSplitConn
    SetPG oMSPrimary, oPlane
  
    If bNewSplit Or bNewPlane Then
        Set oSplitHelper = oSplitConn.Helper
        
        Set elesPlane = New JObjectCollection
        elesPlane.Add oPlane
        splitStatus = oSplitHelper.SetParents(oSplitConn, oMSPrimary, elesPlane)

        Set oILC = oSplitConn
        oILC.SplitParentStatus = ssSplitFirst
    End If

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function        ' CreateSplitConnection


Public Sub GetDiameters(oCanRule As IJGeometricConstructionMacro, portIndex As SPSMemberAxisPortIndex, _
                                ByRef dID As Double, ByRef dOD As Double)

    Const METHOD = "GetDiameter"
    On Error GoTo ErrorHandler
    
    Dim oObj As Object
    Dim oAttrColl As Object
    Dim varValue As Variant
    Dim dThickness As Double
    
    Dim oGC As IJGeometricConstruction
    Set oGC = oCanRule

    If portIndex = SPSMemberAxisAlong Then      ' output member is the BUCan
        Set oObj = oCanRule.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).Item("1")

    Else                                        ' get the neighbor.
        
        Dim elesNbors As IJElements
            
        Set oGC = oCanRule
        Set elesNbors = oGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_Neighbors)

        If elesNbors.count < 2 Then     ' disregard which end.  just get the nbor.  portIndex is not "AxisAlong"
            Set oObj = elesNbors.Item("1")

        ElseIf portIndex = SPSMemberAxisStart Then      ' two neighbors exist.   use portIndex to tell which one.
            Set oObj = elesNbors.Item("1")              ' make sure Migrate assigns the neighbors in this manner.
        Else
            Set oObj = elesNbors.Item("2")
        End If
    End If

    'TubeDiameter is available in BuiltUpTube's occurence attributes.
    Set oAttrColl = GetAttributeCollection(oObj, InterfaceName_IUABuiltUpTube)
    If GetAttributeValue(oAttrColl, attrTubeDiameter, varValue) Then
        dOD = varValue
    Else
        WriteToErrorLog E_FAIL, MODULE, METHOD, "failed to get TubeDiameter"
        Err.Raise E_FAIL, METHOD, "failed to get TubeDiameter"
    End If

    If GetAttributeValue(oAttrColl, attrTubeThickness, varValue) Then
        dThickness = varValue
    Else
        WriteToErrorLog E_FAIL, MODULE, METHOD, "failed to get TubeThickness"
        Err.Raise E_FAIL, METHOD, "failed to get TubeThickness"
    End If
    
    If portIndex <> SPSMemberAxisAlong Then                 ' get neighbors plate to get its real thickness
        On Error Resume Next
        Dim iIJPlate As IJPlate
        Dim iIJDMemberObjects As IJDMemberObjects
        Set iIJDMemberObjects = oObj
        If Not iIJDMemberObjects Is Nothing Then
            Set iIJPlate = iIJDMemberObjects.Item("Tube")
            If Not iIJPlate Is Nothing Then
                dThickness = iIJPlate.thickness
            End If
        End If
    End If

    dID = dOD - 2 * dThickness

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub FindClosestMember(x As Double, y As Double, z As Double, elesMembers As IJElements, oMemberToSkip As Object, ByRef oClosestMember As Object)

    Const METHOD = "FindClosestMember"
    On Error GoTo ErrorHandler

    Dim position As IJDPosition
    Dim ii As Long, count As Long
    Dim oCurve As IJCurve
    Dim distEle As Double, distMin As Double
    Dim pointX As Double, pointY As Double, pointZ As Double
    Dim curveX As Double, curveY As Double, curveZ As Double
    
    Set position = New DPosition
    position.x = x
    position.y = y
    position.z = z

    ' given a list of members, get the one closest to the given position.
    ' skip the [optional] oMemberToSkip

    count = elesMembers.count
    distMin = 100000000
    
    For ii = 1 To count
        Set oCurve = elesMembers(ii)
        If Not oCurve Is oMemberToSkip Then
            oCurve.DistanceBetween position, distEle, pointX, pointY, pointZ, curveX, curveY, curveZ
            If distEle < distMin Then
                distMin = distEle
                Set oClosestMember = oCurve
            End If
        End If
    Next

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

' given a designed member, set its definition to be the BUCan type

Public Sub SetBUCanDefinition(oDMSection As ISPSCrossSection)
    Const METHOD = "SetBUCanDefinition"
    On Error GoTo ErrorHandler

    Dim varDefName As Variant
    Dim oDefinition As Object
    Dim oCatalogConnection As IJDPOM
    
    If oDMSection.SectionName <> crname_BuiltUpCan Then
        
        ' keep the existing standard, and set to BUCan_Custom
        varDefName = oDMSection.SectionStandard & ", BUCan, " & crname_BuiltUpCan
        
        Set oCatalogConnection = GetCatalogDBConnection
        Set oDefinition = oCatalogConnection.GetObject(varDefName)
        If oDefinition Is Nothing Then
            Err.Raise E_FAIL, METHOD, "failed to bind to " & varDefName
            'Exit Sub
        End If
        
        Set oDMSection.Definition = oDefinition
    End If
    
    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub UpdatePlane(iPlane As IJPlane, pos As IJDPosition, vecX As Double, vecY As Double, vecZ As Double)

    Dim posCurrX As Double, posCurrY As Double, posCurrZ As Double
    Dim vecCurrX As Double, vecCurrY As Double, vecCurrZ As Double
    Dim tol As Double

    tol = 0.000001

    iPlane.GetRootPoint posCurrX, posCurrY, posCurrZ
    iPlane.GetNormal vecCurrX, vecCurrY, vecCurrZ
    
    If (Abs(pos.x - posCurrX) > tol Or _
        Abs(pos.y - posCurrY) > tol Or _
        Abs(pos.z - posCurrZ) > tol Or _
        Abs(vecX - vecCurrX) > tol Or _
        Abs(vecY - vecCurrY) > tol Or _
        Abs(vecZ - vecCurrZ) > tol) Then
        
        iPlane.DefineByPointNormal pos.x, pos.y, pos.z, vecX, vecY, vecZ
    
    End If

    Exit Sub

End Sub

' conditionally create the splitting plane and the split connection.  make the plane the output of the GC

Public Function CreateOffsetBoundingSurface(MyGCMacro As IJGeometricConstructionMacro, _
                    dSplitOffset As Double, outputKey As String) As SPSCanRuleStatus

    Const METHOD = "CreateOffsetBoundingSurface"
    On Error GoTo ErrorHandler

    Dim oGC As IJGeometricConstruction
    Dim oCanRule As ISPSCanRule
    Dim oPOM As Object
    Dim oMSPrimary As iSPSMemberSystem
    Dim oVecPrimary As IJDVector
    
    Dim elesOffsetSurf As IJElements

    Dim oGeomFactory As GeometryFactory
    Dim oPoint As IJPoint
    Dim uvecX As Double, uvecY As Double, uvecZ As Double
    Dim PosX As Double, PosY As Double, PosZ As Double
    
    Dim iControlFlags As IJControlFlags
    Dim oPlane As IJPlane
    
    Dim lKey As Long

    Set oPOM = GetPOM(MyGCMacro)
   
    Set oGC = MyGCMacro
    Set oCanRule = MyGCMacro
    
    oCanRule.GetPrimaryMemberSystem oMSPrimary
    GetUnitVectorTangent oMSPrimary.PhysicalAxis, uvecX, uvecY, uvecZ
    
    ' get the CanRule position and offset it by the specified amount
    Set oPoint = MyGCMacro

    oPoint.GetPoint PosX, PosY, PosZ
    PosX = PosX + dSplitOffset * uvecX
    PosY = PosY + dSplitOffset * uvecY
    PosZ = PosZ + dSplitOffset * uvecZ

    Set elesOffsetSurf = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_OffsetSurface)

    lKey = CLng(outputKey)
    
    If lKey <= elesOffsetSurf.count Then
        Set oPlane = elesOffsetSurf.Item(outputKey)
    End If
    If oPlane Is Nothing Then
        Set oGeomFactory = New GeometryFactory
        Set oPlane = oGeomFactory.Planes3d.CreateByPointNormal(oPOM, PosX, PosY, PosZ, uvecX, uvecY, uvecZ)
        Set iControlFlags = oPlane
        iControlFlags.ControlFlags(CTL_FLAG_NO_DISPLAY) = CTL_FLAG_NO_DISPLAY
        iControlFlags.ControlFlags(CTL_FLAG_NOT_IN_SPATIAL_INDEX) = CTL_FLAG_NOT_IN_SPATIAL_INDEX
        elesOffsetSurf.Add oPlane, outputKey
    Else
        oPlane.DefineByPointNormal PosX, PosY, PosZ, uvecX, uvecY, uvecZ
    End If

    
    SetPG oMSPrimary, oPlane
  

    Exit Function

ErrorHandler:
    HandleError MODULE, METHOD
End Function


Public Sub WriteToErrorLog(errNumber As Long, sModule As String, sMethod As String, sMessage As String)
    Dim oEditErrors As IJEditErrors
    
    Set oEditErrors = New JServerErrors
    If Not oEditErrors Is Nothing Then
        oEditErrors.Write errNumber, sModule & ":" & sMethod, sMessage
    End If
    Set oEditErrors = Nothing
End Sub

Public Sub GetSplitNeighbors(oSplitConnection As ISPSSplitMemberConnection, ByRef oObj1 As Object, ByRef oObj2 As Object)

    On Error Resume Next
    Dim elesPorts As IJElements
    Dim iPort As IJPort

    Set oObj1 = Nothing
    Set oObj2 = Nothing

    Set elesPorts = oSplitConnection.PartPorts

    If Not elesPorts Is Nothing Then
        If elesPorts.count = 2 Then             'a can split should be zero ( if in error ) or 2
            Set iPort = elesPorts.Item(1)
            Set oObj1 = iPort.Connectable
            Set iPort = elesPorts.Item(2)
            Set oObj2 = iPort.Connectable
        End If
    End If
  
    Exit Sub

End Sub

Public Function ObjectOnTDL(oObj As Object) As Boolean

    On Error GoTo ErrorHandler

    Dim elesTDLList As IJElements

    ObjectOnTDL = False

    Set elesTDLList = Entity_GetElementsOfRelatedEntities(oObj, "IJDObject", "toErrorList")
    
    If Not elesTDLList Is Nothing Then
        If elesTDLList.count > 0 Then
            ObjectOnTDL = True
        End If
    End If

    Set elesTDLList = Nothing

    Exit Function

ErrorHandler:
    Err.Clear
    ObjectOnTDL = False
    Exit Function
End Function

Public Sub RemoveOutputs(oGCMacro As IJGeometricConstructionMacro)

    On Error Resume Next
    Dim elesOutputs As IJElements
    Dim oIJDObject As IJDObject
    Dim ii As Long, count As Long

    Set elesOutputs = oGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_Planes)
    count = elesOutputs.count
    For ii = 1 To count
        Set oIJDObject = elesOutputs.Item(ii)
        oIJDObject.Remove
    Next ii

    Set elesOutputs = oGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_SplitConnections)
    count = elesOutputs.count
    For ii = 1 To count
        Set oIJDObject = elesOutputs.Item(ii)
        oIJDObject.Remove
    Next ii
    Set elesOutputs = oGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_OffsetSurface)
    count = elesOutputs.count
    For ii = 1 To count
        Set oIJDObject = elesOutputs.Item(ii)
        oIJDObject.Remove
    Next ii

    Exit Sub

End Sub

Public Sub FindActualPlateThickness(oGCCanRule As IJGeometricConstruction, strParamPrefix As String, ByRef dPlateThickness As Double)

    On Error GoTo ErrorHandler
    Const METHOD = "FindActualPlateThickness"
   
    Dim oBUHelper As BUHelperUtils
    Dim oPlateDims As IJDPlateDimensions
    Dim dNominalThickness As Double
    Dim strMaterialType As String, strMaterialgrade As String

    dNominalThickness = oGCCanRule.Parameter(strParamPrefix & "Thickness")
    strMaterialType = oGCCanRule.Parameter(strParamPrefix & "Material")
    strMaterialgrade = oGCCanRule.Parameter(strParamPrefix & "Grade")
    
    dPlateThickness = dNominalThickness         ' initialize actual thickness to user-specified value.
    
    ' Find closest ( or next larger ) thickness based on available thicknesses for given material
    ' This method is also used by the BUCan, which is supposed to synch CanRule computations with the BUCan plates

    Set oBUHelper = New SM3DBUHelper.BUHelperUtils
    oBUHelper.FindPlateDimension strMaterialType, strMaterialgrade, dNominalThickness, oPlateDims
    If Not oPlateDims Is Nothing Then
       dPlateThickness = oPlateDims.thickness
    End If

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'TR #163724
'Always Use CatalogDefaultNameRule for BUCan, because this Name Rule is in SPSNameRule for SmartPlantStructure
Public Sub SetNameRule(oObject As Object)
Const METHOD = "SetNameRule"
On Error GoTo ErrHandler

    Dim NamingRules As IJElements
    Dim oNameRuleHolder As GSCADGenericNamingRulesFacelets.IJDNameRuleHolder
    Dim oNameRuleHlpr As GSCADNameRuleSemantics.IJDNamingRulesHelper
    Dim oNameRuleAE As GSCADGenNameRuleAE.IJNameRuleAE
    
    Set oNameRuleHlpr = New GSCADNameRuleHlpr.NamingRulesHelper

    oNameRuleHlpr.GetEntityNamingRulesGivenProgID DESIGNEDMEMBPROGID, NamingRules
 
    If NamingRules.count > 0 Then
        Dim i As Long, j As Long
        i = 0
        For j = 1 To NamingRules.count
            Set oNameRuleHolder = Nothing
            Set oNameRuleHolder = NamingRules.Item(j)
            If Not oNameRuleHolder Is Nothing Then
                If oNameRuleHolder.Name = "CatalogDefaultNameRule" Then
                    i = j
                End If
            End If
        Next
        If i = 0 Then
            Set oNameRuleHolder = NamingRules.Item(1)
        End If
        If Not oNameRuleHolder Is Nothing Then
            Call oNameRuleHlpr.AddNamingRelations(oObject, oNameRuleHolder, oNameRuleAE)
        End If
    End If
    Set oNameRuleHlpr = Nothing
    Set oNameRuleHolder = Nothing
    Set oNameRuleAE = Nothing
Exit Sub
ErrHandler:
    HandleError MODULE, METHOD
End Sub

Public Function GetCodeListErrorNumber(eStatus As SPSCanRuleStatus) As Long
    Select Case (eStatus)
        Case SPSCanRule_BadGCMacro_InputCount:
            GetCodeListErrorNumber = MISSING_MANDATORY_INPUT
        Case SPSCanRule_BadGCMacro_NoBuiltUpCan:
            GetCodeListErrorNumber = MISSING_CAN_DEF
        Case SPSCanRule_BadGCMacro_NoEndPort:
            GetCodeListErrorNumber = MISSING_MANDATORY_INPUT
        Case SPSCanRule_BadGCMacro_NoPrimary:
            GetCodeListErrorNumber = MISSING_MANDATORY_INPUT
        Case SPSCanRule_BadGCMacro_NoSecondary:
            GetCodeListErrorNumber = MISSING_MANDATORY_INPUT
        Case SPSCanRule_BadGCMacro_Unexpected:
            GetCodeListErrorNumber = UNEXPECTED_ERROR
        Case SPSCanRule_BadGeometry:
            GetCodeListErrorNumber = INVALID_PARAMETERS
        Case SPSCanRule_BadNumberOfObjects:
            GetCodeListErrorNumber = INVALID_PARAMETERS
        Case SPSCanRule_CannotAccess_GC_DefinitionService:
            GetCodeListErrorNumber = UNEXPECTED_ERROR
        Case SPSCanRule_CannotAccess_GC_OutputPoint:
            GetCodeListErrorNumber = MISSING_MANDATORY_INPUT
        Case SPSCanRule_CannotAccessCatalogDefinition:
            GetCodeListErrorNumber = UNEXPECTED_ERROR
        Case SPSCanRule_DuplicateObject:
            GetCodeListErrorNumber = UNEXPECTED_ERROR
        Case SPSCanRule_InconsistentRelations:
            GetCodeListErrorNumber = UNEXPECTED_ERROR
        Case SPSCanRule_InvalidTypeOfObject:
            GetCodeListErrorNumber = UNEXPECTED_ERROR
        Case SPSCanRule_MembersAreNotConnected:
            GetCodeListErrorNumber = INVALID_PARAMETERS
        Case SPSCanRule_UnexpectedError:
            GetCodeListErrorNumber = UNEXPECTED_ERROR
        Case Else
            GetCodeListErrorNumber = UNEXPECTED_ERROR
    End Select
End Function

Public Function UpdateOutputCrossSectionDimensions(bIsInLine As Boolean, oCanRule As ISPSCanRule) As SPSCanRuleStatus

    Dim oBUCan As Object
    Dim oAttrColl As Object
    Dim MyGC As IJGeometricConstruction
    Dim MyGCMacro As IJGeometricConstructionMacro
    Dim diaRule As Long
    Dim diaTube As Double, dThickness As Double
    Dim dNborStartID As Double, dNborStartOD As Double
    Dim dNborEndID As Double, dNborEndOD As Double
    Dim dThicknessParameter As Double

    UpdateOutputCrossSectionDimensions = SPSCanRule_BadGCMacro_Unexpected

    Set MyGC = oCanRule
    Set MyGCMacro = oCanRule

    If MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).count > 0 Then              ' split migration has taken place

        diaRule = MyGC.Parameter(attrDiameterRule)

        FindActualPlateThickness MyGC, "Can", dThickness

        dThicknessParameter = MyGC.Parameter("CanThickness")
        If Abs(dThicknessParameter - dThickness) > dTol Then
            MyGC.Parameter("CanThickness") = dThickness
        End If

        If diaRule = DiameterRule_User Then
            diaTube = MyGC.Parameter("CanOD")
    
            ' update the dependent CanRule parameter
            MyGC.Parameter("CanID") = diaTube - 2 * dThickness
        
        Else
            
            GetDiameters MyGC, SPSMemberAxisStart, dNborStartID, dNborStartOD
            
            If bIsInLine Then

                GetDiameters MyGC, SPSMemberAxisEnd, dNborEndID, dNborEndOD
                
                If diaRule = DiameterRule_ByOD Then
                
                    If dNborStartOD > dNborEndOD Then   ' set diaTube to be the larger one
                        diaTube = dNborStartOD
                    Else
                        diaTube = dNborEndOD
                    End If
                            
                Else
                    
                    If dNborStartID > dNborEndID Then   ' set diaTube to be the larger one
                        diaTube = dNborStartID + 2 * dThickness
                    Else
                        diaTube = dNborEndID + 2 * dThickness
                    End If
            
                End If
            
            Else
    
                If diaRule = DiameterRule_ByOD Then
                    diaTube = dNborStartOD
                Else
                    diaTube = dNborStartID + 2 * dThickness
                End If

            End If
                
            ' update the dependent CanRule parameters
            MyGC.Parameter("CanOD") = diaTube
            MyGC.Parameter("CanID") = diaTube - 2 * dThickness

        End If
    
        Set oBUCan = MyGCMacro.Outputs(StructCanRuleCollectionNames.StructCanRule_BuiltUpCan).Item("1")
        
         ' writing to IUABuiltUpTube
        Set oAttrColl = GetAttributeCollection(oBUCan, crName_BuiltUpTubeInterface)
        Call SetAttributeValue(oAttrColl, "TubeDiameter", diaTube)
        Call SetAttributeValue(oAttrColl, "TubeThickness", dThickness)

    End If
    
    UpdateOutputCrossSectionDimensions = SPSCanRule_Ok
    Exit Function

End Function

Public Sub CheckAndAddSupportingSecondary(oCR As ISPSCanRule)

    On Error Resume Next
    Dim elesSS As IJElements
    Dim oGC As IJGeometricConstruction
    Dim oMS As iSPSMemberSystem
    Dim oFC As ISPSFrameConnection
    Dim oCRSupping1 As Object, oCRSupping2 As Object, oFCSupping1 As Object, oFCSupping2 As Object
    Dim crStatus As SPSCanRuleStatus
    Dim fcStatus As SPSFCInputHelperStatus
    Dim iAxisPort As ISPSAxisEndPort

    ' get the current set of supporting secondary members in relation with the canRule
    Set oGC = oCR
    Set elesSS = oGC.ControlledInputs(StructCanRuleCollectionNames.StructCanRule_SupportingSecondary)
    If elesSS.count > 0 Then
        Set oCRSupping1 = elesSS(1)
        If Not oCRSupping1 Is Nothing And elesSS.count > 1 Then
            Set oCRSupping2 = elesSS(2)
        End If
    End If
        
    Set iAxisPort = oGC.Inputs(StructCanRuleCollectionNames.StructCanRule_EndPort).Item("1")

    If Not iAxisPort Is Nothing Then
    
        Dim oConnObj As Object
    
        Set oConnObj = iAxisPort.ILC
        If TypeOf oConnObj Is ISPSFrameConnection Then
            ' get the current set of supporting members from the FC
            crStatus = oCR.GetPrimaryMemberSystem(oMS)
            Set oFC = oMS.FrameConnectionAtEnd(oCR.portIndex)
            fcStatus = oFC.InputHelper.GetRelatedObjects(oFC, oFCSupping1, oFCSupping2)
            
            ' if FC's supporting object is not MemberSystem, then it is not to be input to the CanRule
            If Not oFCSupping1 Is Nothing Then
                If Not TypeOf oFCSupping1 Is iSPSMemberSystem Then
                    Set oFCSupping1 = Nothing
                End If
            End If
            If Not oFCSupping2 Is Nothing Then
                If Not TypeOf oFCSupping2 Is iSPSMemberSystem Then
                    Set oFCSupping2 = Nothing
                End If
            End If
            
            ' if they are the same objects, leave the CanRule's collection as is, otherwise reset to current FC supping
            If oFCSupping1 Is oCRSupping1 And oFCSupping2 Is oCRSupping2 Then
                crStatus = SPSCanRule_Ok
            ElseIf oFCSupping1 Is oCRSupping2 And oFCSupping2 Is oCRSupping1 Then
                crStatus = SPSCanRule_Ok
            Else
                elesSS.Clear
                If Not oFCSupping1 Is Nothing Then
                    elesSS.Add oFCSupping1, "1"
                    If Not oFCSupping2 Is Nothing Then
                        elesSS.Add oFCSupping2, "2"
                    End If
                End If
            End If
        ElseIf TypeOf oConnObj Is ISPSSplitMemberConnection Then
            'add the splittor as Supporting secondary
            Dim oSC As ISPSSplitMemberConnection
            Dim oSplitHelper As IJStructILCHelper
            Dim oPrimary As Object, oSecondary As Object
            Dim colOtherParents As IJElements
            Dim oMyMS As ISPSMemberSystemLinear, oPrimaryMS As ISPSMemberSystemLinear
            Dim oSecondaryMS As ISPSMemberSystemLinear

            
            Set oSC = oConnObj
            Set oSplitHelper = oSC.Helper
            oSplitHelper.GetParents oSC, oPrimary, colOtherParents
            
            If TypeOf oPrimary Is ISPSMemberSystemLinear Then
                Set oPrimaryMS = oPrimary
            End If
            If colOtherParents.count > 0 Then
                Set oSecondary = colOtherParents.Item(1)
            End If
            
            If Not oSecondary Is Nothing Then
                If TypeOf oSecondary Is ISPSMemberSystemLinear Then
                    Set oSecondaryMS = oSecondary
                End If
            End If
            Set oMyMS = iAxisPort.MemberSystem
            
            
            If oPrimaryMS Is oMyMS Then
                If Not oSecondaryMS Is Nothing Then
                    If Not oSecondaryMS Is oCRSupping1 Then
                        elesSS.Add oSecondaryMS, "1"
                    End If
                End If
            ElseIf oSecondaryMS Is oMyMS Then
                If Not oPrimaryMS Is Nothing Then
                    If Not oPrimaryMS Is oCRSupping1 Then
                        elesSS.Add oPrimaryMS, "1"
                    End If
                End If
            End If
            
        End If
    End If
End Sub

'When a secondary member is added to a CanRule, that secondary member's FC's supporting member
'was set to be this CanRule in lieu of a MemberSystem.    That is how that FC knows what object
'and hence what size to use to compute the offsets for that secondary member.  When the CanRule is deleted,
'this function resets the related FC's to be in relation to the memberSystem again.
'
'If the primary member-system of this CanRule is also being deleted, this sets that FC to unsupported.
'This is necessary to avoid error handling by the FC code since SPSMembSysSuppingNotifyDCS is not called,
'and that is because SPSMembSysSuppingNotify1RCRln is used instead of SPSMembSysSuppingNotify2RCRln as normal
'for the supporting object to reference collection relation.   The disconnect semantic for SPSMembSysSuppingNotify1RCRln
'is IMSSymbolEntities.DInputRCDisconnectStc.1 instead of SPSMembers.SPSMembSysSuppingNotifyDCS.1.
'
Public Sub ResetMemberFCs(oCanRule As ISPSCanRule, bIsPrimaryDeleted As Boolean)

    Const METHOD = "ResetMemberFCs"
    On Error GoTo ErrorHandler

    Dim oFC As ISPSFrameConnection
    Dim oSecMembCol As IJElements
    Dim lngSecMemb As Long, lngIdx As Long
    Dim oSecMembSys As iSPSMemberSystem
    Dim bCheckOtherFC As Boolean
    Dim oObjCrossSection As Object

    Dim crStatus As SPSCanRuleStatus
    Dim ePortId As SPSMemberAxisPortIndex
    Dim oPrimaryMemberSystem As iSPSMemberSystem
    
    oCanRule.GetSecondaryMemberSystems oSecMembCol
    lngSecMemb = oSecMembCol.count
    
    ePortId = oCanRule.portIndex
    
    'any frame connection that is watching this CanRule needs to instead watch the CanRule's primary MemberSystem.
    For lngIdx = 1 To lngSecMemb
        
        bCheckOtherFC = True
        Set oSecMembSys = oSecMembCol.Item(lngIdx)
        
        Set oFC = oSecMembSys.FrameConnectionAtEnd(SPSMemberAxisStart)
        If Not oFC Is Nothing Then
            Set oObjCrossSection = oFC.GetCrossSectionObject(SPSFCPrimary)
            If oObjCrossSection Is oCanRule Then
                bCheckOtherFC = False
                If Not oFC.IsMarkedForDelete Then
                    If bIsPrimaryDeleted Then
                        Set oFC.Definition = Nothing
                    Else
                        oFC.SetCrossSectionObject SPSFCPrimary, Nothing
                    End If
                End If
                'a End or StubEnd Can might be a secondary to VCB or Gap
            ElseIf ePortId = SPSMemberAxisStart Or ePortId = SPSMemberAxisEnd Then
                Set oObjCrossSection = oFC.GetCrossSectionObject(SPSFCSecondary)
                If oObjCrossSection Is oCanRule Then
                    bCheckOtherFC = False
                    If Not oFC.IsMarkedForDelete Then
                        If bIsPrimaryDeleted Then
                            Set oFC.Definition = Nothing
                        Else
                            oFC.SetCrossSectionObject SPSFCSecondary, Nothing
                        End If
                    End If
                End If
            End If
        End If
        
        If bCheckOtherFC Then
            Set oFC = oSecMembSys.FrameConnectionAtEnd(SPSMemberAxisEnd)
            If Not oFC Is Nothing Then
                If Not oFC.IsMarkedForDelete Then
                    Set oObjCrossSection = oFC.GetCrossSectionObject(SPSFCPrimary)
                    If oObjCrossSection Is oCanRule Then
                        If bIsPrimaryDeleted Then
                            Set oFC.Definition = Nothing
                        Else
                            oFC.SetCrossSectionObject SPSFCPrimary, Nothing
                        End If
                        'a End or StubEnd Can might be a secondary to VCB or Gap
                    ElseIf ePortId = SPSMemberAxisStart Or ePortId = SPSMemberAxisEnd Then
                        Set oObjCrossSection = oFC.GetCrossSectionObject(SPSFCSecondary)
                        If oObjCrossSection Is oCanRule Then
                            If bIsPrimaryDeleted Then
                                Set oFC.Definition = Nothing
                            Else
                                oFC.SetCrossSectionObject SPSFCSecondary, Nothing
                            End If
                        End If
                    End If
                End If
            End If
        End If
       
    Next lngIdx
    
    'if this CanRule is at the end of a MemberSystem, that FC will be using it as its supported object.
    If Not bIsPrimaryDeleted Then
        If ePortId = SPSMemberAxisStart Or ePortId = SPSMemberAxisEnd Then
        
            crStatus = oCanRule.GetPrimaryMemberSystem(oPrimaryMemberSystem)
            If Not oPrimaryMemberSystem Is Nothing Then
            
                Set oFC = oPrimaryMemberSystem.FrameConnectionAtEnd(ePortId)
               
                If Not oFC Is Nothing Then
                    If Not oFC.IsMarkedForDelete Then
                        Set oObjCrossSection = oFC.GetCrossSectionObject(SPSFCSupported)
                        If oObjCrossSection Is oCanRule Then
                            oFC.SetCrossSectionObject SPSFCSupported, Nothing
                        End If
                    End If
                End If  ' if MemberSystem was able to get the end FC
            End If  ' if CanRule was able to retrieve the Primary MemberSystem
        End If  ' if CanRule is not inLine
    End If

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub



Public Function GetTubeDiameter(oCanRule As ISPSCanRule, bIsInLine As Boolean, tubeDiameter As Double) As SPSCanRuleStatus



    Dim MyGC As IJGeometricConstruction
    Dim MyGCMacro As IJGeometricConstructionMacro
    Dim diaRule As Long
    Dim dNborStartID As Double, dNborStartOD As Double
    Dim dNborEndID As Double, dNborEndOD As Double

    GetTubeDiameter = SPSCanRule_BadGCMacro_Unexpected

    Set MyGC = oCanRule
    Set MyGCMacro = oCanRule


    diaRule = MyGC.Parameter(attrDiameterRule)

    If diaRule = DiameterRule_User Then
        tubeDiameter = MyGC.Parameter("CanOD")
    Else
        
        Dim dThicknessParameter As Double
        Dim dThickness As Double
        
        FindActualPlateThickness MyGC, "Can", dThickness

        dThicknessParameter = MyGC.Parameter("CanThickness")
        If Abs(dThicknessParameter - dThickness) > dTol Then
            MyGC.Parameter("CanThickness") = dThickness
        End If
        
        
        GetDiameters MyGC, SPSMemberAxisStart, dNborStartID, dNborStartOD
        
        If bIsInLine Then

            GetDiameters MyGC, SPSMemberAxisEnd, dNborEndID, dNborEndOD
            
            If diaRule = DiameterRule_ByOD Then
            
                If dNborStartOD > dNborEndOD Then   ' set diaTube to be the larger one
                    tubeDiameter = dNborStartOD
                Else
                    tubeDiameter = dNborEndOD
                End If
                        
            Else
                
                If dNborStartID > dNborEndID Then   ' set diaTube to be the larger one
                    tubeDiameter = dNborStartID + 2 * dThickness
                Else
                    tubeDiameter = dNborEndID + 2 * dThickness
                End If
        
            End If
        
        Else

            If diaRule = DiameterRule_ByOD Then
                tubeDiameter = dNborStartOD
            Else
                tubeDiameter = dNborStartID + 2 * dThickness
            End If

        End If
            
    
    End If
    
    GetTubeDiameter = SPSCanRule_Ok
    Exit Function

End Function

