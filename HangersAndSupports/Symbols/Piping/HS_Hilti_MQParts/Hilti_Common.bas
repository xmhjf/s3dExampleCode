Attribute VB_Name = "Hilti_Common"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPipeSupport.cls
'   ProgID:         HS_Hilti_MIParts.Hilti_Common
'   Author:         Amlan
'   Creation Date:  05.July.2007
'   Description:
'    AssemblySelectionRule for PipeSupports.
'    This ASR will be used to select Assembly based on SupportDiscipline
'
'   Change History:
'   14/09/11          VSP         TR-CP-193697  Incorrect BOM Description and Warnings were observed in part placement
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
'Create constant to use for error messages
Private Const MODULE = "Hilti_Common"

Public oICH As IJHgrInputConfigHlpr 'The IJHgrInputConfigHlpr object
Public oIOH As IJHgrInputObjectInfo 'The IJHgrInputObjectInfo object
Public oPOC As IJElements           'The PartOccCollection object

Public oPipeCollection As IJElements   'The Pipe Collection
Public oStructCollection As IJElements 'The Structure Collection
Public oPortCollection As IJElements   'The Port Collection
Public oUomServices As UnitsOfMeasureServicesLib.UomVBInterface    'Unit of measure services

Public CatalogPartCollection As New Collection 'The Collection Of Parts To Be Added to the Part Occ Collection
Public ImpliedPartCollection As New Collection 'Collection of Implied parts
Private PartProxy As Object

Private JointCollection As New Collection
Private JointFactory As New HgrJointFactory
Private AssemblyJoint As Object

'The log information for Rigid function.
Type RigidConfigIndex
    lNewPart As Integer
    sNewPartPort As String
    lConnectToPart As Integer
    sConnectToPartPort As String
    dOffsetX As Double
    dOffsetY As Double
    dOffsetZ As Double
    dRotX As Double
    dRotY As Double
    dRotZ As Double
End Type

Public Type hsSteelMember
    ' Names and Descriptions
    sSectionStandard As String
    sSectionType As String
    sSectionName As String
    sSectionDescription As String
    sPartNumber As String
    ' Dimensions and Data
    dUnitWeight As Double
    dDepth As Double
    dWidth As Double
    dWebThickness As Double
    dWebDepth As Double
    dFlangeThickness As Double
    dFlangeWidth As Double
    dCentroidX As Double
    dCentroidY As Double
    ' Back to Back Data
    nB2B_Config As Integer
    dB2B_Spacing As Double
    dB2B_SingleFlangeWidth As Double
    ' HSS
    dHSS_NominalWallThickness As Double
    dHSS_DesignWallThickness As Double
    ' HSSR
    dHSSR_RatioWidthperThickness As Double
    dHSSR_RatioHeightperThickness As Double
    ' HSSC
    dHSSC_OuterDiameter As Double
    dHSSC_RatioDepthPerThickness As Double
    ' Flanged Bolt Gage
    dFB_FlangeGage As Double
    dFB_WebGage As Double
    ' Angle Bolt Gage
    dAB_LongSideGage As Double
    dAB_LongSideGage1 As Double
    dAB_LongSideGage2 As Double
    dAB_ShortSideGage As Double
    dAB_ShortSideGage1 As Double
    dAB_ShortSideGage2 As Double
End Type

Private Type AttributeList
    AttributeName As String
    AttributeValue As Variant
End Type

Dim oAttributeList() As AttributeList

Public Function Hilti_MultipleInterfaceDataLookUp(sDestColName As String, sDestColInterface As String, sPartInterface As String, _
                                                  sRefColName As String, sRefInterfance As String, sRefValue As String, _
                                                  Optional sMaxRefValue As String = "None", Optional bEnforceFailure As Boolean = True) As Variant
    Const METHOD = "Hilti_MultipleInterfaceDataLookup"

    If bEnforceFailure Then
        On Error GoTo ErrorHandler
    Else
        On Error GoTo AllowFailure
    End If

    'Build the complex query
    Dim sComplexQuery As String
    Dim sErrorMessage As String

    If sMaxRefValue = "None" Then
        sComplexQuery = "Select " & sDestColName & " from " & sDestColInterface & " where oid in " _
                      & "(Select oid from " & sPartInterface & " where oid in " _
                      & "(Select oid from " & sRefInterfance & " where " & sRefColName & "= " & sRefValue & "))"
    Else
        sComplexQuery = "Select " & sDestColName & " from " & sDestColInterface & " where oid in " _
                      & "(Select oid from " & sPartInterface & " where oid in " _
                      & "(Select oid from " & sRefInterfance & " where " & sRefColName & "> " & sRefValue & " and " & sRefColName & "< " & sMaxRefValue & "))"
    End If

    sErrorMessage = "The query - " & sComplexQuery & " - Failed."

    Dim vAnswer As Variant
    vAnswer = Hilti_RunDBQuery(sComplexQuery)

    Hilti_MultipleInterfaceDataLookUp = vAnswer

    Exit Function

AllowFailure:
    '    HH.DebugMsg.LogMsg "An error has been allowed to happen in " & MODULE & "." & METHOD
    '    HH.DebugMsg.LogMsg sErrorMessage
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sErrorMessage).Number
End Function

' ---------------------------------------------------------------------------
' Name: GetSectionDataFromDatabase()
' Description: Will get the data for the specified steel section type, standard and name.
' Date - Author: May 17 2006 - JRM
'
' Inputs: secType As String - Type of Section for example "WT"
'         SecStandard As String - Steel standard to be used for Section
'         SecName as String - The name of the Standard for example "L100X100X10
' Outputs: An array of doubles containing the data for the specified section size.
' ---------------------------------------------------------------------------
Public Function Hilti_GetSectionDataFromDatabase(secType As String, SecStandard As String, SecName As String) As Double()
    Const METHOD = "Hilti_GetSectionDataFromDatabase"

    On Error GoTo ErrorHandler

    Dim strCatlogDB As String
    Dim CatalogDef As Object
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim jContext As IJContext
    Dim oConnectMiddle As IJDAccessMiddle
    Dim oCatResMgr As IUnknown
    Dim dAnswer() As Double

    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    strCatlogDB = oDBTypeConfig.get_DataBaseFromDBType("Catalog")
    Set oCatResMgr = oConnectMiddle.GetResourceManager(strCatlogDB)
    Dim m_xService As SP3DStructGenericTools.CrossSectionServices
    Set m_xService = New SP3DStructGenericTools.CrossSectionServices

    'Use the secion type, section standard & section name that you want.
    m_xService.GetStructureCrossSectionDefinition oCatResMgr, _
                                                  SecStandard, secType, _
                                                  SecName, CatalogDef
    Dim Var As Variant
    ReDim dAnswer(4) As Double
    On Error Resume Next

    m_xService.GetCrossSectionAttributeValue CatalogDef, "width", Var
    dAnswer(1) = Var
    m_xService.GetCrossSectionAttributeValue CatalogDef, "tf", Var
    dAnswer(2) = Var
    m_xService.GetCrossSectionAttributeValue CatalogDef, "tw", Var
    dAnswer(3) = Var
    m_xService.GetCrossSectionAttributeValue CatalogDef, "depth", Var
    dAnswer(4) = Var

    Hilti_GetSectionDataFromDatabase = dAnswer

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Function Hilti_GetPartNumber(sPartClass As String, dPipeNomDia As Double) As String

    Const METHOD = "Hilti_GetPartNumber"

    On Error GoTo ErrorHandler

    'Build the complex query
    Dim sComplexQuery As String

    sComplexQuery = "Select PartNumber from JDPart where oid in (Select oid from JCUHAS" & sPartClass _
                  & " where oid in (Select oid from JHgrDiameterSelection where NDFrom <= " & Trim(dPipeNomDia) _
                  & " and NDTo >= " & Trim(dPipeNomDia) & "))"

    Dim vAnswer As Variant
    vAnswer = Hilti_RunDBQuery(sComplexQuery)

    Hilti_GetPartNumber = Trim(vAnswer)

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Function Hilti_RunDBQuery(sComplexQuery As String) As Variant

    Const METHOD = "Hilti_RunDBQuery"

    On Error GoTo ErrorHandler

    Dim oIJPartSelHlpr As IJHgrPartSelectionDBHlpr
    Set oIJPartSelHlpr = New PartSelectionHlpr

    Dim vAnswer() As Variant
    vAnswer = oIJPartSelHlpr.GetValuesFromDBQuery(sComplexQuery)

    Hilti_RunDBQuery = vAnswer(0)

    Set oIJPartSelHlpr = Nothing

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Sub Hilti_InitializeMyHH(ByVal pInputObjectInfo As IJHgrInputConfigHlpr)

    Const METHOD = "Hilti_InitializeMyHH"

    On Error GoTo ErrorHandler
    Initialize pInputObjectInfo
    Exit Sub

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub

Public Function Hilti_DestroyMyHH(ByVal pInputObjectInfo As IJHgrInputConfigHlpr)

    Const METHOD = "Hilti_DestroyMyHH"

    On Error GoTo ErrorHandler

    Class_Terminate

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Function Hilti_BuildBom(ByVal pSupportComp As Object, Optional ByVal sWeb As String) As String
    Const METHOD = "Hilti_BuildBom"

    On Error GoTo ErrorHandler

    Dim sVendor As String
    Dim sGroupName As String
    Dim sPartDesc As String
    Dim sHiltiBomInfo As String
    'Dim dWeight As Double
    Dim sTableName As String
    Dim iLocation As Integer
    Dim sRightString As String
    Dim sLeftString As String
    Dim istringLength As Integer
    Dim oTemp As Object
    Dim sPartNumber As String
    Dim dMaxRodLength As Double
    Dim dRodLength As Double

    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Set oPartOcc = pSupportComp    ' The part occurence
    oPartOcc.GetPart oPart      ' The associated catalog part

    Dim sItemNo As String
    Dim sPartClassName As String

    Dim sStructCon As String
    Dim sInterfaceName As String
    sPartNumber = Trim(oPart.PartNumber)
    sItemNo = GetAttributeFromObject(oPart, "ItemNo")
    sPartClassName = oPart.GetRelatedPartClassName
    sInterfaceName = "JCUHAS" & sPartClassName
    
     Dim pParam_Flange(0 To 0) As Parameters
    'pParam_Flange(0) = FillParameterForDouble("Item_Nb", sItemNo)
    pParam_Flange(0) = FillParameterForBSTR("Item_Nb", sItemNo)

    sHiltiBomInfo = ReadParametricData("JUAHILTIBOMDESC", "where Item_Nb= ?")
    sVendor = GetSParamUsingParametricQuery(sHiltiBomInfo, "Vendor", pParam_Flange)
    sGroupName = GetSParamUsingParametricQuery(sHiltiBomInfo, "Group_name", pParam_Flange)
    sPartDesc = GetSParamUsingParametricQuery(sHiltiBomInfo, "Part_Description", pParam_Flange)

    If Mid(Trim(sInterfaceName), 7, 16) = "MI_Girders" Or sInterfaceName = "JCUHASMI_Girders" Or sInterfaceName = "JCUHASMQ_StrutGALV" Or sInterfaceName = "JCUHASHiltiMQTKGALV" Or Mid(sInterfaceName, 7, 13) = "MQ_SingleChan" Or Mid(sInterfaceName, 7, 13) = "MQ_DoubleChan" Or sInterfaceName = "JCUHASMQ_BackToBackStrutGALV" Or sInterfaceName = "JCUHASMQ_StrutGALV" Or sInterfaceName = "JCUHASMQ_BracketDoubleChGALV" Or sInterfaceName = "JCUHASMQ_BracketDoubleChanSS" Or sInterfaceName = "JCUHASMQ_BracketDoubleChanHDG" Or sInterfaceName = "JCUHASMQ_ThreadedRodMGALV" Or sInterfaceName = "JCUHASMQ_ThreadedRodMHDG" Or sInterfaceName = "JCUHASMQ_ThreadedRodMSS" Or sInterfaceName = "JCUHASMQ_ThreadedRodINGALV" Or sInterfaceName = "JCUHASMQ_ThreadedPipeGALV" Or sInterfaceName = "JCUHASMQ_ThreadedPipeHDG" Or sInterfaceName = "JCUHASMQ_ThreadedPipeSS" Then
        sPartNumber = GetSParamUsingParametricQuery(sHiltiBomInfo, "Part_Name", pParam_Flange)

        If sWeb = "Yes" Then
            sPartNumber = sPartNumber & "-web"
        End If

        'dWeight = GetAttributeFromObject(pSupportComp, "DryWeight")
    Else
        If InStr(sPartNumber, "?") <> 0 Then
            sPartNumber = sPartNumber
        Else
            sPartNumber = GetSParamUsingParametricQuery(sHiltiBomInfo, "Part_Name", pParam_Flange)
        End If

        If sWeb = "Yes" Then
            sPartNumber = sPartNumber & "-web"
        End If
        'dWeight = Hilti_MultipleInterfaceDataLookUp("DryWeight", "JCatalogWtAndCG", sInterfaceName, "PartNumber", "JDPart", "'" & sPartNumber & "'")
    End If
    
    sPartNumber = Replace(sPartNumber, "Hilti ", "")
    
    If InStr(UCase(sPartDesc), "-WEB") = 0 And sWeb = "Yes" Then
        sPartDesc = sPartDesc + " Web"
    End If
    
    sPartDesc = Replace(sPartDesc, "Hilti ", "")


   ' If dWeight < 0.001 Then
        Hilti_BuildBom = sVendor & ", " & sItemNo & ", " & sPartDesc
    'Else
    '    Hilti_BuildBom = sVendor & ", " & sItemNo & ", " & sPartDesc
    'End If

    If sInterfaceName = "JCUHASMQ_TurnbuckleRod" Then

        dMaxRodLength = GetAttributeFromObject(oPart, "Length")
        dRodLength = GetAttributeFromObject(pSupportComp, "Length")

        If dRodLength <= dMaxRodLength + 0.000001 Then
            Hilti_BuildBom = sVendor & ", " & sItemNo & ", " & sPartDesc & ", Rod included with turnbuckle"
        Else
            Hilti_BuildBom = sVendor & ", " & sItemNo & ", " & sPartDesc & ", Rod to be ordered separately"
        End If
    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Function Hilti_StructIsSloped() As Boolean
    Const METHOD = "Hilti_StructIsSloped"

    On Error GoTo ErrorHandler

    Dim dAngleZ As Double

    dAngleZ = Deg(GetPortOrientationAngle(ORIENT_GLOBAL_Z, "Structure", HGRPORT_Z))

    If (dAngleZ > -0.99999 And dAngleZ < 0.99999) Or (dAngleZ > 180 - 0.99999 And dAngleZ < 180 + 0.99999 Or dAngleZ > 90 - 0.99999 And dAngleZ < 90 + 0.99999) Then
        Hilti_StructIsSloped = False
    Else
        Hilti_StructIsSloped = True
    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Function Hilti_PipeIsSloped() As Boolean
    Const METHOD = "Hilti_StructIsSloped"

    On Error GoTo ErrorHandler

    Dim dAngleZ As Double
    Dim dAngleX As Double
    Dim bParallel As Boolean
    Dim dCompareAngle As Double


    dAngleX = Deg(GetPortOrientationAngle(ORIENT_GLOBAL_X, "Route", HGRPORT_X))

    bParallel = Hilti_PipeStructParallel()

    dCompareAngle = dAngleX
    '    End If

    If (dCompareAngle > -0.00001 And dCompareAngle < 0.00001) Or (dCompareAngle > 180 - 0.00001 And dCompareAngle < 180 + 0.00001) Or (dCompareAngle > 90 - 0.00001 And dCompareAngle < 90 + 0.00001) Or (dCompareAngle > 270 - 0.00001 And dCompareAngle < 270 + 0.00001) Then
        Hilti_PipeIsSloped = False
    Else
        Hilti_PipeIsSloped = True
    End If


    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Function Hilti_PipeStructParallel() As Boolean
    Const METHOD = "Hilti_StructIsSloped"

    On Error GoTo ErrorHandler

    Dim dStructAngleX As Double
    Dim dRouteAngleX As Double

    dStructAngleX = Deg(GetPortOrientationAngle(ORIENT_GLOBAL_X, "Structure", HGRPORT_X))
    dRouteAngleX = Deg(GetPortOrientationAngle(ORIENT_GLOBAL_X, "Route", HGRPORT_X))

    If (dStructAngleX > (dRouteAngleX - 0.99999)) And (dStructAngleX < (dRouteAngleX + 0.99999)) Then
        Hilti_PipeStructParallel = True
    Else
        Hilti_PipeStructParallel = False
    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Function Hilti_StructRouteYPar() As Boolean
    Const METHOD = "Hilti_StructIsSloped"

    On Error GoTo ErrorHandler

    Dim dYDifference As Double
    dYDifference = Deg(GetPortOrientationAngle(ORIENT_DIRECT, "Structure", HGRPORT_Y, "Route", HGRPORT_Y))

    If (dYDifference > -0.99999 And dYDifference < 0.99999) Or (dYDifference > 180 - 0.99999 And dYDifference < 180 + 0.99999) Then
        Hilti_StructRouteYPar = True
    Else
        Hilti_StructRouteYPar = False
    End If
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

Public Function Hilti_RouteRunParallelSlopedSlab() As Boolean
    Const METHOD = "Hilti_RouteRunParallelSlopedSlab"

    On Error GoTo ErrorHandler

    Dim sSupportingType As String
    Dim dStructuAngleZ As Double
    Dim dStructuAngleX As Double
    Dim dRouteAngleZ As Double
    Dim dRouteAngelX As Double

    If GetSupportingTypes(1) = "Slab" Then

        dStructuAngleZ = GetPortOrientationAngle(ORIENT_GLOBAL_Z, "Structure", HGRPORT_Z)
        dStructuAngleX = GetPortOrientationAngle(ORIENT_GLOBAL_X, "Structure", HGRPORT_X)
        dRouteAngleZ = GetPortOrientationAngle(ORIENT_GLOBAL_Z, "Route", HGRPORT_Z)
        dRouteAngelX = GetPortOrientationAngle(ORIENT_GLOBAL_X, "Route", HGRPORT_X)

        If (Deg(dStructuAngleZ) > 180 - 0.00001 And Deg(dStructuAngleZ) < 180 + 0.00001) Or (Deg(dStructuAngleZ) > 0 - 0.00001 And Deg(dStructuAngleZ) < 0 + 0.00001) Then
            Hilti_RouteRunParallelSlopedSlab = False
        Else
            If (Deg(dStructuAngleX) > 180 - 0.00001 And Deg(dStructuAngleX) < 180 + 0.00001) Or (Deg(dStructuAngleX) > 0 - 0.00001 And Deg(dStructuAngleX) < 0 + 0.00001) Then
                Hilti_RouteRunParallelSlopedSlab = False
            Else
                If (Deg(dStructuAngleX) > Deg(dRouteAngelX) - 0.00001 And Deg(dStructuAngleX) < Deg(dRouteAngelX) + 0.00001) Or (Deg(dStructuAngleX) > Deg(dRouteAngelX) - 0.00001 And Deg(dStructuAngleX) < Deg(dRouteAngelX) + 0.00001) Then
                    Hilti_RouteRunParallelSlopedSlab = True
                Else
                    Hilti_RouteRunParallelSlopedSlab = False
                End If
            End If
        End If
    Else
        Hilti_RouteRunParallelSlopedSlab = False
    End If

    Exit Function:
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function


Public Sub Hilti_RouteSizeIsValid(ByVal sPartNum As String)
    Const METHOD = "Hilti_RouteSizeIsValid"

    On Error GoTo ErrorHandler

    Dim dPipeND As Double
    Dim sUnit As String
    Dim sHiltiPipeInfo As String
    Dim dMinPipeSize As Double
    Dim dMaxPipeSize As Double

    GetNOMPipeDiaAndUnits 1, dPipeND, sUnit

    Dim pParam_Flange(0 To 0) As Parameters
    pParam_Flange(0) = FillParameterForBSTR("Part", sPartNum)
    sHiltiPipeInfo = ReadParametricData("JUAMQAssyPipePart", "where Part = ?")

    If sUnit = "mm" Then
        dMinPipeSize = GetNParamUsingParametricQuery(sHiltiPipeInfo, "MinPipeSizeM", pParam_Flange) * 1000
        dMaxPipeSize = GetNParamUsingParametricQuery(sHiltiPipeInfo, "MaxPipeSizeM", pParam_Flange) * 1000
    Else
        dMinPipeSize = GetNParamUsingParametricQuery(sHiltiPipeInfo, "MinPipeSizeIn", pParam_Flange) * 39.3700787401575
        dMaxPipeSize = GetNParamUsingParametricQuery(sHiltiPipeInfo, "MaxPipeSizeIn", pParam_Flange) * 39.3700787401575
    End If

    If dPipeND > dMaxPipeSize Or dPipeND < dMinPipeSize - 0.99999 Then
        PF_EventHandler "Pipe size is not an appropriate size for assembly.", Err, MODULE, METHOD, False
    End If
    Exit Sub:

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Sub
Public Function AddHexNut(dWidth1 As Double, dThickness1 As Double, bFlatSide As Boolean, strmatrix As String, outputcoll As IJDOutputCollection, output As String, Optional name As String = "None")
    Const METHOD = "AddHexNut"
    On Error GoTo ErrorHandler

    Dim x As Double
    Dim y As Double
    Dim z As Double

    If bFlatSide = False Then
        y = SinDeg(60) * dWidth1 / 2
        x = CosDeg(60) * dWidth1 / 2
    Else
        y = CosDeg(60) * dWidth1 / 2
        x = SinDeg(60) * dWidth1 / 2
    End If

    z = dThickness1

    If LCase(name) = "none" Then
        name = output
    End If

    Dim oGeomFactory As New GeometryFactory
    Dim matrix As IJDT4x4
    GetMatrixFromString strmatrix, matrix

    Dim dblPoints(0 To 20) As Double

    Dim numPtsLineStr As Integer
    Dim vector(4) As Double

    vector(1) = -dWidth1 / 2
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    dblPoints(0) = vector(1)
    dblPoints(1) = vector(2)
    dblPoints(2) = vector(3)

    vector(1) = -x
    vector(2) = y
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    dblPoints(3) = vector(1)
    dblPoints(4) = vector(2)
    dblPoints(5) = vector(3)

    vector(1) = x
    vector(2) = y
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    dblPoints(6) = vector(1)
    dblPoints(7) = vector(2)
    dblPoints(8) = vector(3)

    vector(1) = dWidth1 / 2
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    dblPoints(9) = vector(1)
    dblPoints(10) = vector(2)
    dblPoints(11) = vector(3)

    vector(1) = x
    vector(2) = -y
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    dblPoints(12) = vector(1)
    dblPoints(13) = vector(2)
    dblPoints(14) = vector(3)

    vector(1) = -x
    vector(2) = -y
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    dblPoints(15) = vector(1)
    dblPoints(16) = vector(2)
    dblPoints(17) = vector(3)

    vector(1) = -dWidth1 / 2
    vector(2) = 0
    vector(3) = 0
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector
    dblPoints(18) = vector(1)
    dblPoints(19) = vector(2)
    dblPoints(20) = vector(3)

    numPtsLineStr = 7

    Dim oLineStr As LineString3d
    Set oLineStr = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, _
                                                             numPtsLineStr, dblPoints)

    vector(1) = -dWidth1 / 2
    vector(2) = 0
    vector(3) = 1
    vector(4) = 1
    MultiplyVectorByMatrix vector, matrix, vector

    Dim oHex As Projection3d
    Set oHex = oGeomFactory.Projections3d.CreateByCurve(outputcoll.ResourceManager, _
                                                        oLineStr, vector(1) - dblPoints(0), vector(2) - dblPoints(1), vector(3) - dblPoints(2), z, True)

    Set oLineStr = Nothing

    outputcoll.AddOutput output, oHex

    Set oHex = Nothing
    Set oGeomFactory = Nothing


    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function


Public Function Hilti_BuildBomForBracket(ByVal pSupportComp As Object, ByVal dCutLen As Double, Optional ByVal sWeb As String) As String
    Const METHOD = "Hilti_BuildBomForBracket"

    On Error GoTo ErrorHandler

    Dim sVendor As String
    Dim sGroupName As String
    Dim sPartDesc As String
    Dim sHiltiBomInfo As String
    'Dim dWeight As Double
    Dim sTableName As String
    Dim iLocation As Integer
    Dim sRightString As String
    Dim sLeftString As String
    Dim istringLength As Integer
    Dim oTemp As Object
    Dim sPartNumber As String
    Dim dMaxRodLength As Double
    Dim dRodLength As Double

    Dim oPartOcc As PartOcc
    Dim oPart As IJDPart
    Set oPartOcc = pSupportComp    ' The part occurence
    oPartOcc.GetPart oPart      ' The associated catalog part

    Dim sItemNo As String
    Dim sPartClassName As String

    Dim sStructCon As String
    Dim sInterfaceName As String
    sPartNumber = Trim(oPart.PartNumber)
    sItemNo = GetAttributeFromObject(oPart, "ItemNo")
    sPartClassName = oPart.GetRelatedPartClassName
    sInterfaceName = "JCUHAS" & sPartClassName
    
    Dim oUomServices As UnitsOfMeasureServicesLib.UomVBInterface
    Set oUomServices = New UnitsOfMeasureServicesLib.UomVBInterface
    
    Dim dCutLength As Double
    Dim sCutlen As String
    
    Dim xomFormat As IJUomVBFormat
    Set xomFormat = New UomVBFormat
    xomFormat.PrecisionType = PRECISIONTYPE_DECIMAL
    xomFormat.FractionalPrecision = 2
    xomFormat.UnitsDisplayed = True
    xomFormat.ReduceFraction = True
                
    oUomServices.FormatUnit UNIT_DISTANCE, dCutLen, sCutlen, xomFormat, DISTANCE_MILLIMETER
    
    Dim pParam_Flange(0 To 0) As Parameters
    'pParam_Flange(0) = FillParameterForDouble("Item_Nb", sItemNo)
    pParam_Flange(0) = FillParameterForBSTR("Item_Nb", sItemNo)

    sHiltiBomInfo = ReadParametricData("JUAHILTIBOMDESC", "where Item_Nb= ?")
    sVendor = GetSParamUsingParametricQuery(sHiltiBomInfo, "Vendor", pParam_Flange)
    sGroupName = GetSParamUsingParametricQuery(sHiltiBomInfo, "Group_name", pParam_Flange)
    sPartDesc = GetSParamUsingParametricQuery(sHiltiBomInfo, "Part_Description", pParam_Flange)

    If Mid(Trim(sInterfaceName), 7, 16) = "MI_Girders" Or sInterfaceName = "JCUHASMI_Girders" Or sInterfaceName = "JCUHASMQ_StrutGALV" Or sInterfaceName = "JCUHASHiltiMQTKGALV" Or Mid(sInterfaceName, 7, 13) = "MQ_SingleChan" Or Mid(sInterfaceName, 7, 13) = "MQ_DoubleChan" Or sInterfaceName = "JCUHASMQ_BackToBackStrutGALV" Or sInterfaceName = "JCUHASMQ_StrutGALV" Or sInterfaceName = "JCUHASMQ_BracketDoubleChGALV" Or sInterfaceName = "JCUHASMQ_BracketDoubleChanSS" Or sInterfaceName = "JCUHASMQ_BracketDoubleChanHDG" Or sInterfaceName = "JCUHASMQ_ThreadedRodMGALV" Or sInterfaceName = "JCUHASMQ_ThreadedRodMHDG" Or sInterfaceName = "JCUHASMQ_ThreadedRodMSS" Or sInterfaceName = "JCUHASMQ_ThreadedRodINGALV" Or sInterfaceName = "JCUHASMQ_ThreadedPipeGALV" Or sInterfaceName = "JCUHASMQ_ThreadedPipeHDG" Or sInterfaceName = "JCUHASMQ_ThreadedPipeSS" Then
        sPartNumber = GetSParamUsingParametricQuery(sHiltiBomInfo, "Part_Name", pParam_Flange)

        If sWeb = "Yes" Then
            sPartNumber = sPartNumber & "-web"
        End If

        'dWeight = GetAttributeFromObject(pSupportComp, "DryWeight")
    Else
        If InStr(sPartNumber, "?") <> 0 Then
            sPartNumber = sPartNumber
        Else
            sPartNumber = GetSParamUsingParametricQuery(sHiltiBomInfo, "Part_Name", pParam_Flange)
        End If

        If sWeb = "Yes" Then
            sPartNumber = sPartNumber & "-web"
        End If
        'dWeight = Hilti_MultipleInterfaceDataLookUp("DryWeight", "JCatalogWtAndCG", sInterfaceName, "PartNumber", "JDPart", "'" & sPartNumber & "'")
    End If
    
    sPartNumber = Replace(sPartNumber, "Hilti ", "")
    
    If InStr(UCase(sPartDesc), "-WEB") = 0 And sWeb = "Yes" Then
        sPartDesc = sPartDesc + " Web"
    End If
    
    sPartDesc = Replace(sPartDesc, "Hilti ", "")
    sPartDesc = sPartDesc & "; Cut Length = " & sCutlen
    
    'If dWeight < 0.01 Then
        Hilti_BuildBomForBracket = sVendor & ", " & sItemNo & ", " & sPartDesc
    'Else
    '    Hilti_BuildBomForBracket = sVendor & ", " & sItemNo & ", " & sPartDesc
    'End If

    If sInterfaceName = "JCUHASMQ_TurnbuckleRod" Then

        dMaxRodLength = GetAttributeFromObject(oPart, "Length")
        dRodLength = GetAttributeFromObject(pSupportComp, "Length")

        If dRodLength <= dMaxRodLength + 0.000001 Then
            Hilti_BuildBomForBracket = sVendor & ", " & sItemNo & ", " & sPartDesc & ", Rod included with turnbuckle"
        Else
            Hilti_BuildBomForBracket = sVendor & ", " & sItemNo & ", " & sPartDesc & ", Rod to be ordered separately"
        End If
    End If

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function
Public Function GetPortOrientationAngle(eOrientType As HgrPortOrientType, sPort1Name As String, eAxis1 As HgrPortAixsType, Optional sPort2Name As String = "NONE", Optional eAxis2 As HgrPortAixsType) As Double
Const METHOD = "GetPortOrientationAngle"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler

    If UCase(sPort2Name) = "NONE" Then
        GetPortOrientationAngle = oICH.GetPortOrientAngle(eOrientType, sPort1Name, eAxis1)
    Else
        GetPortOrientationAngle = oICH.GetPortOrientAngle(eOrientType, sPort1Name, eAxis1, sPort2Name, eAxis2)
    End If

    Exit Function

ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function

' ---------------------------------------------------------------------------
' Method:   GetSupportingTypes
' Desc:     Return the type of structure being connected to.
' Param:
'           lStructureIndex As Long
'               Index of the supporting structure
' Return:   String representation of the supporting type
' ---------------------------------------------------------------------------
Public Function GetSupportingTypes(lStructureIndex As Long) As String
Const METHOD = "GetSupportingTypes"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler

'Steel - ISPSMemberPartPrismatic
'Wall - ISPSWallPart
'Slab - ISPSSlabEntity
'Equipment - IJSmartEquipment
'Designed Equipment Shape - IJShape

'What other objects can supports have for supporting objects???
'       - ISPSMemberPartCurve
'       - ISPSMemberPartLinear

    Dim oSupportingColl As Object
    Set oSupportingColl = oIOH.GetSupportingObjects
    
    If Not oSupportingColl Is Nothing Then
        Select Case TypeName(oSupportingColl.Item(lStructureIndex))
        Case "ISPSMemberPartPrismatic"
            GetSupportingTypes = "Steel"
        Case "ISPSWallPart"
            GetSupportingTypes = "Wall"
        Case "ISPSSlabEntity"
            GetSupportingTypes = "Slab"
        Case "IJSmartEquipment"
            GetSupportingTypes = "Equipment"
        Case "IJShape"
            GetSupportingTypes = "Shape"
        Case "IJHgrConnComponent"
            GetSupportingTypes = "HgrBeam"
        Case Else
            GetSupportingTypes = "Unknown"
        End Select
    Else
        'Place-By-Reference or No Supporting Object
        GetSupportingTypes = "None"
    End If

    Exit Function
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function

' ---------------------------------------------------------------------------
' Method:   GetSupportingSectionData
' Desc:     Looks up all the relevent data for the steel cross section of the
'           supporting object.
' Param:
'           iIndex As Integer
'               The index of the supporting object you want to look up the cross
'               sectional data of.
' Return:   The cross section information in the form of a hsSteelMember user defined
'           type.
' Common Use:
'           Used in the AIR to get any required information about a supporting steel object.
'               Dim tStructure as hsSteelMember
'               tStructure = hsHlpr.GetSupportingSectionData(1)
'               dFlangeThickness = tStructure.dFlangeThickness
' ---------------------------------------------------------------------------
Public Function GetSupportingSectionData(iIndex As Integer) As hsSteelMember
Const METHOD = "GetSupportingSectionData"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler
    
    Dim tResultMember As hsSteelMember
    
    Dim StructColl As Object
    Dim oStructObject As Object
    Set StructColl = oIOH.GetSupportingObjects
    Set oStructObject = StructColl.Item(iIndex)

    If GetSupportingTypes(CLng(iIndex)) = "Steel" Then
        Dim oMember As ISPSMemberPartPrismatic
        Dim oSPSCrossSection As ISPSCrossSection
        Set oMember = oStructObject
        Set oSPSCrossSection = oMember.CrossSection
        
        Dim sSectionName As String
        Dim sSectionType As String
        Dim sSectionStandard As String
    
        sSectionName = oSPSCrossSection.SectionName
        sSectionType = oSPSCrossSection.SectionType
        sSectionStandard = oSPSCrossSection.SectionStandard
    
        tResultMember = GetSectionDataFromSection(sSectionStandard, sSectionType, sSectionName)
    ElseIf GetSupportingTypes(CLng(iIndex)) = "HgrBeam" Then
        Dim oPartOcc As IJPartOcc
        Set oPartOcc = oStructObject
        Dim oPart As IJDPart
        oPartOcc.GetPart oPart
        
        tResultMember = GetSectionDataFromPart(oPart.PartNumber)

    End If

    GetSupportingSectionData = tResultMember

    Exit Function

ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function
Public Sub Initialize(ICH As IJHgrInputConfigHlpr)
Const METHOD = "Initialize"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler

    If Not oPipeCollection Is Nothing Then
        oPipeCollection.Clear
        Set oPipeCollection = Nothing
    End If
    If Not oStructCollection Is Nothing Then
        oStructCollection.Clear
        Set oStructCollection = Nothing
    End If
    If Not oPortCollection Is Nothing Then
        oPortCollection.Clear
        Set oPortCollection = Nothing
    End If
    
    Set oIOH = Nothing
    Set JointCollection = Nothing
    Set CatalogPartCollection = Nothing
    Set ImpliedPartCollection = Nothing
    
    Set oICH = Nothing
    Set oUomServices = Nothing
    
    Set JointCollection = New Collection
    ReDim CILog(1) As RigidConfigIndex
    CILog(1).lNewPart = -1
    CILog(1).sNewPartPort = "Route"
    CILog(1).lConnectToPart = -1
    CILog(1).sConnectToPartPort = ""
    CILog(1).dOffsetX = 0
    CILog(1).dOffsetY = 0
    CILog(1).dOffsetZ = 0
    CILog(1).dRotX = 0
    CILog(1).dRotY = 0
    CILog(1).dRotZ = 0

    ReDim PartClasses(0) As String
    ReDim oAttributeList(0) As AttributeList
    
    'Initialize the Input Config Helper and the Input Object Helper
    Set oICH = ICH
    Set oIOH = ICH

    'Initialize the Route Objects we need
    oICH.GetPipes oPipeCollection
    oICH.GetSupportingCollection oStructCollection
   
   'Initialize the Units of Measure Services incase we need it
    Set oUomServices = New UnitsOfMeasureServicesLib.UomVBInterface
    
    'Initialise the Port collection
    Set oPortCollection = oICH.GetRefPorts

    Exit Sub

ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
    Case 0
        sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Sub


' ---------------------------------------------------------------------------
' Method:   Class_Terminate
' Desc:     Use this method to properly dispose of the hsHlpr object
'               - destroys Input Config Helper and Input Object Info
'               - destroys Part Occurence Collection
'               - destroys Port Collection
'               - destroys Pipe Collection
'               - destroys Units of Measure Service
' Param:
' Return:
' Common Use:
'               - Call this function when you have completed the placement of the assembly
' ---------------------------------------------------------------------------
Public Sub Class_Terminate()
Const METHOD = "Class_Terminate"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler

    Set oICH = Nothing
    Set oIOH = Nothing
    Set oUomServices = Nothing
    
    Set oPOC = Nothing
    
    If Not oPipeCollection Is Nothing Then
        oPipeCollection.Clear
        Set oPipeCollection = Nothing
    End If
    If Not oStructCollection Is Nothing Then
        oStructCollection.Clear
        Set oStructCollection = Nothing
    End If
    
    Set oPortCollection = Nothing
    Set CatalogPartCollection = Nothing
    Set ImpliedPartCollection = Nothing
    Set JointCollection = Nothing
    
    Exit Sub

ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
    Case 0
        sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Sub

' ---------------------------------------------------------------------------
' Method:   GetSectionDataFromSection
' Desc:     Using the supplied section standard, type and name, looks up all the
'           relevent data for the steel cross section.
' Param:
'           sSectionStandard as String
'               The Standard of the Steel Section
'           sSectionType as String
'               The Type of the Steel Section
'           sSectionName as String
'               The Name of the Steel Section
' Return:   The cross section information in the form of a hsSteelMember user defined
'           type.
' Common Use:
'           Used in the AIR to get any required information about a steel section.
'               Dim tSteel as hsSteelMember
'               tSteel = hsHlpr.GetSectionDataFromSection("AISC-LRFD-3.1", "W", "W4X13")
'               dFlangeThickness = tSteel.dFlangeThickness
' ---------------------------------------------------------------------------
Public Function GetSectionDataFromSection(sSectionStandard As String, sSectionType As String, sSectionName As String) As hsSteelMember
Const METHOD = "GetSectionDataFromSection"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler

    Dim tResultMember As hsSteelMember

    tResultMember.sPartNumber = ""
    tResultMember.sSectionStandard = sSectionStandard
    tResultMember.sSectionType = sSectionType
    tResultMember.sSectionName = sSectionName
    
    Dim strCatlogDB As String
    Dim CatalogDef As Object
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim jContext As IJContext
    Dim oConnectMiddle As IJDAccessMiddle
    Dim oCatResMgr As IUnknown
    
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    strCatlogDB = oDBTypeConfig.get_DataBaseFromDBType("Catalog")

    Set oCatResMgr = oConnectMiddle.GetResourceManager(strCatlogDB)
    Dim m_xService As SP3DStructGenericTools.CrossSectionServices
    Set m_xService = New SP3DStructGenericTools.CrossSectionServices

    m_xService.GetStructureCrossSectionDefinition oCatResMgr, _
                                                  sSectionStandard, sSectionType, _
                                                  sSectionName, CatalogDef
    Dim Var As Variant
    On Error Resume Next
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "Depth", Var
    tResultMember.dDepth = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "Width", Var
    tResultMember.dWidth = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "tw", Var
    tResultMember.dWebThickness = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "tf", Var
    tResultMember.dFlangeThickness = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "UnitWeight", Var
    tResultMember.dUnitWeight = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "d", Var
    tResultMember.dWebDepth = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "bf", Var
    tResultMember.dFlangeWidth = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "CentroidX", Var
    tResultMember.dCentroidX = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "CentroidY", Var
    tResultMember.dCentroidY = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "bbConfiguration", Var
    tResultMember.nB2B_Config = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "bb", Var
    tResultMember.dB2B_Spacing = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "b", Var
    tResultMember.dB2B_SingleFlangeWidth = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "tnom", Var
    tResultMember.dHSS_NominalWallThickness = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "tdes", Var
    tResultMember.dHSS_DesignWallThickness = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "b_t", Var
    tResultMember.dHSSR_RatioWidthperThickness = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "h_t", Var
    tResultMember.dHSSR_RatioHeightperThickness = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "OD", Var
    tResultMember.dHSSC_OuterDiameter = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "D_t", Var
    tResultMember.dHSSC_RatioDepthPerThickness = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "gf", Var
    tResultMember.dFB_FlangeGage = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "gw", Var
    tResultMember.dFB_WebGage = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "lsg", Var
    tResultMember.dAB_LongSideGage = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "lsg1", Var
    tResultMember.dAB_LongSideGage1 = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "lsg2", Var
    tResultMember.dAB_LongSideGage2 = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "ssg", Var
    tResultMember.dAB_ShortSideGage = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "ssg1", Var
    tResultMember.dAB_ShortSideGage1 = Var
    Set Var = Empty
    m_xService.GetCrossSectionAttributeValue CatalogDef, "ssg2", Var
    tResultMember.dAB_ShortSideGage2 = Var
    
    GetSectionDataFromSection = tResultMember

    Set jContext = Nothing
    Set oDBTypeConfig = Nothing
    Set oConnectMiddle = Nothing
    Set oCatResMgr = Nothing
    Set m_xService = Nothing
    
    Exit Function

ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function

' ---------------------------------------------------------------------------
' Method:   GetSectionDataFromPart
' Desc:     Using the supplied part number, looks up all the relevent data for the
'           steel cross section..
' Param:
'           sPartNumber as String
'               The part number of a steel member, either a Hanger Beam or a Smart
'               Steel part.
' Return:   The cross section information in the form of a hsSteelMember user defined
'           type.
' Common Use:
'           Used in the AIR to get any required information about a steel part.
'               Dim tSteel as hsSteelMember
'               tSteel = hsHlpr.GetSectionDataFromPart("HsSteel_L2X2X1/8")
'               dFlangeThickness = tSteel.dFlangeThickness
' ---------------------------------------------------------------------------
Public Function GetSectionDataFromPart(ByRef sPartNumber As String) As hsSteelMember
Const METHOD = "GetSectionDataFromPart"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler

    Dim tResultMember As hsSteelMember

    If sPartNumber = "" Then
        tResultMember.sPartNumber = ""
        GetSectionDataFromPart = tResultMember
        Exit Function
    End If

    Dim sQueryString As String
    Dim sSectionName As String
    Dim sSectionType As String
    Dim sSectionStandard As String
    Dim sSectionDescription As String

    sQueryString = "select SectionName from REFDATCrossSection where oid in (select oidTarget from CORERelationOrigin where oid in(select oid from JDPart where PartNumber = '" & sPartNumber & "'))"
    sSectionName = RunDBQuery(sQueryString)

    sQueryString = "select PartClassName from REFDATCrossSection where oid in (select oidTarget from CORERelationOrigin where oid in(select oid from JDPart where PartNumber = '" & sPartNumber & "'))"
    sSectionType = RunDBQuery(sQueryString)

    sQueryString = "select ReferenceStandardName from REFDATCrossSection where oid in (select oidTarget from CORERelationOrigin where oid in(select oid from JDPart where PartNumber = '" & sPartNumber & "'))"
    sSectionStandard = RunDBQuery(sQueryString)

    sQueryString = "select PartDescription from JDPart where PartNumber = '" & sPartNumber & "'"
    sSectionDescription = RunDBQuery(sQueryString)

    tResultMember = GetSectionDataFromSection(sSectionStandard, sSectionType, sSectionName)
    tResultMember.sPartNumber = sPartNumber
    tResultMember.sSectionDescription = sSectionDescription

    GetSectionDataFromPart = tResultMember

    Exit Function

ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function

Public Function RunDBQuery(sComplexQuery As String) As Variant
Const METHOD = "RunDBQuery"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler

    Dim oIJPartSelHlpr As IJHgrPartSelectionDBHlpr
    Set oIJPartSelHlpr = New PartSelectionHlpr

    Dim vAnswer() As Variant
    vAnswer = oIJPartSelHlpr.GetValuesFromDBQuery(sComplexQuery)
    
    'Shin May03-2012: Fix crash if result was null using OracleDB
    'RunDBQuery = vAnswer(0)
    If VarType(vAnswer(0)) = vbNull Then
        RunDBQuery = Empty
    Else
        RunDBQuery = vAnswer(0)
    End If

    Set oIJPartSelHlpr = Nothing

    Exit Function

ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function

' ---------------------------------------------------------------------------
' Method:   GetNOMPipeDiaAndUnits
' Desc:     Gets the nominal pipe diameter of the pipe in the default units as defined
'           by the pipe spec. Also get those units
'
' Param:
'           iIndex As Integer
'               Index of the route object for which the diameter will be determined.
'           ByRef dPipeNomDia As Double
'               Variable that is set to the nominal diameter of the pipe
'           ByRef sPipeDiaUnits As String
'               Variable that is set to the units of the pipe spec.
' Returns:
' Common Use:
'           hsHlpr.GetNOMPipeDiaAndUnits 1, dPipeND, sPipeUnits
' ---------------------------------------------------------------------------
Public Sub GetNOMPipeDiaAndUnits(iIndex As Integer, ByRef dPipeNomDia As Double, ByRef sPipeDiaUnits As String)
Const METHOD = "GetNOMPipeDiaAndUnits"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler

    Dim oRoute As Object
    Set oRoute = oPipeCollection.Item(iIndex)

    dPipeNomDia = oICH.GetPipeDiameter(oRoute, sPipeDiaUnits)
    
    Set oRoute = Nothing

    Exit Sub
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Sub

' ---------------------------------------------------------------------------
' Method:   GetPipeDiameterAndInsulatTh()
' Desc:     Gets the diameter of a pipe and the insulation thickness in meters.
'
' Param:
'           iIndex As Integer
'               Index of the route object for which the diameter will be determined.
'           ByRef dPipeDia As Double
'               Variable that is set to the diameter of the pipe
'           ByRef dInsulationThickness As Double
'               Variable that is set to the thickness of the insulation
' Returns:
' Common Use:
'           hsHlpr.GetPipeDiameterAndInsulatTh 1, dPipeOD, dInsulationTh
' ---------------------------------------------------------------------------
Public Sub GetPipeDiameterAndInsulatTh(iIndex As Integer, ByRef dPipeDia As Double, ByRef dInsulationThickness As Double)
Const METHOD = "GetPipeDiameterAndInsulatTh"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler
    
    Dim oRoute As Object
    Set oRoute = oPipeCollection.Item(iIndex)

    Dim dTotalDia As Double
    dTotalDia = GetExternalPipeDiameter(oICH, iIndex)
    
    Dim oInsulationObject As IJRteInsulation
    Set oInsulationObject = oRoute
    
    dInsulationThickness = oInsulationObject.Thickness
    
    dPipeDia = dTotalDia - 2 * dInsulationThickness
    
    Set oInsulationObject = Nothing
    Set oRoute = Nothing
    
    Exit Sub
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Sub

'*************************************************************************************
' Method:   GetExternalPipeDiameter
'
' Desc:    This method retrieved external diameter of pipe.


' Param:
'           Output: diameter of pipe
'
'*************************************************************************************
Public Function GetExternalPipeDiameter(ByVal pDispInputConfigHlpr As Object, ByVal RouteIndex As Integer) As Double
    
    Const METHOD = "GetExternalPipeDiameter:"
    On Error GoTo ErrorHandler
    
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Dim my_IJHgrInputObjectInfo As IJHgrInputObjectInfo
    Set my_IJHgrInputConfigHlpr = pDispInputConfigHlpr
    Set my_IJHgrInputObjectInfo = pDispInputConfigHlpr
      
    Dim lInsulPurpose           As Long
    Dim lInsulMat               As Long
    Dim dInsulThk               As Double
    Dim bInsulFlag              As Boolean
    Dim dblPipeDia              As Double
    Dim oPipes                  As Object
    Dim pPipeColl               As IJElements
    
    my_IJHgrInputObjectInfo.GetInsulationData RouteIndex, lInsulPurpose, lInsulMat, dInsulThk, bInsulFlag
    my_IJHgrInputConfigHlpr.GetPipes oPipes
    Set pPipeColl = oPipes
    
    dblPipeDia = my_IJHgrInputConfigHlpr.GetExternalPipeDiameter(pPipeColl.Item(RouteIndex))
    If bInsulFlag <> False Then
        dblPipeDia = dblPipeDia + 2# * dInsulThk
    End If

    GetExternalPipeDiameter = dblPipeDia
    
    Set my_IJHgrInputConfigHlpr = Nothing
    Set my_IJHgrInputObjectInfo = Nothing
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

