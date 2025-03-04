Attribute VB_Name = "HiltiAssy_Common"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPipeSupport.cls
'   ProgID:         HS_Hilti_MQ.HiltiAssy_Common
'   Author:         Chethan
'   Creation Date:  14.November.2014
'   Description:
'    AssemblySelectionRule for PipeSupports.
'    This ASR will be used to select Assembly based on SupportDiscipline
'
'   Change History:
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
'Create constant to use for error messages
Private Const MODULE = "HiltiAssy_Common"

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

'Define attribute list
Private Type AttributeList
    AttributeName As String
    AttributeValue As Variant
End Type

Dim oAttributeList() As AttributeList

Private Type POCDelayInputColl
    oAttributeList() As AttributeList
End Type

Dim POCDelayInputCollections() As POCDelayInputColl

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
Private CILog() As RigidConfigIndex

'Used in the AddPart function to get the Rule Type
Enum RuleType
HgrPartSelectionRule = 1
HgrSupportRule = 2
None = 0
End Enum

'Used in the GetAttr function to flag whether we want the long or short desription of a codelist
Enum GetCodelistValue
    LongDesc = 1
    ShortDesc = 2
End Enum

'Data type to hold steel cross section attributes
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

Private JointCollection As New Collection
Private JointFactory As New HgrJointFactory
Private AssemblyJoint As Object
Private oPartClasses As Object
Private PartClasses() As String




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
'    DebugMsg.LogMsg "An error has been allowed to happen in " & MODULE & "." & METHOD
'    DebugMsg.LogMsg sErrorMessage
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
    Dim dWeight As Double
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
    Set oPartOcc = pSupportComp ' The part occurence
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

    If Mid(Trim(sInterfaceName), 7, 16) = "HiltiMIGIRDERS" Or sInterfaceName = "JCUHASGirders" Or sInterfaceName = "JCUHASStrut12And14Gauge" Or sInterfaceName = "JCUHASHiltiMQTKGALV" Or Mid(sInterfaceName, 7, 10) = "SingleChan" Or Mid(sInterfaceName, 7, 10) = "DoubleChan" Or sInterfaceName = "JCUHASBackToBackStrut12Ga" Or sInterfaceName = "JCUHASStrut12And14Ga" Or sInterfaceName = "JCUHASBracketDoubleChanGALV" Or sInterfaceName = "JCUHASBracketDoubleChanSS" Or sInterfaceName = "JCUHASBracketDoubleChanHDG" Or sInterfaceName = "JCUHASThreadedRodMGALV" Or sInterfaceName = "JCUHASThreadedRodMHDG" Or sInterfaceName = "JCUHASThreadedRodMSS" Or sInterfaceName = "JCUHASThreadedRodINGALV" Or sInterfaceName = "JCUHASThreadedPipeGALV" Or sInterfaceName = "JCUHASThreadedPipeHDG" Or sInterfaceName = "JCUHASThreadedPipeSS" Then
        sPartNumber = GetSParamUsingParametricQuery(sHiltiBomInfo, "Part_Name", pParam_Flange)
        
        If sWeb = "Yes" Then
            sPartNumber = sPartNumber & "-web"
        End If
        
        dWeight = GetAttributeFromObject(pSupportComp, "DryWeight")
    Else
        If InStr(sPartNumber, "?") <> 0 Then
            sPartNumber = sPartNumber
        Else
            sPartNumber = GetSParamUsingParametricQuery(sHiltiBomInfo, "Part_Name", pParam_Flange)
        End If
        
        If sWeb = "Yes" Then
            sPartNumber = sPartNumber & "-web"
        End If
        
        dWeight = Hilti_MultipleInterfaceDataLookUp("DryWeight", "JCatalogWtAndCG", sInterfaceName, "PartNumber", "JDPart", "'" & sPartNumber & "'")
    End If

    If dWeight < 0.01 Then
        Hilti_BuildBom = sVendor & ", " & sItemNo & ", " & sGroupName & ", " & sPartNumber & ", " & sPartDesc
    Else
        Hilti_BuildBom = sVendor & ", " & sItemNo & ", " & sGroupName & ", " & sPartNumber & ", " & sPartDesc & ", " & ConvertValueToLongStringValue(pSupportComp, "IJUAHILTIBOMDESC", "Part_Weight", dWeight)
    End If

    If sInterfaceName = "JCUHASTurnbuckleRod" Then

        dMaxRodLength = GetAttributeFromObject(oPart, "Length")
        dRodLength = GetAttributeFromObject(pSupportComp, "Length")
        
        If dRodLength <= dMaxRodLength + 0.000001 Then
            Hilti_BuildBom = sVendor & ", " & sItemNo & ", " & sGroupName & ", " & sPartNumber & ", " & sPartDesc & ", " & ConvertValueToLongStringValue(pSupportComp, "IJUAHILTIBOMDESC", "Part_Weight", dWeight) & " (Rod included with turnbuckle)"
        Else
            Hilti_BuildBom = sVendor & ", " & sItemNo & ", " & sGroupName & ", " & sPartNumber & ", " & sPartDesc & ", " & ConvertValueToLongStringValue(pSupportComp, "IJUAHILTIBOMDESC", "Part_Weight", dWeight) & " (Rod to be ordered separately)"
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
    
    'dAngleZ = Deg(GetPortOrientationAngle(ORIENT_GLOBAL_Z, "Route", HGRPORT_Z))
    dAngleX = Deg(GetPortOrientationAngle(ORIENT_GLOBAL_X, "Route", HGRPORT_X))
    
    bParallel = Hilti_PipeStructParallel()
    
'    If bParallel = True Then
'        dCompareAngle = dAngleZ
'    Else
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
        
        If Hilti_GetSupportingTypes(1) = "Slab" Then
            
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
Public Function Hilti_GetSupportingTypes(lStructureIndex As Long) As String
Const METHOD = "Hilti_GetSupportingTypes"
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
            Hilti_GetSupportingTypes = "Steel"
        Case "ISPSWallPart"
            Hilti_GetSupportingTypes = "Wall"
        Case "ISPSSlabEntity"
            Hilti_GetSupportingTypes = "Slab"
        Case "IJSmartEquipment"
            Hilti_GetSupportingTypes = "Equipment"
        Case "IJShape"
            Hilti_GetSupportingTypes = "Shape"
        Case "IJHgrConnComponent"
            Hilti_GetSupportingTypes = "HgrBeam"
        Case Else
            Hilti_GetSupportingTypes = "Unknown"
        End Select
    Else
        'Place-By-Reference or No Supporting Object
        Hilti_GetSupportingTypes = "None"
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
Public Function AddPart(spart As String, Optional sRule As String = "", Optional oObjectForRule As Object = Nothing, Optional ReferenceIndex As Integer = 1, Optional eRuleType As RuleType = HgrPartSelectionRule) As Integer
Const METHOD = "AddPart"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler

    Dim sErr As String
    sErr = ""
    If oICH Is Nothing Then GoTo ErrorHandler
    
    oICH.RouteIndex = ReferenceIndex

    If eRuleType = HgrSupportRule Then
        Dim vRuleResults1() As Variant
        sErr = "May be a problem with HgrRule: " & sRule
        oICH.SetReferenceInput (spart)
        vRuleResults1 = oICH.GetDataByRule(sRule, oObjectForRule)

        If IsEmpty(vRuleResults1) = False And vRuleResults1(0) <> "" Then
            sErr = "May be a problem with Part: " & CStr(vRuleResults1(0))
            Set PartProxy = oICH.GetPartProxyFromName(CStr(vRuleResults1(0)))
            CatalogPartCollection.Add PartProxy
            AddPart = CatalogPartCollection.Count
        Else
            AddPart = 0
        End If
    Else
        If spart <> "" Then
            If sRule = "" Then
                sErr = "May be a problem with Part: " & spart
                Set PartProxy = oICH.GetPartProxyFromName(spart)
                CatalogPartCollection.Add PartProxy
                AddPart = CatalogPartCollection.Count
            Else
                sErr = "May be a problem with Part: " & spart & " or PartSelectionRule: " & sRule
                Set PartProxy = oICH.GetPartProxyFromName(spart, sRule)
                CatalogPartCollection.Add PartProxy
                AddPart = CatalogPartCollection.Count
            End If
        ElseIf sRule <> "" Then
            'Just a Rule is specifed
            Dim vRuleResults2() As Variant
            sErr = "May be a problem with HgrRule: " & sRule
            vRuleResults2 = oICH.GetDataByRule(sRule, oObjectForRule)
            If IsEmpty(vRuleResults2) = False And vRuleResults2(0) <> "" Then
                sErr = "May be a problem with Part: " & CStr(vRuleResults2(0))
                Set PartProxy = oICH.GetPartProxyFromName(CStr(vRuleResults2(0)))
                CatalogPartCollection.Add PartProxy
                AddPart = CatalogPartCollection.Count
            Else
                AddPart = 0
            End If
        Else
            'No Part Specified
            AddPart = 0
        End If
    End If
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function
Public Sub InitializeJoints(pDispPartOccCollection As Object)
Const METHOD = "InitializeJoints"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler
    Set oPOC = Nothing
    Set oPOC = pDispPartOccCollection
    
    On Error Resume Next
    If oPOC.Count > 0 Then
        ReDim POCDelayInputCollections((oPOC.Count)) As POCDelayInputColl
    End If
    
    Exit Sub

ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
    Case 0
        sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Sub
Public Function GetPartAttribute(iPartIndex As Integer, sAttributeName As String, Optional sValueType As String, Optional bEnforceFailure As Boolean = True, Optional bGetAttributeFromCatalog As Boolean = False) As Variant
Const METHOD = "GetPartAttribute"
Dim iErrCodeToUse As Integer
    If bEnforceFailure Then
        On Error GoTo ErrorHandler
    Else
        On Error GoTo AllowFailure
    End If

    Dim vValue As Variant
    Dim oPartOcc As IJPartOcc
    Dim oPart As IJDPart
    
    Set oPartOcc = oPOC.Item(iPartIndex)
    oPartOcc.GetPart oPart
    
    iErrCodeToUse = 1
    
    If bGetAttributeFromCatalog Then
        oICH.GetAttributeValue sAttributeName, oPart, vValue
    Else
        'First try to get the attribute from the Occurence. If it does not
        'exist on the occurence, then grab it from the catalog part.
        On Error GoTo GetAttributeFromCatalog
        oICH.GetAttributeValue sAttributeName, oPartOcc, vValue
        GoTo GotAttribute
GetAttributeFromCatalog:
        If bEnforceFailure Then
            On Error GoTo ErrorHandler
        Else
            On Error GoTo AllowFailure
        End If
        oICH.GetAttributeValue sAttributeName, oPart, vValue
    End If
GotAttribute:
    
    If UCase(sValueType) = "DOUBLE" Then
        GetPartAttribute = CDbl(vValue)
    ElseIf UCase(sValueType) = "STRING" Then
        GetPartAttribute = CStr(vValue)
    ElseIf UCase(sValueType) = "LONG" Then
        GetPartAttribute = CLng(vValue)
    Else
        GetPartAttribute = vValue
    End If
    
    Exit Function

AllowFailure:
    #If DEBUG_ON Then
        DebugMsg.LogMsg "An error has been allowed to happen in " & MODULE & "." & METHOD
        DebugMsg.LogMsg "Problem getting part attribute: iPartIndex = " & iPartIndex & ", sAttributeName = " & sAttributeName & ", sValueType = " & sValueType
    #End If

    Exit Function
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
        Case 1
            sErr = "Problem getting part attribute: iPartIndex = " & iPartIndex & ", sAttributeName = " & sAttributeName & ", sValueType = " & sValueType
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function
Public Sub SetPartAttribute(iPartIndex As Integer, sAttributeName As String, vValue As Variant, Optional sValueType As String, Optional bEnforceFailure As Boolean = True, Optional bIsSymbolInput As Boolean = True)
Const METHOD = "SetPartAttribute"
Dim iErrCodeToUse As Integer
    If bEnforceFailure Then
        On Error GoTo ErrorHandler
    Else
        On Error GoTo AllowFailure
    End If

    iErrCodeToUse = 1
    If bIsSymbolInput Then
        If UCase(sValueType) = "DOUBLE" Then
            oICH.SetSymbolInputByName oPOC.Item(iPartIndex), sAttributeName, CDbl(vValue)
        ElseIf UCase(sValueType) = "STRING" Then
            oICH.SetSymbolInputByName oPOC.Item(iPartIndex), sAttributeName, CStr(vValue)
        ElseIf UCase(sValueType) = "LONG" Then
            oICH.SetSymbolInputByName oPOC.Item(iPartIndex), sAttributeName, CLng(vValue)
        Else
            GoTo ErrorHandler
        End If
    Else
        If UCase(sValueType) = "DOUBLE" Then
            oICH.SetAttributeValue sAttributeName, oPOC.Item(iPartIndex), CDbl(vValue)
        ElseIf UCase(sValueType) = "STRING" Then
            oICH.SetAttributeValue sAttributeName, oPOC.Item(iPartIndex), CStr(vValue)
        ElseIf UCase(sValueType) = "LONG" Then
            oICH.SetAttributeValue sAttributeName, oPOC.Item(iPartIndex), CLng(vValue)
        Else
            oICH.SetAttributeValue sAttributeName, oPOC.Item(iPartIndex), vValue
        End If
    End If

    Exit Sub
    
AllowFailure:
    #If DEBUG_ON Then
        DebugMsg.LogMsg "An error has been allowed to happen in " & MODULE & "." & METHOD
        DebugMsg.LogMsg "Problem setting part attribute: iPartIndex = " & iPartIndex & ", sAttributeName = " & sAttributeName & ", vValue = " & vValue & ", sValueType = " & sValueType
    #End If

    Exit Sub
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
        Case 1
            sErr = "Problem setting part attribute: iPartIndex = " & iPartIndex & ", sAttributeName = " & sAttributeName & ", vValue = " & vValue & ", sValueType = " & sValueType
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Sub
Public Function Rigid(NewPart As Integer, NewPartPort As String, _
                      Optional ConnectToPart As Integer = -1, Optional ConnectToPartPort As String = "Route", _
                      Optional OffsetX As Double = 0, Optional OffsetY As Double = 0, Optional OffsetZ As Double = 0, _
                      Optional RotX As Double = 0, Optional RotY As Double = 0, Optional RotZ As Double = 0, _
                      Optional flagShowWarnings As Integer = 1)
Const METHOD = "Rigid"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler
    
    If flagShowWarnings = 1 Then 'check whether the part&port are used more than once.
        If Rigid_DetectWarning(NewPart, NewPartPort, ConnectToPart, ConnectToPartPort) = False Then
            PF_EventHandler "Error: The part Index = " + NewPart + " port = " + NewPartPort + ", can be used as the NewPart only once in all callls to " + MODULE, Err, MODULE, METHOD, False
            Exit Function
        End If
    End If
    If (RotX Mod 90) <> 0 Or (RotY Mod 90) <> 0 Or (RotZ Mod 90) <> 0 Or Abs(RotX) >= 360 Or Abs(RotY) >= 360 Or Abs(RotZ) >= 360 Then
        PF_EventHandler "Error: Rotation angle must be one of -270/-180/-90/0/90/180/270!", Err, MODULE, METHOD, False
        Exit Function
    End If
    
    'if rot degree is -270/-180/-90, convert them to be +value
    If RotX < 0 Then
        RotX = 360 + RotX
    End If
    If RotY < 0 Then
        RotY = 360 + RotY
    End If
    If RotZ < 0 Then
        RotZ = 360 + RotZ
    End If
    
    'add the NewPart & NewPartPort to the dynamic structure array.
    ReDim Preserve CILog(UBound(CILog) + 1) As RigidConfigIndex
    CILog(UBound(CILog)).lNewPart = NewPart
    CILog(UBound(CILog)).sNewPartPort = NewPartPort
    CILog(UBound(CILog)).lConnectToPart = ConnectToPart
    CILog(UBound(CILog)).sConnectToPartPort = ConnectToPartPort
    CILog(UBound(CILog)).dOffsetX = OffsetX
    CILog(UBound(CILog)).dOffsetY = OffsetY
    CILog(UBound(CILog)).dOffsetZ = OffsetZ
    CILog(UBound(CILog)).dRotX = RotX
    CILog(UBound(CILog)).dRotY = RotY
    CILog(UBound(CILog)).dRotZ = RotZ
    
    'get the ConfigIndex from private funtion RotateXYZ(...)
    Dim lConfigIndex As Long
    lConfigIndex = RotateXYZ(CILog(UBound(CILog)).dRotX, CILog(UBound(CILog)).dRotY, CILog(UBound(CILog)).dRotZ)
    'MsgBox ConnectToPart & "," & ConnectToPartPort & "                   " & NewPart & "," & NewPartPort & "                  " & lConfigIndex
    'connect the NewPart&NewPartPort with ConnecttoPart&ConnectToPartPort, and set ConfigIndex and offset value for X/Y/Z axis
    Set AssemblyJoint = JointFactory.MakeRigidJoint(ConnectToPart, ConnectToPartPort, NewPart, NewPartPort, lConfigIndex, OffsetZ, OffsetY, OffsetX)
    JointCollection.Add AssemblyJoint 'add Joint to Joint Collection, which will be retrived by Function GetJoints() As Collection

    Exit Function
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function
Private Function Rigid_DetectWarning(NewPart As Integer, NewPartPort As String, _
                                    Optional ConnectToPart As Integer = -1, Optional ConnectToPartPort As String = "Route") As Boolean
Const METHOD = "Rigid_DetectWarning"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler
    
    Dim I As Long
    Rigid_DetectWarning = True 'set defalut return value is TRUE
    'check all the connected parts, whether the same part&port are used before.
    For I = 1 To UBound(CILog) Step 1
        If CILog(I).lNewPart = NewPart And CILog(I).sNewPartPort = NewPartPort Then
            Rigid_DetectWarning = False 'if the part&part is used before, return FALSE and exit function
            Exit Function
        End If
    Next
    Exit Function
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function
Public Function GetJoints() As Collection
Const METHOD = "GetJoints"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler
    
    Set GetJoints = JointCollection
    Exit Function
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function
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
Private Function RotateXYZ(dRot_X As Double, dRot_Y As Double, dRot_Z As Double) As Long
Const METHOD = "RotateXYZ"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler
    
    '----------Define and set 4 Vectors, 3 for X/Y/Z axis, the 4th for exchange usage.
    Dim Aix_X As IJDVector
    Dim Aix_Y As IJDVector
    Dim Aix_Z As IJDVector
    Dim Aix_temp As IJDVector
    Set Aix_X = New DVector
    Set Aix_Y = New DVector
    Set Aix_Z = New DVector
    Set Aix_temp = New DVector
    'set X/Y/Z axis vectors
    Aix_X.Set 1, 0, 0
    Aix_Y.Set 0, 1, 0
    Aix_Z.Set 0, 0, 1
    
    '----------Transform X/Y/Z axis, according to the degree of ratation
    If dRot_X = 270 Then     'Rotate X axis 270 degree, X=X & Z=Y & Y=-Z
        Aix_temp.Set Aix_Y.x, Aix_Y.y, Aix_Y.z
        Aix_Y.Set -Aix_Z.x, -Aix_Z.y, -Aix_Z.z
        Aix_Z.Set Aix_temp.x, Aix_temp.y, Aix_temp.z
    ElseIf dRot_X = 180 Then 'Rotate X axis 180 degree, X=X & Y=-Y & Z=-Z
        Aix_Y.Set -Aix_Y.x, -Aix_Y.y, -Aix_Y.z
        Aix_Z.Set -Aix_Z.x, -Aix_Z.y, -Aix_Z.z
    ElseIf dRot_X = 90 Then  'Rotate X axis 90  degree, X=X & Z=-Y & Y=Z
        Aix_temp.Set Aix_Y.x, Aix_Y.y, Aix_Y.z
        Aix_Y.Set Aix_Z.x, Aix_Z.y, Aix_Z.z
        Aix_Z.Set -Aix_temp.x, -Aix_temp.y, -Aix_temp.z
    End If
    If dRot_Y = 270 Then     'Rotate Y axis 270 degree, X=Z & Y=Y & Z=-X
        'Set Aix_Y = Aix_Y
        Aix_temp.Set Aix_X.x, Aix_X.y, Aix_X.z
        Aix_X.Set Aix_Z.x, Aix_Z.y, Aix_Z.z
        Aix_Z.Set -Aix_temp.x, -Aix_temp.y, -Aix_temp.z
    ElseIf dRot_Y = 180 Then 'Rotate Y axis 180 degree, X=-Z & Y=Y & Z=-X
        Aix_X.Set -Aix_X.x, -Aix_X.y, -Aix_X.z
        Aix_Z.Set -Aix_Z.x, -Aix_Z.y, -Aix_Z.z
    ElseIf dRot_Y = 90 Then  'Rotate Y axis 90  degree, X=-Z & Y=Y & Z=X
        Aix_temp.Set Aix_X.x, Aix_X.y, Aix_X.z
        Aix_X.Set -Aix_Z.x, -Aix_Z.y, -Aix_Z.z
        Aix_Z.Set Aix_temp.x, Aix_temp.y, Aix_temp.z
    End If
    If dRot_Z = 270 Then    'Rotate Z axis 270 degree, X=-Y & Y=X & Z=Z
        Aix_temp.Set Aix_X.x, Aix_X.y, Aix_X.z
        Aix_X.Set -Aix_Y.x, -Aix_Y.y, -Aix_Y.z
        Aix_Y.Set Aix_temp.x, Aix_temp.y, Aix_temp.z
    ElseIf dRot_Z = 180 Then 'Rotate Z axis 180 degree, X=-X & Y=-Y & Z=Z
        Aix_X.Set -Aix_X.x, -Aix_X.y, -Aix_X.z
        Aix_Y.Set -Aix_Y.x, -Aix_Y.y, -Aix_Y.z
    ElseIf dRot_Z = 90 Then  'Rotate Z axis 90  degree, X=Y & Y=-X & Z=Z
        'Set Aix_Z = Aix_Z
        Aix_temp.Set Aix_X.x, Aix_X.y, Aix_X.z
        Aix_X.Set Aix_Y.x, Aix_Y.y, Aix_Y.z
        Aix_Y.Set -Aix_temp.x, -Aix_temp.y, -Aix_temp.z
    End If
    
    '----------Combine the ConfigIndex in binary,  according to the X/Y/Z axis rotation
    'combine Type A & B for ConfigIndex
    Dim sConfigIndexAB As String
    sConfigIndexAB = "100" 'set the base Plane is XY, so there will be only 3 possibolities: XY&XY;XY&XZ;XY&YZ
    If (Abs(Aix_X.x) = 1 And Abs(Aix_Y.y) = 1) Or (Abs(Aix_X.y) = 1 And Abs(Aix_Y.x) = 1) Then 'XY&XY
        If Aix_Z.z = 1 Then      'XY & XY
            sConfigIndexAB = "1" & "100" & sConfigIndexAB
        ElseIf Aix_Z.z = -1 Then 'XY & -XY
            sConfigIndexAB = "0" & "100" & sConfigIndexAB
        Else                     'Error handling
            PF_EventHandler "Error: Illegal Z Axis value when try to Align XY with XY plane.", Err, MODULE, METHOD, False
        End If
    ElseIf (Abs(Aix_X.x) = 1 And Abs(Aix_Z.y) = 1) Or (Abs(Aix_X.y) = 1 And Abs(Aix_Z.x) = 1) Then 'XY&XZ
        If Aix_Y.z = 1 Then      'XY & XZ
            sConfigIndexAB = "1" & "101" & sConfigIndexAB
        ElseIf Aix_Y.z = -1 Then 'XY & -XZ
            sConfigIndexAB = "0" & "101" & sConfigIndexAB
        Else                     'Error handling
            PF_EventHandler "Error: Illegal Z Axis value when try to Align XY with XZ plane.", Err, MODULE, METHOD, False
        End If
    ElseIf (Abs(Aix_Y.x) = 1 And Abs(Aix_Z.y) = 1) Or (Abs(Aix_Y.y) = 1 And Abs(Aix_Z.x) = 1) Then 'XY&YZ
        If Aix_X.z = 1 Then      'XY & YZ
            sConfigIndexAB = "1" & "110" & sConfigIndexAB
        ElseIf Aix_X.z = -1 Then 'XY & -YZ
            sConfigIndexAB = "0" & "110" & sConfigIndexAB
        Else                     'Error handling
            PF_EventHandler "Error: Illegal Z Axis value when try to Align XY with YZ plane." + MODULE, Err, MODULE, METHOD, False
        End If
    Else
        PF_EventHandler "Error: Illegal X/Y/Z Axis value in ", Err, MODULE, METHOD, False
    End If
    
    'combine Type C & D for ConfigIndex
    Dim sConfigIndexCD As String
    sConfigIndexCD = "001"  'set the base Plane is X, so there will be only 6 possibolities: X&X;X&-X;X&Y;X&-Y;X&Z;X&-Z
    If Aix_X.x = 1 Then      'X & X
        sConfigIndexCD = "1" & "001" & sConfigIndexCD
    ElseIf Aix_X.x = -1 Then 'X & -X
        sConfigIndexCD = "0" & "001" & sConfigIndexCD
    ElseIf Aix_Y.x = 1 Then  'X & Y
        sConfigIndexCD = "1" & "010" & sConfigIndexCD
    ElseIf Aix_Y.x = -1 Then 'X & -Y
        sConfigIndexCD = "0" & "010" & sConfigIndexCD
    ElseIf Aix_Z.x = 1 Then  'X & Z
        sConfigIndexCD = "1" & "011" & sConfigIndexCD
    ElseIf Aix_Z.x = -1 Then 'X & -Z
        sConfigIndexCD = "0" & "011" & sConfigIndexCD
    Else
        PF_EventHandler "Error: Illegal X/Y/Z Axis value when try to Align X/Y/Z with X plane.", Err, MODULE, METHOD, False
    End If
    
    '''''MsgBox sConfigIndexCD & sConfigIndexAB
    'Combine the ConfigIndex Type A&B&C&D, and CALL BinaryToDecimal Function to convert configIndex from Binary to Decimal
    RotateXYZ = BinaryToDecimal(sConfigIndexCD & sConfigIndexAB)
        
    Exit Function
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function
Private Function BinaryToDecimal(sBinary As String) As Long
Const METHOD = "BinaryToDecimal"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler
    
    Dim lDecimal As Long
    Dim I As Integer
    For I = 1 To Len(sBinary)
        lDecimal = lDecimal + (Mid(sBinary, Len(sBinary) - I + 1, 1) * (2 ^ (I - 1)))
    Next I
    BinaryToDecimal = lDecimal
    Exit Function

ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function
Public Function GetCodeListDescription(ByVal oObject As Object, sInterface As String, sAttributeName As String, lValue As Long, eLongShortValue As GetCodelistValue) As String
Const METHOD = "GetCodeListDescription"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler

        Dim oCodeListData As IJDCodeListMetaData
        Dim oAttributeData As IJDAttributeMetaData
        Dim oAttributeInfo As IJDAttributeInfo
        Dim InterfaceID As Variant
        Dim tableName As String
        Dim sCodeListDesc As String
        
        Set oCodeListData = oObject
        Set oAttributeData = oObject

        InterfaceID = oAttributeData.iID(sInterface)
        Set oAttributeInfo = oAttributeData.AttributeInfo(InterfaceID, sAttributeName)
        tableName = oAttributeInfo.CodeListTableName
        
        If eLongShortValue = LongDesc Then
            sCodeListDesc = oCodeListData.LongStringValue(tableName, lValue)
        Else
            sCodeListDesc = oCodeListData.ShortStringValue(tableName, lValue)
        End If
        
        GetCodeListDescription = sCodeListDesc
        
        Set oCodeListData = Nothing
        Set oAttributeInfo = Nothing
        Set oAttributeData = Nothing

    Exit Function
ErrorHandler:
    GetCodeListDescription = "<BAD VALUE>"
    Set oCodeListData = Nothing
    Set oAttributeInfo = Nothing
    Set oAttributeData = Nothing
    
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
Public Function GetAttr(ByVal AttributeName As String, Optional sInterface As String = "", Optional eLongShortValue As GetCodelistValue = LongDesc, Optional bEnforceFailure As Boolean = True, Optional bGetAttributeFromCatalog As Boolean = False) As Variant
Const METHOD = "GetAttr"
Dim iErrCodeToUse As Integer
    If bEnforceFailure Then
        On Error GoTo ErrorHandler
    Else
        On Error GoTo AllowFailure
    End If

    iErrCodeToUse = 1

    'Get the Support and Support Occurence
    Dim oSupportOcc As IJPartOccAssembly
    Dim oSupport As IJDPart
    Set oSupportOcc = oICH
    Set oSupport = oSupportOcc.GetPart
    
    Dim varTemp As Variant
    
    'Check the attribute exist on the catalog DB,
    'First check the Occ. If it doesn't exist on the Occ then check the catalog.
    Dim IsAttrExists As Boolean
    If bGetAttributeFromCatalog = True Then
        IsAttrExists = IsAttributeExists(oSupport, AttributeName)
        If IsAttrExists = True Then
            GoTo GetAttributeFromCatalog
        End If
    Else        'Check the attr in Part Occ
        IsAttrExists = IsAttributeExists(oSupportOcc, AttributeName)
        If IsAttrExists = True Then
            GoTo GetAttributeFromOcc
        ElseIf IsAttrExists = False Then            'Check the attr in Catalog Part
            IsAttrExists = IsAttributeExists(oSupport, AttributeName)
            If IsAttrExists = True Then
                GoTo GetAttributeFromCatalog
            End If
        End If
    End If
    
    If IsAttrExists = False Then
        If bEnforceFailure Then
            GoTo ErrorHandler
        Else
            GoTo AllowFailure
        End If
    End If
    
Dim lSelectedValue As Long 'get codelist value
    
GetAttributeFromCatalog:
        oICH.GetAttributeValue AttributeName, oSupport, varTemp
        If sInterface <> "" Then
            lSelectedValue = varTemp
            varTemp = GetCodeListDescription(oICH, sInterface, AttributeName, lSelectedValue, eLongShortValue)
        End If
        GetAttr = varTemp
    Exit Function
GetAttributeFromOcc:
        oICH.GetAttributeValue AttributeName, oSupportOcc, varTemp
        If sInterface <> "" Then
            lSelectedValue = varTemp
            varTemp = GetCodeListDescription(oICH, sInterface, AttributeName, lSelectedValue, eLongShortValue)
        End If
        GetAttr = varTemp
    Exit Function
    
AllowFailure:
    #If DEBUG_ON Then
        DebugMsg.LogMsg "An error has been allowed to happen in " & MODULE & "." & METHOD
        DebugMsg.LogMsg "Problem getting attribute: Name = " & AttributeName & ", Interface = " & sInterface & ", LongShortValue = " & eLongShortValue & ", EnforceFailure = " & bEnforceFailure
    #End If
    
    Exit Function
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
        Case 1
            sErr = "Problem getting attribute: Name = " & AttributeName & ", Interface = " & sInterface & ", LongShortValue = " & eLongShortValue & ", EnforceFailure = " & bEnforceFailure
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function
Public Function GetSupportingSectionData(iIndex As Integer) As hsSteelMember
Const METHOD = "GetSupportingSectionData"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler
    
    Dim tResultMember As hsSteelMember
    
    Dim StructColl As Object
    Dim oStructObject As Object
    Set StructColl = oIOH.GetSupportingObjects
    Set oStructObject = StructColl.Item(iIndex)

    If Hilti_GetSupportingTypes(CLng(iIndex)) = "Steel" Then
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
    ElseIf Hilti_GetSupportingTypes(CLng(iIndex)) = "HgrBeam" Then
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
Public Function GetNOMPipeDiaByUnit(iIndex As Integer, Optional oOutUnit As Units = NPD_INCH) As Double
Const METHOD = "GetNOMPipeDiaByUnit"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler
    
    Dim dPipeNomDia As Double
    Dim sInUnit As String
    Dim uInputUnits As Units
    
    GetNOMPipeDiaAndUnits iIndex, dPipeNomDia, sInUnit

    uInputUnits = oUomServices.GetUnitId(UNIT_NPD, sInUnit)
    dPipeNomDia = oUomServices.ConvertUnitToUnit(UNIT_NPD, dPipeNomDia, uInputUnits, oOutUnit)
        
    GetNOMPipeDiaByUnit = dPipeNomDia

    Exit Function
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function
Public Sub SetAttr(sAttributeName As String, vValue As Variant, Optional bEnforceFailure As Boolean = True)
Const METHOD = "SetAttr"
Dim iErrCodeToUse As Integer
    If bEnforceFailure Then
        On Error GoTo ErrorHandler
    Else
        On Error GoTo AllowFailure
    End If
    
    iErrCodeToUse = 1
    oICH.SetAttributeValue sAttributeName, Nothing, vValue

    Exit Sub
    
AllowFailure:
    #If DEBUG_ON Then
        DebugMsg.LogMsg "An error has been allowed to happen in " & MODULE & "." & METHOD
        DebugMsg.LogMsg "Problem setting attribute: sAttributeName = " & sAttributeName & ", vValue = " & Trim(vValue) & ", bEnforceFailure = " & bEnforceFailure
    #End If

    Exit Sub
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
        Case 0
            sErr = "Problem setting attribute: sAttributeName = " & sAttributeName & ", vValue = " & Trim(vValue) & ", bEnforceFailure = " & bEnforceFailure
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Sub
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

Public Function AddPartsByName(PartClass() As String) As Object
Const METHOD = "AddPartsByName"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler

    If oICH Is Nothing Then GoTo ErrorHandler
    
    iErrCodeToUse = 1
    Set AddPartsByName = AddPartsByClassName(oICH, PartClass)
    Set oPartClasses = AddPartsByName
    PartClasses = PartClass

    Exit Function
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
        Case 1
            Dim iCnt As Integer
            For iCnt = 1 To UBound(PartClass)
                sErr = sErr + PartClass(iCnt)
                If iCnt < UBound(PartClass) Then sErr = sErr + ", "
            Next
            sErr = "May be problem with one of the following parts: " + sErr
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function
Public Function AddJoint(ByRef oJoint As Object)
Const METHOD = "AddJoint"
Dim iErrCodeToUse As Integer
On Error GoTo ErrorHandler
    
    JointCollection.Add oJoint
    Exit Function
    
ErrorHandler:
    Dim sErr As String
    Select Case iErrCodeToUse
        Case 0
            sErr = ""
    End Select
    Err.Raise LogError(Err, MODULE, METHOD, sErr).Number
End Function

Public Function GetRouteConnectionValue() As Integer
    Const METHOD = "GetRouteConnectionValue"
    On Error GoTo ErrorHandler

    GetRouteConnectionValue = oICH.GetRouteConnectionValue

    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function


