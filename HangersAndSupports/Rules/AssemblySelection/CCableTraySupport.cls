VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCableTraySupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CCableTraySupport.cls
'   ProgID:         HgrAssmSelRule.CCableTraySupport
'   Author:         Amlan
'   Creation Date:  11.Feb.2002
'   Description:
'    AssemblySelectionRule for CableTraySupports.
'    This ASR will be used to select Assembly based on SupportDiscipline
'
'   Change History:
'       11.Feb.2002             Amlan           Creation Date
'
'   09.Apr.2002         Caroline    Select Assembly based on the Feature of the Support
'                               'straight, turn etc.
'   24.Jun.2002         Caroline    Select Assembly based on the CommandType of the Support
'   19 Jan 2006         Mahesh    TR-CP·109498  Assembly Selection Rule problem with b-line trapeze support
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CCableTraySupport" 'Used for error messages

Implements IJHgrAssmSelectionRule

Public Function IJHgrAssmSelectionRule_GetAssemblies(ByVal pInputObjectInfo As Object) As HNGSUPPartSelectionHlpr.IJElements
Const METHOD = "IJHgrAssmSelectionRule_GetAssemblies"

    On Error GoTo ErrorHandler
    
    'Set the discipline type to be used while calling the services.
    Dim eHgrDiscipline As GSCADHgrSupFilterServices.HgrDisciplineType
    eHgrDiscipline = HgrCableWayDisciplineType
    
    'Get IJHgrInputConfig Hlpr Interface off of passed Helper
    Dim my_IJHgrInputObjectInfo As IJHgrInputObjectInfo
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputObjectInfo = pInputObjectInfo
    Set my_IJHgrInputConfigHlpr = pInputObjectInfo
        
    '=================================================================
    'retrive the first input object as the input to the inputConfigHlpr
    '=================================================================
    'Get A Common Services Object
    Dim HgrComServ As New HgrCommonServices
    Dim inputObjects As Object
    Dim inputObj As Object
    Dim objCount As Integer
    
    On Error Resume Next
    Set inputObjects = my_IJHgrInputObjectInfo.GetSupportedObjects
    On Error GoTo ErrorHandler
    
    ' if there are no input object, exit function
    If inputObjects Is Nothing Then
        Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
        Exit Function
    End If
           
    'If there are no CTrays in the returned collection exit.
    If inputObjects.Count = 0 Then
        Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
        Exit Function
    End If
     
    'retrieve the CTray shape information from the first input
    Dim csType As CrossSectionShapeTypes
    Dim Radius As Double
    Dim Width As Double
    Dim Depth As Double
    
    Set inputObj = inputObjects.Item(1)
    Dim strUnit As String
    HgrComServ.GetRteCrossSectionParams inputObj, csType, strUnit, Radius, Width, Depth
    If Width <= 0 Then
        Dim oUOMService         As IJUomVBInterface
        Dim unitID              As Units
        Set oUOMService = New UnitsOfMeasureServicesLib.UomVBInterface
        On Error GoTo UOMUnitErrorHandler
        unitID = oUOMService.GetUnitId(UNIT_DISTANCE, strUnit)
        On Error GoTo ErrorHandler
        Radius = oUOMService.ConvertUnitToDbu(UNIT_DISTANCE, Radius, unitID)
    
        Width = Radius * 2#
        Depth = Radius * 2#
    End If
    
    '*********************************************************************************************
    'Filter based on CTray Range(Width/Height) and Cross Section
    '*********************************************************************************************
    Dim AssyCollFromCTrays As Object
    Dim AssyCollFromWidth As Object
    Dim AssyCollFromDepth As Object
    Dim AssyCollFromFeature As Object
    Dim PartCollHlpr As New HgrCatPartsCollHlpr
    Dim e_Operator As HgrOperatorType
        e_Operator = OPER_EQUAL
       
    'Get An Assembly Selection Helper
    Dim AssmSelectionHelper As New AssySelectionHlpr
       
    'retrieve a collection of assembly according to width
    Set AssyCollFromWidth = AssmSelectionHelper.GetAssyCollFromBounds("NominalWidthFrom", _
                                                                       "NominalWidthTo", _
                                                                       Width, eHgrDiscipline)

                                              
    'retrieve a collection of assembly according to depth
    Set AssyCollFromDepth = AssmSelectionHelper.GetAssyCollFromBounds("NominalDepthFrom", _
                                                                       "NominalDepthTo", _
                                                                       Depth, eHgrDiscipline)

    'get a Collection of catalog Parts based on CTrayFeature
    Dim strFeatureType As String
    Dim varFeature() As Variant
    varFeature = my_IJHgrInputObjectInfo.GetAttributeByRule("HgrSupFeatureType")
    strFeatureType = varFeature(0)
    
    Set AssyCollFromFeature = AssmSelectionHelper.GetAssyCollByOperator("SupportedFamily", _
                                                                                strFeatureType, _
                                                                                e_Operator, _
                                                                                eHgrDiscipline)

'     'Get a Collection of Parts based on CommandType

    Dim AssyCollFromCmdType         As Object
   ' Dim AssyCollFromAllCmdType      As Object
    Dim lngCommandType              As Long
    Dim eAllCmdType                 As HgrCmdType
    eAllCmdType = HgrAllCmdType

    lngCommandType = my_IJHgrInputObjectInfo.GetCommandType

    Set AssyCollFromCmdType = AssmSelectionHelper.GetAssyCollByOperator("CommandType", _
                                                                                lngCommandType, _
                                                                                OPER_BITWISE_AND, _
                                                                             eHgrDiscipline)


     ' Get a Collection of Parts based applicable for both the CommandTypes.

    'Set AssyCollFromAllCmdType = AssmSelectionHelper.GetAssyCollByOperator("CommandType", _
                                                                                eAllCmdType, _
                                                                                e_Operator, _
                                                                             eHgrDiscipline)


    '======================================================
    'get the collection of parts based on distance
    '======================================================
    Dim StructPortCollection As Object
    Dim pDuct As Object
    
    Set StructPortCollection = my_IJHgrInputObjectInfo.GetSupportingObjectPorts
    
    Dim lngSupportingPortCnt As Long
    If StructPortCollection Is Nothing Then
        lngSupportingPortCnt = 0
    Else
        lngSupportingPortCnt = StructPortCollection.Count
    End If
    
    ' CR#128332 ; get the supporting collection which includes surfaces also
    Dim SupportingCollection As Object
    Set SupportingCollection = my_IJHgrInputObjectInfo.GetSupportingObjects
    Dim lngSupportingCnt As Long
    If SupportingCollection Is Nothing Then
        lngSupportingCnt = 0
    Else
        lngSupportingCnt = SupportingCollection.Count
    End If
    Set SupportingCollection = Nothing
    ' CR#128332 ends here
    
    'TR# 109498
    Dim lngSupportedCnt As Long
    lngSupportedCnt = inputObjects.Count
    Dim AssyCollByInputs As Object
    
    Set AssyCollByInputs = AssmSelectionHelper.GetAssembliesByInput(lngCommandType, eHgrDiscipline, strFeatureType, lngSupportedCnt, lngSupportingCnt)
    
    Dim AssyCollFromLength As Object
    If (lngSupportingPortCnt > 0) Then
    
     Dim ductDepth As Double
     Dim eSide As HgrEnumSectionSide
     
     eSide = my_IJHgrInputConfigHlpr.SupportedSectionOrientation(1)

     If eSide = FarSide Or NearSide Or NoSide Then
        ductDepth = Depth
     Else
        ductDepth = Width
     End If
     
     Dim eDistanceType           As HgrPortDistanceType
     Dim RouteStructMinDistance As Double
     Dim RouteStructMaxDistance As Double
     
     eDistanceType = DIRECT
'     RouteStructDistance = my_IJHgrInputConfigHlpr.GetDistanceBetweenPorts("Route", "Structure", eDistanceType)
     'RouteStructDistance = my_IJHgrInputConfigHlpr.GetDistanceBetweenPortCurves(eDistanceType) ' TR#111337 ; suncor problem
     my_IJHgrInputConfigHlpr.GetDistanceBetweenPortCurves eDistanceType, RouteStructMinDistance, RouteStructMaxDistance ' TR#111337 problem

     RouteStructMinDistance = RouteStructMinDistance - ductDepth / 2#
     RouteStructMaxDistance = RouteStructMaxDistance - ductDepth / 2#
     
     Set AssyCollFromLength = AssmSelectionHelper.GetAssyCollFromBounds("MinAssemblyLength", _
                                                                        "MaxAssemblyLength", _
                                                                        RouteStructMinDistance, _
                                                                        HgrCableWayDisciplineType, _
                                                                        , RouteStructMaxDistance)
    
    ' TRs#87104, 125406, 125477 ; now build a union of assemblies list for all the valid supporting faces for the Support Type
    Dim iNextStructPort As IJPort, iStructPort As IJPort
    Dim IJPort_Route As IJPort, IJPort_Struct As IJPort
    Dim iConnObjHlpr As IJHgrConnectableObjHlpr
    Dim iPortHlpr As IJHgrPortHlpr
    Dim AssyCollFromNextPort As Object
     'oPortColl contains all ports of given cross section type
    Dim oPortColl As IJElements
    Set oPortColl = New JObjectCollection
    
    Set IJPort_Route = inputObjects.Item(1)
    Set IJPort_Struct = StructPortCollection.Item(1)
    'get the port from structure
    oPortColl.Add IJPort_Struct
    Set iConnObjHlpr = CreateObject("HgrSupFilterServices.HgrConnectableObjHlpr")
    Set iPortHlpr = iConnObjHlpr.GetConnectableObjPortService(IJPort_Struct.Connectable)
    If (Not (iPortHlpr Is Nothing)) Then
        Set iNextStructPort = iPortHlpr.GetNextPort(IJPort_Struct, my_IJHgrInputObjectInfo, HgrIncrement)
        If (Not (iNextStructPort Is Nothing)) Then
             While (Not (oPortColl.Contains(iNextStructPort)))
                my_IJHgrInputConfigHlpr.GetDistBetweenSpecifiedPortCurves IJPort_Route, iNextStructPort, eDistanceType, RouteStructMinDistance, RouteStructMaxDistance
                RouteStructMinDistance = RouteStructMinDistance - ductDepth / 2#
                Set AssyCollFromNextPort = AssmSelectionHelper.GetAssyCollFromBounds("MinAssemblyLength", _
                                                                            "MaxAssemblyLength", _
                                                                            RouteStructMinDistance, _
                                                                            HgrCableWayDisciplineType, , RouteStructMaxDistance)
                Set AssyCollFromLength = PartCollHlpr.GetMergedCollection(AssyCollFromNextPort, AssyCollFromLength)
                Set iStructPort = iNextStructPort
                'add the this port to portcoll
                oPortColl.Add iNextStructPort ' DM#136232
                Set iNextStructPort = Nothing
                Set AssyCollFromNextPort = Nothing
                Set iNextStructPort = iPortHlpr.GetNextPort(iStructPort, my_IJHgrInputObjectInfo, HgrIncrement)
                Set iStructPort = Nothing
            Wend
        End If
    End If
    Set oPortColl = Nothing
    ' TRs#87104,125406, 125477 ends here
    Set IJPort_Route = Nothing
    Set IJPort_Struct = Nothing
    End If
   
   'Finds out if collection of hangers which can support more than one CTray.
   'Intersect the returned sets
   
    Dim AssyMergedCTrayColl1 As Object
    Dim AssyMergedCTrayColl2 As Object
    Dim AssyMergedCTrayColl3 As Object
    Dim AssyMergedCTrayColl4 As Object
    Dim AssyMergedCTrayColl5 As Object
    Dim AssyMergedCTrayColl6 As Object
    Dim AssyMergedCTrayColl7 As Object
    
   If inputObjects.Count > 1 Then
        
        Dim AssycollFromCTrayCount As Object
        
        Set AssycollFromCTrayCount = AssmSelectionHelper.GetAssyCollByOperator("Supported_Count", "1+", e_Operator, eHgrDiscipline)
        'merged the three collections together
        If Not AssycollFromCTrayCount Is Nothing Then
            Set AssyMergedCTrayColl1 = PartCollHlpr.GetInterferencePartsColl(AssyCollFromWidth, AssycollFromCTrayCount)
            Set AssyMergedCTrayColl2 = PartCollHlpr.GetInterferencePartsColl(AssyCollFromDepth, AssyMergedCTrayColl1)
            Set AssyMergedCTrayColl3 = PartCollHlpr.GetInterferencePartsColl(AssyCollFromFeature, AssyMergedCTrayColl2)
            'get a Union of two collections
            Set AssyMergedCTrayColl4 = PartCollHlpr.GetInterferencePartsColl(AssyMergedCTrayColl3, AssyCollFromCmdType)
            Set AssyMergedCTrayColl5 = PartCollHlpr.GetInterferencePartsColl(AssyMergedCTrayColl4, AssyCollFromLength)
            Set AssyMergedCTrayColl6 = PartCollHlpr.GetInterferencePartsColl(AssyMergedCTrayColl5, AssyCollByInputs)
            Set AssyCollFromCTrays = AssyMergedCTrayColl6
        End If
    Else
        Set AssyMergedCTrayColl1 = PartCollHlpr.GetInterferencePartsColl(AssyCollFromWidth, AssyCollFromDepth)
        Set AssyMergedCTrayColl2 = PartCollHlpr.GetInterferencePartsColl(AssyCollFromFeature, AssyMergedCTrayColl1)
        Set AssyMergedCTrayColl3 = PartCollHlpr.GetInterferencePartsColl(AssyMergedCTrayColl2, AssyCollFromCmdType)
        
        Set AssyMergedCTrayColl4 = PartCollHlpr.GetInterferencePartsColl(AssyMergedCTrayColl3, AssyCollFromLength)
        Set AssyMergedCTrayColl5 = PartCollHlpr.GetInterferencePartsColl(AssyMergedCTrayColl4, AssyCollByInputs)
        Set AssyCollFromCTrays = AssyMergedCTrayColl5
    End If
        
    Set AssyCollFromWidth = Nothing
    Set AssyCollFromDepth = Nothing
    'Set AssyCollFromAllCmdType = Nothing
    Set AssyCollFromCmdType = Nothing
    
    Set AssyMergedCTrayColl1 = Nothing
    Set AssyMergedCTrayColl2 = Nothing
    Set AssyMergedCTrayColl3 = Nothing
    Set AssyMergedCTrayColl4 = Nothing
    Set AssyMergedCTrayColl5 = Nothing
    Set AssyMergedCTrayColl6 = Nothing
    
   '***************************************************************************************
   'Get the Collection of Assemblies which support the current Port Configuration.
   '        Only check against results already returned by the CTray size filter.
   '***************************************************************************************
   Dim AssyCollFromConfigAndCTray As Object
   
   'Get all structures on the support.  In particular we are looking for the
   'port on the input object that is being used by the support.
    
    'If there is no Struct Collection go to next filter.
    If StructPortCollection Is Nothing Then
        Set AssyCollFromConfigAndCTray = AssyCollFromCTrays
        GoTo NON_STRUCT_FILTER
    End If
    
    'If there are no items in the returned collection go to next filter.
    If StructPortCollection.Count = 0 Then
        Set AssyCollFromConfigAndCTray = AssyCollFromCTrays
        GoTo NON_STRUCT_FILTER
    End If
    
    'Get the primary CTray Port.
    'This is a port on the primary CTray.  Since CTray Object and CTray Port are
    'the same object we are able to use the first item returned
    'by GetSupportedObjects() rather than GetSupportedObjectPorts()
    Dim IJPort_CTray As IJPort
    Set IJPort_CTray = inputObjects.Item(1)
    
    'Get the first item in the Struct Port Collection.
    'This is a port on the primary supporting object (beam or slab).
    'Dim IJPort_Struct As IJPort
    Set IJPort_Struct = StructPortCollection.Item(1)
    
    'Get the Target Point.
    Dim IJDPosition_Target As IJDPosition
    Set IJDPosition_Target = my_IJHgrInputObjectInfo.GetTargetPoint
   
    'Get the command type
    Dim eHgrCmdType As HgrCmdType
    eHgrCmdType = my_IJHgrInputObjectInfo.GetCommandType
    
    'Get a Port Configuration Service.
    Dim PortPostureService As New HgrPortPostureService
    
    'Ask the Port Configuration Service for the Configuration of the Ports
    Dim bUseFiniteGeometry As Boolean
    bUseFiniteGeometry = True
    Dim PortConfig As Long
    
    PortConfig = PortPostureService.GetGenericPortConfig(my_IJHgrInputObjectInfo, _
                                            IJPort_CTray, _
                                            IJPort_Struct, _
                                            IJDPosition_Target, _
                                            lngCommandType)
    
    'Clear un-used items.
    Set PortPostureService = Nothing
    
    'Only search for matching assemblies if the Configuration is NOT zero.
    If (PortConfig <> 0) And (Not AssyCollFromCTrays Is Nothing) Then
        'Get a collection of Catalog Parts with a Configuration Matching that
        'returned by the Port Posture Service.  Only check against the items already
        'returned by the CTray filter.
        Set AssyCollFromConfigAndCTray = AssmSelectionHelper.GetAssyCollByOperator( _
                                                "TypeSelectionRule", _
                                                PortConfig, OPER_BITWISE_AND, _
                                                eHgrDiscipline, AssyCollFromCTrays)
                                                
        Set AssyCollFromConfigAndCTray = PartCollHlpr.GetInterferencePartsColl(AssyCollFromConfigAndCTray, AssyCollFromCTrays)
    Else
        Set AssyCollFromConfigAndCTray = AssyCollFromCTrays
    End If
    
    Set AssyCollFromConfigAndCTray = AssmSelectionHelper.GetAssyCollFromSupportingFace(IJPort_Struct, pInputObjectInfo, AssyCollFromConfigAndCTray)
    
    '***************************************************************************************
NON_STRUCT_FILTER:
    
    Dim AssemblyMergedCollection As Object
    Set AssemblyMergedCollection = AssyCollFromConfigAndCTray 'AssyCollFromCTrays
 
    
    'The method needs to return IJElements.  Get this off of the object returned.
    Dim IJElements_AssemblyCollection As HNGSUPPartSelectionHlpr.IJElements
    Set IJElements_AssemblyCollection = AssemblyMergedCollection
    
    'Return the Collection of Assembly Definitions
    Set IJHgrAssmSelectionRule_GetAssemblies = IJElements_AssemblyCollection
    
    Set AssmSelectionHelper = Nothing
    Set AssyCollFromCTrays = Nothing
    Set AssyCollFromConfigAndCTray = Nothing
    Set PartCollHlpr = Nothing
    Set AssemblyMergedCollection = Nothing
    Set IJDPosition_Target = Nothing
    Set StructPortCollection = Nothing
    Set AssyCollFromFeature = Nothing
    Set inputObjects = Nothing
    Set inputObj = Nothing
    
    Exit Function
    
UOMUnitErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "GetUnitId method of IJUOMVBInterface failed.Support type list may not be available.Could be because of localised string mismatch").Number
    
ErrorHandler:
     Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function




