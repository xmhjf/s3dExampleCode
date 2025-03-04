VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDuctSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CDuctSupport.cls
'   ProgID:         HgrAssmSelRule.CDuctSupport
'   Author:         Amlan
'   Creation Date:  11.Feb.2002
'   Description:
'    AssemblySelectionRule for DuctSupports.
'    This ASR will be used to select Assembly based on SupportDiscipline
'
'   Change History:
'    09.Apr.2002        Caroline        Select Assembly based on the Feature of the Support
'                               straight, turn etc.
'    21.Jun.2002        Caroline        Select Assembly Based on the Shape of
'                               Duct Rectangle, Oval etc
'    24.Jun.2002        Caroline        Select Assembly based on the CommandType of the Support
'     16 Jan 2008       Mahesh          TR134890
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "CDuctSupport" 'Used for error messages

Implements IJHgrAssmSelectionRule

Public Function IJHgrAssmSelectionRule_GetAssemblies(ByVal pInputObjectInfo As Object) As HNGSUPPartSelectionHlpr.IJElements
Const METHOD = "IJHgrAssmSelectionRule_GetAssemblies"

    On Error GoTo ErrorHandler
    
    'Set the discipline type to be used while calling the services.
    Dim eHgrDiscipline As GSCADHgrSupFilterServices.HgrDisciplineType
    eHgrDiscipline = HgrHVACDisciplineType
    
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
           
    'If there are no ducts in the returned collection exit.
    If inputObjects.Count = 0 Then
        Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
        Exit Function
    End If
     
    'retrieve the duct shape information from the first input
    Dim csType As CrossSectionShapeTypes
    Dim Radius As Double
    Dim Width As Double
    Dim Depth As Double
    
    Set inputObj = inputObjects.Item(1)
    Dim strUnit As String
    HgrComServ.GetRteCrossSectionParams inputObj, csType, strUnit, Radius, Width, Depth
    If Width <= 0 Then
        Width = Radius / 2#
        Depth = Radius / 2#
    End If
    
    '*********************************************************************************************
    'Filter based on Duct Range(Width/Height) and Cross Section
    '*********************************************************************************************
    
    Dim AssyCollFromDucts   As Object
    Dim AssyCollFromWidth   As Object
    Dim AssyCollFromDepth   As Object
    Dim AssyCollFromFeature As Object
    Dim AssyCollFromShape   As Object
    Dim PartCollHlpr As New HgrCatPartsCollHlpr
    Dim e_Operator As HgrOperatorType
        e_Operator = OPER_EQUAL
       
    'Get An Assembly Selection Helper
    Dim AssmSelectionHelper As New AssySelectionHlpr
    
    'retrieve a collection of assembly according to width
    Set AssyCollFromWidth = AssmSelectionHelper.GetAssyCollFromBounds("NominalWidthFrom", _
                                                                       "NominalWidthTo", _
                                                                       Width, eHgrDiscipline)
    If Not AssyCollFromWidth Is Nothing Then
        If AssyCollFromWidth.Count = 0 Then
            Set AssyCollFromWidth = Nothing
        End If
    End If
                                              
    'retrieve a collection of assembly according to depth
    Set AssyCollFromDepth = AssmSelectionHelper.GetAssyCollFromBounds("NominalDepthFrom", _
                                                                       "NominalDepthTo", _
                                                                       Depth, eHgrDiscipline)
    If Not AssyCollFromDepth Is Nothing Then
        If AssyCollFromDepth.Count = 0 Then
            Set AssyCollFromDepth = Nothing
        End If
    End If

      'Get a Collection of catalog Parts based on DuctFeature
    Dim strFeatureType As String
    Dim varFeature() As Variant
    varFeature = my_IJHgrInputObjectInfo.GetAttributeByRule("HgrSupFeatureType")
    strFeatureType = varFeature(0)
    
    Set AssyCollFromFeature = AssmSelectionHelper.GetAssyCollByOperator("SupportedFamily", _
                                                                                strFeatureType, _
                                                                                e_Operator, _
                                                                                eHgrDiscipline)
    If Not AssyCollFromFeature Is Nothing Then
        If AssyCollFromFeature.Count = 0 Then
            Set AssyCollFromFeature = Nothing
        End If
    End If
    
     'Get a Collection of catalog Parts based on DuctShape
     'CrossSection Shapes are Codelisted, SupportedShapeType is a longProperty
     'defined in HgrSupportDefinition of Refdata\SupportEntities

    Set AssyCollFromShape = AssmSelectionHelper.GetAssyCollByOperator("SupportedShapeType", _
                                                                                csType, _
                                                                                e_Operator, _
                                                                                eHgrDiscipline)
    If Not AssyCollFromShape Is Nothing Then
        If AssyCollFromShape.Count = 0 Then
            Set AssyCollFromShape = Nothing
        End If
    End If
   
    'Get a Collection of Parts based on CommandType, this code will be used in Future
    'when we have Assemblies specific for Commmand( By Point, By Struct)
    
    Dim AssyCollFromCmdType         As Object
    'Dim AssyCollFromAllCmdType      As Object
    Dim lngCommandType              As Long
    Dim eAllCmdType                 As HgrCmdType
    eAllCmdType = HgrAllCmdType

    lngCommandType = my_IJHgrInputObjectInfo.GetCommandType

    Set AssyCollFromCmdType = AssmSelectionHelper.GetAssyCollByOperator("CommandType", _
                                                                                lngCommandType, _
                                                                                OPER_BITWISE_AND, _
                                                                             eHgrDiscipline)

    If Not AssyCollFromCmdType Is Nothing Then
        If AssyCollFromCmdType.Count = 0 Then
            Set AssyCollFromCmdType = Nothing
        End If
    End If


     ' Get a Collection of Parts based applicable for both the CommandTypes.

    'Set AssyCollFromAllCmdType = AssmSelectionHelper.GetAssyCollByOperator("CommandType", _
                                                                                eAllCmdType, _
                                                                                e_Operator, _
                                                                             eHgrDiscipline)


    'If Not AssyCollFromAllCmdType Is Nothing Then
        'If AssyCollFromAllCmdType.Count = 0 Then
            'Set AssyCollFromAllCmdType = Nothing
        'End If
    'End If


    '======================================================
    'get the collection of parts based on distance
    '======================================================
    Dim StructPortCollection As Object
    Dim pDuct As Object
    
    Set StructPortCollection = my_IJHgrInputObjectInfo.GetSupportingObjectPorts
    
    Dim lngSupportingCnt As Long
    If StructPortCollection Is Nothing Then
        lngSupportingCnt = 0
    Else
        lngSupportingCnt = StructPortCollection.Count
    End If
    
    Dim AssyCollFromLength As Object
    If (lngSupportingCnt > 0) Then
    
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
                                                                        HgrHVACDisciplineType, _
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
                                                                            HgrHVACDisciplineType, , RouteStructMaxDistance)
                Set AssyCollFromLength = PartCollHlpr.GetMergedCollection(AssyCollFromNextPort, AssyCollFromLength)
                Set iStructPort = iNextStructPort
                'add the this port to portcoll
                oPortColl.Add iNextStructPort ' TR#136231, 136232, 136233
                Set iNextStructPort = Nothing
                Set AssyCollFromNextPort = Nothing
                Set iNextStructPort = iPortHlpr.GetNextPort(iStructPort, my_IJHgrInputObjectInfo, HgrIncrement)
                Set iStructPort = Nothing
            Wend
        End If
    End If
    ' TRs#87104,125406, 125477 ends here
    Set oPortColl = Nothing
    Set IJPort_Route = Nothing
    Set IJPort_Struct = Nothing
    End If
    
   'Finds out if collection of hangers which can support more than one Duct.
   'Intersect the returned sets
    Dim AssyMergedDuctColl1 As Object
    Dim AssyMergedDuctColl2 As Object
    Dim AssyMergedDuctColl3 As Object
    Dim AssyMergedDuctColl4 As Object
    Dim AssyMergedDuctColl5 As Object
    Dim AssyMergedDuctColl6 As Object
    Dim AssyMergedDuctColl7 As Object
    
   If inputObjects.Count > 1 Then
        Dim AssycollFromDuctCount As Object
        Set AssycollFromDuctCount = AssmSelectionHelper.GetAssyCollByOperator("Supported_Count", "1+", e_Operator, eHgrDiscipline)
        'merged the four collections together
        If Not AssycollFromDuctCount Is Nothing Then
            Set AssyMergedDuctColl1 = PartCollHlpr.GetInterferencePartsColl(AssyCollFromWidth, AssycollFromDuctCount)
            Set AssyMergedDuctColl2 = PartCollHlpr.GetInterferencePartsColl(AssyCollFromDepth, AssyMergedDuctColl1)
            Set AssyMergedDuctColl3 = PartCollHlpr.GetInterferencePartsColl(AssyCollFromFeature, AssyMergedDuctColl2)
            Set AssyMergedDuctColl4 = PartCollHlpr.GetInterferencePartsColl(AssyCollFromShape, AssyMergedDuctColl3)
            'Get a Union of both Collections
            'Set AssyMergedDuctColl5 = PartCollHlpr.GetMergedCollection(AssyCollFromAllCmdType, AssyCollFromCmdType)
            Set AssyMergedDuctColl5 = PartCollHlpr.GetInterferencePartsColl(AssyMergedDuctColl4, AssyCollFromCmdType)
            Set AssyMergedDuctColl6 = PartCollHlpr.GetInterferencePartsColl(AssyMergedDuctColl5, AssyCollFromLength)
            Set AssyCollFromDucts = AssyMergedDuctColl6
        End If
    Else
        Set AssyMergedDuctColl1 = PartCollHlpr.GetInterferencePartsColl(AssyCollFromWidth, AssyCollFromDepth)
        Set AssyMergedDuctColl2 = PartCollHlpr.GetInterferencePartsColl(AssyCollFromFeature, AssyMergedDuctColl1)
        Set AssyMergedDuctColl3 = PartCollHlpr.GetInterferencePartsColl(AssyCollFromShape, AssyMergedDuctColl2)
        'Set AssyMergedDuctColl4 = PartCollHlpr.GetMergedCollection(AssyCollFromAllCmdType, AssyCollFromCmdType)
        Set AssyMergedDuctColl4 = PartCollHlpr.GetInterferencePartsColl(AssyMergedDuctColl3, AssyCollFromCmdType)
        Set AssyMergedDuctColl5 = PartCollHlpr.GetInterferencePartsColl(AssyMergedDuctColl4, AssyCollFromLength)
        Set AssyCollFromDucts = AssyMergedDuctColl5
    End If
        
    Set AssyCollFromWidth = Nothing
    Set AssyCollFromDepth = Nothing
    Set AssyCollFromFeature = Nothing
    Set AssyCollFromShape = Nothing
    'Set AssyCollFromAllCmdType = Nothing
    Set AssyCollFromCmdType = Nothing
    
    Set AssyMergedDuctColl1 = Nothing
    Set AssyMergedDuctColl2 = Nothing
    Set AssyMergedDuctColl3 = Nothing
    Set AssyMergedDuctColl4 = Nothing
    Set AssyMergedDuctColl5 = Nothing
    Set AssyMergedDuctColl6 = Nothing
    
   '***************************************************************************************
   'Get the Collection of Assemblies which support the current Port Configuration.
   '        Only check against results already returned by the Duct size filter.
   '***************************************************************************************
   Dim AssyCollFromConfigAndDuct As Object
   
   'Get all structures on the support.  In particular we are looking for the
   'port on the input object that is being used by the support.
'    Dim StructPortCollection As Object
'    Set StructPortCollection = my_IJHgrInputObjectInfo.GetSupportingObjectPorts
    
    'If there is no Struct Collection go to next filter.
    If StructPortCollection Is Nothing Then
        Set AssyCollFromConfigAndDuct = AssyCollFromDucts
        GoTo NON_STRUCT_FILTER
    End If
    
    'If there are no items in the returned collection go to next filter.
    If StructPortCollection.Count = 0 Then
        Set AssyCollFromConfigAndDuct = AssyCollFromDucts
        GoTo NON_STRUCT_FILTER
    End If
    
    'Get the primary Duct Port.
    'This is a port on the primary Duct.  Since Duct Object and Duct Port are
    'the same object we are able to use the first item returned
    'by GetSupportedObjects() rather than GetSupportedObjectPorts()
    Dim IJPort_Duct As IJPort
    Set IJPort_Duct = inputObjects.Item(1)
    
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
'    PortPostureService.GetPortConfiguration my_IJHgrInputObjectInfo, _
'                                            IJPort_Duct, _
'                                            IJPort_Struct, _
'                                            IJDPosition_Target, _
'                                            eHgrCmdType, _
'                                            bUseFiniteGeometry, _
'                                            PortConfig

    On Error Resume Next
    PortConfig = PortPostureService.GetGenericPortConfig(my_IJHgrInputConfigHlpr, _
                                                            IJPort_Duct, _
                                                            IJPort_Struct, _
                                                            IJDPosition_Target, _
                                                            eHgrCmdType)

    On Error GoTo ErrorHandler
    
'    If (PortConfig = 0) Then
'        'Didn't find a configuration using finite geometry.
'        'Try again with infinite geometry
'        bUseFiniteGeometry = False
'
'        PortPostureService.GetPortConfiguration my_IJHgrInputObjectInfo, _
'                                            IJPort_Duct, _
'                                            IJPort_Struct, _
'                                            IJDPosition_Target, _
'                                            eHgrCmdType, _
'                                            bUseFiniteGeometry, _
'                                            PortConfig
'    End If
        
    
    'Clear un-used items.
    Set PortPostureService = Nothing
    
    'Only search for matching assemblies if the Configuration is NOT zero.
    If (PortConfig <> 0) And (Not AssyCollFromDucts Is Nothing) Then
        'Get a collection of Catalog Parts with a Configuration Matching that
        'returned by the Port Posture Service.  Only check against the items already
        'returned by the Duct filter.
        
        Dim AssyCollFromConfigAndDuct1 As Object
        Set AssyCollFromConfigAndDuct1 = AssmSelectionHelper.GetAssyCollByOperator( _
                                                "TypeSelectionRule", _
                                                PortConfig, OPER_BITWISE_AND, _
                                                eHgrDiscipline, AssyCollFromDucts)
                                                
        'Null the collection if it contains no items
        Set AssyCollFromConfigAndDuct = PartCollHlpr.GetInterferencePartsColl(AssyCollFromDucts, _
                                                                              AssyCollFromConfigAndDuct1)
    Else
        Set AssyCollFromConfigAndDuct = AssyCollFromDucts
    End If
    
    
    Set AssyCollFromConfigAndDuct = AssmSelectionHelper.GetAssyCollFromSupportingFace(IJPort_Struct, pInputObjectInfo, AssyCollFromConfigAndDuct)
    
    
   '***************************************************************************************
   'Filter based on the Family of the beam cross section.
   '***************************************************************************************
'    Dim AssyCollFromBeams As Object
    
'    'Get all structures.
'    Dim StructCollection As Object
'    Set StructCollection = my_IJHgrInputObjectInfo.GetSupportingObjects
'
'    Set AssyCollFromBeams = AssmSelectionHelper.GetAssyCollFromSupporting(StructCollection)
'
'    If Not AssyCollFromBeams Is Nothing Then
'        If AssyCollFromBeams.Count = 0 Then
'            Set AssyCollFromBeams = Nothing
'        End If
'    End If
    
    '***************************************************************************************
NON_STRUCT_FILTER:

    ' Merge the Three Lists.
    ' If a list is empty, it is ignored when merging.
    
    Dim AssemblyMergedCollection As Object
    Set AssemblyMergedCollection = AssyCollFromConfigAndDuct
    
'    Set AssemblyMergedCollection = PartCollHlpr.GetInterferencePartsColl(AssyCollFromConfigAndDuct, _
'                                                                     AssyCollFromBeams)
    
'    If AssyCollFromConfigAndDuct Is Nothing And AssyCollFromBeams Is Nothing Then
'        'Both Empty
'        MsgBox "both nothing"
'        Set AssemblyMergedCollection = Nothing
'    ElseIf AssyCollFromBeams Is Nothing Then
'        'No Beam Results
'        MsgBox "duct = " & AssyCollFromConfigAndDuct.Count
'        Set AssemblyMergedCollection = AssyCollFromConfigAndDuct
'    ElseIf AssyCollFromConfigAndDuct Is Nothing Then
'        'No Duct and Configuration Results
'        MsgBox "Beam = " & AssyCollFromBeams.Count
'        Set AssemblyMergedCollection = AssyCollFromBeams
'    Else
'        'Merge Beam, Duct, and Configuration Results
'        MsgBox "duct = " & AssyCollFromConfigAndDuct.Count
'        MsgBox "Beam = " & AssyCollFromBeams.Count
'        Set AssemblyMergedCollection = _
'                  PartCollHlpr.GetInterferencePartsColl(AssyCollFromConfigAndDuct, AssyCollFromBeams)
'    End If
    
    'The method needs to return IJElements.  Get this off of the object returned.
    Dim IJElements_AssemblyCollection As HNGSUPPartSelectionHlpr.IJElements
    Set IJElements_AssemblyCollection = AssemblyMergedCollection
    
    'Return the Collection of Assembly Definitions
    Set IJHgrAssmSelectionRule_GetAssemblies = IJElements_AssemblyCollection
    
    Set AssmSelectionHelper = Nothing
    Set AssyCollFromDucts = Nothing
    Set AssyCollFromConfigAndDuct = Nothing
    Set PartCollHlpr = Nothing
    Set AssemblyMergedCollection = Nothing
    Set IJDPosition_Target = Nothing
    Set StructPortCollection = Nothing
    Set HgrComServ = Nothing
    Set inputObjects = Nothing
    Set inputObj = Nothing
    
    Exit Function
    
ErrorHandler:
     Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function


