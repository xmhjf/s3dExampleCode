VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConduitSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CConduitSupport.cls
'   ProgID:         HgrAssmSelRule.CConduitSupport
'   Author:         Rama Rao
'   Creation Date:  07.Nov.2002
'   Description:
'    AssemblySelectionRule for ConduitSupports.
'    This ASR will be used to select Assembly based on SupportDiscipline
'
'   Change History:
'   07.Nov.2002     Rama Rao             Creation Date
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CConduitSupport" 'Used for error messages

Implements IJHgrAssmSelectionRule

Public Function IJHgrAssmSelectionRule_GetAssemblies(ByVal pInputObjectInfo As Object) As HNGSUPPartSelectionHlpr.IJElements
Const METHOD = "IJHgrAssmSelectionRule_GetAssemblies"

    On Error GoTo ErrorHandler
    
    'Set the discipline type to be used while calling the services.
    Dim eHgrDiscipline As GSCADHgrSupFilterServices.HgrDisciplineType
    eHgrDiscipline = HgrConduitDisciplineType
    
    'Get IJHgrInputConfig Hlpr Interface off of passed Helper
    Dim my_IJHgrInputObjectInfo As IJHgrInputObjectInfo
    Set my_IJHgrInputObjectInfo = pInputObjectInfo
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pInputObjectInfo
    
    'Get all Conduits on the support.
    Dim ConduitCollection As Object
    On Error Resume Next
    Set ConduitCollection = my_IJHgrInputObjectInfo.GetSupportedObjects
    On Error GoTo ErrorHandler
    
    'If there is no Conduits Collection exit.
    If ConduitCollection Is Nothing Then
        Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
        Exit Function
    End If
    
    'If there are no Conduits in the returned collection exit.
    If ConduitCollection.Count = 0 Then
        Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
        Exit Function
    End If
    
    '*********************************************************************************************
    'Filter based on Conduit Size
    '*********************************************************************************************

    Dim AssyCollFromConduits As Object
'    'Get A Common Services Object
    Dim HgrComServ As New HgrCommonServices
'    'Get An Assembly Selection Helper
    Dim AssmSelectionHelper As New AssySelectionHlpr

'    'Get the Size of the Primary Conduit
    Dim PrimaryConduitSize As Double
    Dim strUnit As String
    PrimaryConduitSize = HgrComServ.PrimaryPipeDiameter(ConduitCollection, strUnit)
    '
    Dim strFeatureType As String
    Dim varFeature() As Variant
     'Get a Collection of catalog Parts based on ConduitFeature
    varFeature = my_IJHgrInputObjectInfo.GetAttributeByRule("HgrSupFeatureType")
    strFeatureType = varFeature(0)
    Dim AssyCollFromFeature As Object
    Dim PartCollHlpr As New HgrCatPartsCollHlpr
    Dim e_Operator As HgrOperatorType
        e_Operator = OPER_EQUAL
    Set AssyCollFromFeature = AssmSelectionHelper.GetAssyCollByOperator("SupportedFamily", _
                                                                                strFeatureType, _
                                                                                e_Operator, _
                                                                             eHgrDiscipline)
     

    
    'Get a collection of Catalog Parts with a Min and Max Conduit Size bound that
    'of the Primary Conduit.
    
    Dim AssyCollFromCmdType         As Object
    'Dim AssyCollFromAllCmdType      As Object
    Dim lngCommandType              As Long
    Dim eAllCmdType                 As HgrCmdType
    eAllCmdType = HgrAllCmdType
    
    lngCommandType = my_IJHgrInputObjectInfo.GetCommandType
    
    Set AssyCollFromConduits = AssmSelectionHelper.GetHgrSupportDefByND(lngCommandType, _
                                                                     eHgrDiscipline, _
                                                                     strFeatureType, _
                                                                     PrimaryConduitSize, _
                                                                     strUnit)

    
    ' Get a Collection of Parts based on the CommandType (By Point, By Struct)
    Set AssyCollFromCmdType = AssmSelectionHelper.GetAssyCollByOperator("CommandType", _
                                                                                lngCommandType, _
                                                                                 OPER_BITWISE_AND, _
                                                                             eHgrDiscipline)
                                                                             
                                                                          
    
    ' Get a Collection of Parts based applicable for both the CommandTypes.
   ' Set AssyCollFromAllCmdType = AssmSelectionHelper.GetAssyCollByOperator("CommandType", _
                                                                                eAllCmdType, _
                                                                                e_Operator, _
                                                                             eHgrDiscipline)
        
                                                                                                                                          
                                                                             
                                                                         
    Dim AssyMergedConduitColl As Object
    Dim AssyMergedConduitColl1 As Object
    Dim AssyMergedConduitColl2 As Object
                        
    ' Get a union of two collections
    'Set AssyMergedConduitColl = PartCollHlpr.GetMergedCollection(AssyCollFromAllCmdType, _
                                                                   AssyCollFromCmdType)
    Set AssyMergedConduitColl1 = PartCollHlpr.GetInterferencePartsColl(AssyCollFromCmdType, _
                                                                    AssyCollFromFeature)
    Set AssyMergedConduitColl2 = PartCollHlpr.GetInterferencePartsColl(AssyMergedConduitColl1, _
                                                                    AssyCollFromConduits)
    Set AssyCollFromConduits = AssyMergedConduitColl2

    
   
   'Finds out if collection of hangers which can support more than one Conduit.
   If ConduitCollection.Count > 1 Then
        Dim AssycollFromConduitCount As Object
        e_Operator = OPER_EQUAL
        Set AssycollFromConduitCount = AssmSelectionHelper.GetAssyCollByOperator("Supported_Count", _
                                                                                "1+", _
                                                                                e_Operator, _
                                                                                eHgrDiscipline)
        
        
        Set AssyMergedConduitColl = PartCollHlpr.GetInterferencePartsColl(AssyCollFromConduits, _
                                                                        AssycollFromConduitCount)
        Set AssyCollFromConduits = AssyMergedConduitColl
    End If
    
    
    '=======================================
    'get support based on distance
    '=======================================
    Dim TmpAssemblyMergedCollection As Object
    Dim StructPortCollection As Object
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
    
    If (lngSupportingPortCnt > 0) Then
     Dim AssyCollFromLength As Object
    
     Dim oUOMService As IJUomVBInterface
     Dim unitID As Units
     Dim ConduitRadius As Double
     
     Set oUOMService = New UnitsOfMeasureServicesLib.UomVBInterface
     
     On Error GoTo UOMUnitErrorHandler
     unitID = oUOMService.GetUnitId(UNIT_DISTANCE, strUnit)
     On Error GoTo ErrorHandler
     
     ConduitRadius = oUOMService.ConvertUnitToDbu(UNIT_DISTANCE, PrimaryConduitSize, unitID) / 2#
     
     Dim eDistanceType           As HgrPortDistanceType
     Dim RouteStructMinDistance As Double
     Dim RouteStructMaxDistance As Double
     
     eDistanceType = VERTICAL
'     RouteStructDistance = my_IJHgrInputConfigHlpr.GetDistanceBetweenPorts("Route", "Structure", eDistanceType)
     'RouteStructDistance = my_IJHgrInputConfigHlpr.GetDistanceBetweenPortCurves(eDistanceType) ' TR#111337 ; suncor problem
     my_IJHgrInputConfigHlpr.GetDistanceBetweenPortCurves eDistanceType, RouteStructMinDistance, RouteStructMaxDistance  ' TR#111337 problem

     RouteStructMinDistance = RouteStructMinDistance - ConduitRadius
     RouteStructMaxDistance = RouteStructMaxDistance - ConduitRadius
     
     Set AssyCollFromLength = AssmSelectionHelper.GetAssyCollFromBounds("MinAssemblyLength", _
                                                                        "MaxAssemblyLength", _
                                                                        RouteStructMinDistance, _
                                                                        HgrConduitDisciplineType, _
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
    
    Set IJPort_Route = ConduitCollection.Item(1)
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
                RouteStructMinDistance = RouteStructMinDistance - ConduitRadius
                Set AssyCollFromNextPort = AssmSelectionHelper.GetAssyCollFromBounds("MinAssemblyLength", _
                                                                            "MaxAssemblyLength", _
                                                                            RouteStructMinDistance, _
                                                                            HgrConduitDisciplineType, , RouteStructMaxDistance)
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
     
     Set TmpAssemblyMergedCollection = PartCollHlpr.GetInterferencePartsColl(AssyCollFromLength, AssyCollFromConduits)
     
     Set AssyCollFromConduits = Nothing
     Set AssyCollFromConduits = TmpAssemblyMergedCollection
     Set IJPort_Route = Nothing
     Set IJPort_Struct = Nothing
    End If
    
    Set TmpAssemblyMergedCollection = Nothing
    Set AssyCollFromLength = Nothing
   '***************************************************************************************
   'Get the Collection of Assemblies which support the current Port Configuration.
   '        Only check against results already returned by the Conduit size filter.
   '***************************************************************************************
   Dim AssyCollFromConfigAndConduit As Object
   
   'Get all structures on the support.  In particular we are looking for the
   'port on the input object that is being used by the support.
    
    'If there is no Struct Collection go to next filter.
    Dim AssyCollFromSupporting As Object
    Dim AssyFromConduitAndBeam As Object
    
    e_Operator = OPER_EQUAL
'    If StructPortCollection Is Nothing Then
'        Set AssyCollFromSupporting = AssmSelectionHelper.GetAssyCollByOperator("Supporting_Count", _
'                                                                                "0", _
'                                                                                e_Operator, _
'                                                                                eHgrDiscipline)
'
'        Set AssyFromConduitAndBeam = PartCollHlpr.GetInterferencePartsColl(AssyCollFromSupporting, _
'                                                                         AssyCollFromConduits)
'        Set AssyCollFromConfigAndConduit = AssyFromConduitAndBeam
'        GoTo NON_STRUCT_FILTER
'    End If
    
    'If there are no items in the returned collection go to next filter.
    If lngSupportingCnt = 0 Then ' CR#128332
        Set AssyCollFromSupporting = AssmSelectionHelper.GetAssyCollByOperator("Supporting_Count", _
                                                                                "0", _
                                                                                e_Operator, _
                                                                                eHgrDiscipline)
        
        Set AssyFromConduitAndBeam = PartCollHlpr.GetInterferencePartsColl(AssyCollFromSupporting, _
                                                                         AssyCollFromConduits)
        Set AssyCollFromConfigAndConduit = AssyFromConduitAndBeam
        GoTo NON_STRUCT_FILTER
    Else
        Dim AssyCollFromSupporting1 As Object
        Dim AssyCollFromSupporting2 As Object
        Set AssyCollFromSupporting1 = AssmSelectionHelper.GetAssyCollByOperator("Supporting_Count", _
                                                                                "1", _
                                                                                e_Operator, _
                                                                                eHgrDiscipline)
        
        Set AssyCollFromSupporting2 = AssmSelectionHelper.GetAssyCollByOperator("Supporting_Count", _
                                                                                "1+", _
                                                                                e_Operator, _
                                                                                eHgrDiscipline)
            Set AssyCollFromSupporting = PartCollHlpr.GetMergedCollection(AssyCollFromSupporting1, _
                                                                        AssyCollFromSupporting2)
                                                                                
            Set AssyFromConduitAndBeam = PartCollHlpr.GetInterferencePartsColl(AssyCollFromSupporting, _
                                                                             AssyCollFromConduits)
        If lngSupportingPortCnt = 0 Then GoTo NON_STRUCT_FILTER ' CR#128332
    End If
    
    'Get the primary Conduit Port.
    'This is a port on the primary Conduit.  Since Conduit Object and Conduit Port are
    'the same object we are able to use the first item returned
    'by GetSupportedObjects() rather than GetSupportedObjectPorts()
    Dim IJPort_Conduit As IJPort
    Set IJPort_Conduit = ConduitCollection.Item(1)
    
    'Get the first item in the Struct Port Collection.
    'This is a port on the primary supporting object (beam or slab).
    ' Dim IJPort_Struct As IJPort
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

        
    On Error Resume Next
    PortConfig = PortPostureService.GetGenericPortConfig(my_IJHgrInputObjectInfo, _
                                            IJPort_Conduit, _
                                            IJPort_Struct, _
                                            IJDPosition_Target, _
                                            eHgrCmdType)
                                            
    'Clear un-used items.
    
    On Error GoTo ErrorHandler
    
    Set PortPostureService = Nothing
    
    'Only search for matching assemblies if the Configuration is NOT zero.
    If (PortConfig <> 0) And (Not AssyFromConduitAndBeam Is Nothing) Then
        'Get a collection of Catalog Parts with a Configuration Matching that
        'returned by the Port Posture Service.  Only check against the items already
        'returned by the Conduit filter.
        Set AssyCollFromConfigAndConduit = AssmSelectionHelper.GetAssyCollByOperator( _
                                                "TypeSelectionRule", _
                                                PortConfig, OPER_BITWISE_AND, _
                                                eHgrDiscipline, _
                                                AssyFromConduitAndBeam)
                                                
    Else
        Set AssyCollFromConfigAndConduit = AssyCollFromConduits
    End If
    
    Set AssyCollFromConfigAndConduit = AssmSelectionHelper.GetAssyCollFromSupportingFace(IJPort_Struct, pInputObjectInfo, AssyCollFromConfigAndConduit)
    
    '***************************************************************************************
NON_STRUCT_FILTER:

    ' Merge the Three Lists.
    ' If a list is empty, it is ignored when merging.
    
    Dim AssemblyMergedCollection As Object
    
    'The method needs to return IJElements.  Get this off of the object returned.
    Dim IJElements_AssemblyCollection As HNGSUPPartSelectionHlpr.IJElements
    Set IJElements_AssemblyCollection = AssyCollFromConfigAndConduit
    
    'Return the Collection of Assembly Definitions
    Set IJHgrAssmSelectionRule_GetAssemblies = IJElements_AssemblyCollection
    
    Set AssyCollFromFeature = Nothing
    Set AssyCollFromCmdType = Nothing
   ' Set AssyCollFromAllCmdType = Nothing
    Set AssmSelectionHelper = Nothing
    Set AssyCollFromConduits = Nothing
    Set AssyCollFromConfigAndConduit = Nothing
    Set PartCollHlpr = Nothing
    Set AssemblyMergedCollection = Nothing
    Set IJDPosition_Target = Nothing
    Set StructPortCollection = Nothing
    Set HgrComServ = Nothing
    Set ConduitCollection = Nothing
    
    Exit Function
    
UOMUnitErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "GetUnitId method of IJUOMVBInterface failed.Support type list may not be available.Could be because of localised string mismatch").Number
    
ErrorHandler:
     Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function




