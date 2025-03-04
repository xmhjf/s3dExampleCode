VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPipeSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPipeSupport.cls
'   ProgID:         HgrAssmSelRule.CPipeSupport
'   Author:         Amlan
'   Creation Date:  11.Feb.2002
'   Description:
'    AssemblySelectionRule for PipeSupports.
'    This ASR will be used to select Assembly based on SupportDiscipline
'
'   Change History:
'               09.Apr.2002             Caroline        Select Assembly based on the Feature of the Support
'                                                                      straight, turn etc.
'               24.Jun.2002             Caroline        Select Assembly based on the CommandType of the Support
'
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "CPipeSupport" 'Used for error messages

Implements IJHgrAssmSelectionRule

Public Function IJHgrAssmSelectionRule_GetAssemblies(ByVal pInputObjectInfo As Object) As HNGSUPPartSelectionHlpr.IJElements
Const METHOD = "IJHgrAssmSelectionRule_GetAssemblies"
On Error GoTo ErrorHandler
   
    'Set the discipline type to be used while calling the services.
    Dim eHgrDiscipline As GSCADHgrSupFilterServices.HgrDisciplineType
    eHgrDiscipline = HgrPipingDisciplineType
    
    'Get IJHgrInputConfig Hlpr Interface off of passed Helper
    Dim my_IJHgrInputObjectInfo As IJHgrInputObjectInfo
    Set my_IJHgrInputObjectInfo = pInputObjectInfo
    
    Dim my_IJHgrInputConfigHlpr As IJHgrInputConfigHlpr
    Set my_IJHgrInputConfigHlpr = pInputObjectInfo

    'Get all pipes on the support.
    Dim PipeCollection As Object
    On Error Resume Next
    Set PipeCollection = my_IJHgrInputObjectInfo.GetSupportedObjects
    On Error GoTo ErrorHandler
    
    'If there is no Pipes Collection exit.
    If PipeCollection Is Nothing Then
        Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
        Exit Function
    End If

    'If there are no Pipes in the returned collection exit.
    If PipeCollection.Count = 0 Then
        Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
        Exit Function
    End If

    '*********************************************************************************************
    'Filter based on Pipe Size, command type, discipline type
    '*********************************************************************************************
    Dim AssyCollFromPipeDia As Object
    'Get A Common Services Object
    Dim HgrComServ As New HgrCommonServices
    'Get An Assembly Selection Helper
    Dim AssmSelectionHelper As New AssySelectionHlpr
    
    'Get the Size of the Primary Pipe
    Dim PrimaryPipeSize As Double
    Dim strUnit As String
    PrimaryPipeSize = HgrComServ.PrimaryPipeDiameter(PipeCollection, strUnit)
    '
    Dim strFeatureType As String
    Dim varFeature() As Variant
     'Get a Collection of catalog Parts based on PipeFeature
    varFeature = my_IJHgrInputObjectInfo.GetAttributeByRule("HgrSupFeatureType")
    strFeatureType = varFeature(0)

    Dim PartCollHlpr                As New HgrCatPartsCollHlpr
    
    Dim lInsulationPurpose       As Long
    Dim lInsulationMaterial      As Long
    Dim dInsulationThickness     As Double
    Dim bIsInsulated             As Boolean
    Dim iLoop                    As Integer
    Dim bFlag                    As Boolean
    Dim AssyCollFromInsualtion   As Object
    bFlag = True

''' Uncomment following code if Insulation Purpose is needed to be considered for filtering out assemblies
'    For iLoop = 1 To PipeCollection.Count
'        my_IJHgrInputObjectInfo.GetInsulationData iLoop, lInsulationPurpose, lInsulationMaterial, dInsulationThickness, bIsInsulated
'
'        If bIsInsulated <> True Then
'            bFlag = False
'        End If
'    Next
'
'    If bFlag <> False Then
'     'Get assemblies based on the insulation criteria
'        Set AssyCollFromInsualtion = AssmSelectionHelper.GetAssembliesByInsulationInfo(lInsulationPurpose)
'    Else
'        lInsulationPurpose = 2
'        Set AssyCollFromInsualtion = AssmSelectionHelper.GetAssembliesByInsulationInfo(lInsulationPurpose)
'    End If
   ''' Uncomment above code if Insulation Purpose is needed to be considered for filtering out assemblies
   
    Dim lngCommandType              As Long
    lngCommandType = my_IJHgrInputObjectInfo.GetCommandType

    Set AssyCollFromPipeDia = AssmSelectionHelper.GetHgrSupportDefByND(lngCommandType, eHgrDiscipline, _
                             strFeatureType, PrimaryPipeSize, strUnit)
                             
'''Uncomment following code if Insulation Purpose is needed to be considered for filtering out assemblies
'    Dim AssyMergedInsulationColl As Object
'    If Not AssyCollFromInsualtion Is Nothing Then
'        'If AssyCollFromInsualtion.Count > 0 Then ' this check is not required
'        Set AssyMergedInsulationColl = PartCollHlpr.GetInterferencePartsColl(AssyCollFromPipeDia, _
'                                                                            AssyCollFromInsualtion)
'        Set AssyCollFromPipeDia = AssyMergedInsulationColl
'        'End If ' this check is not required
'   End If
''' Uncomment above code if Insulation Purpose is needed to be considered for filtering out assemblies


   '=====================================================
   'Filter based on Min/MaxAssemblyLength
   'First route object is taken as reference to calculate the distance
   'between route and structure
   '=====================================================
    Dim AssemblyMergedCollection As Object
    Set AssemblyMergedCollection = Nothing
    
    'Get all structures on the support.  In particular we are looking for the
   'port on the input object that is being used by the support.
   
    Dim StructPortCollection As Object
    Dim lngSupportingPortCnt As Long
    
    If lngCommandType = HgrByReferencePointCmdType Then
        Set StructPortCollection = my_IJHgrInputConfigHlpr.GetUserReferencePort
    Else
        Set StructPortCollection = my_IJHgrInputObjectInfo.GetSupportingObjectPorts
    End If
    
    If Not StructPortCollection Is Nothing Then
        lngSupportingPortCnt = StructPortCollection.Count
    Else
         lngSupportingPortCnt = 0
    End If
        
    ' CR#128332 ; get the supporting collection which includes surfaces also
    Dim SupportingCollection As Object
    Dim lngSupportingCnt As Long
    
    If lngCommandType = HgrByReferencePointCmdType Then
        Set SupportingCollection = my_IJHgrInputConfigHlpr.GetUserReferencePort
    Else
        Set SupportingCollection = my_IJHgrInputObjectInfo.GetSupportingObjects
    End If
    
    If Not SupportingCollection Is Nothing Then
            lngSupportingCnt = SupportingCollection.Count
    Else
        lngSupportingCnt = 0
    End If
    
    Set SupportingCollection = Nothing
    ' CR#128332 ends here
    
   '***************************************************************************************
   'Get the Collection of Assemblies based on the input count.
   ' Supporting count and Supported count
   '***************************************************************************************
    Dim AssyFromPipeAndBeam As Object
    Set AssyFromPipeAndBeam = AssmSelectionHelper.GetAssembliesByInput(lngCommandType, eHgrDiscipline, strFeatureType, PipeCollection.Count, lngSupportingCnt)
    
    Set AssemblyMergedCollection = PartCollHlpr.GetInterferencePartsColl(AssyFromPipeAndBeam, AssyCollFromPipeDia)

    If lngSupportingPortCnt = 0 Then GoTo NON_STRUCT_FILTER
    
        '***************************************************************************************
        'Get the Collection of Assemblies which support the current Port Configuration.
        '        Only check against results already returned by the pipe size filter.
        '***************************************************************************************
        'Get the primary Pipe Port.
        'This is a port on the primary pipe.  Since Pipe Object and Pipe Port are
        'the same object we are able to use the first item returned
        'by GetSupportedObjects() rather than GetSupportedObjectPorts()
        Dim IJPort_Pipe As IJPort
        Set IJPort_Pipe = PipeCollection.Item(1)
    
        'Get the first item in the Struct Port Collection.
        'This is a port on the primary supporting object (beam or slab).
        Dim IJPort_Struct As IJPort
        
        Set IJPort_Struct = StructPortCollection.Item(1)
        'Get the Target Point.
        Dim IJDPosition_Target As IJDPosition
        Set IJDPosition_Target = my_IJHgrInputObjectInfo.GetTargetPoint
    
        'Get a Port Configuration Service.
        Dim PortPostureService As New HgrPortPostureService
    
        'Ask the Port Configuration Service for the Configuration of the Ports
        Dim bUseFiniteGeometry As Boolean
        bUseFiniteGeometry = True
        Dim PortConfig As Long
        
        On Error Resume Next
        PortConfig = PortPostureService.GetGenericPortConfig(my_IJHgrInputObjectInfo, _
                                            IJPort_Pipe, _
                                            IJPort_Struct, _
                                            IJDPosition_Target, _
                                            lngCommandType)
                                            

        'Clear un-used items.
        On Error GoTo ErrorHandler
    
        Set IJPort_Pipe = Nothing
        Set IJDPosition_Target = Nothing
        Set PortPostureService = Nothing
    
        'Only search for matching assemblies if the Configuration is NOT zero.
        If (PortConfig <> 0) Then
            'Get a collection of Catalog Parts with a Configuration Matching that
            'returned by the Port Posture Service.
            Dim AssyCollFromConfigAndPipe As Object
            Set AssyCollFromConfigAndPipe = AssmSelectionHelper.GetAssembliesByTypeSelRule(PortConfig)
            Set AssemblyMergedCollection = PartCollHlpr.GetInterferencePartsColl(AssyCollFromConfigAndPipe, AssemblyMergedCollection)
        
            AssyCollFromConfigAndPipe.Clear
            Set AssyCollFromConfigAndPipe = Nothing
        Else
            Set AssyCollFromConfigAndPipe = AssemblyMergedCollection
        End If
        
        If (lngSupportingPortCnt > 0) Then
            Dim AssyCollFromLength As Object
    
            Dim oUOMService As IJUomVBInterface
            Dim unitID As Units
            Dim PipeRadius As Double
     
            Set oUOMService = New UnitsOfMeasureServicesLib.UomVBInterface
     
            On Error GoTo UOMUnitErrorHandler
            unitID = oUOMService.GetUnitId(UNIT_DISTANCE, strUnit)
            On Error GoTo ErrorHandler
     
            PipeRadius = oUOMService.ConvertUnitToDbu(UNIT_DISTANCE, PrimaryPipeSize, unitID) / 2#
     
            Dim eDistanceType           As HgrPortDistanceType
            'Dim RouteStructDistance  As Double
            Dim RouteStructMinDistance As Double
            Dim RouteStructMaxDistance As Double
     
            eDistanceType = VERTICAL
     
            'need to handle intersection case
            If PortConfig = 8 Then
                RouteStructMinDistance = 0#
            Else
                If PortConfig = 4 Then
                eDistanceType = HORIZONTAL
            End If
            If strFeatureType = "TURN" Then
                eDistanceType = VERTICAL
                'RouteStructDistance = my_IJHgrInputConfigHlpr.GetDistanceBetweenPorts("TurnRef", "Structure", eDistanceType)
                'RouteStructDistance = my_IJHgrInputConfigHlpr.GetDistanceBetweenPortCurves(eDistanceType) ' TR#111337 problem
                my_IJHgrInputConfigHlpr.GetDistanceBetweenPortCurves eDistanceType, RouteStructMinDistance, RouteStructMaxDistance ' TR#111337 problem
            Else
                'RouteStructDistance = my_IJHgrInputConfigHlpr.GetDistanceBetweenPorts("Route", "Structure", eDistanceType)
                'RouteStructDistance = my_IJHgrInputConfigHlpr.GetDistanceBetweenPortCurves(eDistanceType) ' TR#111337 problem
                my_IJHgrInputConfigHlpr.GetDistanceBetweenPortCurves eDistanceType, RouteStructMinDistance, RouteStructMaxDistance  ' TR#111337 problem
            End If
            RouteStructMinDistance = RouteStructMinDistance - PipeRadius
            RouteStructMaxDistance = RouteStructMaxDistance - PipeRadius
        End If
     
        Set AssyCollFromLength = AssmSelectionHelper.GetAssyCollFromBounds("MinAssemblyLength", _
                                                                        "MaxAssemblyLength", _
                                                                        RouteStructMinDistance, _
                                                                        HgrPipingDisciplineType, _
                                                                        , RouteStructMaxDistance)
                                                                        
        Set AssemblyMergedCollection = PartCollHlpr.GetInterferencePartsColl(AssyCollFromLength, AssemblyMergedCollection)
    
        If Not (lngCommandType = HgrByReferencePointCmdType) Then
            ' TRs#87104, 125406, 125477 ; now build a union of assemblies list for all the valid supporting faces for the Support Type
            Dim iNextStructPort As IJPort, iStructPort As IJPort
            Dim iConnObjHlpr As IJHgrConnectableObjHlpr
            Dim iPortHlpr As IJHgrPortHlpr
            Dim AssyCollFromNextPort As Object
            'oPortColl contains all ports of given cross section type
            Dim oPortColl As IJElements
            Set oPortColl = New JObjectCollection
        
            'get the port from structure
            oPortColl.Add IJPort_Struct
            
            Set IJPort_Pipe = PipeCollection.Item(1)
            Set iConnObjHlpr = CreateObject("HgrSupFilterServices.HgrConnectableObjHlpr")
            Set iPortHlpr = iConnObjHlpr.GetConnectableObjPortService(IJPort_Struct.Connectable)
            If (Not (iPortHlpr Is Nothing)) Then
                Set iNextStructPort = iPortHlpr.GetNextPort(IJPort_Struct, my_IJHgrInputObjectInfo, HgrIncrement)
            
                If (Not (iNextStructPort Is Nothing)) Then
                    While (Not (oPortColl.Contains(iNextStructPort)))
                        my_IJHgrInputConfigHlpr.GetDistBetweenSpecifiedPortCurves IJPort_Pipe, iNextStructPort, eDistanceType, RouteStructMinDistance, RouteStructMaxDistance
                        RouteStructMinDistance = RouteStructMinDistance - PipeRadius
                        Set AssyCollFromNextPort = AssmSelectionHelper.GetAssyCollFromBounds("MinAssemblyLength", _
                                                                "MaxAssemblyLength", _
                                                                RouteStructMinDistance, _
                                                                HgrPipingDisciplineType, , RouteStructMaxDistance)
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
        
            Set AssemblyMergedCollection = AssmSelectionHelper.GetAssyCollFromSupportingFace(IJPort_Struct, pInputObjectInfo, AssemblyMergedCollection)
         
            Set IJPort_Struct = Nothing
            Set IJPort_Pipe = Nothing ' TRs#87104
     
        End If
    End If
    
    '***************************************************************************************
NON_STRUCT_FILTER:

    '========================================================================================'
    ' Next step, is to order the support definition based on its old definition custom number
    ' This need to be done only during modification of the support.
    '========================================================================================'
    Dim IsPlacement As Boolean
    Dim oInputConfigHlpr As IJHgrInputConfigHlpr
    Set oInputConfigHlpr = my_IJHgrInputObjectInfo
    IsPlacement = oInputConfigHlpr.IsPlacement
    If Not AssemblyMergedCollection Is Nothing And IsPlacement = False Then
        Dim oSupportDefinition As Object
        Set oSupportDefinition = my_IJHgrInputObjectInfo.GetSupportDefinition
        If Not oSupportDefinition Is Nothing Then
            
            Dim varOldValue            As Variant
            
            varOldValue = HgrComServ.GetPropertyFromPart(oSupportDefinition, 1, "Number1")
            If (varOldValue > 0) Then
                ' Here we implement the logic to order the return collection
                ' such that the return collection contains the definitions of the
                ' same type as the active one in the begining.
                Dim oMatchingDefinition     As IJElements
                Dim oNonMatchingDefinition  As IJElements
                
                Set oMatchingDefinition = New JObjectCollection
                Set oNonMatchingDefinition = New JObjectCollection
                
                Dim oIJElements  As IJElements
                Set oIJElements = AssemblyMergedCollection
                
                For iLoop = 1 To oIJElements.Count
                    Dim oDefItem As Object
                    Dim varNewValue As Variant
                    
                    Set oDefItem = oIJElements.Item(iLoop)
                    varNewValue = HgrComServ.GetPropertyFromPart(oDefItem, 1, "Number1")
                    
                    If varNewValue = varOldValue Then
                        ' Add to the matching collection
                        oMatchingDefinition.Add oDefItem
                    Else
                        ' Add to the non-matching collection
                        oNonMatchingDefinition.Add oDefItem
                    End If
                Next
                
                ' merge the matching and non-matching, with matching in the front
                oMatchingDefinition.AddElements oNonMatchingDefinition
                Set AssemblyMergedCollection = oMatchingDefinition
            End If
        End If
    End If
    ' Return the Collection of Assembly Definitions
    
    Dim strAttribName As String
    Dim strInterfaceName As String

    strAttribName = "RulePriority"
    strInterfaceName = "IJUAHgrRulePriority"
               
    On Error Resume Next
    Set AssemblyMergedCollection = AssmSelectionHelper.GetSortedAssmCollByAttribute(strAttribName, strInterfaceName, AssemblyMergedCollection) 'This API returns the error E_MD_OBJECTNOTINMETADATA
    ' if the input interface does not exist in the metadata.
    
    Set IJHgrAssmSelectionRule_GetAssemblies = AssemblyMergedCollection
    
    If Not AssyCollFromPipeDia Is Nothing Then
        AssyCollFromPipeDia.Clear
        Set AssyCollFromPipeDia = Nothing
    End If
    
    If Not AssyFromPipeAndBeam Is Nothing Then
        AssyFromPipeAndBeam.Clear
        Set AssyFromPipeAndBeam = Nothing
    End If
    
    If Not AssyCollFromLength Is Nothing Then
        AssyCollFromLength.Clear
        Set AssyCollFromLength = Nothing
    End If
    
    Set AssmSelectionHelper = Nothing
    Set AssyCollFromConfigAndPipe = Nothing
    Set PartCollHlpr = Nothing
    Set AssemblyMergedCollection = Nothing
    Set IJDPosition_Target = Nothing
    Set StructPortCollection = Nothing
    Set HgrComServ = Nothing
    Set PipeCollection = Nothing
    Set my_IJHgrInputObjectInfo = Nothing
    
    
    
    Exit Function
UOMUnitErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "GetUnitId method of IJUOMVBInterface failed.Support type list may not be available.Could be because of localised string mismatch").Number

ErrorHandler:

    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function



