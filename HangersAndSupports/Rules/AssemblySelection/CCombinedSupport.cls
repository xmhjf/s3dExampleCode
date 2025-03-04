VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCombinedSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CCombinedSupport.cls
'   ProgID:         HgrAssmSelRule.CCombinedSupport
'   Author:         Yan Ji
'   Creation Date:  12.Mar.2003
'   Description:
'    AssemblySelectionRule for CombinedSupport.
'    This ASR will be used to select Assembly based on SupportDiscipline
'
'   Change History:
'       12.Mar.2003             Yan Ji          Creation Date
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "CCombinedSupport" 'Used for error messages

Implements IJHgrAssmSelectionRule

Public Function IJHgrAssmSelectionRule_GetAssemblies(ByVal pInputObjectInfo As Object) As HNGSUPPartSelectionHlpr.IJElements
Const METHOD = "IJHgrAssmSelectionRule_GetAssemblies"

    On Error GoTo ErrorHandler
    
    'Set the discipline type to be used while calling the services.
    Dim eHgrDiscipline As GSCADHgrSupFilterServices.HgrDisciplineType
    eHgrDiscipline = HgrCombinedDisciplineType
    
    'Get IJHgrInputConfig Hlpr Interface off of passed Helper
    Dim my_IJHgrInputObjectInfo As IJHgrInputObjectInfo
    Set my_IJHgrInputObjectInfo = pInputObjectInfo
    
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
    Set AssyCollFromPipeDia = Nothing
    Dim AssmSelectionHelper As New AssySelectionHlpr

    Dim strFeatureType As String
    Dim varFeature() As Variant
'     'Get a Collection of catalog Parts based on PipeFeature
    varFeature = my_IJHgrInputObjectInfo.GetAttributeByRule("HgrSupFeatureType")
    strFeatureType = varFeature(0)

    Dim PartCollHlpr                As New HgrCatPartsCollHlpr
    Dim lngCommandType              As Long

    lngCommandType = my_IJHgrInputObjectInfo.GetCommandType

   '***************************************************************************************
   'Get the Collection of Assemblies based on the input count.
   ' Supporting count and Supported count
   '***************************************************************************************
    Dim AssemblyMergedCollection As Object

   'Get all structures on the support.  In particular we are looking for the
   'port on the input object that is being used by the support.
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
    
    Dim AssyFromPipeAndBeam As Object
    Set AssyFromPipeAndBeam = AssmSelectionHelper.GetAssembliesByInput(lngCommandType, eHgrDiscipline, strFeatureType, PipeCollection.Count, lngSupportingCnt)
    Set AssemblyMergedCollection = AssyFromPipeAndBeam 'PartCollHlpr.GetInterferencePartsColl(AssyFromPipeAndBeam, AssyCollFromPipeDia)

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
    
    Set AssyCollFromConfigAndPipe = AssmSelectionHelper.GetAssyCollFromSupportingFace(IJPort_Struct, pInputObjectInfo, AssyCollFromConfigAndPipe)
    
    Set IJPort_Struct = Nothing
    '***************************************************************************************
NON_STRUCT_FILTER:

    ' Return the Collection of Assembly Definitions
    Set IJHgrAssmSelectionRule_GetAssemblies = AssemblyMergedCollection
    
    'AssyCollFromPipeDia.Clear
    'Set AssyCollFromPipeDia = Nothing
    AssyFromPipeAndBeam.Clear
    Set AssyFromPipeAndBeam = Nothing
        
    Set AssmSelectionHelper = Nothing
    Set AssyCollFromConfigAndPipe = Nothing
    Set PartCollHlpr = Nothing
    Set AssemblyMergedCollection = Nothing
    Set IJDPosition_Target = Nothing
    Set StructPortCollection = Nothing
'    Set HgrComServ = Nothing
    Set PipeCollection = Nothing
    Set my_IJHgrInputObjectInfo = Nothing
    
    Exit Function
    
ErrorHandler:
     Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function




