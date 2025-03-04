VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAssemblySelectionRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CAssemblySelectionRule.cls
'   ProgID:         HgrAssmSelRule.CAssemblySelectionRule
'   Author:         Amlan
'   Creation Date:  11.Feb.2002
'   Description:
'
'   Change History:
'       11.Feb.2002             Amlan           Creation Date
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CAssemblySelectionRule" 'Used for error messages

Implements IJHgrAssmSelectionRule

Public Function IJHgrAssmSelectionRule_GetAssemblies(ByVal pInputObjectInfo As Object) As HNGSUPPartSelectionHlpr.IJElements
Const METHOD = "IJHgrAssmSelectionRule_GetAssemblies"

    On Error GoTo ErrorHandler
    
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
    'Filter based on Pipe Size
    '*********************************************************************************************
    Dim AssyCollFromPipes As Object
    
    'Get A Common Services Object
    Dim HgrComServ As New HgrCommonServices
    'Get An Assembly Selection Helper
    Dim AssmSelectionHelper As New AssySelectionHlpr
    
    'Get the Size of the Primary Pipe
    Dim PrimaryPipeSize As Double
    Dim strUnit As String
    PrimaryPipeSize = HgrComServ.PrimaryPipeDiameter(PipeCollection, strUnit)
    
    'Get a collection of Catalog Parts with a Min and Max Pipe Size bound that
    'of the Primary Pipe.
'    Set AssyCollFromPipes = AssmSelectionHelper.GetPartCollFromBounds("NDFrom", _
'                                                                       "NDTo", _
'                                                                       PrimaryPipeSize)

    'Get the command type
    Dim eHgrCmdType As HgrCmdType
    eHgrCmdType = my_IJHgrInputObjectInfo.GetCommandType
    
    Set AssyCollFromPipes = AssmSelectionHelper.GetHgrSupportDefByND(eHgrCmdType, HgrAllDisciplineType, _
                             vbNullString, PrimaryPipeSize, strUnit)

    If Not AssyCollFromPipes Is Nothing Then
        If AssyCollFromPipes.Count = 0 Then
            Set AssyCollFromPipes = Nothing
        End If
    End If
                                                      
   Dim PartCollHlpr As New HgrCatPartsCollHlpr
   
   
   'Finds out if collection of hangers which can support more than one pipe.
   If PipeCollection.Count > 1 Then
        Dim AssycollFromPipeCount As Object
        Dim e_Operator As HgrOperatorType
        e_Operator = OPER_EQUAL
        Set AssycollFromPipeCount = AssmSelectionHelper.GetPartCollByOperator("Supported_Count", "1+", e_Operator)
        
        
        
        If Not AssycollFromPipeCount Is Nothing Then

            Dim AssyMergedPipeColl As Object
            Set AssyMergedPipeColl = PartCollHlpr.GetInterferencePartsColl(AssyCollFromPipes, AssycollFromPipeCount)
            Set AssyCollFromPipes = AssyMergedPipeColl

        End If
            
    End If
        
    
   '***************************************************************************************
   'Get the Collection of Assemblies which support the current Port Configuration.
   '        Only check against results already returned by the pipe size filter.
   '***************************************************************************************
   Dim AssyCollFromConfigAndPipe As Object
   
   'Get all structures on the support.  In particular we are looking for the
   'port on the input object that is being used by the support.
    Dim StructPortCollection As Object
    Set StructPortCollection = my_IJHgrInputObjectInfo.GetSupportingObjectPorts
    
    'If there is no Struct Collection go to next filter.
    If StructPortCollection Is Nothing Then
        Set AssyCollFromConfigAndPipe = AssyCollFromPipes
        GoTo NON_STRUCT_FILTER
    End If
    
    'If there are no items in the returned collection go to next filter.
    If StructPortCollection.Count = 0 Then
        Set AssyCollFromConfigAndPipe = AssyCollFromPipes
        GoTo NON_STRUCT_FILTER
    End If
    
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
'    PortPostureService.GetPortConfiguration IJPort_Pipe, _
'                                            IJPort_Struct, _
'                                            IJDPosition_Target, _
'                                            eHgrCmdType, _
'                                            bUseFiniteGeometry, _
'                                            PortConfig
'
'    If (PortConfig = 0) Then
'        'Didn't find a configuration using finite geometry.
'        'Try again with infinite geometry
'        bUseFiniteGeometry = False
'
'        PortPostureService.GetPortConfiguration IJPort_Pipe, _
'                                            IJPort_Struct, _
'                                            IJDPosition_Target, _
'                                            eHgrCmdType, _
'                                            bUseFiniteGeometry, _
'                                            PortConfig
'    End If
        
    PortConfig = PortPostureService.GetGenericPortConfig(my_IJHgrInputObjectInfo, _
                                            IJPort_Pipe, _
                                            IJPort_Struct, _
                                            IJDPosition_Target, _
                                            eHgrCmdType)
                                            
    'Clear un-used items.
    Set PortPostureService = Nothing
    
    'Only search for matching assemblies if the Configuration is NOT zero.
    If (PortConfig <> 0) Then
        'Get a collection of Catalog Parts with a Configuration Matching that
        'returned by the Port Posture Service.  Only check against the items already
        'returned by the Pipe filter.
        Set AssyCollFromConfigAndPipe = AssmSelectionHelper.GetPartCollByOperator( _
                                                "TypeSelectionRule", _
                                                PortConfig, OPER_BITWISE_AND, _
                                                AssyCollFromPipes)
                                                
        'Null the collection if it contains no items
        If Not AssyCollFromConfigAndPipe Is Nothing Then
            If AssyCollFromConfigAndPipe.Count = 0 Then
                Set AssyCollFromConfigAndPipe = Nothing
            End If
        End If
    Else
        Set AssyCollFromConfigAndPipe = AssyCollFromPipes
    End If
    
    Set AssyCollFromConfigAndPipe = AssmSelectionHelper.GetAssyCollFromSupportingFace(IJPort_Struct, pInputObjectInfo, AssyCollFromConfigAndPipe)
    
    
   '***************************************************************************************
   'Filter based on the Family of the beam cross section.
   '***************************************************************************************
'    Dim AssyCollFromBeams As Object
'
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
    Set AssemblyMergedCollection = AssyCollFromConfigAndPipe
    
    'Merge
    '   AssyCollFromConfigAndPipe
    'with
    '   AssyCollFromBeams
    'into
    '   AssemblyMergedCollection
    
    
'    If AssyCollFromConfigAndPipe Is Nothing And AssyCollFromBeams Is Nothing Then
'        'Both Empty
'        Set AssemblyMergedCollection = Nothing
'    ElseIf AssyCollFromBeams Is Nothing Then
'        'No Beam Results
'        Set AssemblyMergedCollection = AssyCollFromConfigAndPipe
'    ElseIf AssyCollFromConfigAndPipe Is Nothing Then
'        'No Pipe and Configuration Results
'        Set AssemblyMergedCollection = AssyCollFromBeams
'    Else
'        'Merge Beam, Pipe, and Configuration Results
'        Set AssemblyMergedCollection = _
'                  PartCollHlpr.GetInterferencePartsColl(AssyCollFromConfigAndPipe, AssyCollFromBeams)
'    End If
    
    'The method needs to return IJElements.  Get this off of the object returned.
    Dim IJElements_AssemblyCollection As HNGSUPPartSelectionHlpr.IJElements
    Set IJElements_AssemblyCollection = AssemblyMergedCollection
    
    'Return the Collection of Assembly Definitions
    Set IJHgrAssmSelectionRule_GetAssemblies = IJElements_AssemblyCollection
    
    Set AssmSelectionHelper = Nothing
    Set AssyCollFromPipes = Nothing
    Set AssyCollFromConfigAndPipe = Nothing
    Set PartCollHlpr = Nothing
    Set AssemblyMergedCollection = Nothing
    Set IJDPosition_Target = Nothing
    Set StructPortCollection = Nothing
    Set HgrComServ = Nothing
    Set PipeCollection = Nothing
    
    Exit Function
    
ErrorHandler:
     Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function
