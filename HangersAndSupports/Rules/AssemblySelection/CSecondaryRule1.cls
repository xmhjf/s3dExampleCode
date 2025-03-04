VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSecondaryRule1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSecondaryRule1.cls
'   ProgID:         HgrAssmSelRule.CSecondaryRule1
'   Author:         Amlan
'   Creation Date:  28.May.2002
'   Description:
'    AssemblySelectionRule for Secondary Supports for pipes.
'    This ASR will be used to select specific assemblies for pipes for the secondary supports.
'    The collection returned will consist of just one kind of support.
'
'   Change History:
'   28.May.2002         Amlan           Creation Date
'
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CSecondaryRule1" 'Used for error messages
Implements IJHgrAssmSelectionRule

Private Function IJHgrAssmSelectionRule_GetAssemblies(ByVal pInputObjectInfo As Object) As HNGSUPPartSelectionHlpr.IJElements
Const METHOD = "IJHgrAssmSelectionRule_GetAssemblies"

        On Error GoTo ErrorHandler
        
        'Declare all the necessary variables used in this implementation.
        
        Dim my_InputConfigHlpr      As IJHgrInputConfigHlpr  'for all input information and services
        Dim PortPostureService      As New HgrPortPostureService
        Dim my_InputObjectInfo      As IJHgrInputObjectInfo
        
        'Collections being used
        Dim PipeCollection          As Object
        Dim StructPortCollection    As Object
        Dim IJDPosition_Target      As IJDPosition
        Dim eHgrCmdType             As HgrCmdType
        
        'Pipe and structure ports.
        Dim IJPort_Struct           As IJPort
        Dim IJPort_Pipe             As IJPort
        
        'Measurements
        Dim PipeRadius              As Double
        Dim RouteStructDistance     As Double
        Dim RouteStructMaxDistance     As Double
        Dim PortConfig              As Long
        Dim eDistanceType           As HgrPortDistanceType
        
        'Temporary variables
        Dim oAssycollection         As IJElements
        Set oAssycollection = New JObjectCollection
        Dim strAssyName             As String
        Dim oAssyPart               As Object
        Dim Tolerance               As Double
        
        Dim RouteIndex              As Long
        Dim lInsulPurpose           As Long
        Dim lInsulMat               As Long
        Dim dInsulThk               As Double
        Dim bInsulFlag              As Boolean
        Dim dblPipeDia              As Double

'        Dim RouteIndex As Long
        
        Tolerance = 0.1
        
        'Set the input interfaces
        Set my_InputConfigHlpr = pInputObjectInfo
        Set my_InputObjectInfo = pInputObjectInfo
        
        If ((my_InputConfigHlpr Is Nothing) Or (my_InputObjectInfo Is Nothing)) Then
            Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
            Exit Function
        End If
        
             
        'get the pipe and the structure port collection
        On Error Resume Next
        Set PipeCollection = my_InputObjectInfo.GetSupportedObjects
        Set StructPortCollection = my_InputObjectInfo.GetSupportingObjectPorts
        On Error GoTo ErrorHandler
        
        If ((PipeCollection Is Nothing) Or (StructPortCollection Is Nothing)) Then
        Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
        Exit Function
        End If
        
        'check whether the input counts are ok
        If ((PipeCollection.Count = 0) Or (StructPortCollection.Count = 0)) Then
            Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
            Exit Function
        End If
        
        'Set the route and the struct ports
        Set IJPort_Pipe = PipeCollection(1)
        Set IJPort_Struct = StructPortCollection(1)

        If ((IJPort_Pipe Is Nothing) Or (IJPort_Struct Is Nothing)) Then
            Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
            Exit Function
        End If
        
        'Get the required measurements
        Dim strUnit As String
        'PipeRadius = my_InputConfigHlpr.GetPipeDiameter(IJPort_Pipe, strUnit)
        RouteIndex = 1
        RouteIndex = my_InputConfigHlpr.RouteIndex
        my_InputObjectInfo.GetInsulationData RouteIndex, lInsulPurpose, lInsulMat, dInsulThk, bInsulFlag
        
        If bInsulFlag <> False Then
            Dim pIJHgrCatalogSrv As IJHgrCatalogServices
            Set pIJHgrCatalogSrv = New CatalogServices
        
            Dim pipeOD As Double
            pipeOD = my_InputConfigHlpr.GetExternalPipeDiameter(IJPort_Pipe)
            pIJHgrCatalogSrv.GetEqualNPDValue pipeOD, dInsulThk, UPPER_BOUND_EQUAL_NPD, dblPipeDia, strUnit
        Else
            dblPipeDia = my_InputConfigHlpr.GetPipeDiameter(IJPort_Pipe, strUnit)
        End If
        PipeRadius = dblPipeDia / 2#
        
        Dim oUOMService         As IJUomVBInterface
        Dim unitID              As Units
        Set oUOMService = New UnitsOfMeasureServicesLib.UomVBInterface
        On Error GoTo UOMUnitErrorHandler
        unitID = oUOMService.GetUnitId(UNIT_DISTANCE, strUnit)
        On Error GoTo ErrorHandler
        PipeRadius = oUOMService.ConvertUnitToDbu(UNIT_DISTANCE, PipeRadius, unitID)
                
        Set IJDPosition_Target = my_InputObjectInfo.GetTargetPoint
        eHgrCmdType = my_InputObjectInfo.GetCommandType
        PortConfig = PortPostureService.GetGenericPortConfig(my_InputConfigHlpr, _
                                                                IJPort_Pipe, _
                                                                IJPort_Struct, _
                                                                IJDPosition_Target, _
                                                                eHgrCmdType)
        If (PortConfig = 0) Then
            Set IJHgrAssmSelectionRule_GetAssemblies = Nothing
            Exit Function
        End If
        
        If PortConfig = 4 Then
            eDistanceType = HORIZONTAL
        Else
            eDistanceType = VERTICAL
        End If
'        RouteStructDistance = my_InputConfigHlpr.GetDistanceBetweenPorts("Route", "Structure", eDistanceType)
        'RouteStructDistance = my_InputConfigHlpr.GetDistanceBetweenPortCurves(eDistanceType) ' TR#111337 ; suncor problem
        my_InputConfigHlpr.GetDistanceBetweenPortCurves eDistanceType, RouteStructDistance, RouteStructMaxDistance  ' TR#111337 problem

        ' if pipe is below structure - only UBolt is to be placed
        ' If pipe is above and distance more than radius - variable shoe
        ' If pipe is above and distance less than or equal to radius - standard shoe
        
        
        If ((PortConfig = 1) Or (PortConfig = 4)) Then 'pipe below or beside structure
            strAssyName = "G4G_1450_06_U1"
        End If
        
        If (PortConfig = 2) Then
            If (RouteStructDistance <= (PipeRadius + Tolerance)) Then
                strAssyName = "G4G_1450_06_U1"
            Else
                If (RouteStructDistance <= (PipeRadius * 2#)) Then
                    strAssyName = "VariableShoe_01"
                Else
                    strAssyName = "G4G_1402_01_B1"
                End If
            End If
        End If
        
        On Error Resume Next
        
        Set oAssyPart = my_InputConfigHlpr.GetPartProxyFromName(strAssyName)
        
        
        oAssycollection.Add oAssyPart
        
        Set IJHgrAssmSelectionRule_GetAssemblies = oAssycollection
        
        Exit Function
        
UOMUnitErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "GetUnitId method of IJUOMVBInterface failed.Support type list may not be available.Could be because of localised string mismatch").Number
        
ErrorHandler:
        Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function
