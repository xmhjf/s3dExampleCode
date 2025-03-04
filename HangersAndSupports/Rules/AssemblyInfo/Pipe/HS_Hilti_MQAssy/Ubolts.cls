VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ubolts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'   Copyright   (c) 2008, Intergraph Corporation.  All Rights reserved.
'
'   ChannelConnRod.cls
'   ProgID:         HS_Hilti_MQAssy.Ubolts
'   Author:         Chethan
'   Createion Date: Nov 14, 2014
'   Description:    Fixed Point Set assembly for Hilti MQ library
'
'   Change History:
'   Date            who         change description
'   14-11-2014      Chethan             DI-CP-231162  Merge recent Hilti eCustomer changes into Product version
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Ubolts"
Private Const ENFORCE_FAILURE = False

Implements IJHgrAssmInfo

Private PIPE_ATT As Integer
Private WINGNUT() As Integer
Private BOLT() As Integer
Private WASHER() As Integer
Private sPipeAtt As String
Private sBolt As String
Private sWingNut As String
Private sWasher As String
Private sPartNum As String

'General
Private dPipeDia As Double
Private dInsulTh As Double

'----------------------------------------------------------------------------------------------
'   Name:           IJHgrAssmInfo_GetAssemblyCatalogParts()
'   Description:    Will Define part class and select the parts.
'   Date - Author:  March 4, 2013 - SK
'
'   Inputs:         pDispInputConfigHlpr as Object
'   Outputs:        A collection of object implementing IJDPart interface
'----------------------------------------------------------------------------------------------

Private Function IJHgrAssmInfo_GetAssemblyCatalogParts(ByVal pDispInputConfigHlpr As Object) As Object
    Const METHOD = "IJHgrAssmInfo_GetAssemblyCatalogParts"

    On Error GoTo ErrorHandler
    
    Initialize pDispInputConfigHlpr
    
    sPipeAtt = GetAttr("Ubolt", , , ENFORCE_FAILURE)
    GetPipeDiameterAndInsulatTh 1, dPipeDia, dInsulTh
    sWingNut = GetAttr("WingNut", , , ENFORCE_FAILURE)
    sBolt = GetAttr("Bolt", , , ENFORCE_FAILURE)
    sWasher = GetAttr("Washer", , , ENFORCE_FAILURE)
    sPartNum = GetAttr("PartNum")
    
    Hilti_RouteSizeIsValid sPartNum
    
    
    If sPipeAtt <> "" Then
        PIPE_ATT = AddPart(sPipeAtt)
    End If
    
    Dim nIndex As Long
    
    ReDim WINGNUT(1 To 2) As Integer
    For nIndex = 1 To 2
    WINGNUT(nIndex) = AddPart(sWingNut)
    Next
    
    ReDim BOLT(1 To 2) As Integer
    For nIndex = 1 To 2
        BOLT(nIndex) = AddPart(sBolt)
    Next
    
    ReDim WASHER(1 To 2) As Integer
    For nIndex = 1 To 2
        WASHER(nIndex) = AddPart(sWasher)
    Next
    
    
    'Return the Catalog Parts
    Set IJHgrAssmInfo_GetAssemblyCatalogParts = CatalogPartCollection
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function

' ---------------------------------------------------------------------------
' Name: IJHgrAssmInfo_GetAssemblyJoints()
' Description: Will specify the joint data to the assembly symbols
' Date - Author: November 3, 2008 -
'
' Inputs: pDispInputConfigHlpr As Object
'         pDispPartOccCollection As Object
' Outputs: A collection of objects implementing the IJHgrJoint interface
' ---------------------------------------------------------------------------

Private Function IJHgrAssmInfo_GetAssemblyJoints(ByVal pDispInputConfigHlpr As Object, _
                                                 ByVal pDispPartOccCollection As Object) As Object
    Const METHOD = "IJHgrAssmInfo_GetAssemblyJoints"

    On Error GoTo ErrorHandler:
    
    'Initialize hsHlpr
    InitializeJoints pDispPartOccCollection
    
    'Joints
    Dim JointFactory As New HgrJointFactory

    'Get the Pipe Attachment Type
    Dim sPartProgId As String
    Dim sQuery As String
    Dim sPipeAttClass As String
    Dim oPipeAttOcc As IJPartOcc
    Dim oPipeAttPart As IJDPart
    Dim dH0 As Double
    Dim db2 As Double
    Dim dRodDia As Double
    Dim de As Double
    Dim dBoltxOffset As Double
    Dim dBoltyOffset As Double
    Dim nIndex As Long
    Dim dExtPipeDia As Double
    Dim dPipeDia2 As Double
    Dim dInsulTh2 As Double
        
    GetPipeDiameterAndInsulatTh 1, dPipeDia2, dInsulTh2
    de = GetPartAttribute(PIPE_ATT, "E", "Double")
    'dExtPipeDia = GetPartAttribute(PIPE_ATT, "ExtPipeDia", "Double")
   
    SetPartAttribute PIPE_ATT, "ExtPipeDia", dPipeDia2, "Double"
    
    dRodDia = GetPartAttribute(PIPE_ATT, "RodDia", "Double")
    dH0 = GetPartAttribute(PIPE_ATT, "H0", "Double")
    db2 = dH0 - (dPipeDia / 2)
    
    If sBolt <> "" Then
        For nIndex = 1 To 2
             dBoltxOffset = 0.0375
            If nIndex = 1 Then
             dBoltyOffset = de / 2
            Else
            dBoltyOffset = -de / 2
            End If
        Rigid BOLT(nIndex), "RodEnd1", PIPE_ATT, "Route", -dBoltxOffset, dBoltyOffset, (dPipeDia / 2) - 0.006
        Next
    End If
    
    If sWasher <> "" Then
    For nIndex = 1 To 2
        Rigid WASHER(nIndex), "Bottom", BOLT(nIndex), "RodEnd1"
        Next
    End If
        
    If sWingNut <> "" Then
        For nIndex = 1 To 2
        
        Rigid WINGNUT(nIndex), "Structure", BOLT(nIndex), "RodEnd1", , , 0.006
        Next
    End If
    
    If sPipeAtt <> "" Then
        Rigid PIPE_ATT, "Route", -1, "Route", dBoltxOffset
       'AddJoint JointFactory.MakePrismaticJoint(PIPE_ATT, "Route", -1, "Route")
       'AddJoint JointFactory.MakePlanarJoint(BOLT(1), "RodEnd1", -1, "Structure", 8302)
    End If
    
    Set IJHgrAssmInfo_GetAssemblyJoints = GetJoints
    Class_Terminate

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function
    
    Private Function IJHgrAssmInfo_GetMaxRouteConnectionValue() As Long
    IJHgrAssmInfo_GetMaxRouteConnectionValue = 2
End Function

Private Function IJHgrAssmInfo_GetRouteConnections(ByVal pDispInputConfigHlpr As Object) As Object
    Const METHOD = "IJHgrAssmInfo_GetRouteConnections"
    On Error GoTo ErrorHandler

    Dim RouteConnColl As New Collection
    Dim PartConnInfo(2) As Integer

    PartConnInfo(1) = 1
    PartConnInfo(2) = 1

    RouteConnColl.Add PartConnInfo

    Set IJHgrAssmInfo_GetRouteConnections = RouteConnColl

    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function


Private Function IJHgrAssmInfo_GetStructConnections(ByVal pDispInputConfigHlpr As Object) As Object
    Const METHOD = "IJHgrAssmInfo_GetStructConnections"

    On Error GoTo ErrorHandler
    Dim StructConnColl As New Collection
    Dim PartConnInfo(2) As Integer

    PartConnInfo(1) = 1
    PartConnInfo(2) = 1

    StructConnColl.Add PartConnInfo

    Set IJHgrAssmInfo_GetStructConnections = StructConnColl
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, "").Number
End Function




