VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CornerFeature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const sSOURCEFILE As String = "S:\StructDetail\Data\Include\SMRefDataHelper\CornerFeature.cls"

'***********************************************************************
' METHOD:  GetAngleBetweenCornerEdgePorts
'
' DESCRIPTION:  Gets the angles between two edge ports of corner feature
'               (this method can also be used for getting the angle between corner feature
'                face port and any edge port)
'
' INPUTS:  Any two input ports of the corner feature
'
' OUPUT:  Angle between the two ports
'***********************************************************************
Public Function GetAngleBetweenCornerEdgePorts(oEdgePort1 As IJPort, oEdgePort2 As IJPort) As Double

    On Error GoTo ErrorHandler
    Const sMETHOD = "GetAngleBetweenCornerEdgePorts"
    
    If oEdgePort1 Is Nothing Or _
       oEdgePort2 Is Nothing Then
      
        GoTo ErrorHandler
    
    End If
    
    
    Dim oPartInfo As New PartInfo
    Dim oNormalOnEdgePort1 As IJDVector
    Dim oNormalOnEdgePort2 As IJDVector
    
    Dim bApproxmationUsed As Boolean
    
    Set oNormalOnEdgePort1 = oPartInfo.GetPortNormal(oEdgePort1, bApproxmationUsed)
    Set oNormalOnEdgePort2 = oPartInfo.GetPortNormal(oEdgePort2, bApproxmationUsed)
    
    
    
    Dim x As Double
    Dim y As Double
    Dim z As Double
    
    
    Dim dAngle As Double
    

    oNormalOnEdgePort1.Get x, y, z
    oNormalOnEdgePort1.Set -x, -y, -z
          
    oNormalOnEdgePort2.Get x, y, z
    oNormalOnEdgePort2.Set -x, -y, -z
          
    dAngle = oNormalOnEdgePort1.Angle(oNormalOnEdgePort2, oNormalOnEdgePort1)
    dAngle = PI - dAngle
    
    
    GetAngleBetweenCornerEdgePorts = dAngle
    
    Set oNormalOnEdgePort1 = Nothing
    Set oNormalOnEdgePort1 = Nothing
    Set oPartInfo = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, sSOURCEFILE, "sMETHOD").Number
End Function

'***********************************************************************
' METHOD: GetSlotAtFeature
'
' DESCRIPTION: Gets the slot feature on support1/support 2 to which input
'              corner feature is related i.e. if slot height changes, the
'              corner feature geometry changes in a corresponding manner.
'
' INPUTS: Corner feature Smart Occurrence
'
' OUPUT: Slot object
'***********************************************************************
Public Function GetSlotAtFeature(oCornerFeatureSO As Object) As Object

Const sMETHOD = "GetSlotAtFeature"
On Error GoTo ErrHandler

    Set GetSlotAtFeature = Nothing 'incase no slot is found

    Const MIN_DISTANCE = 0.0001

'   Declare variables
    Dim oStructFeature As IJStructFeature
    Dim oMemberObjects As IJDMemberObjects
    Dim i As Long
    Dim j As Long
    Dim iC As Long
    Dim n As Long
    Dim oObj As Object
    Dim oModelBody As IJDModelBody
    Dim dMinimumDistance As Double
    Dim oBasePort1 As IJPort
    Dim oBasePort2 As IJPort
    Dim oCornerPos As IJDPosition
    Dim oCornerFeature As StructDetailObjects.CornerFeature
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim nPort1CFACs As Long
    Dim nPort2CFACs As Long
    Dim aPort1CFACData() As ConnectionData
    Dim aPort2CFACData() As ConnectionData
    Dim nSupport1ACs As Long
    Dim nSupport2ACs As Long
    Dim aSupport1ACData() As ConnectionData
    Dim aSupport2ACData() As ConnectionData
    Dim pIJStructFeatUtils As IJSDFeatureAttributes
    Dim oReqPortOnSupp1 As IJPort
    Dim oCheckPort As IJPort
    Dim AppConnection As IJAppConnectionType
    Dim bFoundSlot As Boolean
    Dim oPart As Object
    Dim oReqPortOnSupp2 As IJPort
    
    ' ----------------------
    ' Check for valid inputs
    ' ----------------------
    Set oCornerFeature = New StructDetailObjects.CornerFeature
    Set oCornerFeature.object = oCornerFeatureSO
    If oCornerFeature.object Is Nothing Then
        Exit Function
    ElseIf TypeOf oCornerFeature.object Is IJStructFeature Then
        Set oStructFeature = oCornerFeature.object
        If oStructFeature.get_StructFeatureType <> SF_CornerFeature Then
            Exit Function
        End If
    Else
        Exit Function
    End If

    oCornerFeature.GetLocationOfCornerFeature oCornerPos
    ' -------------------------------------------------------------------------------
    ' Retreive base ports of support 1 and support 2
    ' -------------------------------------------------------------------------------
    Set oSDO_Helper = New StructDetailObjects.Helper
    Set pIJStructFeatUtils = New SDFeatureUtils
    
    'Check type of part and get BasePort
    Set oPart = oCornerFeature.GetPartObject

    If TypeOf oPart Is IJProfile Then
        Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
        Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
        Set oSDO_ProfilePart.object = oPart
        Set oBasePort1 = oSDO_ProfilePart.BasePortBeforeTrim(BPT_Lateral)
        Set oBasePort2 = oSDO_ProfilePart.BasePortBeforeTrim(BPT_Base)
        Set oSDO_ProfilePart = Nothing
    ElseIf TypeOf oPart Is IJPlate Then
        Dim oSDO_PlatePart As StructDetailObjects.PlatePart
        Set oSDO_PlatePart = New StructDetailObjects.PlatePart
        Set oSDO_PlatePart.object = oPart
        Set oBasePort1 = oSDO_PlatePart.BasePort(BPT_Lateral)
        Set oBasePort2 = oSDO_PlatePart.BasePort(BPT_Base)
        Set oSDO_PlatePart = Nothing
    End If
    
'   Get all connections using the Corner Feature’s port 1
    oSDO_Helper.Object_AppConnections oBasePort1, AppConnectionType_Assembly, nPort1CFACs, aPort1CFACData()

'  Loop through all connections using the Corner Feature’s port 1
'  Use the AC connection geometry and corner feature location to find the closest AC (discard other ACs)
    Set oReqPortOnSupp1 = Nothing
    For i = 1 To nPort1CFACs
        If TypeOf aPort1CFACData(i).ToConnectedPort Is IJPort Then
            Set oCheckPort = aPort1CFACData(i).ToConnectedPort
            Set oModelBody = oCheckPort.Geometry
            oModelBody.GetMinimumDistanceFromPosition oCornerPos, Nothing, dMinimumDistance
            If dMinimumDistance < MIN_DISTANCE Then
                Set oReqPortOnSupp1 = aPort1CFACData(i).ToConnectedPort
                Exit For
            End If
        End If
    Next i
    If Not oReqPortOnSupp1 Is Nothing Then
        oSDO_Helper.Object_AppConnections oReqPortOnSupp1.Connectable, AppConnectionType_Assembly, nSupport1ACs, aSupport1ACData()
    Else
        Exit Function
    End If
    Set oReqPortOnSupp1 = Nothing
    
'   Get all connections using the Corner Feature’s port 2
    oSDO_Helper.Object_AppConnections oBasePort2, AppConnectionType_Assembly, nPort2CFACs, aPort2CFACData()

'  Loop through all connections using the Corner Feature’s port 2
'  Use the AC connection geometry and corner feature location to find the closest AC (discard other ACs)
    Set oReqPortOnSupp2 = Nothing
    For i = 1 To nPort2CFACs
        If TypeOf aPort2CFACData(i).ToConnectedPort Is IJPort Then
            Set oCheckPort = aPort2CFACData(i).ToConnectedPort
            Set oModelBody = oCheckPort.Geometry
            oModelBody.GetMinimumDistanceFromPosition oCornerPos, Nothing, dMinimumDistance
            If dMinimumDistance < MIN_DISTANCE Then
                Set oReqPortOnSupp2 = aPort2CFACData(i).ToConnectedPort
                Exit For
            End If
        End If
    Next i
    If Not oReqPortOnSupp2 Is Nothing Then
        oSDO_Helper.Object_AppConnections oReqPortOnSupp2.Connectable, AppConnectionType_Assembly, nSupport2ACs, aSupport2ACData()
    Else
        Exit Function
    End If
    Set oReqPortOnSupp2 = Nothing
    Set oCornerPos = Nothing
    
    ' -------------------------------------------------------------------
    ' Get common AC between support 1 and support 2
    ' -------------------------------------------------------------------
    bFoundSlot = False
    For i = 1 To nSupport1ACs
        For j = 1 To nSupport2ACs
            If aSupport1ACData(i).ConnectingPort Is aSupport2ACData(j).ToConnectedPort Then
                Set AppConnection = aSupport1ACData(i).AppConnection
                If AppConnection.Behavior = ConnectionPenetration Then
                    Set oMemberObjects = aSupport1ACData(i).AppConnection
                    n = oMemberObjects.Count
                    For iC = 1 To n
                        Set oObj = oMemberObjects.Item(iC)
                        If TypeOf oObj Is IJStructFeature Then
                            Set oStructFeature = oObj
                            'Check for slot feature
                            If oStructFeature.get_StructFeatureType = SF_Slot Then
                                Set GetSlotAtFeature = oObj
                                bFoundSlot = True
                                Exit For
                            End If
                        End If
                    Next iC
                End If
            End If
            If bFoundSlot = True Then Exit For
        Next j
        If bFoundSlot = True Then Exit For
    Next i
    
    'Clean up
    Set oBasePort1 = Nothing
    Set oBasePort2 = Nothing
    Set oCheckPort = Nothing
    Set oCornerFeature = Nothing
    Set oMemberObjects = Nothing
    Set oModelBody = Nothing
    Set oObj = Nothing
    Set oPart = Nothing
    Set oSDO_Helper = Nothing
    Set oStructFeature = Nothing
    Set pIJStructFeatUtils = Nothing

    Exit Function
    
ErrHandler:
        Err.Raise LogError(Err, sSOURCEFILE, "sMETHOD").Number
End Function
