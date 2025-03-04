Attribute VB_Name = "ChamferEdgeUtils"
Option Explicit

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\AssyConnRules\ChamferEdgeUtils.bas"
'

'***********************************************************************
' METHOD:  AC_PlateEdgeChamferedByStiffenerEdge
'
' DESCRIPTION:
'
'***********************************************************************
Public Function AC_PlateEdgeChamferedByStiffenerEdge(oAssemblyConn As Object) As Boolean
                                                              
Const METHOD = "::AC_PlateEdgeChamferedByStiffenerEdge"
On Error GoTo ErrorHandler
    
    Dim bExists As Boolean
    Dim bEdgeChamfer As Boolean
    
    Dim oTestPart1 As Object
    Dim oTestPart2 As Object
    Dim oPlatePort1 As Object
    Dim oStiffenerPort2 As Object
    Dim oSDOex_Chamfer As StructDetailObjectsex.Chamfer
    Dim oSDO_AssemblyConn As StructDetailObjects.AssemblyConn

    AC_PlateEdgeChamferedByStiffenerEdge = False
    
    Set oSDO_AssemblyConn = New StructDetailObjects.AssemblyConn
    Set oSDO_AssemblyConn.object = oAssemblyConn

    Set oTestPart1 = oSDO_AssemblyConn.ConnectedObject1
    Set oTestPart2 = oSDO_AssemblyConn.ConnectedObject2
    If TypeOf oTestPart1 Is IJPlate Then
        If TypeOf oTestPart2 Is IJProfile Then
            Set oPlatePort1 = oSDO_AssemblyConn.Port1
            Set oStiffenerPort2 = oSDO_AssemblyConn.Port2
        Else
            Exit Function
        End If
        
    ElseIf TypeOf oTestPart2 Is IJPlate Then
        If TypeOf oTestPart1 Is IJProfile Then
            Set oPlatePort1 = oSDO_AssemblyConn.Port2
            Set oStiffenerPort2 = oSDO_AssemblyConn.Port1
        Else
            Exit Function
        End If
        
    Else
        Exit Function
    End If
    
    bEdgeChamfer = PlateEdgeChamferedByStiffenerEdge(oAssemblyConn, oPlatePort1, oStiffenerPort2)

    If bEdgeChamfer Then
        ' Verify that the required/expected Asseembly Connection Smart class/items have been bulkloaded
        ' ... only check if at least one Smartitem exists (then assume all exists)
        bExists = CheckSmartItemExists(SMARTTYPE_ASSY_CONNECTION_CLASS, 7, _
                                       "PlateByStiffener", "ChamferEdgeToStiffenerEdge")
        If bExists Then
            ' Verify that the required/expected Chamfer Smart class/items have been bulkloaded
            ' ... only check if at least one Smartitem exists (then assume all exists)
            bExists = CheckSmartItemExists(SMARTTYPE_CHAMFER, 8, _
                                           "RootEdgeChamfer", "ChamferEdgeBase")
        End If
        
        bEdgeChamfer = bExists
    End If
    
    AC_PlateEdgeChamferedByStiffenerEdge = bEdgeChamfer

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

'***********************************************************************
' METHOD:  PlateEdgeChamferedByStiffenerEdge
'
' DESCRIPTION:
'
'***********************************************************************
Public Function PlateEdgeChamferedByStiffenerEdge(oAssemblyConn As Object, _
                                                  oPlatePort1 As Object, _
                                                  oStiffenerPort2 As Object) As Boolean
                                                              
Const METHOD = "::PlateEdgeChamferedByStiffenerEdge"
On Error GoTo ErrorHandler
    
    Dim sTrace As String
    
    Dim eCtx As eUSER_CTX_FLAGS
    Dim eJxsec As JXSEC_CODE
    Dim eBasicCtx As IMSStructConnection.eUSER_CTX_FLAGS
    Dim eObjectType As sdwObjectType
    Dim eProfileFaceName As IMSProfileEntity.ProfileFaceName
    
    Dim oPort As IJPort
    Dim oStructPort As IJStructPort
    Dim oConnectable As IJConnectable
    
    Dim oERPart As IJProfileER
    Dim oPlatePart As IJPlatePart
    
    Dim oSDO_ProfilePart As StructDetailObjects.ProfilePart
    Dim oSDO_AssemblyConn As StructDetailObjects.AssemblyConn

    sTrace = METHOD
    PlateEdgeChamferedByStiffenerEdge = False
    
    Set oSDO_AssemblyConn = New StructDetailObjects.AssemblyConn
    Set oSDO_AssemblyConn.object = oAssemblyConn

    ' Verify oPlatePort1 is Plate Part Lateral Edge Port
    If Not TypeOf oPlatePort1 Is IJPort Then
        sTrace = sTrace & vbCrLf & "NOT oPlatePort1 Is IJPort"
        GoTo Zexit
    End If
    
    Set oPort = oPlatePort1
    Set oConnectable = oPort.Connectable
    If Not TypeOf oConnectable Is IJPlatePart Then
        sTrace = sTrace & vbCrLf & "NOT oConnectable1 Is IJPlatePart"
        GoTo Zexit
    End If
    
    If Not TypeOf oPlatePort1 Is IJStructPort Then
        sTrace = sTrace & vbCrLf & "NOT oPlatePort1 Is IJStructPort"
        GoTo Zexit
    End If
    
    Set oStructPort = oPlatePort1
    eBasicCtx = PortBasicContext(oStructPort.ContextID)
    If eBasicCtx <> PORT_BASIC_CONTEXT_LATERAL Then
        sTrace = sTrace & vbCrLf & "NOT eBasicCtx1 <> : " & eBasicCtx
        GoTo Zexit
    End If
    
    ' Verify oStiffenerPort2 is Edge Reinforcement Base/Offset/Lateral Edge Port
    If Not TypeOf oStiffenerPort2 Is IJPort Then
        sTrace = sTrace & vbCrLf & "NOT oPlatePort1 Is IJPort"
        GoTo Zexit
    End If
    
    Set oPort = oStiffenerPort2
    Set oConnectable = oPort.Connectable
    If Not TypeOf oConnectable Is IJProfilePart Then
        sTrace = sTrace & vbCrLf & "NOT oConnectable2 Is IJProfilePart"
        GoTo Zexit
    ElseIf Not TypeOf oConnectable Is IJProfileER Then
        sTrace = sTrace & vbCrLf & "NOT oConnectable2 Is IJProfileER"
        GoTo Zexit
    End If
    
    If Not TypeOf oStiffenerPort2 Is IJStructPort Then
        sTrace = sTrace & vbCrLf & "NOT oStiffenerPort2 Is IJStructPort"
        GoTo Zexit
    End If
    
    Set oStructPort = oStiffenerPort2
    eBasicCtx = PortBasicContext(oStructPort.ContextID)
    If eBasicCtx <> PORT_BASIC_CONTEXT_BASE And _
       eBasicCtx <> PORT_BASIC_CONTEXT_OFFSET And _
       eBasicCtx <> PORT_BASIC_CONTEXT_LATERAL Then
        sTrace = sTrace & vbCrLf & "NOT eBasicCtx2 <> : " & eBasicCtx
        GoTo Zexit
    End If
    
    ' Verify Edge Reinforcement Cross Section Type is "FlatBar"
    Set oSDO_ProfilePart = New StructDetailObjects.ProfilePart
    Set oSDO_ProfilePart.object = oConnectable
    If oSDO_ProfilePart.sectionType <> "FB" Then
        sTrace = sTrace & vbCrLf & "NOT .sectionType <> : " & oSDO_ProfilePart.sectionType
        GoTo Zexit
    End If
    
    ' Verify Edge Reinforcement Mounting Face is Web Left/Right
    eProfileFaceName = oSDO_ProfilePart.MountingFaceName
    If eProfileFaceName <> LeftWeb And _
       eProfileFaceName <> RightWeb Then
        sTrace = sTrace & vbCrLf & "NOT eProfileFaceName <> : " & eProfileFaceName
        GoTo Zexit
    End If
    
    ' Have Edge Reinforcement Bounded by Plate Lateral Edge configuration
    'OR
    ' Have Plate Edge Bounded by Edge Reinforcement Lateral Edge configuration
        
    PlateEdgeChamferedByStiffenerEdge = True
    sTrace = sTrace & vbCrLf & "PlateEdgeChamferedByStiffenerEdge = True "

Zexit:
'''MsgBox sTrace
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

'***********************************************************************
' METHOD:  AC_IsPlateEdgeChamferNeeded
'
' DESCRIPTION:
'
'***********************************************************************
Public Function AC_IsPlateEdgeChamferNeeded(oAssemblyConn As Object) As Boolean
                                                              
Const METHOD = "::AC_IsPlateEdgeChamferNeeded"
On Error GoTo ErrorHandler
    
    Dim dAngle As Double
    Dim dChamferForBase As Double
    Dim dChamferForOffset As Double
    Dim bExists As Boolean
    
    Dim oTestPart1 As Object
    Dim oTestPart2 As Object
    Dim oPlatePort1 As Object
    Dim oStiffenerPort2 As Object
    Dim oSDOex_Chamfer As StructDetailObjectsex.Chamfer
    Dim oSDO_AssemblyConn As StructDetailObjects.AssemblyConn

    AC_IsPlateEdgeChamferNeeded = False
    
    ' Verify that the Assembly Connection is between a Plate and Profile
    Set oSDO_AssemblyConn = New StructDetailObjects.AssemblyConn
    Set oSDO_AssemblyConn.object = oAssemblyConn

    Set oTestPart1 = oSDO_AssemblyConn.ConnectedObject1
    Set oTestPart2 = oSDO_AssemblyConn.ConnectedObject2
    If TypeOf oTestPart1 Is IJPlate Then
        If TypeOf oTestPart2 Is IJProfile Then
            Set oPlatePort1 = oSDO_AssemblyConn.Port1
            Set oStiffenerPort2 = oSDO_AssemblyConn.Port2
        Else
            Exit Function
        End If
    
    ElseIf TypeOf oTestPart2 Is IJPlate Then
        If TypeOf oTestPart1 Is IJProfile Then
            Set oPlatePort1 = oSDO_AssemblyConn.Port2
            Set oStiffenerPort2 = oSDO_AssemblyConn.Port1
        Else
            Exit Function
        End If
    
    Else
        Exit Function
    End If
    
    ' Check if there is any Chamfer to be applied to the Plate Edge
    ' ... mimimum Chamfer is currently set to 5mm difference
    Set oSDOex_Chamfer = New StructDetailObjectsex.Chamfer
    oSDOex_Chamfer.GetStiffToPlateEdgeChamferData oPlatePort1, oStiffenerPort2, dChamferForBase, dChamferForOffset
    If dChamferForBase < 0.005 Then
        If dChamferForOffset < 0.005 Then
            Exit Function
        End If
    End If

    ' for valid Stiff End to Plate Edge Chamfer cases
    ' verify that there is no twist between the Chamfered surfaces
    ' ... angle between port surfaces should be less then 1 degree
    oSDOex_Chamfer.IsAngleBetweenPlateEdgeAndStiffenerValidForChamfer oPlatePort1, oStiffenerPort2, dAngle
    If Abs(dAngle) > (Atn(1) / 45#) Then
        Exit Function
    End If
    
    AC_IsPlateEdgeChamferNeeded = True

    Exit Function
    
ErrorHandler:
    LogError Err, MODULE, METHOD
    Err.Clear
End Function


