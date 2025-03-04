VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AssyConnDefCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Const MODULE = "S:\StructDetail\Data\SmartOccurrence\AssyConnRules\AssyConnDefCM.cls"

Private sError As String
Private sMETHOD As String

Public Function CMConstruct_PhysConnFromACPorts(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_PhysConnFromACPorts"
    
    ' Get Wrapper Class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    sError = "Setting Assembly Connection Inputs"
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting Assembly Connection Ports"
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    Set oPort1 = pAssyConn.Port1
    Set oPort2 = pAssyConn.Port2
    
    ' Get the Assembly connection, since it is the parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the slot's ca
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    ' The Create Physical Cnnection now requires a connection behavior.
    ' In the future, it problably should be ConnectionStandard but to avoid
    ' geometry checking between non-overlapping ports, set it to ConnectionPenetration
    ' Create physical connection
    Dim oPhysicalConnection As New PhysicalConn
    sError = "Creating Physical Connection"
    
    Dim oSLPort1 As IJPort
    Dim oSLPort2 As IJPort
    
    Dim oStructDetailObjectHelper As New StructDetailObjects.Helper
    
    Set oSLPort1 = oStructDetailObjectHelper.GetEquivalentLastPort(oPort1)
    Set oSLPort2 = oStructDetailObjectHelper.GetEquivalentLastPort(oPort2)
    Set oStructDetailObjectHelper = Nothing
    Set oPort1 = Nothing
    Set oPort2 = Nothing
    
    Call oPhysicalConnection.Create(pResourceManager, oSLPort1, oSLPort2, _
                                    strStartClass, pSystemParent, ConnectionStandard)
                               
    sError = "Setting Physical Connection to private variable"
    Set CMConstruct_PhysConnFromACPorts = oPhysicalConnection.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CMConstruct_PhysConnFromSubToACPort(ByVal pMemberDescription As IJDMemberDescription, _
                                                    ByVal pResourceManager As IUnknown, _
                                                    ByVal ePort1Sub As JXSEC_CODE, _
                                                    strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_PhysConnFromSubToACPort"
    
    ' Create helper
    sError = "Create Helper"
    Dim oStructDetailObjectHelper As New StructDetailObjects.Helper
        
    ' Get Wrapper Class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class
    sError = "Setting Assembly Connection Inputs"
    Set pAssyConn.object = pMemberDescription.CAO
    
    
    ' Get port1 from Sub port of connected object
    sError = "Getting sub Port 1"
    Dim oPort1 As IJPort
    
    If pAssyConn.ConnectedObject1Type = SDOBJECT_PLATE Then
        Dim oPlate As New StructDetailObjects.PlatePart
        Set oPlate.object = pAssyConn.ConnectedObject1
        Set oPort1 = oPlate.SubPort(ePort1Sub)
    ElseIf pAssyConn.ConnectedObject1Type = SDOBJECT_STIFFENER Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = pAssyConn.ConnectedObject1
        Set oPort1 = oProfile.SubPort(ePort1Sub)
    Else
        sError = "Unable to determine type of connected object"
        GoTo ErrorHandler
    End If
    
    
    ' Get Port2 from Assembly connection
    sError = "Getting Assembly Connection Port 2"
    Dim oPort2 As IJPort
    Set oPort2 = pAssyConn.Port2
    Dim oSLPort2 As IJPort
    Set oSLPort2 = oStructDetailObjectHelper.GetEquivalentLastPort(oPort2)
    
    ' Get the Assembly connection, since it is the parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the slot's ca
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    ' The Create Physical Cnnection now requires a connection behavior.
    ' In the future, it problably should be ConnectionStandard but to avoid
    ' geometry checking between non-overlapping ports, set it to ConnectionPenetration
    ' Create physical connection
    Dim oPhysicalConnection As New PhysicalConn
    sError = "Creating Physical Connection"
    
    Call oPhysicalConnection.Create(pResourceManager, oPort1, oSLPort2, _
                                    strStartClass, pSystemParent, ConnectionStandard)
                               
    sError = "Setting Physical Connection to private variable"
    Set CMConstruct_PhysConnFromSubToACPort = oPhysicalConnection.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function
Public Function CMConstruct_PhysConnFromBaseToACPort(ByVal pMemberDescription As IJDMemberDescription, _
                                                    ByVal pResourceManager As IUnknown, _
                                                    ByVal ePort2 As Base_Port_Types, _
                                                    strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_PhysConnFromBaseToACPort"
    
    ' Create helper
    sError = "Create Helper"
    Dim oStructDetailObjectHelper As New StructDetailObjects.Helper
        
    ' Get Wrapper Class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class
    sError = "Setting Assembly Connection Inputs"
    Set pAssyConn.object = pMemberDescription.CAO
    
    
    ' Get port1 from Sub port of connected object
    ' Get Port2 from Assembly connection
    sError = "Getting Assembly Connection Port 1"
    Dim oPort1 As IJPort
    Set oPort1 = pAssyConn.Port1
    Dim oSLPort1 As IJPort
    Set oSLPort1 = oStructDetailObjectHelper.GetEquivalentLastPort(oPort1)
   
   
    ' Get port 2 base port passed in
    Dim oPort2 As IJPort
    sError = "Getting base Port 2"
    
    If pAssyConn.ConnectedObject2Type = SDOBJECT_PLATE Then
        Dim oPlate As New StructDetailObjects.PlatePart
        Set oPlate.object = pAssyConn.ConnectedObject2
        Set oPort2 = oPlate.BasePort(ePort2)
    ElseIf pAssyConn.ConnectedObject2Type = SDOBJECT_STIFFENER Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = pAssyConn.ConnectedObject2
        Set oPort2 = oProfile.BasePort(ePort2)
    Else
        sError = "Unable to determine type of connected object"
        GoTo ErrorHandler
    End If
    
    
     
    ' Get the Assembly connection, since it is the parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the slot's ca
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    ' The Create Physical Cnnection now requires a connection behavior.
    ' In the future, it problably should be ConnectionStandard but to avoid
    ' geometry checking between non-overlapping ports, set it to ConnectionPenetration
    ' Create physical connection
    Dim oPhysicalConnection As New PhysicalConn
    sError = "Creating Physical Connection"
    
    Call oPhysicalConnection.Create(pResourceManager, oSLPort1, oPort2, _
                                    strStartClass, pSystemParent, ConnectionStandard)
                               
    sError = "Setting Physical Connection to private variable"
    Set CMConstruct_PhysConnFromBaseToACPort = oPhysicalConnection.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CMConstruct_PhysConnBetweenLastPorts(ByVal pMemberDescription As IJDMemberDescription, _
                                                    ByVal pResourceManager As IUnknown, _
                                                    ByVal ePort2 As Base_Port_Types, _
                                                    strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_PhysConnBetweenLastPorts"
    
    ' Create helper
    sError = "Create Helper"
    Dim oStructDetailObjectHelper As New StructDetailObjects.Helper
        
    ' Get Wrapper Class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class
    sError = "Setting Assembly Connection Inputs"
    Set pAssyConn.object = pMemberDescription.CAO
    
    ' Get port1 from lateral ports of  connected object
    ' Get Port2 from Assembly connection
    sError = "Getting Assembly Connection Port 1"
    Dim oPlate1 As New StructDetailObjects.PlatePart
    Set oPlate1.object = pAssyConn.ConnectedObject1
    
    Dim oAggtor As IJDMemberObjects
    Set oAggtor = pMemberDescription.CAO
    Dim oCuttingTool As Object
    Set oCuttingTool = oAggtor.Item(1)
    
    Dim oPort1 As IJPort
    Set oPort1 = oPlate1.SolidCutoutPort(oCuttingTool)
      
    ' Get port 2 base port passed in
    Dim oPort2 As IJPort
    sError = "Getting base Port 2"
    
    If pAssyConn.ConnectedObject2Type = SDOBJECT_PLATE Then
        Dim oPlate2 As New StructDetailObjects.PlatePart
        Set oPlate2.object = pAssyConn.ConnectedObject2
        Set oPort2 = oPlate2.BasePort(ePort2)
    ElseIf pAssyConn.ConnectedObject2Type = SDOBJECT_STIFFENER Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = pAssyConn.ConnectedObject2
        Set oPort2 = oProfile.BasePort(ePort2)
    Else
        sError = "Unable to determine type of connected object"
        GoTo ErrorHandler
    End If
    
    
     
    ' Get the Assembly connection, since it is the parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the slot's ca
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    ' The Create Physical Cnnection now requires a connection behavior.
    ' In the future, it problably should be ConnectionStandard but to avoid
    ' geometry checking between non-overlapping ports, set it to ConnectionPenetration
    ' Create physical connection
    Dim oPhysicalConnection As New PhysicalConn
    sError = "Creating Physical Connection"
    
    Call oPhysicalConnection.Create(pResourceManager, oPort1, oPort2, _
                                    strStartClass, pSystemParent, ConnectionStandard)
                               
    sError = "Setting Physical Connection to private variable"
    Set CMConstruct_PhysConnBetweenLastPorts = oPhysicalConnection.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CMConstruct_WebCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_WebCut"
    
    ' Get wrapper class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    
    sError = "Setting assembly connection inputs."
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting assembly connection ports."
    Set oPort1 = pAssyConn.Port1
    Set oPort2 = pAssyConn.Port2
           
    ' Create webcut
    Dim oWebcut As New WebCut
    sError = "Creating web cut."
    Call oWebcut.Create(pResourceManager, oPort2, oPort1, "WebCuts", pAssyConn.object)
                               
    sError = "Return the webcut"
    Set pObject = oWebcut.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CMConstruct_WebCutEndToEndPart1(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
zMsgBox "AssyConnDefCM::CMConstruct_WebCutEndToEndPart1"
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_WebCut"
    
    ' Get wrapper class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    
    sError = "Setting assembly connection inputs."
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting assembly connection ports."
    Set oPort1 = pAssyConn.Port1
    Set oPort2 = pAssyConn.Port2
           
    ' Create webcut
    Dim oWebcut As New WebCut
    sError = "Creating web cut."
    Call oWebcut.Create(pResourceManager, oPort2, oPort1, "WebCutsEndToEnd", pAssyConn.object)
zMsgBox "   Webcut.Create WebCutsEndToEnd"
                               
    sError = "Return the webcut"
    Set pObject = oWebcut.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CMConstruct_WebCutEndToEndPart2(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
zMsgBox "AssyConnDefCM::CMConstruct_WebCutEndToEndPart2"
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_WebCutEndToEnd"
    
    ' Get wrapper class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    
    sError = "Setting assembly connection inputs."
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting assembly connection ports."
    Set oPort1 = pAssyConn.Port1
    Set oPort2 = pAssyConn.Port2
           
    ' Create webcut
    Dim oWebcut As New WebCut
    sError = "Creating web cut."
    Call oWebcut.Create(pResourceManager, oPort1, oPort2, "WebCutsEndToEnd", pAssyConn.object)
zMsgBox "   Webcut.Create WebCutsEndToEnd"
                               
    sError = "Return the webcut"
    Set pObject = oWebcut.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CMConstruct_FlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object)
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_FlangeCut"
     
    CreateFlangeCut pMemberDescription, pResourceManager, pObject
    
    ' Add code to set the "TopFlangeCut" symbol parameter to True
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

Public Function CMConstruct_FlangeCutEndToEndPart1(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object)
zMsgBox "AssyConnDefCM::CMConstruct_FlangeCutEndToEndPart1"
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_FlangeCutEndToEnd"
     
    CreateFlangeCutEndToEnd pMemberDescription, pResourceManager, pObject
    
    ' Add code to set the "TopFlangeCut" symbol parameter to True
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

Public Function CMConstruct_FlangeCutEndToEndPart2(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object)
zMsgBox "AssyConnDefCM::CMConstruct_FlangeCutEndToEndPart2"
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_FlangeCutEndToEnd"
     
    CreateFlangeCutEndToEnd2 pMemberDescription, pResourceManager, pObject
    
    ' Add code to set the "TopFlangeCut" symbol parameter to True
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

Public Function CMConstruct_BottomFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                            ByVal pResourceManager As IUnknown, _
                                            ByRef pObject As Object)
    On Error GoTo ErrorHandler

    CreateFlangeCut pMemberDescription, pResourceManager, pObject
    
    ' Add code to set the "TopFlangeCut" symbol parameter to False
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

Public Function CMConstruct_BottomFlangeCutEndToEndPart1(ByVal pMemberDescription As IJDMemberDescription, _
                                            ByVal pResourceManager As IUnknown, _
                                            ByRef pObject As Object)
zMsgBox "AssyConnDefCM::CMConstruct_BottomFlangeCutEndToEndPart1"
    On Error GoTo ErrorHandler

    CreateFlangeCutEndToEnd pMemberDescription, pResourceManager, pObject
    
    ' Add code to set the "TopFlangeCut" symbol parameter to False
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

Public Function CMConstruct_BottomFlangeCutEndToEndPart2(ByVal pMemberDescription As IJDMemberDescription, _
                                            ByVal pResourceManager As IUnknown, _
                                            ByRef pObject As Object)
zMsgBox "AssyConnDefCM::CMConstruct_BottomFlangeCutEndToEndPart2"
    On Error GoTo ErrorHandler

    CreateFlangeCutEndToEnd2 pMemberDescription, pResourceManager, pObject
    
    ' Add code to set the "TopFlangeCut" symbol parameter to False
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function

Private Function CreateFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                 ByVal pResourceManager As IUnknown, _
                                 ByRef pObject As Object)
    On Error GoTo ErrorHandler

    sMETHOD = "CreateFlangeCut"
     ' Get wrapper class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    Dim oBoundedPort As IJPort, oBoundingPort As IJPort
    
    sError = "Setting assembly connection inputs."
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting assembly connection ports."
    Set oBoundedPort = pAssyConn.Port1
    Set oBoundingPort = pAssyConn.Port2
    
    ' Create flange cut if one is required
    Dim oFlangeCut As New FlangeCut
    sError = "Creating flange cut."
    
    ' create the flange when the web cut is found
    Dim pMemberObjects As IJDMemberObjects
    Set pMemberObjects = pMemberDescription.CAO
    Dim oWebcut As IJStructFeature
    Dim bFound As Boolean
    bFound = False
    ' not all members are webcuts so allow the Set to fail
    On Error Resume Next
    Dim iCount As Long, iIndex As Long
    iCount = pMemberObjects.Count
    For iIndex = 1 To iCount
        If (bFound = False) Then
            Set oWebcut = pMemberObjects.Item(iIndex)
            If (Not (oWebcut Is Nothing)) Then
                If (oWebcut.get_StructFeatureType = SF_WebCut) Then
                   Call oFlangeCut.Create(pResourceManager, oBoundingPort, oBoundedPort, oWebcut, "FlangeCuts", _
                                          pAssyConn.object)
                   sError = "Return the flange cut "
                   Set pObject = oFlangeCut.object
                   bFound = True
               End If
               Set oWebcut = Nothing
            End If
       End If
    Next iIndex
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Private Function CreateFlangeCutEndToEnd(ByVal pMemberDescription As IJDMemberDescription, _
                                 ByVal pResourceManager As IUnknown, _
                                 ByRef pObject As Object)
    On Error GoTo ErrorHandler

    sMETHOD = "CreateFlangeCutEndToEnd"
     ' Get wrapper class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    Dim oBoundedPort As IJPort, oBoundingPort As IJPort
    
    sError = "Setting assembly connection inputs."
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting assembly connection ports."
    Set oBoundedPort = pAssyConn.Port1
    Set oBoundingPort = pAssyConn.Port2
    
    ' Create flange cut if one is required
    Dim oFlangeCut As New FlangeCut
    sError = "Creating flange cut."
    
    ' create the flange when the web cut is found
    Dim pMemberObjects As IJDMemberObjects
    Set pMemberObjects = pMemberDescription.CAO
    Dim oWebcut As IJStructFeature
    ' The Assembly defines six(6) possible Item members (StiffEndToStiffEndDef.cls)
    ' WebCut between part_A and Part_B
    ' Top FlangeCut between part_A and Part_B
    ' Bottom FlangeCut between part_A and Part_B
    ' WebCut between part_B and Part_A
    ' Top FlangeCut between part_B and Part_A
    ' Bottom FlangeCut between part_B and Part_A
    '
    ' For FlangeCuts between Part_A and Part_B
    ' use the first(1st) item as the WebCut input
    sError = "WebCut required for FlangeCut not valid"
    Set oWebcut = pMemberObjects.Item(1)
    If (Not (oWebcut Is Nothing)) Then
        If (oWebcut.get_StructFeatureType = SF_WebCut) Then
            Call oFlangeCut.Create(pResourceManager, _
                                   oBoundingPort, oBoundedPort, oWebcut, _
                                   "FlangeCutsEndToEnd", pAssyConn.object)
            sError = "Return the flange cut "
            Set pObject = oFlangeCut.object
        End If
    End If
                   
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Private Function CreateFlangeCutEndToEnd2(ByVal pMemberDescription As IJDMemberDescription, _
                                 ByVal pResourceManager As IUnknown, _
                                 ByRef pObject As Object)
    On Error GoTo ErrorHandler

    sMETHOD = "CreateFlangeCutEndToEnd2"
     ' Get wrapper class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    Dim oBoundedPort As IJPort, oBoundingPort As IJPort
    
    sError = "Setting assembly connection inputs."
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting assembly connection ports."
    Set oBoundedPort = pAssyConn.Port1
    Set oBoundingPort = pAssyConn.Port2
    
    ' Create flange cut if one is required
    Dim oFlangeCut As New FlangeCut
    sError = "Creating flange cut."
    
    ' create the flange when the web cut is found
    Dim pMemberObjects As IJDMemberObjects
    Set pMemberObjects = pMemberDescription.CAO
    Dim oWebcut As IJStructFeature
    
    ' The Assembly defines six(6) possible Item members (StiffEndToStiffEndDef.cls)
    ' WebCut between part_A and Part_B
    ' Top FlangeCut between part_A and Part_B
    ' Bottom FlangeCut between part_A and Part_B
    ' WebCut between part_B and Part_A
    ' Top FlangeCut between part_B and Part_A
    ' Bottom FlangeCut between part_B and Part_A
    '
    ' For FlangeCuts between Part_B and Part_A
    ' use the fourth(4th) item as the WebCut input
    sError = "WebCut required for FlangeCut not valid"
    Set oWebcut = pMemberObjects.Item(4)
    If (Not (oWebcut Is Nothing)) Then
        If (oWebcut.get_StructFeatureType = SF_WebCut) Then
            Call oFlangeCut.Create(pResourceManager, _
                                   oBoundedPort, oBoundingPort, oWebcut, _
                                   "FlangeCutsEndToEnd", pAssyConn.object)
            sError = "Return the flange cut "
            Set pObject = oFlangeCut.object
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

'  This function determines if a top flange cut is required
' all cross-sections without a flange return false
' all other cross-sections return true

Public Sub CMCreateTopFlangeCut(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMCreateTopFlangeCut"
    
    bIsNeeded = True
    
    ' FlangeCut not required for Split/Angled SplitEndToEndCase :
    Dim bSplitAngleEndToEndCase As Boolean
    bSplitAngleEndToEndCase = IsSplitAngleEndToEndCase(pMD)
    If bSplitAngleEndToEndCase Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim strSectionType As String
    strSectionType = Get_CrossSectionType(pMD)
        
    If ((strSectionType = "FB") Or (strSectionType = "HalfR") Or _
        (strSectionType = "P") Or (strSectionType = "R") Or _
        (strSectionType = "SB") Or (strSectionType = "SqTu") Or _
        (strSectionType = "RT")) Then bIsNeeded = False
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub


'  This function determines if a bottom flange cut is required
' only I cross-ssections return true
' all other cross-sections return false

Public Sub CMCreateBottomFlangeCut(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMCreateBottomFlangeCut"
    
    bIsNeeded = False
    
    ' FlangeCut not required for Split/Angled SplitEndToEndCase :
    Dim bSplitAngleEndToEndCase As Boolean
    bSplitAngleEndToEndCase = IsSplitAngleEndToEndCase(pMD)
    If bSplitAngleEndToEndCase Then
        bIsNeeded = False
        Exit Sub
    End If

    Dim strSectionType As String
    strSectionType = Get_CrossSectionType(pMD)
        
    If ((strSectionType = "I") Or (strSectionType = "ISType") Or _
        (strSectionType = "C_SS") Or (strSectionType = "CSType")) Then bIsNeeded = True
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

Private Function GetCrossSectionType(ByRef pMD As IJDMemberDescription, _
                        Optional bConnectedObject1 As Boolean = True) As String
    On Error GoTo ErrorHandler

    ' For Compatibility,
    ' Do Not remove/rename any existing methods
    sMETHOD = "GetCrossSectionType"
    
    GetCrossSectionType = Get_CrossSectionType(pMD, bConnectedObject1)
    
        
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CMConstruct_Slot(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_Slot"
    
    ' Get Wrapper Class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    sError = "Setting Assembly Connection Inputs"
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting Assembly Connection Parts associated to ports"
    Dim oPort1Part As Object
    Dim oPort2Part As Object
    Set oPort1Part = pAssyConn.ConnectedObject1 ' must be the penetrated by the rules
    Set oPort2Part = pAssyConn.ConnectedObject2 ' must be the penetrating by rules
    
' Manual Slot
' Verify that order of graphic inputs are valid
' required if not Plate bounded by Profile case
' (testing only)
' (Allow Profile bounded by Plate case)
' (Profile bounded by Plate case, switch order of inputs to Slot)
    Dim zObjectType1 As sdwObjectType
    Dim zObjectType2 As sdwObjectType
    zObjectType1 = pAssyConn.ConnectedObject1Type
    zObjectType2 = pAssyConn.ConnectedObject2Type
    If zObjectType2 = SDOBJECT_PLATE Then
        If zObjectType1 = SDOBJECT_STIFFENER Then
            Set oPort1Part = pAssyConn.ConnectedObject2
            Set oPort2Part = pAssyConn.ConnectedObject1
        End If
    End If
    
    ' Get the Assembly connection, since it is the parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the slot's ca
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    ' Create physical connection
    Dim oSlot As New StructDetailObjects.Slot
    sError = "Creating Slot"
    Call oSlot.Create(pResourceManager, oPort2Part, oPort1Part, strStartClass, pSystemParent)
                               
    sError = "Setting Slot to private variable"
    Set CMConstruct_Slot = oSlot.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function CMConstruct_Bracket2S(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_Bracket2S"
    
    ' Get wrapper class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMemberDescription.CAO
    
    ' Initialize wrapper class and get the 2 ports
    Dim oPort1 As IJPort
    Dim oPort2 As IJPort
    
    Dim oProfile1 As New StructDetailObjects.ProfilePart
    Set oProfile1.object = pAssyConn.ConnectedObject1
    Set oPort1 = oProfile1.SubPort(JXSEC_TOP)
    
    Dim oProfile2 As New StructDetailObjects.ProfilePart
    Set oProfile2.object = pAssyConn.ConnectedObject2
    Set oPort2 = oProfile2.SubPort(JXSEC_TOP)
    
    ' Get Location for Bracket Plane from Assembly Connection
    Dim oLocation As IJDPosition
    If pAssyConn.ConnectionBehavior = ConnectionStandard Then
        Set oLocation = pAssyConn.BoundGlobalShipLocation
    Else
        Set oLocation = pAssyConn.PenetrationGlobalShipLocation
    End If
    
    Dim oPortOnWebLeft As IJPort
    Set oPortOnWebLeft = oProfile2.SubPort(JXSEC_WEB_LEFT)
    
    Dim oNewLocation As IJDPosition
    Set oNewLocation = PointOnWebLeft(oPortOnWebLeft, oLocation)
    
    Set oLocation = Nothing
    Set oPortOnWebLeft = Nothing
    
    ' Get plane of Web_Left
    Dim pRefPlane As IJPlane
    Set pRefPlane = oProfile2.EdgePlane(JXSEC_WEB_LEFT, oNewLocation, pResourceManager)
    Set oNewLocation = Nothing
    
    Dim Normal_X As Double
    Dim Normal_Y As Double
    Dim Normal_Z As Double
    
    pRefPlane.GetNormal Normal_X, Normal_Y, Normal_Z

    pRefPlane.SetNormal -Normal_X, -Normal_Y, -Normal_Z
 
    ' Create bracket
    Dim oBracket As New StructDetailObjects.Bracket
    sError = "Creating bracket"
    Call oBracket.CreateBracket_2S(pResourceManager, "Bracket2S", pRefPlane, oPort2, oPort1, pAssyConn.object)
    
    Dim oBracketPart As New StructDetailObjects.PlatePart
    Set oBracketPart.object = oBracket.object
    oBracketPart.PlateThickness = 0.01
    
    'Set PlateThicknessOffset value based on the profile Cross-section
    'Find the profile cross-section first
    Dim oIJDProfileSection As IJDProfileSection
    Dim oIJCrossSection As IJCrossSection
    Dim strCrossSectionType As String
    
    Set oIJDProfileSection = oProfile2.object
    Set oIJCrossSection = oIJDProfileSection.crossSection
    strCrossSectionType = oIJCrossSection.Type
    
    Set oIJDProfileSection = Nothing
    Set oIJCrossSection = Nothing
    
    Dim dThicknessOffset As Double
    
    If (strCrossSectionType = "FB") Or (strCrossSectionType = "SB") Or _
        (strCrossSectionType = "T_XType") Or (strCrossSectionType = "TSType") Or _
        (strCrossSectionType = "I") Or (strCrossSectionType = "ISType") Or _
        (strCrossSectionType = "BUT") Or (strCrossSectionType = "BUTL2") Or (strCrossSectionType = "BUTL3") Then
        'dThicknessOffset = 0# 'default value = 0 on the object
    Else
        dThicknessOffset = 0.005
        oBracketPart.thicknessOffset = dThicknessOffset
    End If
    
    Set oBracketPart = Nothing
     
    sError = "Return the bracket"
    Set pObject = oBracket.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function


Private Function PointOnWebLeft(oPortOnWebLeft As IJPort, oPlanePnt As IJDPosition) As IJDPosition

On Error GoTo ErrorHandler
sMETHOD = "PlaneOnWebLeft"

    On Error Resume Next
    
    Dim oNormalVector As IJDVector
    Dim oProjectedPnt As IJDPosition
    Dim oPort As IJPort
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    
    Set oPort = oPortOnWebLeft
    If Not oPort Is Nothing Then
        Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
        oTopologyLocate.GetProjectedPointOnModelBody oPort.Geometry, oPlanePnt, _
                                                     oProjectedPnt, oNormalVector
                                                     
        Set PointOnWebLeft = oProjectedPnt
    End If
    
    Set oProjectedPnt = Nothing
    Set oNormalVector = Nothing
    Set oPort = Nothing
    Set oTopologyLocate = Nothing
    
Exit Function
ErrorHandler:
   Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Function


Public Function CMConstruct_Chamfer(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_Chamfer"
    
    ' Get Wrapper Class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    sError = "Setting Assembly Connection Inputs"
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting Assembly Connection Parts associated to ports"
    Dim oPart1 As IJPlate
    Dim oPart2 As IJPlate
    Dim oPort1Part As Object
    Dim oPort2Part As Object
    Set oPart1 = pAssyConn.ConnectedObject1 ' must be the chamfered by the rules
    Set oPart2 = pAssyConn.ConnectedObject2 ' must be the one driving the chamfer by rules
    Set oPort1Part = pAssyConn.Port1
'    If pAssyConn.FromIntersectionSeam Then
'        Dim oElem As IJElements
'        Set oElem = pAssyConn.PartsOnOtherSide
'        If oElem.Count() > 0 Then
'            Set oPort2Part = oElem.Item(1)
'        End If
'        Set oElem = Nothing
'    Else
        Set oPort2Part = pAssyConn.Port2
'    End If
    
    ' Get the Assembly connection, since it is the parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the chamfer's ca
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    ' Create physical connection
    Dim oChamfer As New StructDetailObjects.Chamfer
    sError = "Creating Chamfer"
    Call oChamfer.Create(pResourceManager, oPort1Part, oPort2Part, strStartClass, pSystemParent)
                               
    sError = "Setting Chamfer to private variable"
    Set CMConstruct_Chamfer = oChamfer.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function
Public Function CMConstruct_Chamfer2(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_Chamfer"
    
    ' Get Wrapper Class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    sError = "Setting Assembly Connection Inputs"
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting Assembly Connection Parts associated to ports"
    Dim oPart1 As IJPlate
    Dim oPart2 As IJPlate
    Dim oPort1Part As Object
    Dim oPort2Part As Object
    Set oPart1 = pAssyConn.ConnectedObject1 ' must be the chamfered by the rules
    Set oPart2 = pAssyConn.ConnectedObject2 ' must be the one driving the chamfer by rules
    Set oPort1Part = pAssyConn.Port1
'    If pAssyConn.FromIntersectionSeam Then
'        Dim oElem As IJElements
'        Set oElem = pAssyConn.PartsOnOtherSide
'        If oElem.Count() > 0 Then
'            Set oPort2Part = oElem.Item(1)
'        End If
'        Set oElem = Nothing
'    Else
        Set oPort2Part = pAssyConn.Port2
'    End If
    
    ' Get the Assembly connection, since it is the parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the chamfer's ca
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    ' Create physical connection
    Dim oChamfer As New StructDetailObjects.Chamfer
    sError = "Creating Chamfer"
    Call oChamfer.Create(pResourceManager, oPort2Part, oPort1Part, strStartClass, pSystemParent)
                               
    sError = "Setting Chamfer to private variable"
    Set CMConstruct_Chamfer2 = oChamfer.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function


'***********************************************************************
' METHOD:  CMConstruct_Collar
'
' DESCRIPTION:  Creates a collar part.
'
'***********************************************************************
Public Function CMConstruct_Collar(ByVal pSlot As IJStructFeature, strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMETHOD = "CMConstruct_Collar"

    ' Create New Collar
    Dim oCollar As New StructDetailObjects.Collar
    sError = "Creating Collar"
    oCollar.Create pSlot, strStartClass
                               
    sError = "Setting Collar to private variable"
    Set CMConstruct_Collar = oCollar.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function
'***********************************************************************
' METHOD:  CMCreateCutout
'
' DESCRIPTION:  Create a cutting tool and set the SDCut_AE operation
'
'***********************************************************************
Public Function CMCreateCutout(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMETHOD = "CMCreateCutout"
    On Error GoTo ErrorHandler
    
    ' Get Wrapper Class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
 
    ' Initialize wrapper class and get the 2 ports
    sError = "Setting Assembly Connection Inputs"
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting Assembly Connection Parts associated to ports"
    Dim oPort1Part As Object
    Dim oPort2Part As Object
    Set oPort1Part = pAssyConn.ConnectedObject1 ' must be the penetrated by the rules
    Set oPort2Part = pAssyConn.ConnectedObject2 ' must be the penetrating by rules
           
    ' Create the cutout
    Dim oCutout As StructDetailObjects.Cutout
    Set oCutout = New StructDetailObjects.Cutout
    sError = "Creating Cutout"
    Call oCutout.Create(pResourceManager, oPort2Part, oPort1Part, pAssyConn.object)
    
    sError = "Adding Cutout"
    Call oCutout.AddCutout
    
    Set CMCreateCutout = oCutout.object
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function '***********************************************************************
' METHOD:  CMCreateBracket
'
' DESCRIPTION:  Conditional that determines if there is a bracket
'
'***********************************************************************
Public Sub CMCreateBracket(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMCreateBracket"
    
    bIsNeeded = False
    
    Dim pAssyConn As New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMD.CAO
    
    '********************************************************************************
    Const strIASelectionRuleRoot As String = "IAAssyConnRules_StiffEndByStiffFaceSel"
    Dim pHelper As New StructDetailObjects.Helper
    
    Dim retValue As Variant
    Dim retValueType As EResultType
    Dim sTableName As String
    Dim lCodeList As Long
    Dim sShortDesc As String
    Dim sLongDesc As String
    
    Call pHelper.GetCustomAttributeParamValue(pAssyConn.object, strIASelectionRuleRoot, "PlaceBracket", _
            retValue, retValueType, sTableName, lCodeList, sShortDesc, sLongDesc)
    
    Dim sPlaceBracket As String
    
    If Len(Trim(sTableName)) > 0 Then
        sPlaceBracket = sShortDesc
    Else
        sPlaceBracket = retValue
    End If
    '***********************************************************************************
    
    If sPlaceBracket = "Yes" Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

'***********************************************************************
' METHOD:  CMCreateChamfer1
'
' DESCRIPTION:  Conditional that determines if there is a chamfer 1
'
'***********************************************************************
Public Sub CMCreateChamfer1(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMCreateChamfer1"
    
    bIsNeeded = True
    
    Dim pAssyConn As New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMD.CAO

    ' Check if there are Muiltple Assembly Connections between the two parts
    ' if there are Muiltple Assembly Connections this indicates that is
    ' is a special case of Edge to Edge that does NOT require a Chamfer
    '
    '                           |---------------------
    ' Connect2(offset)          |
    '   ======================= |
    ' Connect2 (base)        || |  Connect1 Edge
    '   ======================= |
    '   ------------------------|
    '
    ' In this case (Built up bounded by Built up)
    ' There are Assembly Connections:
    '       Between Connect1 Edge and Connect2 Edge
    '       Between Connect1 Edge and Connect2 (base)
    ' Therefore,
    ' we only want a Physical Connection Between Connect1 Edge and Connect2 Edge
    '
    Dim bMultiple As Boolean
    bMultiple = CheckMultipleAssemblyConnections(pAssyConn.ConnectedObject1, _
                                                 pAssyConn.Port1, _
                                                 pAssyConn.ConnectedObject2, _
                                                 pAssyConn.Port2)
    If bMultiple Then
        bIsNeeded = False
        Exit Sub
    End If
    


    'define interface from which to get the question
    Const strIASelectionRuleRoot As String = "IAAssyConnRules_PlateEdgeByPlateEdgeSel"
    Dim pHelper As New StructDetailObjects.Helper
    Dim retValue As Variant
    Dim retValueType As EResultType
    Dim sTableName As String
    Dim lCodeList As Long
    Dim sShortDesc As String
    Dim sLongDesc As String

    'use helper to get the answer from the question
    Call pHelper.GetCustomAttributeParamValue(pAssyConn.object, strIASelectionRuleRoot, "ChamferType", _
            retValue, retValueType, sTableName, lCodeList, sShortDesc, sLongDesc)
    
    Dim sChamferType As String
    
    'check if answer is from code list; if so, get short desc, if not, use retValue
    If Len(Trim(sTableName)) > 0 Then
        sChamferType = sShortDesc
    Else
        sChamferType = retValue
    End If
        
    If (sChamferType = "Obj1Double") Or (sChamferType = "Obj1Base") Or (sChamferType = "Obj1Offset") Or _
        (sChamferType = "Obj1BaseObj2Offset") Or (sChamferType = "Obj1OffsetObj2Base") Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub
'***********************************************************************
' METHOD:  CMCreateChamfer2
'
' DESCRIPTION:  Conditional that determines if there is a chamfer 2.
'
'***********************************************************************
Public Sub CMCreateChamfer2(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMCreateChamfer2"
    
    bIsNeeded = False
    
    Dim pAssyConn As New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMD.CAO
    
    ' Check if there are Muiltple Assembly Connections between the two parts
    ' if there are Muiltple Assembly Connections this indicates that is
    ' is a special case of Edge to Edge that does NOT require a Chamfer
    '
    '                           |---------------------
    ' Connect2(offset)          |
    '   ======================= |
    ' Connect2 (base)        || |  Connect1 Edge
    '   ======================= |
    '   ------------------------|
    '
    ' In this case (Built up bounded by Built up)
    ' There are Assembly Connections:
    '       Between Connect1 Edge and Connect2 Edge
    '       Between Connect1 Edge and Connect2 (base)
    ' Therefore,
    ' we only want a Physical Connection Between Connect1 Edge and Connect2 Edge
    '
    Dim bMultiple As Boolean
    bMultiple = CheckMultipleAssemblyConnections(pAssyConn.ConnectedObject1, _
                                                 pAssyConn.Port1, _
                                                 pAssyConn.ConnectedObject2, _
                                                 pAssyConn.Port2)
    If bMultiple Then
        bIsNeeded = False
        Exit Sub
    End If
    
    
    'define interface from which to get the question
    Const strIASelectionRuleRoot As String = "IAAssyConnRules_PlateEdgeByPlateEdgeSel"
    Dim pHelper As New StructDetailObjects.Helper
    Dim retValue As Variant
    Dim retValueType As EResultType
    Dim sTableName As String
    Dim lCodeList As Long
    Dim sShortDesc As String
    Dim sLongDesc As String

    'use helper to get the answer from the question
    Call pHelper.GetCustomAttributeParamValue(pAssyConn.object, strIASelectionRuleRoot, "ChamferType", _
            retValue, retValueType, sTableName, lCodeList, sShortDesc, sLongDesc)
    
    Dim sChamferType As String
    
    'check if answer is from code list; if so, get short desc, if not, use retValue
    If Len(Trim(sTableName)) > 0 Then
        sChamferType = sShortDesc
    Else
        sChamferType = retValue
    End If
    
    If (sChamferType = "Obj2Double") Or (sChamferType = "Obj2Base") Or (sChamferType = "Obj2Offset") Or _
        (sChamferType = "Obj1BaseObj2Offset") Or (sChamferType = "Obj1OffsetObj2Base") Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub
'***********************************************************************
' METHOD:  CMChamferPhysConn
'
' DESCRIPTION:  Conditional that determines if there is a chamfer.  If not, place a PC
'
'***********************************************************************

Public Sub CMChamferPhysConn(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMChamferPhysConn"
    
    bIsNeeded = True
    
    Dim pAssyConn As New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMD.CAO
    'define interface from which to get the question
    Const strIASelectionRuleRoot As String = "IAAssyConnRules_PlateEdgeByPlateEdgeSel"
    Dim pHelper As New StructDetailObjects.Helper
    Dim retValue As Variant
    Dim retValueType As EResultType
    Dim sTableName As String
    Dim lCodeList As Long
    Dim sShortDesc As String
    Dim sLongDesc As String

    'use helper to get the answer from the question
    Call pHelper.GetCustomAttributeParamValue(pAssyConn.object, strIASelectionRuleRoot, "ChamferType", _
            retValue, retValueType, sTableName, lCodeList, sShortDesc, sLongDesc)
    
    Dim sChamferType As String
    
    'check if answer is from code list; if so, get short desc, if not, use retValue
    If Len(Trim(sTableName)) > 0 Then
        sChamferType = sShortDesc
    Else
        sChamferType = retValue
    End If
        
    If (sChamferType = "Obj1Base") Or (sChamferType = "Obj1Offset") Or (sChamferType = "Obj1Double") Or _
        (sChamferType = "Obj2Base") Or (sChamferType = "Obj2Offset") Or (sChamferType = "Obj2Double") Or _
        (sChamferType = "Obj1BaseObj2Offset") Or (sChamferType = "Obj1OffsetObj2Base") Then
        bIsNeeded = False
    Else
        bIsNeeded = True
    End If
        
    Set pAssyConn = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub


'***********************************************************************
' METHOD:  CMCopyParentMatAndGrade
'
' DESCRIPTION:  Copies the parent base plate thinkness (dimensions).
'
'***********************************************************************
Public Sub CMCopyParentMatAndGrade(pCollar As Object)

 On Error GoTo ErrorHandler
     
    Dim oCollar As New Collar
    Set oCollar.object = pCollar
    
    Dim oSlot As New Slot
    Set oSlot.object = oCollar.Slot
        
    Dim pParentPlate As IJStructureMaterial
    Set pParentPlate = oSlot.Penetrated
    
    Dim oMaterial As Object
    Set oMaterial = pParentPlate.Material
    
    Dim oCollarPlate As IJStructureMaterial
    Set oCollarPlate = pCollar
    
    ' Set the collar to the dimension retrieved off parent
    oCollarPlate.Material = oMaterial
    
    Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMCopyParentMatAndGrade").Number
End Sub

'***********************************************************************
' METHOD:  CMCopyParentThickness
'
' DESCRIPTION:  copies parent material and grade.
'
'***********************************************************************
Public Sub CMCopyParentThickness(pCollar As Object)

 On Error GoTo ErrorHandler
     
    Dim oCollar As New Collar
    Set oCollar.object = pCollar

    Dim oSlot As New Slot
    Set oSlot.object = oCollar.Slot
    
    Dim oCollarPlate As IJPlate
    Set oCollarPlate = pCollar

    'get penetrated parent
    On Error Resume Next
    Dim dThickness As Double
    Dim pPenetratedParent As IJPlate
    Dim oDimensions As IJDPlateDimensions
    
    Set pPenetratedParent = oSlot.Penetrated
    If Not pPenetratedParent Is Nothing Then
        Set pPenetratedParent.object = oSlot.Penetrated
        Set oDimensions = pPenetratedParent.Dimensions
        oCollarPlate.Dimensions = oDimensions
    Else
        Dim pProfilePart As New StructDetailObjects.ProfilePart
        Set pProfilePart.object = oSlot.Penetrated
        dThickness = pProfilePart.WebThickness
        Set oDimensions = oCollarPlate.Dimensions
        oDimensions.thickness = dThickness
    End If
            
    Set oDimensions = Nothing
    Set pPenetratedParent = Nothing
    Set pProfilePart = Nothing
            
    Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMCopyParentThickness").Number
End Sub

Public Function CMSetSideOfPlate(pCollar As Object)
    On Error GoTo ErrorHandler

    sMETHOD = "CMSetSideOfPlate"
    sError = "Setting Side of Plate"
      
    ' Get Collar
    Dim oCollar As New Collar
    Set oCollar.object = pCollar
    
    ' Get Penetrated Plate
    Dim oSlot As New Slot
    Set oSlot.object = oCollar.Slot
    
    On Error Resume Next
    Dim pPenetratedPlate As IJPlate
    Dim lCandidateSideOfPlate As Long
    Set pPenetratedPlate = oSlot.Penetrated
    If Not pPenetratedPlate Is Nothing Then
        ' Get what side of the plate to set the collar on
        Dim pPartInfoUtils As GSCADStructGeomUtilities.IJDPartInfo
        Set pPartInfoUtils = New GSCADStructGeomUtilities.PartInfo
        lCandidateSideOfPlate = pPartInfoUtils.IdentifyCandidateSideOfPlate(pPenetratedPlate)
    Else
        lCandidateSideOfPlate = 1
    End If
    
    ' Set the Side Of Plate
    Dim pCollarPart As IJCollarPart
    Set pCollarPart = pCollar
    If lCandidateSideOfPlate = 1 Then  ' Anti Molded
        pCollarPart.SideOfPlate = PLATE_ANTI_MOLDED
    Else
       pCollarPart.SideOfPlate = PLATE_MOLDED
    End If
      
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function ConvertMountingFaceToPort(eFaceName As ProfileFaceName) As JXSEC_CODE
    
    Dim ePort As JXSEC_CODE
    ePort = JXSEC_UNKNOWN
    
    On Error GoTo ErrorHandler
    sMETHOD = "ConvertMountingFaceToPort"
    
    
    sError = "Case on face name"
    
    Select Case eFaceName
        Case BeamVirtual
            ePort = JXSEC_UNKNOWN
        Case BottomInnerRadiusOfSuperiorFlange
            ePort = JXSEC_WEB_LEFT_TOP_CORNER
        Case BottomOuterRadiusOfInferiorFlange
            ePort = JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM_CORNER
        Case BottomOuterRadiusOfSuperiorFlange
            ePort = JXSEC_TOP_FLANGE_LEFT_BOTTOM_CORNER
        Case BottomSideOfInferiorFlange
            ePort = JXSEC_BOTTOM
        Case LeftInteriorSideOfInferiorFlange
            ePort = JXSEC_BOTTOM_FLANGE_LEFT_TOP
        Case LeftInteriorSideOfSuperiorFlange
            ePort = JXSEC_TOP_FLANGE_LEFT_BOTTOM
        Case LeftSideOfInferiorFlange
            ePort = JXSEC_BOTTOM_FLANGE_LEFT
        Case LeftSideOfSuperiorFlange
            ePort = JXSEC_TOP_FLANGE_LEFT
        Case LeftWeb
            ePort = JXSEC_WEB_LEFT
        Case RightSideOfSuperiorFlange
            ePort = JXSEC_TOP_FLANGE_RIGHT
        Case RightInteriorSideOfSuperiorFlange
            ePort = JXSEC_TOP_FLANGE_RIGHT_BOTTOM
        Case RightWeb
            ePort = JXSEC_WEB_RIGHT
        Case RightInteriorSideOfInferiorFlange
            ePort = JXSEC_BOTTOM_FLANGE_RIGHT_TOP
        Case RightSideOfInferiorFlange
            ePort = JXSEC_BOTTOM_FLANGE_RIGHT
        Case TopSideOfSuperiorFlange
            ePort = JXSEC_TOP
        Case TopInnerRadiusOfInferiorFlange
            ePort = JXSEC_WEB_LEFT_BOTTOM_CORNER
        Case TopOuterRadiusOfInferiorFlange
            ePort = JXSEC_BOTTOM_FLANGE_LEFT_TOP_CORNER
        Case TopOuterRadiusOfSuperiorFlange
            ePort = JXSEC_TOP_FLANGE_LEFT_TOP_CORNER
        Case Else
            sError = "Invalid Face type"
            GoTo ErrorHandler
    End Select
    
    ConvertMountingFaceToPort = ePort
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
    ConvertMountingFaceToPort = ePort
End Function

Public Function IsSplitAngleEndToEndCase(ByRef pMD As IJDMemberDescription) As Boolean
sMETHOD = "IsSplitAngleEndToEndCase"
    
    On Error GoTo ErrorHandler
    IsSplitAngleEndToEndCase = False
    
    ' FlangeCut not required for SplitAngle SplitEndToEndCase
    Dim sEndToEndCase As String
    sEndToEndCase = GetAnswer_SplitEndToEndCase(pMD)
    
    If LCase(Trim(sEndToEndCase)) = LCase("AngleWebSquareFlange") Or _
       LCase(Trim(sEndToEndCase)) = LCase("AngleWebBevelFlange") Or _
       LCase(Trim(sEndToEndCase)) = LCase("AngleWebAngleFlange") Or _
       LCase(Trim(sEndToEndCase)) = LCase("DistanceWebDistanceFlange") Or _
       LCase(Trim(sEndToEndCase)) = LCase("OffsetWebOffsetFlange") Then
        IsSplitAngleEndToEndCase = True
        Exit Function
    End If
    
    If LCase(Trim(sEndToEndCase)) = LCase("AngleWebSquareFlange_Flip") Or _
       LCase(Trim(sEndToEndCase)) = LCase("AngleWebBevelFlange_Flip") Or _
       LCase(Trim(sEndToEndCase)) = LCase("AngleWebAngleFlange_Flip") Or _
       LCase(Trim(sEndToEndCase)) = LCase("DistanceWebDistanceFlange_Flip") Or _
       LCase(Trim(sEndToEndCase)) = LCase("OffsetWebOffsetFlange_Flip") Then
        IsSplitAngleEndToEndCase = True
        Exit Function
    End If
    
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function GetAnswer_SplitEndToEndCase(ByRef pMD As IJDMemberDescription) As String
sMETHOD = "GetAnswer_SplitEndToEndCase"
    
    On Error GoTo ErrorHandler
    
    Dim retValue As Variant
    Dim lCodeList As Long
    Dim sLongDesc As String
    Dim sShortDesc As String
    Dim sTableName As String
    Dim sSplitEndToEndCase As String
    
    Dim e_ResultType As EResultType
    Dim oSDO_Helper As StructDetailObjects.Helper
    Const strIASelectionRuleRoot As String = "IAAssyConnRules_StiffEndByStiffEndSel"
    
    sSplitEndToEndCase = ""
    Set oSDO_Helper = New StructDetailObjects.Helper
    Call oSDO_Helper.GetCustomAttributeParamValue(pMD.CAO, _
                                              strIASelectionRuleRoot, _
                                              "SplitEndToEndCase", _
                                              retValue, _
                                              e_ResultType, _
                                              sTableName, _
                                              lCodeList, _
                                              sShortDesc, _
                                              sLongDesc)
    
    If Len(Trim(sTableName)) > 0 Then
        sSplitEndToEndCase = sShortDesc
    Else
        sSplitEndToEndCase = retValue
    End If
    
    GetAnswer_SplitEndToEndCase = sSplitEndToEndCase
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Function GetAnswer_EndToEndCase(ByRef pMD As IJDMemberDescription) As String
sMETHOD = "GetAnswer_EndToEndCase"
    GetAnswer_EndToEndCase = ""
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Sub CMCreateClip_Primary(ByRef pMD As IJDMemberDescription, _
                                ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

' Manual Slot
' Check if Primary Clip creation is to be allowed
    sMETHOD = "CMCreateClip_Primary"
    bIsNeeded = True
    
    Dim sAllowedClips As String
    sAllowedClips = GetAnswer_AllowedClips(pMD)
    If LCase(sAllowedClips) = LCase("None") Then
        bIsNeeded = False
    ElseIf LCase(sAllowedClips) = LCase("Secondary") Then
        bIsNeeded = False
    Else
        bIsNeeded = True
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

Public Sub CMCreateClip_Secondary(ByRef pMD As IJDMemberDescription, _
                                ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

' Manual Slot
' Check if Secondary Clip creation is to be allowed
    sMETHOD = "CMCreateClip_Secondary"
    bIsNeeded = True
    
    Dim sAllowedClips As String
    sAllowedClips = GetAnswer_AllowedClips(pMD)
    If LCase(sAllowedClips) = LCase("None") Then
        bIsNeeded = False
    ElseIf LCase(sAllowedClips) = LCase("Primary") Then
        bIsNeeded = False
    Else
        bIsNeeded = True
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

Public Function GetAnswer_AllowedClips(ByRef pMD As IJDMemberDescription) As String
sMETHOD = "GetAnswer_AllowedClips"
    
    On Error GoTo ErrorHandler
    
    Dim retValue As Variant
    Dim lCodeList As Long
    Dim sLongDesc As String
    Dim sShortDesc As String
    Dim sTableName As String
    Dim sPenertrationClip As String
    Dim e_ResultType As EResultType
    Dim oSDO_Helper As StructDetailObjects.Helper
    Const sIASelectionRule As String = "IAAssyConnRules_StructPenetrationSel"
    
    sPenertrationClip = ""
    Set oSDO_Helper = New StructDetailObjects.Helper
    Call oSDO_Helper.GetCustomAttributeParamValue(pMD.CAO, _
                                              sIASelectionRule, _
                                              "AllowedClips", _
                                              retValue, _
                                              e_ResultType, _
                                              sTableName, _
                                              lCodeList, _
                                              sShortDesc, _
                                              sLongDesc)
    
    If Len(Trim(sTableName)) > 0 Then
        sPenertrationClip = sShortDesc
    Else
        sPenertrationClip = retValue
    End If
    
    GetAnswer_AllowedClips = sPenertrationClip
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

'***********************************************************************
' METHOD:  CMCreatePlatePenePhysConn1
'
' DESCRIPTION:  Conditional that determines if there is a physical conn
'               between the base of the penetrating plate part and
'               the penetrated plate part.
'
'***********************************************************************
Public Sub CMCreatePlatePenePhysConn1(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMCreatePlatePenePhysConn1"
    
    bIsNeeded = True
    
    Dim pAssyConn As New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMD.CAO
    
    If pAssyConn.PlatePeneReqPhysConn = ePortInPlatePenePhysConn.Offset Then
        bIsNeeded = False
    End If
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub
'***********************************************************************
' METHOD:  CMCreatePlatePenePhysConn2
'
' DESCRIPTION:  Conditional that determines if there is a physical conn
'               between the offset of the penetrating plate part and
'               the penetrated plate part.
'
'***********************************************************************
Public Sub CMCreatePlatePenePhysConn2(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMCreatePlatePenePhysConn2"
    
    bIsNeeded = True
    
    Dim pAssyConn As New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMD.CAO
    
    If pAssyConn.PlatePeneReqPhysConn = ePortInPlatePenePhysConn.base Then
        bIsNeeded = False
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

'  This function determines if a top flange cut is required
' all cross-sections without a flange return false
' all other cross-sections return true

Public Sub CMCreateTopFlangeCut2(ByRef pMD As IJDMemberDescription, _
                                 ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMCreateTopFlangeCut2"
    
    bIsNeeded = True
    
    ' FlangeCut not required for Split/Angled SplitEndToEndCase :
    Dim bSplitAngleEndToEndCase As Boolean
    bSplitAngleEndToEndCase = IsSplitAngleEndToEndCase(pMD)
    If bSplitAngleEndToEndCase Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim strSectionType As String
    strSectionType = GetCrossSectionType(pMD, False)
        
    If ((strSectionType = "FB") Or (strSectionType = "HalfR") Or _
        (strSectionType = "P") Or (strSectionType = "R") Or _
        (strSectionType = "SB") Or (strSectionType = "SqTu") Or _
        (strSectionType = "RT")) Then bIsNeeded = False
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

'  This function determines if a bottom flange cut is required
' only I cross-ssections return true
' all other cross-sections return false
Public Sub CMCreateBottomFlangeCut2(ByRef pMD As IJDMemberDescription, _
                                    ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMCreateBottomFlangeCut2"
    
    bIsNeeded = False
    
    ' FlangeCut not required for Split/Angled SplitEndToEndCase :
    Dim bSplitAngleEndToEndCase As Boolean
    bSplitAngleEndToEndCase = IsSplitAngleEndToEndCase(pMD)
    If bSplitAngleEndToEndCase Then
        bIsNeeded = False
        Exit Sub
    End If

    Dim strSectionType As String
    strSectionType = GetCrossSectionType(pMD, False)
        
    If ((strSectionType = "I") Or (strSectionType = "ISType") Or _
        (strSectionType = "C_SS") Or (strSectionType = "CSType")) Then bIsNeeded = True
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

'***********************************************************************
' METHOD:  CMChamferTeePhysConn
'
' DESCRIPTION:  Conditional that determines if there is a chamfer for a tee connection.  If not, place a PC
'
'***********************************************************************

Public Sub CMChamferTeePhysConn(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    sMETHOD = "CMChamferTeePhysConn"
    Dim bChanferIsNeeded As Boolean
    
    bIsNeeded = True
    Check_ChamferTeeIsNeeded pMD, bChanferIsNeeded
    If bChanferIsNeeded Then
        bIsNeeded = False
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

'***********************************************************************
' METHOD:  CMChamferTee
'
' DESCRIPTION:  Conditional that determines if there is a chamfer for a tee connection
'
'***********************************************************************
Public Sub CMChamferTee(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    sMETHOD = "CMChamferTee"
    Check_ChamferTeeIsNeeded pMD, bIsNeeded
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub

'***********************************************************************
' METHOD:  Check_ChamferTeeIsNeeded
'
' DESCRIPTION:
'   Checks if a Chamfer at a Tee Connections is required
'
'***********************************************************************
Private Sub Check_ChamferTeeIsNeeded(ByRef pMD As IJDMemberDescription, _
                                     ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    sMETHOD = "Check_ChamferTeeIsNeeded"
    
    Dim retValue As Variant
    Dim retValueType As EResultType
    
    Dim sLongDesc As String
    Dim sShortDesc As String
    Dim sTableName As String
    
    Dim lCodeList As Long
    Dim dChamferThickness As Double

    Dim pHelper As StructDetailObjects.Helper
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Dim oPlateBounded As StructDetailObjects.PlatePart
    Dim oPlateBounding As StructDetailObjects.PlatePart

    'define interface from which to get the question
    Const strIASelectionRuleRoot As String = "IAAssyConnRules_PlateEdByPlateFaceChSel"
    Const dMinThickness As Double = 0.004
    
    bIsNeeded = True
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMD.CAO

    'use helper to get the answer from the question
    'This is the default chamfer thickness, or the value entered by the user
    Set pHelper = New StructDetailObjects.Helper
    Call pHelper.GetCustomAttributeParamValue(pAssyConn.object, _
                                              strIASelectionRuleRoot, _
                                              "ChamferThickness", _
                                              retValue, retValueType, sTableName, _
                                              lCodeList, sShortDesc, sLongDesc)
    
    dChamferThickness = retValue
    
    'get the thickness of the bounded and bounding parts
    Dim dBoundedThick As Double
    Dim dBoundingThick As Double
    
    If pAssyConn.ConnectedObject1Type = SDOBJECT_PLATE And _
       pAssyConn.ConnectedObject2Type = SDOBJECT_PLATE Then
        
        Set oPlateBounded = New StructDetailObjects.PlatePart
        Set oPlateBounded.object = pAssyConn.ConnectedObject1
        dBoundedThick = oPlateBounded.PlateThickness
        
        ' Know that this Plate Edge Bounded by Plate Face case
        ' Want to determine if the Bounded Plate is Split by the Bounding Plate
        ' such that Leaf Part(s) with the same Root Part as the bounded Part
        ' exist on the "other" side of the Bounding Plate
        Dim bPartsOnOtherSide As Boolean
        Dim oPartOnOtherSide As IJElements
        Set oPartOnOtherSide = pAssyConn.PartsOnOtherSide
        
        If oPartOnOtherSide Is Nothing Then
            bPartsOnOtherSide = False
        ElseIf oPartOnOtherSide.Count < 1 Then
            bPartsOnOtherSide = False
        Else
            bPartsOnOtherSide = True
        End If
        
            'check if the assy conn came from an Split type seam
        If bPartsOnOtherSide Then
            ' The Chamfer is based on
            '   the Thckness of the Part on the "Other Side"
            '   dChamferThickness is the thickness of the Part(s) on the "Other Side"
            If dBoundedThick >= (dChamferThickness + dMinThickness) Then
                'a chamfer should be created, and so no phys conn is needed
                bIsNeeded = True
            Else
                bIsNeeded = False
            End If

        Else
            Set oPlateBounding = New StructDetailObjects.PlatePart
            Set oPlateBounding.object = pAssyConn.ConnectedObject2
            dBoundingThick = oPlateBounding.PlateThickness
        
            If (dBoundedThick - dBoundingThick) >= (dChamferThickness + dMinThickness) Then
                'a chamfer should be created, and so no phys conn is needed
                bIsNeeded = True
            Else
                bIsNeeded = False
            End If
        End If
        
    Else
        'not supported for profiles at this time
        bIsNeeded = False
    End If
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD).Number
End Sub
