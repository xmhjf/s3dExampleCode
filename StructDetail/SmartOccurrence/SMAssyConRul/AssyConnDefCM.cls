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
'*******************************************************************
'
'Copyright (C) 2013-16 Intergraph Corporation. All rights reserved.
'
'File : AssyConnDefCM.cls
'
'Author : Alligators
'
'Description :
'
'
'History :
'    5/Nov/2013 - svsmylav
'          DI-CP-240506 'UpdateDependentCornerSeam' method is called conditinally (e.g. seam movement case) for performance improvement.
'                       Also, in CMConstructPCCon, 'UpdateCFWhenSeamIsDeleted' is called for seam movement cases.
'    8/Nov/2013 - vb/svsmylav
'          DI-CP-240506 Added extracheck using 'IsACisWithinRangeOfChange' for seam-movement case.
'
'    21/Nov/2013 - svsmylav
'          TR-234194 Modified 'CMConstructPCCon' to handle stiffener-face to plate-face AC
'                    to create PC for all cases except for the AC within the CF range.
'    12/Oct/2016 - svsmylav
'         DI-259156 Modified 'CMChamferPhysConn' and 'CMConstructPCCon' to delete geometry checks
'         which were added to avoid missing PC TDR. Also cleaned up unused code, revamped
'         'CMCreateBracket' and 'CMCreateChamfer1'.
'*****************************************************************************
Option Explicit

Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "AssyConRul\AssyConnDefCM.cls"


Private sError As String
Private sMethod As String

Public Function CMConstruct_PhysConnFromACPorts(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_PhysConnFromACPorts"
    
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
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Public Function CMConstruct_PhysConnFromSubToACPort(ByVal pMemberDescription As IJDMemberDescription, _
                                                    ByVal pResourceManager As IUnknown, _
                                                    ByVal ePort1Sub As JXSEC_CODE, _
                                                    strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_PhysConnFromSubToACPort"
    
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
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function
Public Function CMConstruct_PhysConnFromBaseToACPort(ByVal pMemberDescription As IJDMemberDescription, _
                                                    ByVal pResourceManager As IUnknown, _
                                                    ByVal ePort2 As Base_Port_Types, _
                                                    strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_PhysConnFromBaseToACPort"
    
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
        Set oPort2 = oPlate.baseport(ePort2)
    ElseIf pAssyConn.ConnectedObject2Type = SDOBJECT_STIFFENER Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = pAssyConn.ConnectedObject2
        Set oPort2 = oProfile.baseport(ePort2)
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
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Public Function CMConstruct_PhysConnBetweenLastPorts(ByVal pMemberDescription As IJDMemberDescription, _
                                                    ByVal pResourceManager As IUnknown, _
                                                    ByVal ePort2 As Base_Port_Types, _
                                                    strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_PhysConnBetweenLastPorts"
        
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
        Set oPort2 = oPlate2.baseport(ePort2)
    ElseIf pAssyConn.ConnectedObject2Type = SDOBJECT_STIFFENER Then
        Dim oProfile As New StructDetailObjects.ProfilePart
        Set oProfile.object = pAssyConn.ConnectedObject2
        Set oPort2 = oProfile.baseport(ePort2)
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
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Public Function CMConstruct_WebCut(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_WebCut"
    
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
    Dim oWebCut As New WebCut
    sError = "Creating web cut."
    If Not pAssyConn.ConnectedObject1Type = SDOBJECT_BEAM Then
        If IsLapped(oPort1) Then
            Call oWebCut.Create(pResourceManager, oPort2, oPort1, "WebCutsLapped", pAssyConn.object)
        Else
            Call oWebCut.Create(pResourceManager, oPort2, oPort1, "WebCuts", pAssyConn.object)
        End If
    ElseIf pAssyConn.ConnectedObject1Type = SDOBJECT_BEAM Then
        Call oWebCut.Create(pResourceManager, oPort2, oPort1, "WebCuts", pAssyConn.object)
    End If
                               
    sError = "Return the webcut"
    
    'CF update when  seam is deleted
    
    UpdateCFWhenSeamIsDeleted pAssyConn.object

    Set pObject = oWebCut.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Public Function CMConstruct_WebCutEndToEndPart1( _
               ByVal pMemberDescription As IJDMemberDescription, _
               ByVal pResourceManager As IUnknown, _
               ByRef pObject As Object)
zMsgBox "AssyConnDefCM::CMConstruct_WebCutEndToEndPart1"
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_WebCut"
    
    ' Get wrapper class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    Dim oACBoundedPort As IJPort
    Dim oACBoundingPort As IJPort
    
    sError = "Setting assembly connection inputs."
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting assembly connection ports."
    Set oACBoundedPort = pAssyConn.Port1
    Set oACBoundingPort = pAssyConn.Port2
           
    Dim oCutBoundingObject As Object
    Dim oCutBoundedObject As Object
    
    On Error Resume Next
    CheckAndCreateBoundingPlane _
                         pMemberDescription, _
                         oACBoundedPort, _
                         oACBoundingPort, _
                         oCutBoundingObject
                         
    On Error GoTo ErrorHandler
    If oCutBoundingObject Is Nothing Then
       Set oCutBoundingObject = oACBoundingPort
    End If
    Set oCutBoundedObject = oACBoundedPort
           
    ' Create webcut
    Dim oWebCut As New WebCut
    sError = "Creating web cut."
    Call oWebCut.Create(pResourceManager, _
                        oCutBoundingObject, _
                        oCutBoundedObject, _
                        "WebCutsEndToEnd", _
                        pAssyConn.object)
zMsgBox "   Webcut.Create WebCutsEndToEnd"
                               
    sError = "Return the webcut"
    Set pObject = oWebCut.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Public Function CMConstruct_WebCutEndToEndPart2(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
zMsgBox "AssyConnDefCM::CMConstruct_WebCutEndToEndPart2"
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_WebCutEndToEnd"
    
    ' Get wrapper class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    Dim oACBoundedPort As IJPort
    Dim oACBoundingPort As IJPort
    
    sError = "Setting assembly connection inputs."
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting assembly connection ports."
    Set oACBoundedPort = pAssyConn.Port1
    Set oACBoundingPort = pAssyConn.Port2

    Dim oCutBoundingObject As Object
    Dim oCutBoundedObject As Object
    
    On Error Resume Next
    CheckAndCreateBoundingPlane _
                         pMemberDescription, _
                         oACBoundedPort, _
                         oACBoundingPort, _
                         oCutBoundingObject
    
    On Error GoTo ErrorHandler
    If oCutBoundingObject Is Nothing Then
       Set oCutBoundingObject = oACBoundedPort
    End If
    Set oCutBoundedObject = oACBoundingPort
    
    ' Create webcut
    Dim oWebCut As New WebCut
    sError = "Creating web cut."
    Call oWebCut.Create(pResourceManager, _
                        oCutBoundingObject, _
                        oCutBoundedObject, _
                        "WebCutsEndToEnd", _
                        pAssyConn.object)
zMsgBox "   Webcut.Create WebCutsEndToEnd"
                               
    sError = "Return the webcut"
    Set pObject = oWebCut.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Public Function CMConstruct_FlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object)
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_FlangeCut"
     
    CreateFlangeCut pMemberDescription, pResourceManager, pObject
    
    ' Add code to set the "TopFlangeCut" symbol parameter to True
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Function

Public Function CMConstruct_FlangeCutEndToEndPart1(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object)
zMsgBox "AssyConnDefCM::CMConstruct_FlangeCutEndToEndPart1"
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_FlangeCutEndToEnd"
     
    CreateFlangeCutEndToEnd pMemberDescription, pResourceManager, pObject
    
    ' Add code to set the "TopFlangeCut" symbol parameter to True
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Function

Public Function CMConstruct_FlangeCutEndToEndPart2(ByVal pMemberDescription As IJDMemberDescription, _
                                      ByVal pResourceManager As IUnknown, _
                                      ByRef pObject As Object)
zMsgBox "AssyConnDefCM::CMConstruct_FlangeCutEndToEndPart2"
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_FlangeCutEndToEnd"
     
    CreateFlangeCutEndToEnd2 pMemberDescription, pResourceManager, pObject
    
    ' Add code to set the "TopFlangeCut" symbol parameter to True
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Function

Public Function CMConstruct_BottomFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                            ByVal pResourceManager As IUnknown, _
                                            ByRef pObject As Object)
    On Error GoTo ErrorHandler

    CreateFlangeCut pMemberDescription, pResourceManager, pObject
    
    ' Add code to set the "TopFlangeCut" symbol parameter to False
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
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
    Err.Raise LogError(Err, MODULE, sMethod).Number
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
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Function

Private Function CreateFlangeCut(ByVal pMemberDescription As IJDMemberDescription, _
                                 ByVal pResourceManager As IUnknown, _
                                 ByRef pObject As Object)
    On Error GoTo ErrorHandler

    sMethod = "CreateFlangeCut"
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
    Dim oWebCut As IJStructFeature
    Dim bFound As Boolean
    bFound = False
    ' not all members are webcuts so allow the Set to fail
    On Error Resume Next
    Dim iCount As Long, iIndex As Long
    iCount = pMemberObjects.Count
    For iIndex = 1 To iCount
        If (bFound = False) Then
            Set oWebCut = pMemberObjects.Item(iIndex)
            If (Not (oWebCut Is Nothing)) Then
                If (oWebCut.get_StructFeatureType = SF_WebCut) Then
                   Call oFlangeCut.Create(pResourceManager, oBoundingPort, oBoundedPort, oWebCut, "FlangeCuts", _
                                          pAssyConn.object)
                   sError = "Return the flange cut "
                   Set pObject = oFlangeCut.object
                   bFound = True
               End If
               Set oWebCut = Nothing
            End If
       End If
    Next iIndex
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Private Function CreateFlangeCutEndToEnd(ByVal pMemberDescription As IJDMemberDescription, _
                                 ByVal pResourceManager As IUnknown, _
                                 ByRef pObject As Object)
    On Error GoTo ErrorHandler

    sMethod = "CreateFlangeCutEndToEnd"
     ' Get wrapper class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    Dim oACBoundedPort As IJPort
    Dim oACBoundingPort As IJPort
    
    sError = "Setting assembly connection inputs."
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting assembly connection ports."
    Set oACBoundedPort = pAssyConn.Port1
    Set oACBoundingPort = pAssyConn.Port2
    
    ' Create flange cut if one is required
    Dim oFlangeCut As New FlangeCut
    sError = "Creating flange cut."
    
    ' create the flange when the web cut is found
    Dim pMemberObjects As IJDMemberObjects
    Set pMemberObjects = pMemberDescription.CAO
    Dim oWebCut As IJStructFeature
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
    Set oWebCut = pMemberObjects.Item(1)
    If (Not (oWebCut Is Nothing)) Then
        If (oWebCut.get_StructFeatureType = SF_WebCut) Then
            Dim oCutBoundingObject As Object
            Dim oCutBoundedObject As Object
            
            On Error Resume Next
            CheckAndCreateBoundingPlane _
                                 pMemberDescription, _
                                 oACBoundedPort, _
                                 oACBoundingPort, _
                                 oCutBoundingObject
                                 
            On Error GoTo ErrorHandler
            If oCutBoundingObject Is Nothing Then
               Set oCutBoundingObject = oACBoundingPort
            End If
            Set oCutBoundedObject = oACBoundedPort
            
            Call oFlangeCut.Create(pResourceManager, _
                                   oCutBoundingObject, _
                                   oCutBoundedObject, _
                                   oWebCut, _
                                   "FlangeCutsEndToEnd", _
                                   pAssyConn.object)
            sError = "Return the flange cut "
            Set pObject = oFlangeCut.object
        End If
    End If
                   
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Private Function CreateFlangeCutEndToEnd2(ByVal pMemberDescription As IJDMemberDescription, _
                                 ByVal pResourceManager As IUnknown, _
                                 ByRef pObject As Object)
    On Error GoTo ErrorHandler

    sMethod = "CreateFlangeCutEndToEnd2"
     ' Get wrapper class
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    Dim oACBoundedPort As IJPort
    Dim oACBoundingPort As IJPort
    
    sError = "Setting assembly connection inputs."
    Set pAssyConn.object = pMemberDescription.CAO
    sError = "Getting assembly connection ports."
    Set oACBoundedPort = pAssyConn.Port1
    Set oACBoundingPort = pAssyConn.Port2
    
    ' Create flange cut if one is required
    Dim oFlangeCut As New FlangeCut
    sError = "Creating flange cut."
    
    ' create the flange when the web cut is found
    Dim pMemberObjects As IJDMemberObjects
    Set pMemberObjects = pMemberDescription.CAO
    Dim oWebCut As IJStructFeature
    
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
    Set oWebCut = pMemberObjects.Item(4)
    If (Not (oWebCut Is Nothing)) Then
        If (oWebCut.get_StructFeatureType = SF_WebCut) Then
            Dim oCutBoundingObject As Object
            Dim oCutBoundedObject As Object
            
            On Error Resume Next
            CheckAndCreateBoundingPlane _
                                 pMemberDescription, _
                                 oACBoundedPort, _
                                 oACBoundingPort, _
                                 oCutBoundingObject
                                 
            On Error GoTo ErrorHandler
            If oCutBoundingObject Is Nothing Then
               Set oCutBoundingObject = oACBoundedPort
            End If
            Set oCutBoundedObject = oACBoundingPort
            
            Call oFlangeCut.Create(pResourceManager, _
                                   oCutBoundingObject, _
                                   oCutBoundedObject, _
                                   oWebCut, _
                                   "FlangeCutsEndToEnd", _
                                   pAssyConn.object)
            sError = "Return the flange cut "
            Set pObject = oFlangeCut.object
        End If
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

'  This function determines if a top flange cut is required
' all cross-sections without a flange return false
' all other cross-sections return true

Public Sub CMCreateTopFlangeCut(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMethod = "CMCreateTopFlangeCut"
    
    bIsNeeded = True

    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
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
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Sub


'  This function determines if a bottom flange cut is required
' only I cross-ssections return true
' all other cross-sections return false

Public Sub CMCreateBottomFlangeCut(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMethod = "CMCreateBottomFlangeCut"
    
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
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
        (strSectionType = "C_SS") Or (strSectionType = "CSType") Or _
        (strSectionType = "H")) Then
        bIsNeeded = True
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Sub

Private Function GetCrossSectionType(ByRef pMD As IJDMemberDescription, _
                        Optional bConnectedObject1 As Boolean = True) As String
    On Error GoTo ErrorHandler

    ' For Compatibility,
    ' Do Not remove/rename any existing methods
    sMethod = "GetCrossSectionType"
    
    GetCrossSectionType = Get_CrossSectionType(pMD, bConnectedObject1)
    
        
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Public Function CMConstruct_Slot(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_Slot"
    
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
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Public Function CMConstruct_Bracket2S(ByVal pMemberDescription As IJDMemberDescription, _
                                   ByVal pResourceManager As IUnknown, _
                                   ByRef pObject As Object)
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_Bracket2S"
    
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
    
    If pAssyConn.ConnectedObject2Type = SDOBJECT_PLATE Then
        Set oPort2 = pAssyConn.Port2
    Else
        Dim oProfile2 As New StructDetailObjects.ProfilePart
        Set oProfile2.object = pAssyConn.ConnectedObject2
        Set oPort2 = oProfile2.SubPort(JXSEC_TOP)
    End If
    
    ' Get Location for Bracket Plane from Assembly Connection
    Dim oLocation As IJDPosition
    If pAssyConn.ConnectionBehavior = ConnectionStandard Then
        Set oLocation = pAssyConn.BoundGlobalShipLocation
    Else
        Set oLocation = pAssyConn.PenetrationGlobalShipLocation
    End If
    
    Dim oPortOnWebLeft As IJPort
    Set oPortOnWebLeft = oProfile1.SubPort(JXSEC_WEB_LEFT)
    
    Dim oNewLocation As IJDPosition
    Set oNewLocation = PointOnWebLeft(oPortOnWebLeft, oLocation)
    
    Set oLocation = Nothing
    Set oPortOnWebLeft = Nothing
    
    ' Get plane of Web_Left
    Dim pRefPlane As IJPlane
    Set pRefPlane = oProfile1.EdgePlane(JXSEC_WEB_LEFT, oNewLocation, pResourceManager)
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
    
    Set oIJDProfileSection = oProfile1.object
    Set oIJCrossSection = oIJDProfileSection.crossSection
    strCrossSectionType = oIJCrossSection.Type
    
    Set oIJDProfileSection = Nothing
    Set oIJCrossSection = Nothing
    
    Dim dThicknessOffset As Double
    
    If (strCrossSectionType = "FB") Or (strCrossSectionType = "SB") Or _
        (strCrossSectionType = "T_XType") Or (strCrossSectionType = "TSType") Or _
        (strCrossSectionType = "I") Or (strCrossSectionType = "ISType") Or _
        (strCrossSectionType = "H") Or _
        (strCrossSectionType = "BUT") Or (strCrossSectionType = "BUTL2") Or (strCrossSectionType = "BUTL3") Then
        'dThicknessOffset = 0# 'default value = 0 on the object
    Else
        dThicknessOffset = 0.005
        oBracketPart.ThicknessOffset = dThicknessOffset
    End If
    
    Set oBracketPart = Nothing
     
    sError = "Return the bracket"
    Set pObject = oBracket.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function


Private Function PointOnWebLeft(oPortOnWebLeft As IJPort, oPlanePnt As IJDPosition) As IJDPosition

On Error GoTo ErrorHandler
sMethod = "PlaneOnWebLeft"

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
   Err.Raise LogError(Err, MODULE, sMethod).Number
End Function


Public Function CMConstruct_Chamfer(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_Chamfer"
    
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
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function
Public Function CMConstruct_Chamfer2(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_Chamfer"
    
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
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function


'***********************************************************************
' METHOD:  CMConstruct_Collar
'
' DESCRIPTION:  Creates a collar part.
'
'***********************************************************************
Public Function CMConstruct_Collar(ByVal pSlot As IJStructFeature, strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMethod = "CMConstruct_Collar"

    ' Create New Collar
    Dim oCollar As New StructDetailObjects.Collar
    sError = "Creating Collar"
    oCollar.Create pSlot, strStartClass
                               
    sError = "Setting Collar to private variable"
    Set CMConstruct_Collar = oCollar.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function
'***********************************************************************
' METHOD:  CMCreateCutout
'
' DESCRIPTION:  Create a cutting tool and set the SDCut_AE operation
'
'***********************************************************************
Public Function CMCreateCutout(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, strStartClass As String) As Object
    On Error GoTo ErrorHandler

    sMethod = "CMCreateCutout"
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
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function '***********************************************************************
' METHOD:  CMCreateBracket
'
' DESCRIPTION:  Conditional that determines if there is a bracket
'
'***********************************************************************
Public Sub CMCreateBracket(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMethod = "CMCreateBracket"
     
    bIsNeeded = False 'Set the conditional value
        If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_Bracket) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sPlaceBracket As String
    
    'TR-169311 Question's answer to be retrived without hard coding
    'the path by calling GetSelectorAnswer() method
    
    GetSelectorAnswer pMD.CAO, "PlaceBracket", sPlaceBracket
    If sPlaceBracket = "Yes" Then
        bIsNeeded = True
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Sub

'***********************************************************************
' METHOD:  CMCreateChamfer1
'
' DESCRIPTION:  Conditional that determines if there is a chamfer 1
'
'***********************************************************************
Public Sub CMCreateChamfer1(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMethod = "CMCreateChamfer1"
    
    bIsNeeded = False 'Set the conditional value
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_Chamfer) Then
        bIsNeeded = False
        Exit Sub
    End If

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

    If CheckMultipleAssemblyConnections(pAssyConn.ConnectedObject1, _
                                                 pAssyConn.Port1, _
                                                 pAssyConn.ConnectedObject2, _
                                                 pAssyConn.Port2) Then
        Exit Sub 'Note: bIsNeeded is already initialized to False
    End If
   
    Dim sChamferType As String

    'TR-169311 Question's answer to be retrived without hard coding
    'the path by calling below method
    GetSelectorAnswer pMD.CAO, "ChamferType", sChamferType
    If sChamferType = "None" Then Exit Sub 'Note: bIsNeeded is already initialized to False
        
    If (sChamferType = "Obj1Double") Or (sChamferType = "Obj1Base") Or (sChamferType = "Obj1Offset") Or _
        (sChamferType = "Obj1BaseObj2Offset") Or (sChamferType = "Obj1OffsetObj2Base") Then
        bIsNeeded = True
    End If
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Sub
'***********************************************************************
' METHOD:  CMCreateChamfer2
'
' DESCRIPTION:  Conditional that determines if there is a chamfer 2.
'
'***********************************************************************
Public Sub CMCreateChamfer2(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMethod = "CMCreateChamfer2"
    
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_Chamfer) Then
        bIsNeeded = False
        Exit Sub
    End If
    
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
    Const strIASelectionRuleRoot As String = "IA" + CUSTOMERID + "AssyConRul_PlateEdgeByPlateEdgeSel"
    Dim pHelper As New StructDetailObjects.Helper
    Dim retValue As Variant
    Dim retValueType As EResultType
    Dim sTableName As String
    Dim lCodeList As Long
    Dim sShortDesc As String
    Dim sLongDesc As String
    Dim sChamferType As String
    
    
    'TR-169311 Question's answer to be retrived without hard coding
    'the path by calling below method
    GetSelectorAnswer pMD.CAO, "ChamferType", sChamferType
    
    If (sChamferType = "Obj2Double") Or (sChamferType = "Obj2Base") Or (sChamferType = "Obj2Offset") Or _
        (sChamferType = "Obj1BaseObj2Offset") Or (sChamferType = "Obj1OffsetObj2Base") Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Sub
'***********************************************************************
' METHOD:  CMChamferPhysConn
'
' DESCRIPTION:  Conditional that determines if there is a chamfer.  If not, place a PC
'
'***********************************************************************
Public Sub CMChamferPhysConn(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    

    sMethod = "CMChamferPhysConn"
    
    Dim sChamferType As String
        
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    'TR-169311 Question's answer to be retrived without hard coding
    'the path by calling GetSelectorAnswer() method
    GetSelectorAnswer pMD.CAO, "ChamferType", sChamferType
    If sChamferType = "None" Then
        bIsNeeded = True 'Set the conditional value
        
        Dim pAssyConn As StructDetailObjects.AssemblyConn
        Set pAssyConn = New StructDetailObjects.AssemblyConn
        
        Set pAssyConn.object = pMD.CAO
        
        Dim lOpn As Long
        Dim gRngBox As GBox
        
        If IsSeamMovement(pAssyConn.Port2, lOpn, gRngBox) Then
            'Seam-movement case: check if the current AC is within the range-of-change box
            If IsACisWithinRangeOfChange(pAssyConn.object, gRngBox) Then
                UpdateDependentCornersSeam pAssyConn.object
            End If
        Else
            If lOpn <> GEOMETRY_OPERATION_PLATE_GENERATE Then 'for lOpn = 8 i.e. plate thickness change avoid force-update of CF
                UpdateDependentCornersSeam pAssyConn.object
            End If
        End If
        Set pAssyConn = Nothing
    Else
        bIsNeeded = False 'No PC for a valid chamfer
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
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
        dThickness = pProfilePart.webThickness
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

    sMethod = "CMSetSideOfPlate"
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
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Public Function ConvertMountingFaceToPort(eFaceName As ProfileFaceName) As JXSEC_CODE
    
    Dim ePort As JXSEC_CODE
    ePort = JXSEC_UNKNOWN
    
    On Error GoTo ErrorHandler
    sMethod = "ConvertMountingFaceToPort"
    
    
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
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
    ConvertMountingFaceToPort = ePort
End Function

Public Function IsSplitAngleEndToEndCase(ByRef pMD As IJDMemberDescription) As Boolean
sMethod = "IsSplitAngleEndToEndCase"
    
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
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Public Function GetAnswer_SplitEndToEndCase(ByRef pMD As IJDMemberDescription) As String
sMethod = "GetAnswer_SplitEndToEndCase"
    
    On Error GoTo ErrorHandler
    
    Dim retValue As Variant
    Dim lCodeList As Long
    Dim sLongDesc As String
    Dim sShortDesc As String
    Dim sTableName As String
    Dim sSplitEndToEndCase As String
    
    Dim e_ResultType As EResultType
    Dim oSDO_Helper As StructDetailObjects.Helper
    Const strIASelectionRuleRoot As String = "IA" + CUSTOMERID + "AssyConRul_StiffEndByStiffEndSel"
    
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
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Public Function GetAnswer_EndToEndCase(ByRef pMD As IJDMemberDescription) As String
sMethod = "GetAnswer_EndToEndCase"
    GetAnswer_EndToEndCase = ""
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
End Function

Public Sub CMCreateClip_Primary(ByRef pMD As IJDMemberDescription, _
                                ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

' Manual Slot
' Check if Primary Clip creation is to be allowed
    sMethod = "CMCreateClip_Primary"
    bIsNeeded = True
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_Collar) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim bvalid As Boolean
    
    Dim sAllowedClips As String
    sAllowedClips = GetAnswer_AllowedClips(pMD)
    If LCase(sAllowedClips) = LCase("None") Then
        bIsNeeded = False
    ElseIf LCase(sAllowedClips) = LCase("Secondary") Then
        bIsNeeded = False
    Else
        IsClipValid pMD, bvalid, "Primary"
        If bvalid Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Sub

Public Sub CMCreateClip_Secondary(ByRef pMD As IJDMemberDescription, _
                                ByRef bIsNeeded As Boolean)
        On Error GoTo ErrorHandler
        
        

' Manual Slot
' Check if Secondary Clip creation is to be allowed
    sMethod = "CMCreateClip_Secondary"
    bIsNeeded = True
    Dim bvalid As Boolean
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_Collar) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim sAllowedClips As String
    sAllowedClips = GetAnswer_AllowedClips(pMD)
    If LCase(sAllowedClips) = LCase("None") Then
        bIsNeeded = False
    ElseIf LCase(sAllowedClips) = LCase("Primary") Then
        bIsNeeded = False
    Else
        IsClipValid pMD, bvalid, "Secondary"
        If bvalid Then
            bIsNeeded = True
        Else
            bIsNeeded = False
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Sub

Public Function GetAnswer_AllowedClips(ByRef pMD As IJDMemberDescription) As String
sMethod = "GetAnswer_AllowedClips"
    
    On Error GoTo ErrorHandler
    
    Dim retValue As Variant
    Dim lCodeList As Long
    Dim sLongDesc As String
    Dim sShortDesc As String
    Dim sTableName As String
    Dim sPenertrationClip As String
    Dim e_ResultType As EResultType
    Dim oSDO_Helper As StructDetailObjects.Helper
    Const sIASelectionRule As String = "IA" + CUSTOMERID + "AssyConRul_StructPenetrationSel"
    
    sPenertrationClip = ""
    
    'TR-169311 Question's answer to be retrived without hard coding
    'the path by calling below method
    GetSelectorAnswer pMD.CAO, "AllowedClips", sPenertrationClip
    
    GetAnswer_AllowedClips = sPenertrationClip
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
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

    sMethod = "CMCreatePlatePenePhysConn1"
    
    bIsNeeded = True
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim pAssyConn As New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMD.CAO
    
    If pAssyConn.PlatePeneReqPhysConn = ePortInPlatePenePhysConn.Offset Then
        bIsNeeded = False
    End If
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
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

    sMethod = "CMCreatePlatePenePhysConn2"
    
    bIsNeeded = True
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim pAssyConn As New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMD.CAO
    
    If pAssyConn.PlatePeneReqPhysConn = ePortInPlatePenePhysConn.Base Then
        bIsNeeded = False
    End If
        
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Sub

'  This function determines if a top flange cut is required
' all cross-sections without a flange return false
' all other cross-sections return true

Public Sub CMCreateTopFlangeCut2(ByRef pMD As IJDMemberDescription, _
                                 ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMethod = "CMCreateTopFlangeCut2"
    
    bIsNeeded = True
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
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
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Sub

'  This function determines if a bottom flange cut is required
' only I cross-ssections return true
' all other cross-sections return false
Public Sub CMCreateBottomFlangeCut2(ByRef pMD As IJDMemberDescription, _
                                    ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMethod = "CMCreateBottomFlangeCut2"
    
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_FlangeCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
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
        (strSectionType = "H") Or _
        (strSectionType = "C_SS") Or (strSectionType = "CSType")) Then bIsNeeded = True
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Sub

'***********************************************************************
' METHOD:  CMChamferTeePhysConn
'
' DESCRIPTION:  Conditional that determines if there is a chamfer for a tee connection.  If not, place a PC
'
'***********************************************************************

Public Sub CMChamferTeePhysConn(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    sMethod = "CMChamferTeePhysConn"
    Dim bChanferIsNeeded As Boolean
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True
    Check_ChamferTeeIsNeeded pMD, bChanferIsNeeded
    If bChanferIsNeeded Then
        bIsNeeded = False
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Sub

'***********************************************************************
' METHOD:  CMChamferTee
'
' DESCRIPTION:  Conditional that determines if there is a chamfer for a tee connection
'
'***********************************************************************
Public Sub CMChamferTee(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    sMethod = "CMChamferTee"
        
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_Chamfer) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Check_ChamferTeeIsNeeded pMD, bIsNeeded
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
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
    sMethod = "Check_ChamferTeeIsNeeded"
    
    Dim retValue As Variant
    Dim retValueType As EResultType
    
    Dim sLongDesc As String
    Dim sShortDesc As String
    Dim sTableName As String
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_Chamfer) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim lCodeList As Long
    Dim dChamferThickness As Double

    Dim pHelper As StructDetailObjects.Helper
    Dim pAssyConn As StructDetailObjects.AssemblyConn
    Dim oPlateBounded As StructDetailObjects.PlatePart
    Dim oPlateBounding As StructDetailObjects.PlatePart

    'define interface from which to get the question
    Const strIASelectionRuleRoot As String = "IA" + CUSTOMERID + "AssyConRul_PlateEdByPlateFaceChSel"
    Const dMinThickness As Double = 0.004
    
    bIsNeeded = True
    Set pAssyConn = New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMD.CAO

    'use helper to get the answer from the question
    'This is the default chamfer thickness, or the value entered by the user
    Set pHelper = New StructDetailObjects.Helper
    
    'TR-169311 Question's answer to be retrived without hard coding
    'the path by calling GetSelectorAnswer() method
    GetSelectorAnswer pMD.CAO, "ChamferThickness", dChamferThickness
    
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
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Sub

' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' Dummy Methods used to keep "cached" method DispIds the same
' If the DispIds changes, a Synchronize will be required
' "Custom Methods" Disp Ids that are referenced in the symbol defintions (Selection, Definition, parameter rules)
' are "cached" during Bulkload,
' Therefore, existing methods should NEVER be deleted
' and New methods MUST be added at the bottom
'
Private Sub Dummy_DispId6003002F()
End Sub
Private Sub Dummy_DispId60030030()
End Sub
Private Sub Dummy_DispId60030031()
End Sub
Private Sub Dummy_DispId60030032()
End Sub
Private Sub Dummy_DispId60030033()
End Sub
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Sub CMConstructPCCon(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    

    sMethod = "CMConstructPCCon"
    
    Dim pAssyConn As New StructDetailObjects.AssemblyConn
    Set pAssyConn.object = pMD.CAO
    
    
    bIsNeeded = True
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True 'Set the conditional value
    
    'Consider plate-bounded-to-plate test case such that a CF-with-seam (e.g. LongScallopwithSeam) exists on bounded plate.
    'Let the seam is moved away from bounded plate by two meters - in order to update the CF we need to force-update it from here
    '(since force-update from the construct method associated with this conditional will not trigger for above case):
    Dim oBoundingPort As IJPort
    
    Set oBoundingPort = pAssyConn.Port2
    If Not oBoundingPort Is Nothing Then
        If IsSeamMovement(oBoundingPort) Then
            UpdateCFWhenSeamIsDeleted pAssyConn.object
        End If
    End If

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number

End Sub

' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
' Dummy Methods used to keep "cached" method DispIds the same
' If the DispIds changes, a Synchronize will be required
' "Custom Methods" Disp Ids that are referenced in the symbol defintions (Selection, Definition, parameter rules)
' are "cached" during Bulkload,
' Therefore, existing methods should NEVER be deleted
' and New methods MUST be added at the bottom
'
Private Sub Dummy_DispId60030035()
End Sub
Private Sub Dummy_DispId60030036()
End Sub
Private Sub Dummy_DispId60030037()
End Sub
Private Sub Dummy_DispId60030038()
End Sub
Private Sub Dummy_DispId60030039()
End Sub
Private Sub Dummy_DispId6003003A()
End Sub
' $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

Public Sub CMBracketCornerFeatureConn(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
 
  On Error GoTo ErrorHandler
    sMethod = "CMBracketCornerFeatureConn"
  
    
    bIsNeeded = False
    
    Dim oBracketPort As IJPort
    Dim oOtherPort As IJPort
    Dim oBracket As Object
    Dim oSupportRoot As Object
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_CornerFeature) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    GetPortsForBracketInAC pMD, oBracketPort, oOtherPort, oBracket, oSupportRoot
    
    If oBracketPort Is Nothing Then
        Exit Sub
    End If
    
    Dim oS1 As Object
    Dim oS2 As Object
    Dim oS3 As Object
    Dim oS4 As Object
    Dim oS5 As Object
    Dim oS1Port As IJPort
    Dim oS2Port As IJPort
    Dim oS3Port As IJPort
    Dim oS4Port As IJPort
    Dim oS5Port As IJPort
    
    GetBracketSupports oBracket, oS1, oS2, oS3, oS4, oS5, oS1Port, oS2Port, oS3Port, oS4Port, oS5Port
    
    'use helper to get the answer from the question
    'This is the default Answer, or the value entered by the user
    Dim sAnswer As String

    Dim oHelper As New CommonHelper
    Dim oSO As IJSmartOccurrence
    Dim oSI As IJSmartItem
    Dim oSC As IJSmartClass
    
    Dim oPosition_S1S2 As IJDPosition
    Dim oPosition_S1S3 As IJDPosition
    Dim oPosition_S2S4 As IJDPosition
    Dim oPosition_S3S5 As IJDPosition
    
    Dim oS1PortBeforeCut As IJPort
    Dim oS2PortBeforeCut As IJPort
    Dim oS3PortBeforeCut As IJPort
    Dim oS4PortBeforeCut As IJPort
    Dim oS5PortBeforeCut As IJPort
      
    Dim oIntersect_S1S2 As New TopologyLocate
    Dim oIntersect_S1S3 As New TopologyLocate
    Dim oIntersect_S2S4 As New TopologyLocate
    Dim oIntersect_S3S5 As New TopologyLocate
    
    Set oSO = pMD.CAO
    Set oSI = oSO.ItemObject
    Set oSC = oSI.Parent

 
    sAnswer = oHelper.GetAnswer(oSO, _
                            oSC.SelectionRuleDef, _
                            "AddCornerFeature")
    bIsNeeded = False
   
   If Not oS1Port Is Nothing Then
      Set oS1PortBeforeCut = RelatedPortBeforeCut(oS1Port)
   Else
      Set oS1PortBeforeCut = Nothing
   End If
   
   If Not oS2Port Is Nothing Then
      Set oS2PortBeforeCut = RelatedPortBeforeCut(oS2Port)
   Else
      Set oS2PortBeforeCut = Nothing
   End If
   
   If Not oS3Port Is Nothing Then
      Set oS3PortBeforeCut = RelatedPortBeforeCut(oS3Port)
   Else
      Set oS3PortBeforeCut = Nothing
   End If
   
   If Not oS4Port Is Nothing Then
      Set oS4PortBeforeCut = RelatedPortBeforeCut(oS4Port)
   Else
      Set oS4PortBeforeCut = Nothing
   End If
   
   If Not oS5Port Is Nothing Then
      Set oS5PortBeforeCut = RelatedPortBeforeCut(oS5Port)
   Else
      Set oS5PortBeforeCut = Nothing
   End If
   
   If Not oS1PortBeforeCut Is Nothing And _
      Not oS2PortBeforeCut Is Nothing Then
          Set oPosition_S1S2 = oIntersect_S1S2.FindIntersectionPoint(oS1PortBeforeCut, oS2PortBeforeCut)
   Else
         Set oPosition_S1S2 = Nothing
   End If
   
   If Not oS1PortBeforeCut Is Nothing And _
      Not oS3PortBeforeCut Is Nothing Then
          Set oPosition_S1S3 = oIntersect_S1S3.FindIntersectionPoint(oS1PortBeforeCut, oS3PortBeforeCut)
   Else
          Set oPosition_S1S3 = Nothing
   End If
   
   If Not oS2PortBeforeCut Is Nothing And _
      Not oS4PortBeforeCut Is Nothing Then
          Set oPosition_S2S4 = oIntersect_S2S4.FindIntersectionPoint(oS2PortBeforeCut, oS4PortBeforeCut)
   Else
          Set oPosition_S2S4 = Nothing
   End If
   
   If Not oS3PortBeforeCut Is Nothing And _
      Not oS5PortBeforeCut Is Nothing Then
          Set oPosition_S3S5 = oIntersect_S3S5.FindIntersectionPoint(oS3PortBeforeCut, oS5PortBeforeCut)
   Else
          Set oPosition_S3S5 = Nothing
   End If
   
   
   If sAnswer = "Yes" Then
   
    ' AC on support2 has responsibility to create corner between 1 and 2 - CMT changed support1 to support2
    ' AC on support3 has responsibility to create corner between 1 and 3
    ' AC on support4 has responsibility to create corner between 2 and 4
    ' AC on support5 has responsibility to create corner between 3 and 5
    
    'CMT changed oS1 to oS2 so that S2 is responsible for creating the corner feature between S1 and S2
        If IsSupportOrPlateForStiffenerSupport(oSupportRoot, oS2) And _
           Not oPosition_S1S2 Is Nothing And _
           Not oS1Port Is Nothing And _
           Not oS2Port Is Nothing _
        Or _
           IsSupportOrPlateForStiffenerSupport(oSupportRoot, oS3) And _
           Not oPosition_S1S3 Is Nothing And _
           Not oS1Port Is Nothing And _
           Not oS3Port Is Nothing _
        Or _
           IsSupportOrPlateForStiffenerSupport(oSupportRoot, oS4) And _
           Not oPosition_S2S4 Is Nothing And _
           Not oS2Port Is Nothing And _
           Not oS4Port Is Nothing _
        Or _
           IsSupportOrPlateForStiffenerSupport(oSupportRoot, oS5) And _
           Not oPosition_S3S5 Is Nothing And _
           Not oS3Port Is Nothing And _
           Not oS5Port Is Nothing _
        Then
                bIsNeeded = True
        End If
        
    End If
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
 
 End Sub

 Public Sub CMBracketCornerFeature(ByVal oMD As IJDMemberDescription, _
                                  ByVal pResourceManager As IUnknown, _
                                  ByRef oObject As Object)

 
 
  On Error GoTo ErrorHandler
    sMethod = "CMBracketCornerFeature"
 
    
    Dim oBracketPort As IJPort
    Dim oOtherPort As IJPort
    Dim oBracketPart As Object
    Dim oSupportRoot As Object
    
    GetPortsForBracketInAC oMD, oBracketPort, oOtherPort, oBracketPart, oSupportRoot
    
    Dim oS1 As Object
    Dim oS2 As Object
    Dim oS3 As Object
    Dim oS4 As Object
    Dim oS5 As Object
    Dim oS1Port As IJPort
    Dim oS2Port As IJPort
    Dim oS3Port As IJPort
    Dim oS4Port As IJPort
    Dim oS5Port As IJPort
    
    GetBracketSupports oBracketPart, oS1, oS2, oS3, oS4, oS5, oS1Port, oS2Port, oS3Port, oS4Port, oS5Port

    Dim oPlate As New StructDetailObjects.PlatePart
    Dim oFacePort As IJPort
    Dim oEdgePort1 As IJPort
    Dim oEdgePort2 As IJPort
    Set oPlate.object = oBracketPart
       
    Dim oSupportV As Object
    Dim oSupportU As Object
        
   'CMT changed oS1 to oS2 so that S2 is responsible for creating the corner feature between S1 and S2
    If IsSupportOrPlateForStiffenerSupport(oSupportRoot, oS2) Then
        Set oEdgePort1 = oS1Port
        Set oEdgePort2 = oS2Port
        Set oSupportU = oS1
        Set oSupportV = oS2
    ElseIf IsSupportOrPlateForStiffenerSupport(oSupportRoot, oS3) Then
        Set oEdgePort1 = oS1Port
        Set oEdgePort2 = oS3Port
        Set oSupportU = oS1
        Set oSupportV = oS3
    ElseIf IsSupportOrPlateForStiffenerSupport(oSupportRoot, oS4) Then
        Set oEdgePort1 = oS2Port
        Set oEdgePort2 = oS4Port
         Set oSupportU = oS2
        Set oSupportV = oS4
    ElseIf IsSupportOrPlateForStiffenerSupport(oSupportRoot, oS5) Then
        Set oEdgePort1 = oS3Port
        Set oEdgePort2 = oS5Port
        Set oSupportU = oS3
        Set oSupportV = oS5
    Else
        ' Error condition
        Exit Sub
    End If
    
    Dim oEdgePort1ForCut As IJPort
    Dim oEdgePort2ForCut As IJPort
    
    Set oEdgePort1ForCut = RelatedPortBeforeCut(oEdgePort1)
    Set oEdgePort2ForCut = RelatedPortBeforeCut(oEdgePort2)
    
    Dim oPosition As IJDPosition
        
    'get supportV object , if its a profile, get its primary orientation
    If TypeOf oSupportV Is IJStiffener And TypeOf oSupportU Is IJPlate Then
        
        Dim oOtherPort1 As IJPort
        Dim oOtherPort2 As IJPort
        Dim oOtherConnectable1 As IJConnectable
        Dim oOtherConnectable2 As IJConnectable
        
        GetOtherPortAndConnectable oEdgePort1ForCut, oOtherPort1, oOtherConnectable1
        GetOtherPortAndConnectable oEdgePort2ForCut, oOtherPort2, oOtherConnectable2
        
        Dim ppEnumConnections As IJElements
        Dim pConnected As Boolean
        pConnected = False
        Dim pObject As Object
        Dim oProfilehelper As IJProfileAttributes
        Set oProfilehelper = New ProfileUtils
        Dim oSecOrien As IJDVector
        Dim oPrimOrien As IJDVector
        
        If Not oOtherConnectable1 Is Nothing And Not oOtherConnectable2 Is Nothing Then
        
                oOtherConnectable1.isConnectedTo oOtherConnectable2, pConnected, ppEnumConnections
                
                If pConnected Then
                    For Each pObject In ppEnumConnections
                        If TypeOf pObject Is IJAssemblyConnection Then
                            Dim oAC As New StructDetailObjects.AssemblyConn
                            Set oAC.object = pObject
                            If oAC.ConnectionBehavior = ConnectionPenetration Then
                                
                                Set oPosition = oAC.PenetrationGlobalShipLocation
                                oProfilehelper.GetProfileOrientation oSupportV, oPosition, oSecOrien, oPrimOrien
                                Exit For
                            End If
                        End If
                    Next
                End If
        End If
        If Not oPrimOrien Is Nothing Then
            'get the Bracket part normal
            Dim oBracketNormal As IJDVector
            Set oBracketNormal = New dVector
            
            Dim oStructDetailHelper As StructDetailHelper
            Set oStructDetailHelper = New StructDetailHelper
            
            Dim oBracketRoot As IJPlate
            oStructDetailHelper.IsPartDerivedFromSystem oBracketPart, oBracketRoot, True
            
            Dim x As Double:  Dim y As Double:           Dim z As Double
            
            If TypeOf oBracketRoot Is IJPlane Then
                Dim oPlane As IJPlane
                Set oPlane = oBracketRoot
                oPlane.GetNormal x, y, z
            End If
            
            oBracketNormal.x = x:       oBracketNormal.y = y:         oBracketNormal.z = z
            oBracketNormal.Length = 1
            
            'find out the anlge between bracket part and primary orientation of the bracket
            Dim oAngle As Double
            oAngle = oPrimOrien.Angle(oBracketNormal, oPrimOrien)
            If Cos(oAngle) < 0 Then
                Set oFacePort = oPlate.baseport(BPT_Base)
            ElseIf Cos(oAngle) > 0 Then
                Set oFacePort = oPlate.baseport(BPT_Offset)
            Else
                Set oFacePort = oPlate.baseport(BPT_Base)
            End If
        Else
            Set oFacePort = oPlate.baseport(BPT_Base)
        End If
    
    Else
        Set oFacePort = oPlate.baseport(BPT_Base)
    End If
       
    sError = "Setting system parent to Member Description Custom Assembly"
    Dim oSystemParent As IJSystemChild
    Set oSystemParent = oMD.CAO
     
    Dim oFacePortForCut As IJPort
   
    Set oFacePortForCut = RelatedPortBeforeCut(oFacePort, True)
    ' Create Corner Feature
    sError = "Creating Corner Feature"
    Dim oSDO_CornerFeature As StructDetailObjects.CornerFeature
    Set oSDO_CornerFeature = New StructDetailObjects.CornerFeature
    oSDO_CornerFeature.Create pResourceManager, _
                             oFacePortForCut, _
                             oEdgePort1ForCut, _
                             oEdgePort2ForCut, _
                             "RootCorner", _
                              oSystemParent
                               
     sError = "Returning CornerFeature just created"
     Set oObject = oSDO_CornerFeature.object
       
    Set oSupportV = Nothing
    Set oSupportU = Nothing
    Set oOtherPort1 = Nothing
    Set oOtherPort2 = Nothing
    Set oOtherConnectable1 = Nothing
    Set oOtherConnectable2 = Nothing
    Set ppEnumConnections = Nothing
    Set pObject = Nothing
    Set oProfilehelper = Nothing
    Set oAC = Nothing
    Set oStructDetailHelper = Nothing
    Set oPlane = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
 
 End Sub
    

Public Sub CMBracketCornerFeatureOnS3TopConn(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)

  '****************************************************************************
   On Error GoTo ErrorHandler
    sMethod = "CMBracketCornerFeatureOnS3TopConn"
    
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_CornerFeature) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim oBracketPort As IJPort
    Dim oOtherPort As IJPort
    Dim oBracket As Object
    Dim oSupportRoot As Object
    
    GetPortsForBracketInAC pMD, oBracketPort, oOtherPort, oBracket, oSupportRoot
    
    If oBracketPort Is Nothing Then
        Exit Sub
    End If
    
    Dim oS1 As Object
    Dim oS2 As Object
    Dim oS3 As Object
    Dim oS4 As Object
    Dim oS5 As Object
    Dim oS1Port As IJPort
    Dim oS2Port As IJPort
    Dim oS3Port As IJPort
    Dim oS4Port As IJPort
    Dim oS5Port As IJPort
    
    GetBracketSupports oBracket, oS1, oS2, oS3, oS4, oS5, oS1Port, oS2Port, oS3Port, oS4Port, oS5Port
    
    Dim oBracketRoot As IJPlate
    Dim oStructDetailHelper As StructDetailHelper
    Set oStructDetailHelper = New StructDetailHelper
    
    oStructDetailHelper.IsPartDerivedFromSystem oBracket, oBracketRoot, True

    If oBracketRoot Is Nothing Then
          Set oBracketRoot = oBracket
    End If
     
    If oS3Port Is Nothing Then
        Exit Sub
    End If
    
    Dim oBracketPortBeforeCut As IJPort
    Dim oS3PortBeforeCut As IJPort
    Dim oBracketStructPort As IJStructPort
    Dim oS3StructPort As IJStructPort
    Set oBracketPortBeforeCut = RelatedPortBeforeCut(oBracketPort)
    
    If Not oS3Port Is Nothing Then
      Set oS3PortBeforeCut = RelatedPortBeforeCut(oS3Port)
    Else
      Set oS3PortBeforeCut = Nothing
    End If
    
    Set oBracketStructPort = oBracketPortBeforeCut
    Set oS3StructPort = oS3PortBeforeCut
    
    Dim existingType As ShpStrBracketReinforcementType
    Dim bReinforcementExist As Boolean

    Dim oModelTopBody As IJModelBody
    Dim oTopo As New GSCADStructGeomUtilities.TopologyLocate
    
    'Check whether the AC port is the port connected to S3
    If oBracketStructPort.OperatorID = oS3StructPort.OperatorID And _
       oBracketStructPort.OperationID = oS3StructPort.OperationID Then
         bReinforcementExist = IsReinforcedBracket(oBracket, existingType)
         If Not bReinforcementExist _
           Or existingType = BRACKETREINFORCEMENTTYPE_BucklingStiffener Then
             Dim oBracketFreeEdgeLateralFacePort As IJPort
             Set oBracketFreeEdgeLateralFacePort = GetBracketFreeEdgeLateralFacePort(oBracket)
             
             Dim oIntersection_Between_FreeEdgeAndS3 As IJDPosition
             Dim GeomUtils As New TopologyLocate
             Dim oTempModelBody As IJDModelBody
             Dim oBrackerFreeEdgeBeforeCut As IJPort
             
             If Not oBracketFreeEdgeLateralFacePort Is Nothing Then
                Set oBrackerFreeEdgeBeforeCut = RelatedPortBeforeCut(oBracketFreeEdgeLateralFacePort)
                Set oIntersection_Between_FreeEdgeAndS3 = GeomUtils.FindIntersectionPoint(oBrackerFreeEdgeBeforeCut.Geometry, oBracketPortBeforeCut.Geometry)
             End If
             
             If Not oIntersection_Between_FreeEdgeAndS3 Is Nothing Then
             
                If TypeOf oSupportRoot Is IJStiffener Then
                  Dim sSectionType As String
                  Dim oProfile As New StructDetailObjects.ProfilePart
                  Set oProfile.object = oOtherPort.Connectable
                  sSectionType = oProfile.sectionType
                  'if S3 is FB or Flange Pointing away
                  If sSectionType = "FB" Or sSectionType = "B_TSN" Or _
                     sSectionType = "EA" Or sSectionType = "UA" Or _
                     sSectionType = "BUTL3" Or sSectionType = "CSType" Or _
                     sSectionType = "C_SS" Then
                     
                     If IsFlangeTowardsBracket(oOtherPort.Connectable, oBracket) And Not sSectionType = "FB" Then
                         Exit Sub
                     End If
                  Else
                      Exit Sub
                  End If
                End If
                
                Dim oRefPlane As IJPlane
                Dim uVector As IJDVector
                Dim vVector As IJDVector
                Dim lNumOfSupports As Long
                    
               'Get Bracket Plane and U/V vectors of Bracket
                GetBracketPlaneAndUVVectors oBracket, oRefPlane, uVector, vVector, lNumOfSupports
                 
                Dim uCurve As IJCurve
                Dim ppLocation As IJDPosition
                Dim oBoreLIneDir As IJDVector
                Dim oEdgeUVe As IJDVector
                Dim oEdgeVVe As IJDVector
                     
                Dim oPartInfo As New PartInfo
                Dim oPortNormal As IJDVector
                Dim oS3Vec As IJDVector
                Dim oBrackPlaneNormal As IJDVector
                Set oBrackPlaneNormal = New dVector
                     
                Dim x As Double
                Dim y As Double
                Dim z As Double
                     
                Set oPortNormal = oPartInfo.GetPortNormal(oOtherPort, False)
                oRefPlane.GetNormal x, y, z
                    
                oBrackPlaneNormal.Set x, y, z
                Set oS3Vec = oPortNormal.Cross(oBrackPlaneNormal)
                     
                If oS3Vec.Dot(vVector) < 0 Then
                    oS3Vec.Set -oS3Vec.x, -oS3Vec.y, -oS3Vec.z
                End If
                     
                oS3Vec.Set -oS3Vec.x, -oS3Vec.y, -oS3Vec.z
                
                If Not TypeOf oSupportRoot Is IJStiffener Then
                    On Error Resume Next
                    GeomUtils.FindNearestLogicalEdgeGivenBoreLine oOtherPort.Geometry, oIntersection_Between_FreeEdgeAndS3, _
                                                              oS3Vec, ppLocation, oEdgeUVe, oEdgeVVe, uCurve
                                                              
                    Dim dDistFromFreeEdgeToCurve As Double
                     
                    dDistFromFreeEdgeToCurve = ppLocation.DistPt(oIntersection_Between_FreeEdgeAndS3)
                     
                    If dDistFromFreeEdgeToCurve < 0.01 Then
                        bIsNeeded = True
                        Exit Sub
                    End If
                End If
                                                              
                
                Dim oPtToWebVector As IJDVector
                Dim oPointOnWeb As IJDPosition
                Dim oTopPort As IJPort
                Dim oWebLeftPort As IJPort
                Dim oWebRightPort As IJPort
                Dim oModelWebBody As IJModelBody
                Dim dDistFromInToTop As Double
                Dim dDistFromInToWebLeft As Double
                Dim dDistFromInToWebRight As Double
                
                If TypeOf oSupportRoot Is IJStiffener Then
                
                    If uCurve Is Nothing Then
                    
                        Set oTopPort = oProfile.SubPort(JXSEC_TOP)
                        oTopo.GetIJModelBodyFromObject oTopPort, oModelTopBody
                        oTopo.GetProjectedPointOnModelBody oModelTopBody, oIntersection_Between_FreeEdgeAndS3, ppLocation, oEdgeUVe
                        If Not ppLocation Is Nothing Then
                            dDistFromInToTop = ppLocation.DistPt(oIntersection_Between_FreeEdgeAndS3)
                        End If
                        
                        Set oWebLeftPort = oProfile.SubPort(JXSEC_WEB_LEFT)
                        oTopo.GetIJModelBodyFromObject oWebLeftPort, oModelWebBody
                        oTopo.GetProjectedPointOnModelBody oModelWebBody, oIntersection_Between_FreeEdgeAndS3, oPointOnWeb, oPtToWebVector
                        If Not oPointOnWeb Is Nothing Then
                            dDistFromInToWebLeft = oIntersection_Between_FreeEdgeAndS3.DistPt(oPointOnWeb)
                        End If
                                            
                        Set oModelWebBody = Nothing
                        Set oPointOnWeb = Nothing
                        Set oPtToWebVector = Nothing
                        
                        Set oWebRightPort = oProfile.SubPort(JXSEC_WEB_RIGHT)
                        oTopo.GetIJModelBodyFromObject oWebRightPort, oModelWebBody
                        oTopo.GetProjectedPointOnModelBody oModelWebBody, oIntersection_Between_FreeEdgeAndS3, oPointOnWeb, oPtToWebVector
                        If Not oPointOnWeb Is Nothing Then
                            dDistFromInToWebRight = oIntersection_Between_FreeEdgeAndS3.DistPt(oPointOnWeb)
                        End If
                        
                    End If
                                    
                    If dDistFromInToWebLeft > 0.01 Or dDistFromInToWebRight > 0.01 Then 'condiiton where the profile is bounded to flange of the profile
                        bIsNeeded = False
                        GoTo CleanUp
                    End If
                                   
                    If dDistFromInToTop < 0.01 And dDistFromInToWebLeft < 0.01 And dDistFromInToWebRight < 0.01 Then
                        bIsNeeded = True
                        GoTo CleanUp
                    End If
                End If
             End If
       End If
    End If
    
CleanUp:
                Set oPtToWebVector = Nothing
                Set oPointOnWeb = Nothing
                Set oTopPort = Nothing
                Set oWebLeftPort = Nothing
                Set oWebRightPort = Nothing
                Set oModelWebBody = Nothing
                
                
    
    Exit Sub
     '****************************************************************************
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
 
End Sub
 Public Sub CMBracketCornerFeatureOnS3Top(ByVal oMD As IJDMemberDescription, _
                                          ByVal pResourceManager As IUnknown, _
                                          ByRef oObject As Object)
     
      On Error GoTo ErrorHandler
    sMethod = "CMBracketCornerFeature"
 
    
    Dim oBracketPort As IJPort
    Dim oOtherPort As IJPort
    Dim oBracketPart As Object
    Dim oSupportRoot As Object
    
    GetPortsForBracketInAC oMD, oBracketPort, oOtherPort, oBracketPart, oSupportRoot
    
    Dim oS1 As Object
    Dim oS2 As Object
    Dim oS3 As Object
    Dim oS4 As Object
    Dim oS5 As Object
    Dim oS1Port As IJPort
    Dim oS2Port As IJPort
    Dim oS3Port As IJPort
    Dim oS4Port As IJPort
    Dim oS5Port As IJPort
    
    GetBracketSupports oBracketPart, oS1, oS2, oS3, oS4, oS5, oS1Port, oS2Port, oS3Port, oS4Port, oS5Port

    Dim oPlate As New StructDetailObjects.PlatePart
    Dim oFacePort As IJPort
    Dim oEdgePort1 As IJPort
    Dim oEdgePort2 As IJPort
    Set oPlate.object = oBracketPart
    Set oFacePort = oPlate.baseport(BPT_Base)
        
   If IsSupportOrPlateForStiffenerSupport(oSupportRoot, oS3) Then
        Set oEdgePort1 = GetBracketFreeEdgeLateralFacePort(oBracketPart)
        Set oEdgePort2 = oS3Port
    Else
        ' Error condition
        Exit Sub
    End If
    
    sError = "Setting system parent to Member Description Custom Assembly"
    Dim oSystemParent As IJSystemChild
    Set oSystemParent = oMD.CAO
     
    Dim oFacePortForCut As IJPort
    Dim oEdgePort1ForCut As IJPort
    Dim oEdgePort2ForCut As IJPort
    
    Set oFacePortForCut = RelatedPortBeforeCut(oFacePort, True)
    Set oEdgePort1ForCut = RelatedPortBeforeCut(oEdgePort1)
    Set oEdgePort2ForCut = RelatedPortBeforeCut(oEdgePort2)
    

    ' Create Corner Feature
    sError = "Creating Corner Feature"
    Dim oSDO_CornerFeature As StructDetailObjects.CornerFeature
    Set oSDO_CornerFeature = New StructDetailObjects.CornerFeature
    oSDO_CornerFeature.Create pResourceManager, _
                             oFacePortForCut, _
                             oEdgePort1ForCut, _
                             oEdgePort2ForCut, _
                             "RootCorner", _
                              oSystemParent
                               
     sError = "Returning CornerFeature just created"
     Set oObject = oSDO_CornerFeature.object
       
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod, sError).Number
 
End Sub

Public Sub CMCreateWebCut(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMethod = "CMCreateWebCut"
    
    bIsNeeded = True
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_WebCut) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Sub

Public Sub CMCreateSlot(ByRef pMD As IJDMemberDescription, ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMethod = "CMCreateSlot"
    bIsNeeded = True
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_Slot) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMethod).Number
End Sub
