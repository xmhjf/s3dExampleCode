VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollarDefCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULE As String = m_sProjectPath + "CollarClipParm.cls"

Private sERROR As String
Private sMETHOD As String

Public Function CAConstruct_PhysConn(ByVal pMemberDescription As IJDMemberDescription, _
                                     ByVal pResourceManager As IUnknown, _
                                     strPhysicalConnType As String, _
                                     ByVal oPort1 As IJPort, _
                                     ByVal oPort2 As IJPort) As Object

    sMETHOD = "CAConstruct_PhysConn"
            
    ' Get the collar, since it is the parent of the PC
    Dim pSystemParent As IJSystemChild    'Collar Custom Assembly
    sERROR = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    
    ' Create physical connection
    Dim oPhysicalConnection As New PhysicalConn
    sERROR = "Creating Physical Connection"
    oPhysicalConnection.Create pResourceManager, _
                               oPort1, _
                               oPort2, _
                               strPhysicalConnType, _
                               pSystemParent, _
                               ConnectionStandard

                               
    sERROR = "Setting Physical Connection to private variable"
    Set CAConstruct_PhysConn = oPhysicalConnection.object
    
    Set pSystemParent = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
        
End Function
'***********************************************************************
' METHOD:  CAConstruct_PhysConnNames
'
' DESCRIPTION:  Constructs the Physical Connection using the sub port name as input
'               Currently this method assumes that:
'                     Penetrating object is a profile
'                     Collar Object is a plate
'***********************************************************************

Public Function CAConstruct_PhysConnNames(ByVal pMemberDescription As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown, _
                                    strStartClass As String, _
                                    ByVal ePenetratingSubPort As JXSEC_CODE, _
                                    ByVal eCollarSubPort As JXSEC_CODE) As Object

    sMETHOD = "Collar_FinalConstructNames"
    
    ' Get Class
    Dim pCollar As New StructDetailObjects.Collar ' create new slot wrapper class
    sERROR = "Setting Collar Inputs"
    Set pCollar.object = pMemberDescription.CAO
    
    ' Get Penetrating Part's Port
    Dim oPenetratingPart As Object ' This is the penetrating part
    sERROR = "Getting Penetrating object"
    Set oPenetratingPart = pCollar.Penetrating ' Get penetrating part from wrapper
    
    ' If it is a penetrating profile get the profile sub port.
    ' If it is a penetrating plate, get the mapped port.
    Dim oPenetratingPort As IJPort
    If TypeOf oPenetratingPart Is IJProfile Then
    
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        sERROR = "Setting Profile Part wrapper object to penetrating object"
        Set oProfilePart.object = oPenetratingPart
    
        sERROR = "Getting Penetrating object's port"
        Set oPenetratingPort = oProfilePart.SubPort(ePenetratingSubPort)
    ElseIf TypeOf oPenetratingPart Is IJPlate Then
        Dim oBasePort As Object
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oMappedPorts As JCmnShp_CollectionAlias
        Set oMappedPorts = New Collection

        oSlotMappingRule.GetEmulatedPorts oPenetratingPart, pCollar.Penetrated, oBasePort, oMappedPorts
        
        Dim pSDOHelper As New StructDetailObjects.Helper
        Set oPenetratingPort = pSDOHelper.GetEquivalentLastPort(oMappedPorts.Item(CStr(ePenetratingSubPort)))
    Else
        Exit Function
    End If
   
    ' Get collarss Port
    Dim oPenetratedPort As IJPort
    Set oPenetratedPort = pCollar.SubPort(eCollarSubPort)
'        Set oPenetratedPort = pCollar.BasePort(BPT_Offset) ' CHANGED BY RAS *** DELETE THIS ******************
    

        
    ' Get the Collar, since it is the parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the slot's ca
    sERROR = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
 
    ' Create physical connection
    Dim oPhysicalConnection As New PhysicalConn
    sERROR = "Creating Physical Connection"
    oPhysicalConnection.Create pResourceManager, _
                               oPenetratedPort, _
                               oPenetratingPort, _
                               strStartClass, _
                               pSystemParent, ConnectionStandard
                               
    sERROR = "Setting Physical Connection to private variable"
    Set CAConstruct_PhysConnNames = oPhysicalConnection.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
        
End Function
''***********************************************************************
'' METHOD:  CAConstruct_BasePhysConn
''
'' DESCRIPTION:  Constructs the Physical Connection
''               between the collar and the base plate
''
''               This weld is the full lenght of the collar  / Base plate ports.
''               In the future we need to stop this weld in way of the lap weld.
''
''             !!!!!This custom method requires the wrapper class Slot.BasePlatePort
''                  that is not available yet
''
''***********************************************************************
'
'Public Function CAConstruct_BasePhysConn(ByVal pMemberDescription As IJDMemberDescription, _
'                                    ByVal pResourceManager As IUnknown, _
'                                    strStartClass As String, _
'                                    ByVal eCollarSubPort As JXSEC_CODE) As Object
'
'    sMETHOD = "BasePlate Weld Create"
'
'    ' Get Class
'    Dim pCollar As New StructDetailObjects.Collar ' create new slot wrapper class
'    sError = "Setting Collar Inputs"
'    Set pCollar.Object = pMemberDescription.CAO
'
'    ' Get Collar base plate port
'    Dim oBasePlatePort As IJPort
'    sError = "Getting collars's base plate port port"
'    Set oBasePlatePort = pCollar.BasePlatePort  'need new wrapper class here
'
'    ' Get collarss Port
'    Dim oPenetratedPort As IJPort
'    Set oPenetratedPort = pCollar.SubPort(eCollarSubPort)
'
'
'    ' Get the Collar, since it is the parent of the PC
'    Dim pSystemParent As IJSystemChild ' this is the slot's ca
'    sError = "Setting system parent to Member Description Custom Assembly"
'    Set pSystemParent = pMemberDescription.CAO
'
'
'    ' Create physical connection
'    Dim oPhysicalConnection As New PhysicalConn
'    sError = "Creating Physical Connection"
'    oPhysicalConnection.Create pResourceManager, _
'                               oBasePlatePort, _
'                               oPenetratingPort, _
'                               strStartClass, _
'                               pSystemParent, ConnectionStandard
'
'    sError = "Setting Physical Connection to private variable"
'    Set CAConstruct_BasePhysConn = oPhysicalConnection.Object
'
'    Exit Function
'
'ErrorHandler:
'    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
'
'End Function

Public Function CreateLapPCforCollar(ByVal pMemberDescription As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown) As Object
   

    sMETHOD = "CAConstruct_PhysConn"
    
    ' Get Class
    Dim pCollar As New StructDetailObjects.Collar ' create new collar wrapper class
    sERROR = "Setting Collar Inputs"
    Set pCollar.object = pMemberDescription.CAO ' Intialize the wrapper with slot
    
    ' Get Penetrated Part Port
    sERROR = "Getting Penetrated object"
    Dim oSlotClass As New StructDetailObjects.Slot
    sERROR = "Setting Collar Part to penetrating object"
    Set oSlotClass.object = pCollar.Slot
       
    'Get the Penetrated Object
    On Error Resume Next
    Dim oPenetratedPart As IJPlate 'check if the penetrating part is a plate or a profile
    Set oPenetratedPart = oSlotClass.Penetrated ' Get penetrating part from wrapper
    'On Error GoTo ErrorHandler
    Dim oPenetratedPort As IJPort
    If Not (oPenetratedPart Is Nothing) Then
        Dim oPlatePart As New StructDetailObjects.PlatePart
        sERROR = "Setting Plate Part to penetrated object"
        Set oPlatePart.object = oPenetratedPart
        sERROR = "Getting Penetrated object's port"
        Set oPenetratedPort = oPlatePart.BasePort(BPT_Base)
    Else
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        sERROR = "Setting Profile Part to penetrated object"
        Set oProfilePart.object = oSlotClass.Penetrated
        sERROR = "Getting Penetrated object's port"
        Set oPenetratedPort = oProfilePart.BasePort(BPT_Lateral)
    End If

    'Get Collar Port
    'Since the collar is expanded in the direction same direction
    ' as Base to Offset Ports from the plate,
    ' the port used for the lap weld on the collar is always the offset
    Dim oCollarPort As IJPort
    Set oCollarPort = pCollar.BasePort(BPT_Offset)
    
                               
    sERROR = "Setting Physical Connection to private variable"
    Set CreateLapPCforCollar = CAConstruct_PhysConn(pMemberDescription, _
                                                    pResourceManager, "LapWeld", _
                                                    oPenetratedPort, oCollarPort)

    Set oPenetratedPort = Nothing
    Set pCollar = Nothing
    Set oCollarPort = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
        
End Function
Public Function CreateLapPCforCollarOnOtherSide(ByVal pMemberDescription As IJDMemberDescription, _
                                                ByVal pResourceManager As IUnknown) As Object

    sMETHOD = "CAConstruct_PhysConn"
    

    ' Get Class
    Dim pCollar As New StructDetailObjects.Collar ' create new collar wrapper class
    sERROR = "Setting Collar Inputs"
    Set pCollar.object = pMemberDescription.CAO ' Intialize the wrapper with slot
    
    ' Get Penetrated Part Port
    sERROR = "Getting Penetrated object"
    Dim oSlotClass As New StructDetailObjects.Slot
    sERROR = "Setting Collar Part to penetrating object"
    Set oSlotClass.object = pCollar.Slot
       
    'Get the Penetrated Object
    On Error Resume Next
    Dim oPenetratedPart As IJPlate 'check if the penetrating part is a plate or a profile
    Set oPenetratedPart = oSlotClass.Penetrated ' Get penetrating part from wrapper
    'On Error GoTo ErrorHandler
    Dim oPenetratedPort As IJPort
    If Not (oPenetratedPart Is Nothing) Then
        Dim oPlatePart As New StructDetailObjects.PlatePart
        sERROR = "Setting Plate Part to penetrated object"
        Set oPlatePart.object = oPenetratedPart
        sERROR = "Getting Penetrated object's port"
        Set oPenetratedPort = oPlatePart.BasePort(BPT_Offset)
    Else
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        sERROR = "Setting Profile Part to penetrated object"
        Set oProfilePart.object = oSlotClass.Penetrated
        sERROR = "Getting Penetrated object's port"
        Set oPenetratedPort = oProfilePart.BasePort(BPT_Lateral)
    End If

    'Get Collar Port
    'Since the collar is expanded in the direction same direction
    ' as Base to Offset Ports from the plate,
    ' the port used for the lap weld on the collar is always the base
    Dim oCollarPort As IJPort
    Set oCollarPort = pCollar.BasePort(BPT_Base)
    
                               
    sERROR = "Setting Physical Connection to private variable"
    Set CreateLapPCforCollarOnOtherSide = CAConstruct_PhysConn(pMemberDescription, _
                                                    pResourceManager, "LapWeld", _
                                                    oPenetratedPort, oCollarPort)

    Set oPenetratedPort = Nothing
    Set pCollar = Nothing
    Set oCollarPort = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
        
End Function

Public Sub Collar_FinalConstruct(ByVal pAggregatorDescription As IJDAggregatorDescription)

  On Error GoTo ErrorHandler
  sMETHOD = "Collar_FinalConstruct"
  
  ' Get collar plate
  Dim pCollarPlate As IJPlate
  Set pCollarPlate = pAggregatorDescription.CAO
  
  ' Intialize the collar utility and get the penetrated part
    Dim pCollar As New StructDetailObjects.Collar ' create new collar wrapper class
    sERROR = "Setting Collar Inputs"
    Set pCollar.object = pAggregatorDescription.CAO ' Intialize the wrapper with slot
    Dim oSlotClass As New StructDetailObjects.Slot
    Dim oPenetratedPart As Object
    Set oSlotClass.object = pCollar.Slot
    Set oPenetratedPart = oSlotClass.Penetrated

  
  ' Create the SDCreateModifyUtilities and place collar
  Dim oCollarUtils As New GSCADSDCreateModifyUtilities.SDCollarUtils
  oCollarUtils.PlaceCollar pCollarPlate, oPenetratedPart

  Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number

End Sub
'***********************************************************************
' METHOD:  CMCopyParentMatAndGrade
'
' DESCRIPTION:  Copies the parent base plate thinkness (dimensions).
'
'***********************************************************************
Public Sub CMCopyParentMatAndGrade(pPD As IJDPropertyDescription, pCollar As Object)

 On Error GoTo ErrorHandler
    sMETHOD = "CMCopyParentMatAndGrade"
     
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
  Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub

'***********************************************************************
' METHOD:  CMCopyParentThickness
'
' DESCRIPTION:  Sets thickness to 80% of parent thickness for a plate.
'               Sets thickness to web thickness if penetrated is profile
'
'***********************************************************************
Public Sub CMCopyParentThickness(pPD As IJDPropertyDescription, pCollar As Object)

 On Error GoTo ErrorHandler
       
    'Get Collar as collar
    Dim oCollar As StructDetailObjects.Collar
    Set oCollar = New StructDetailObjects.Collar
    Set oCollar.object = pCollar
        
    'Get Slot
    Dim oSlot As New Slot
    Set oSlot.object = oCollar.Slot
       
    'Get Collar as plate part
    Dim oCollarPlate As StructDetailObjects.PlatePart
    Set oCollarPlate = New StructDetailObjects.PlatePart
    Set oCollarPlate.object = pCollar
    
    'Check to see if penetrated is a plate or profile
    Dim pHelper As StructDetailObjects.Helper
    Set pHelper = New StructDetailObjects.Helper
   
    Select Case pHelper.ObjectType(oSlot.Penetrated)
        Case SDOBJECT_PLATE
            Dim oPlatePart As StructDetailObjects.PlatePart
            Set oPlatePart = New StructDetailObjects.PlatePart
            Set oPlatePart.object = oSlot.Penetrated
            oCollarPlate.PlateThickness = oPlatePart.PlateThickness * 0.8
        
        Case SDOBJECT_STIFFENER
            Dim oProfilePart As StructDetailObjects.ProfilePart
            Set oProfilePart = New StructDetailObjects.ProfilePart
            Set oProfilePart.object = oSlot.Penetrated
            oCollarPlate.PlateThickness = oProfilePart.WebThickness
            
        Case Else
            'Penetrated object is not a plate or profile.

    End Select

    Set oPlatePart = Nothing
    Set oProfilePart = Nothing

    
    Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMCopyParentThickness").Number
End Sub
Public Function CMSetSideOfPlate(pPD As IJDPropertyDescription, pCollar As Object)
    On Error GoTo ErrorHandler

    sMETHOD = "CMSetSideOfPlate"
    sERROR = "Setting Side of Plate"
      
    ' Get Collar
    Dim oCollar As New Collar
    Set oCollar.object = pCollar
    
    ' Get Penetrated Plate
    Dim oSlot As New Slot
    Set oSlot.object = oCollar.Slot
    Dim lCandidateSideOfPart As Long
    Dim pCollarPart As IJCollarPart
    
    If TypeOf oSlot.Penetrated Is IJPlate Then
        ' Retreive the current
        ' Collar Symbol Parameter input value for "SideOfPart"
        ' The CollarPart get SideOfPart property method gets the "SideOfPart"
        ' value from the Collar's Symbol Parmater Input values
        Set pCollarPart = pCollar
        lCandidateSideOfPart = pCollarPart.SideOfPlate

        ' The Collar Part put SideOfPart property method sets the "SideOfPart"
        ' value as an Attribute value (strange and weird but TRUE)
        ' Set the Collar's Part "Side Of Plate" attribute value
        If lCandidateSideOfPart = 1 Then  ' Anti Molded
            pCollarPart.SideOfPlate = PLATE_ANTI_MOLDED
        Else
           pCollarPart.SideOfPlate = PLATE_MOLDED
        End If
    Else
        Set pCollarPart = pCollar
        pCollarPart.SideOfPlate = PLATE_MOLDED
    End If
      
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Function

'***********************************************************************
' METHOD:  CMCreateClipNormalSide
'
' DESCRIPTION:  Determines if a physical connection should be placed on
'               the plate molded side
'
'***********************************************************************


Public Sub CMCreateClipNormalSide(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    sMETHOD = "CMCreateClipNormalSide"
    sERROR = "Setting Conditional value"
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Dim sItemName As String
    
    Set pCollar.object = pMD.CAO
    sItemName = pCollar.ItemName
    
    If InStr(sItemName, "_A2") > 0 Or _
       InStr(sItemName, "_B2") > 0 Then
       Exit Sub
    End If

    Dim oCollar As IJCollarPart
    Set oCollar = pMD.CAO
    
    Dim Side As eSideOfPlate
    Side = oCollar.SideOfPlate
        
    If Side = PLATE_MOLDED Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub

'***********************************************************************
' METHOD:  CMCreateClipOppositeSide
'
' DESCRIPTION:  Determines if a physical connection should be placed on
'               the plate anti molded side
'
'***********************************************************************
Public Sub CMCreateClipOppositeSide(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMCreateClipOppositeSide"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Dim sItemName As String
    
    Set pCollar.object = pMD.CAO
    sItemName = pCollar.ItemName
    
    If InStr(sItemName, "_A2") > 0 Or _
       InStr(sItemName, "_B2") > 0 Then
       Exit Sub
    End If
    
    Dim oCollar As IJCollarPart
    Set oCollar = pMD.CAO
    
    Dim Side As eSideOfPlate
    Side = oCollar.SideOfPlate
    
    If Side = PLATE_ANTI_MOLDED Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
'Custom methods below create welds and a conditional for the edgename combination needed.
'Each construction custom method is followed by its conditional
'The conditioals all are based on the profile cross section type.
'Web_Left, Web_Right & Top do not have conditional because every profile cross section has these edges.
'
'***********************************************************************
' METHOD:  CMWebLeftPC
'
' DESCRIPTION:  Create physical connection for WebLeft  # 257
'
'
'***********************************************************************
Public Sub CMWebLeftPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    
    Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                                                    JXSEC_WEB_LEFT, JXSEC_WEB_LEFT)
End Sub
'***********************************************************************
' METHOD:  CMWebRightPC
'
' DESCRIPTION:  Create physical connection for WebRight  # 258
'
'
'***********************************************************************
Public Sub CMWebRightPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                                                   JXSEC_WEB_RIGHT, JXSEC_WEB_RIGHT)
End Sub
'***********************************************************************
' METHOD:  CMTopPC
'
' DESCRIPTION:  Create physical connection for Top  # 514
'
'
'***********************************************************************
Public Sub CMTopPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                                                   JXSEC_TOP, JXSEC_TOP_FLANGE_TOP)
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeLeftBottomPC
'
' DESCRIPTION:  Create physical connection for TopFlangeLeftBottom  # 770
'
'
'***********************************************************************
Public Sub CMTopFlangeLeftBottomPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                        JXSEC_TOP_FLANGE_LEFT_BOTTOM, JXSEC_TOP_FLANGE_LEFT_BOTTOM)
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeLeftBottomCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMTopFlangeLeftBottomCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMTopFlangeLeftBottomCon"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim sXSectionType As String
    If TypeOf pCollar.Penetrating Is IJProfile Then
        'Get Penetrating object
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pCollar.Penetrating
        sXSectionType = oProfilePart.sectionType
        Set oProfilePart = Nothing
    ElseIf TypeOf pCollar.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias pCollar.Penetrating, pCollar.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
        Set oSlotMappingRule = Nothing
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If

    Select Case sXSectionType
        Case "BUT", "BUTL2", "I", "ISType", "TSType", "T_XType"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeRightBottomPC
'
' DESCRIPTION:  Create physical connection for TopFlangeRightBottom  # 772
'
'
'***********************************************************************
Public Sub CMTopFlangeRightBottomPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_TOP_FLANGE_RIGHT_BOTTOM, JXSEC_TOP_FLANGE_RIGHT_BOTTOM)
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeRightBottomCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMTopFlangeRightBottomCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMTopFlangeRightBottomCon"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim sXSectionType As String
    If TypeOf pCollar.Penetrating Is IJProfile Then
        'Get Penetrating object
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pCollar.Penetrating
        sXSectionType = oProfilePart.sectionType
        Set oProfilePart = Nothing
    ElseIf TypeOf pCollar.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias pCollar.Penetrating, pCollar.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
        Set oSlotMappingRule = Nothing
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If
      
    Select Case sXSectionType
        Case "B", "EA", "UA", "BUT", "BUTL2", "BUTL3", "CSType", "C_SS", "I", "ISType", "TSType", "T_XType"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeLeftPC
'
' DESCRIPTION:  Create physical connection for TopFlangeLeft  # 1026
'
'
'***********************************************************************
Public Sub CMTopFlangeLeftPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                                    JXSEC_TOP_FLANGE_LEFT, JXSEC_TOP_FLANGE_LEFT)
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeLeftCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMTopFlangeLeftCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMTopFlangeLeftCon"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim sXSectionType As String
    If TypeOf pCollar.Penetrating Is IJProfile Then
        'Get Penetrating object
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pCollar.Penetrating
        sXSectionType = oProfilePart.sectionType
        Set oProfilePart = Nothing
    ElseIf TypeOf pCollar.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias pCollar.Penetrating, pCollar.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
        Set oSlotMappingRule = Nothing
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If
 
    Select Case sXSectionType
        Case "BUT", "BUTL2", "I", "ISType", "TSType", "T_XType"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
    End Sub
    '***********************************************************************
' METHOD:  CMTopFlangeRightPC
'
' DESCRIPTION:  Create physical connection for TopFlangeRight  # 1028
'
'
'***********************************************************************
Public Sub CMTopFlangeRightPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                            JXSEC_TOP_FLANGE_RIGHT, JXSEC_TOP_FLANGE_RIGHT)
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeRightCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMTopFlangeRightCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMTopFlangeRightCon"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim sXSectionType As String
    If TypeOf pCollar.Penetrating Is IJProfile Then
        'Get Penetrating object
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pCollar.Penetrating
        sXSectionType = oProfilePart.sectionType
        Set oProfilePart = Nothing
    ElseIf TypeOf pCollar.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias pCollar.Penetrating, pCollar.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
        Set oSlotMappingRule = Nothing
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If

    Select Case sXSectionType
        Case "EA", "UA", "BUT", "BUTL2", "BUTL3", "CSType", "C_SS", "I", "ISType", "TSType", "T_XType"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
    End Sub
    '***********************************************************************
' METHOD:  CMWebRightTopCornerPC
'
' DESCRIPTION:  Create physical connection for WebRightTopCorner  # 1281
'
'
'***********************************************************************
Public Sub CMWebRightTopCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                            JXSEC_WEB_RIGHT_TOP_CORNER, JXSEC_WEB_RIGHT_TOP_CORNER)
End Sub
'***********************************************************************
'
' METHOD:  CMWebRightTopCornerCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMWebRightTopCornerCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMWebRightTopCornerCon"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO
    Dim bIsPenetratingObjPlate As Boolean
    Dim sXSectionType As String
    If TypeOf pCollar.Penetrating Is IJProfile Then
        'Get Penetrating object
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pCollar.Penetrating
        sXSectionType = oProfilePart.sectionType
        Set oProfilePart = Nothing
        bIsPenetratingObjPlate = False
    ElseIf TypeOf pCollar.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias pCollar.Penetrating, pCollar.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
        Set oSlotMappingRule = Nothing
        bIsPenetratingObjPlate = True
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If

    Select Case sXSectionType
        Case "B", "EA", "UA", "CSType", "C_SS", "I", "ISType", "TSType", "T_XType"
            If Not bIsPenetratingObjPlate Then
                bIsNeeded = True
            Else
                'Not needed if the penetrating part is a plate because there is no radius
                bIsNeeded = False
            End If
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
'***********************************************************************
' METHOD:  CMWebLeftTopCornerPC
'
' DESCRIPTION:  Create physical connection for WebLeftTopCorner  # 1284
'
'
'***********************************************************************
Public Sub CMWebLeftTopCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                            JXSEC_WEB_LEFT_TOP_CORNER, JXSEC_WEB_LEFT_TOP_CORNER)
End Sub
'***********************************************************************
' METHOD:  CMWebLeftTopCornerCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMWebLeftTopCornerCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMWebLeftTopCornerCon"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim sXSectionType As String
    If TypeOf pCollar.Penetrating Is IJProfile Then
        'Get Penetrating object
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pCollar.Penetrating
        sXSectionType = oProfilePart.sectionType
        Set oProfilePart = Nothing
    ElseIf TypeOf pCollar.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias pCollar.Penetrating, pCollar.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
        Set oSlotMappingRule = Nothing
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If
      
    Select Case sXSectionType
        Case "B", "I", "ISType", "TSType", "T_XType"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeRightBottomCornerPC
'
' DESCRIPTION:  Create physical connection for TopFlangeRightBottomCorner  # 1537
'
'
'***********************************************************************
Public Sub CMTopFlangeRightBottomCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
   Set pObject = CAConstruct_PhysConnNames(pMemberDescription, _
                                           pResourceManager, _
                                           "TeeWeld", _
                                           JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER, _
                                           JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER)
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeRightBottomCornerCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMTopFlangeRightBottomCornerCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMTopFlangeRightBottomCornerCon"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO
    Dim bIsPenetratingObjPlate As Boolean
    Dim sXSectionType As String
    If TypeOf pCollar.Penetrating Is IJProfile Then
        'Get Penetrating object
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pCollar.Penetrating
        sXSectionType = oProfilePart.sectionType
        Set oProfilePart = Nothing
        bIsPenetratingObjPlate = False
    ElseIf TypeOf pCollar.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias pCollar.Penetrating, pCollar.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
        Set oSlotMappingRule = Nothing
        bIsPenetratingObjPlate = True
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If
    
    Select Case sXSectionType
        Case "B", "EA", "UA", "CSType", "C_SS", "I", "ISType", "TSType", "T_XType"
            If Not bIsPenetratingObjPlate Then
                bIsNeeded = True
            Else
                'Not needed if the penetrating part is a plate because there is no radius
                bIsNeeded = False
            End If
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeLeftBottomCornerPC
'
' DESCRIPTION:  Create physical connection for TopFlangeRightBottomCorner  # 1540
'
'
'***********************************************************************
Public Sub CMTopFlangeLeftBottomCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
  JXSEC_TOP_FLANGE_LEFT_BOTTOM_CORNER, JXSEC_TOP_FLANGE_LEFT_BOTTOM_CORNER)
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeLeftBottomCornerCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMTopFlangeLeftBottomCornerCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMTopFlangeLeftBottomCornerCon"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO
    
    Dim sXSectionType As String
    If TypeOf pCollar.Penetrating Is IJProfile Then
        'Get Penetrating object
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pCollar.Penetrating
        sXSectionType = oProfilePart.sectionType
        Set oProfilePart = Nothing
    ElseIf TypeOf pCollar.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias pCollar.Penetrating, pCollar.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
        Set oSlotMappingRule = Nothing
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If
    
    Select Case sXSectionType
        Case "I", "ISType", "TSType", "T_XType"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeRightTopCornerPC
'
' DESCRIPTION:  Create physical connection for TopFlangeRightTopCorner  # 1794
'
'
'***********************************************************************
Public Sub CMTopFlangeRightTopCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
            JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER, JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER)
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeRightTopCornerCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMTopFlangeRightTopCornerCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMTopFlangeRightTopCornerCon"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim sXSectionType As String
    If TypeOf pCollar.Penetrating Is IJProfile Then
        'Get Penetrating object
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pCollar.Penetrating
        sXSectionType = oProfilePart.sectionType
        Set oProfilePart = Nothing
    ElseIf TypeOf pCollar.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias pCollar.Penetrating, pCollar.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
        Set oSlotMappingRule = Nothing
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If
    
    Select Case sXSectionType
        Case "B"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeRightTopPC
'
' DESCRIPTION:  Create physical connection for TopFlangeRightTopCorner  # 2052
'
'
'***********************************************************************
Public Sub CMTopFlangeRightTopPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                JXSEC_TOP_FLANGE_RIGHT_TOP, JXSEC_TOP_FLANGE_RIGHT_TOP)
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeRightTopCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMTopFlangeRightTopCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMTopFlangeRightTopCon"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim sXSectionType As String
    If TypeOf pCollar.Penetrating Is IJProfile Then
        'Get Penetrating object
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pCollar.Penetrating
        sXSectionType = oProfilePart.sectionType
        Set oProfilePart = Nothing
    ElseIf TypeOf pCollar.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias pCollar.Penetrating, pCollar.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
        Set oSlotMappingRule = Nothing
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If
      
    Select Case sXSectionType
        Case "BUTL3"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
'***********************************************************************
' METHOD:  CMWebRightTopPC
'
' DESCRIPTION:  Create physical connection for WebRightTop  # 2564
'
'
'***********************************************************************
Public Sub CMWebRightTopPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
            JXSEC_WEB_RIGHT_TOP, JXSEC_WEB_RIGHT_TOP)
End Sub
'***********************************************************************
' METHOD:  CMWebRightTopCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMWebRightTopCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMWebRightTopCon"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim sXSectionType As String
    If TypeOf pCollar.Penetrating Is IJProfile Then
        'Get Penetrating object
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        Set oProfilePart.object = pCollar.Penetrating
        sXSectionType = oProfilePart.sectionType
        Set oProfilePart = Nothing
    ElseIf TypeOf pCollar.Penetrating Is IJPlate Then
        Dim oSlotMappingRule As IJSlotMappingRule
        Set oSlotMappingRule = CreateSlotMappingRuleSymbolInstance
        
        Dim oWeb As Object
        Dim oFlange As Object
        Dim o2ndWeb As Object
        Dim o2ndFlange As Object
        oSlotMappingRule.GetSectionAlias pCollar.Penetrating, pCollar.Penetrated, sXSectionType, oWeb, oFlange, o2ndWeb, o2ndFlange
        Set oSlotMappingRule = Nothing
    Else
        'Unsupported Penetrating Object
        Exit Sub
    End If
      
    Select Case sXSectionType
        Case "BUTL3"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub
'
'The following Custom methods create welds for collars to slot edges.
'
'
'***********************************************************************
' METHOD:  CMLeftPC
'
' DESCRIPTION:  Create physical connection for Left  # 4353
'
'
'***********************************************************************
Public Sub CMLeftPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                JXSEC_LEFT, JXSEC_LEFT)
End Sub

'***********************************************************************
' METHOD:  CMRightPC
'
' DESCRIPTION:  Create physical connection for Right  # 4354
'
'
'***********************************************************************
Public Sub CMRightPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                JXSEC_RIGHT, JXSEC_RIGHT)
End Sub

'***********************************************************************
' METHOD:  CMCommonPC
'
' DESCRIPTION:  Create physical connection for Common  # 4611
'
'
'***********************************************************************
Public Sub CMCommonPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                                                   JXSEC_COMMON, JXSEC_COMMON)
End Sub
'***********************************************************************
' METHOD:  CMTopLeftCornerPC
'
' DESCRIPTION:  Create physical connection for TopLeftCorner  # 6914
'
'
'***********************************************************************
Public Sub CMTopLeftCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                                        JXSEC_TOP_LEFT_CORNER, JXSEC_TOP_LEFT_CORNER)
End Sub
'***********************************************************************
' METHOD:  CMTopRightCornerPC
'
' DESCRIPTION:  Create physical connection for TopRightCorner  # 6916
'
'
'***********************************************************************
Public Sub CMTopRightCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                                        JXSEC_TOP_RIGHT_CORNER, JXSEC_TOP_RIGHT_CORNER)
End Sub
''***********************************************************************
'' METHOD:  CMBasePlatePC
''
'' DESCRIPTION:  Create physical connection between the base plate and the collar
''
''
''            Waiting on the Collar.BasePlatePort wrapper class.
''
''***********************************************************************
'Public Sub CMBasePlatePC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
'
'  Set pObject = CAConstruct_BasePhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
'                                        JXSEC_BOTTOM)
'End Sub
''***********************************************************************
'' METHOD:  CMBasePlateRPC
''
'' DESCRIPTION:  Create physical connection between the base plate and the collar
''
''
''            Waiting on the Collar.BasePlatePort wrapper class.
''
''***********************************************************************
'Public Sub CMBasePlatePC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
'
'  Set pObject = CAConstruct_BasePhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
'                                        JXSEC_BOTTOM_RIGHT)
'End Sub
''***********************************************************************
'' METHOD:  CMBasePlateLPC
''
'' DESCRIPTION:  Create physical connection between the base plate and the collar
''
''
''            Waiting on the Collar.BasePlatePort wrapper class.
''
''***********************************************************************
'Public Sub CMBasePlatePC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
'
'  Set pObject = CAConstruct_BasePhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
'                                        JXSEC_BOTTOM_LEFT)
'End Sub


Public Sub CMConstructBaseLeftPC(ByVal oMemberDescription As IJDMemberDescription, _
                                 ByVal oResourceManager As IUnknown, _
                                 ByRef oObject As Object)
                                 
   ' Create Physical Connection between Collar Bottom Left Edge and Base Plate
   ConstructBaseEdgePC JXSEC_BOTTOM_LEFT, _
                       oMemberDescription, _
                       oResourceManager, _
                       oObject
End Sub

Public Sub CMConstructBaseRightPC(ByVal oMD As IJDMemberDescription, _
                                  ByVal oResourceManager As IUnknown, _
                                  ByRef oObject As Object)
   ' Create Physical Connection between Collar Bottom Right Edge and Base Plate
   ConstructBaseEdgePC JXSEC_BOTTOM_RIGHT, _
                       oMD, _
                       oResourceManager, _
                       oObject
End Sub

Public Sub CMConstructNormalSideLapPC(ByVal oMD As IJDMemberDescription, _
                                      ByVal oResourceManager As IUnknown, _
                                      ByRef oObject As Object)
   ' Create normal side Lap PC
   Set oObject = CreateLapPCforCollar(oMD, oResourceManager)
End Sub

Public Sub CMConstructOppositeSideLapPC(ByVal oMD As IJDMemberDescription, _
                                      ByVal oResourceManager As IUnknown, _
                                      ByRef oObject As Object)
   ' Create normal side Lap PC
   Set oObject = CreateLapPCforCollarOnOtherSide(oMD, oResourceManager)
End Sub
Public Sub CMConstructLeftPCBetweenCollarAndPenetrated( _
                              ByVal oMemberDescription As IJDMemberDescription, _
                              ByVal oResourceManager As IUnknown, _
                              ByRef oObject As Object)
   ' Create Left PC between collar and slot
   Set oObject = CreatePCBetweenCollarAndPenetrated( _
                                              oMemberDescription, _
                                              oResourceManager, _
                                              JXSEC_LEFT, _
                                              JXSEC_LEFT)
End Sub
Public Sub CMConstructTopLeftCornerPCBetweenCollarAndPenetrated( _
                              ByVal oMemberDescription As IJDMemberDescription, _
                              ByVal oResourceManager As IUnknown, _
                              ByRef oObject As Object)
   ' Create Top Left Corner PC between collar and slot
   Set oObject = CreatePCBetweenCollarAndPenetrated( _
                                              oMemberDescription, _
                                              oResourceManager, _
                                              JXSEC_TOP_FLANGE_LEFT_TOP_CORNER, _
                                              JXSEC_TOP_LEFT_CORNER)
End Sub
Public Sub CMConstructTopLeftPCBetweenCollarAndPenetrated( _
                              ByVal oMemberDescription As IJDMemberDescription, _
                              ByVal oResourceManager As IUnknown, _
                              ByRef oObject As Object)
   ' Create Top Left PC between collar and slot
   Set oObject = CreatePCBetweenCollarAndPenetrated( _
                                              oMemberDescription, _
                                              oResourceManager, _
                                              JXSEC_TOP, _
                                              JXSEC_TOP_FLANGE_LEFT_TOP)

End Sub
Public Sub CMConstructTopRightPCBetweenCollarAndPenetrated( _
                              ByVal oMemberDescription As IJDMemberDescription, _
                              ByVal oResourceManager As IUnknown, _
                              ByRef oObject As Object)
   ' Create Top Right PC between collar and slot
   Set oObject = CreatePCBetweenCollarAndPenetrated( _
                                              oMemberDescription, _
                                              oResourceManager, _
                                              JXSEC_TOP, _
                                              JXSEC_TOP_FLANGE_RIGHT_TOP)

End Sub
Public Sub CMConstructTopRightCornerPCBetweenCollarAndPenetrated( _
                              ByVal oMemberDescription As IJDMemberDescription, _
                              ByVal oResourceManager As IUnknown, _
                              ByRef oObject As Object)
   ' Create Top Right Corner PC between collar and slot
   Set oObject = CreatePCBetweenCollarAndPenetrated( _
                                              oMemberDescription, _
                                              oResourceManager, _
                                              JXSEC_TOP_FLANGE_RIGHT_TOP_CORNER, _
                                              JXSEC_TOP_RIGHT_CORNER)
End Sub
Public Sub CMConstructRightPCBetweenCollarAndPenetrated( _
                              ByVal oMemberDescription As IJDMemberDescription, _
                              ByVal oResourceManager As IUnknown, _
                              ByRef oObject As Object)
   ' Create Right PC between collar and slot
   Set oObject = CreatePCBetweenCollarAndPenetrated( _
                                              oMemberDescription, _
                                              oResourceManager, _
                                              JXSEC_RIGHT, _
                                              JXSEC_RIGHT)
End Sub
Private Function CreatePCBetweenCollarAndPenetrated( _
                              ByVal oMD As IJDMemberDescription, _
                              ByVal oResMgrUnk As IUnknown, _
                              ePenetratedSubPortXid As JXSEC_CODE, _
                              eCollarSubPortXid As JXSEC_CODE) As Object
   ' Get penetrated port
   Dim oCollarWrapper As New StructDetailObjects.Collar
   Dim oPenetratedPart As Object
   Dim oPenetratedPort As IJPort
   Dim oCollarPort As IJPort
   
   Set oCollarWrapper.object = oMD.CAO
   Set oPenetratedPart = oCollarWrapper.Penetrated
   If TypeOf oPenetratedPart Is IJPlatePart Then
      Dim oPlatePartWrapper As New StructDetailObjects.PlatePart
      
      Set oPlatePartWrapper.object = oPenetratedPart
      Set oPenetratedPort = oPlatePartWrapper.CutoutSubPort(oCollarWrapper.Slot, ePenetratedSubPortXid)
   Else
      Dim oProfolePartWrapper As New StructDetailObjects.ProfilePart
      
      Set oProfolePartWrapper.object = oPenetratedPart
      Set oPenetratedPort = oProfolePartWrapper.CutoutSubPort(oCollarWrapper.Slot, ePenetratedSubPortXid)
   End If
   
   ' Get collar port
   Set oCollarPort = oCollarWrapper.SubPort(eCollarSubPortXid)
   
   ' Construct PC
   Dim oPhysicalConnWrapper As New StructDetailObjects.PhysicalConn
   
   oPhysicalConnWrapper.Create oResMgrUnk, _
                               oCollarPort, _
                               oPenetratedPort, _
                               "ButtWeld", _
                               oMD.CAO, _
                               ConnectionStandard
   Set CreatePCBetweenCollarAndPenetrated = oPhysicalConnWrapper.object
End Function

Private Sub CreateFreeEdgeTreatment( _
                              ByVal oMD As IJDMemberDescription, _
                              ByVal oResMgr As IUnknown, _
                              ByVal ePortXid As JXSEC_CODE, _
                              ByRef oObject As Object, _
                              Optional ByVal oBoundary1 As Object = Nothing, _
                              Optional ByVal oBoundary2 As Object = Nothing)
   Dim oCollarWrapper As New StructDetailObjects.Collar
   Dim oFETWrapper As New StructDetailObjects.FreeEdgeTreatment
   Dim oCollarPortForFET As IJPort
   Dim strBevelName As String
   
   Set oCollarWrapper.object = oMD.CAO
   Set oCollarPortForFET = oCollarWrapper.SubPort(ePortXid)
   strBevelName = "Bevel"
   oFETWrapper.Create oResMgr, _
                      oCollarPortForFET, _
                      oBoundary1, _
                      oBoundary2, _
                      strBevelName, _
                      oMD.CAO
   Set oObject = oFETWrapper.object
End Sub

Public Sub CMConstructBottomFlangeRightBottomFET( _
                              ByVal oMD As IJDMemberDescription, _
                              ByVal oResMgr As IUnknown, _
                              ByRef oObject As Object)
   CreateFreeEdgeTreatment oMD, _
                           oResMgr, _
                           JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM, _
                           oObject
End Sub

Public Sub CMConstructBottomFlangeRightBottomCornerFET( _
                              ByVal oMD As IJDMemberDescription, _
                              ByVal oResMgr As IUnknown, _
                              ByRef oObject As Object)
   CreateFreeEdgeTreatment oMD, _
                           oResMgr, _
                           JXSEC_BOTTOM_FLANGE_RIGHT_BOTTOM_CORNER, _
                           oObject
End Sub

Public Sub CMConstructBottomFlangeLeftBottomFET( _
                              ByVal oMD As IJDMemberDescription, _
                              ByVal oResMgr As IUnknown, _
                              ByRef oObject As Object)
   CreateFreeEdgeTreatment oMD, _
                           oResMgr, _
                           JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM, _
                           oObject
End Sub

Public Sub CMConstructBottomFlangeLeftBottomCornerFET( _
                              ByVal oMD As IJDMemberDescription, _
                              ByVal oResMgr As IUnknown, _
                              ByRef oObject As Object)
   CreateFreeEdgeTreatment oMD, _
                           oResMgr, _
                           JXSEC_BOTTOM_FLANGE_LEFT_BOTTOM_CORNER, _
                           oObject
End Sub


'***********************************************************************
' METHOD:  CMCreateClipOppositeSide2
'
' DESCRIPTION:  Determines if a physical connection should be placed on
'               the plate anti molded side. This PC is the second lap PC if valid, that would be
'               between collar and the penetrated part(Cases where the collar overlaps with more than one leaf
'               part of the penetrated part. Assumed a max of 2 lap PCs
'***********************************************************************
Public Sub CMCreateClipOppositeSide2(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
                                  
   
    On Error GoTo ErrorHandler

    sMETHOD = "CMCreateClipOppositeSide2"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Dim sItemName As String
    
    Set pCollar.object = pMD.CAO
    sItemName = pCollar.ItemName
    
    If InStr(sItemName, "_A2") > 0 Or _
       InStr(sItemName, "_B2") > 0 Then
       GoTo Cleanup
    End If

    Dim oCollar As IJCollarPart
    Set oCollar = pMD.CAO
    
    Dim Side As eSideOfPlate
    Side = oCollar.SideOfPlate
    
    Dim bIsValid As Boolean
    bIsValid = False
    
    Dim oPartForPCValid As New StructDetailObjects.PlatePart
    Dim oObject As Object
    
    Dim oSlot As New StructDetailObjects.Slot
    Set oSlot.object = pCollar.Slot
    
    If TypeOf oSlot.Penetrated Is IJStiffener Then
        bIsNeeded = False
        GoTo Cleanup
    End If
    
    If Side = PLATE_ANTI_MOLDED Then
        bIsValid = IsSecondLapPCValid(pMD.CAO, oObject)
        If bIsValid Then
             bIsNeeded = True
        End If
    Else
        bIsNeeded = False
    End If

Cleanup:
    Set oSlot = Nothing
    Set oObject = Nothing
    Set oPartForPCValid = Nothing
    Set pCollar = Nothing
    Set oCollar = Nothing

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
    GoTo Cleanup
End Sub

'***********************************************************************
' METHOD:  CMCreateClipNormalSide2
'
' DESCRIPTION:  Determines if a physical connection should be placed on
'               the plate molded side
'
'***********************************************************************
Public Sub CMCreateClipNormalSide2(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    
    On Error GoTo ErrorHandler

    sMETHOD = "CMCreateClipNormalSide2"
    sERROR = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Dim sItemName As String
    
    Set pCollar.object = pMD.CAO
    sItemName = pCollar.ItemName
    
    If InStr(sItemName, "_A2") > 0 Or _
       InStr(sItemName, "_B2") > 0 Then
       GoTo Cleanup
    End If

    Dim oCollar As IJCollarPart
    Set oCollar = pMD.CAO
    
    Dim Side As eSideOfPlate
    Side = oCollar.SideOfPlate
    
    Dim bIsValid As Boolean
    Dim oPartForPCValid As New StructDetailObjects.PlatePart
    Dim oObject As Object
    
    Dim oSlot As New StructDetailObjects.Slot
    Set oSlot.object = pCollar.Slot
    
    If TypeOf oSlot.Penetrated Is IJStiffener Then
        bIsNeeded = False
        GoTo Cleanup
    End If
    
    If Side = PLATE_MOLDED Then
        bIsValid = IsSecondLapPCValid(pCollar.object, oObject)
        If bIsValid Then
             bIsNeeded = True
        End If
    Else
        bIsNeeded = False
    End If

Cleanup:
    Set oSlot = Nothing
    Set oObject = Nothing
    Set oPartForPCValid = Nothing
    Set oCollar = Nothing
    Set pCollar = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub


'***********************************************************************
' METHOD:  CMCreateBasePC2
'
' DESCRIPTION:  Determines if an extra base pc is needed where a collar overlaps with more than one leaf plate part of Base Plate
'                           (stiffener base plate)
'
'***********************************************************************
Public Sub CMCreateBasePC2(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
                                  
    bIsNeeded = False
    
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO
    Dim pSecondBasePCPort As IJPort
    
    bIsNeeded = IsSecondBasePCValid(pCollar.object, pSecondBasePCPort)
    
    Set pCollar = Nothing

Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub


Public Function CreateLapPC2forCollarOnOtherSide(ByVal pMemberDescription As IJDMemberDescription, _
                                             ByVal pResourceManager As IUnknown) As Object
    On Error GoTo ErrorHandler
    sMETHOD = "CreateLapPC2forCollarOnOtherSide"
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMemberDescription.CAO
    
   
    'get the penetrated plate part associaed with collar and then loop through all its connected leaf plates
    'see if any of the leaf plate other than oAC.connectedObject1 overlaps collar
    'if yes, create PC with it
                   
    Dim bIsValid As Boolean
    Dim oPenetratedPart As New StructDetailObjects.PlatePart
    Dim oPenetratedPort As IJPort
    
    Dim oObject As Object
    bIsValid = IsSecondLapPCValid(pCollar.object, oObject)
    If bIsValid Then
        Set oPenetratedPart.object = oObject
        Set oPenetratedPort = oPenetratedPart.BasePort(BPT_Offset)
    Else
        GoTo Cleanup
    End If
    
    'Get Collar Port
    'Since the collar is expanded in the direction same direction
    ' as Base to Offset Ports from the plate,
    ' the port used for the lap weld on the collar is always the offset
    Dim oCollarPort As IJPort
    Set oCollarPort = pCollar.BasePort(BPT_Base)
    
                               
    sERROR = "Setting Physical Connection to private variable"
    Set CreateLapPC2forCollarOnOtherSide = CAConstruct_PhysConn(pMemberDescription, _
                                                    pResourceManager, "LapWeld", _
                                                    oPenetratedPort, oCollarPort)
                                                    
Cleanup:
    Set oObject = Nothing
    Set oPenetratedPart = Nothing
    Set oPenetratedPort = Nothing
    Set pCollar = Nothing
    Set oCollarPort = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
        
End Function


Public Function CreateLapPC2forCollar(ByVal pMemberDescription As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown) As Object
    On Error GoTo ErrorHandler
    sMETHOD = "CreateLapPC2forCollar"
   
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMemberDescription.CAO
        
    'get the penetrated plate part associaed with collar and then loop through all its connected leaf plates
    'see if any of the leaf plate other than oAC.connectedObject1 overlaps collar
    'if yes, create PC with it
                   
    Dim bIsValid As Boolean
    Dim oPenetratedPart As New StructDetailObjects.PlatePart
    Dim oPenetratedPort As IJPort
    Dim oObject As Object
        
    bIsValid = IsSecondLapPCValid(pMemberDescription.CAO, oObject)
    If bIsValid Then
        Set oPenetratedPart.object = oObject
        Set oPenetratedPort = oPenetratedPart.BasePort(BPT_Base)
    Else
        GoTo Cleanup
    End If
    
        
    'Get Collar Port
    'Since the collar is expanded in the direction same direction
    ' as Base to Offset Ports from the plate,
    ' the port used for the lap weld on the collar is always the offset
    Dim oCollarPort As IJPort
    Set oCollarPort = pCollar.BasePort(BPT_Offset)
    
                               
    sERROR = "Setting Physical Connection to private variable"
    Set CreateLapPC2forCollar = CAConstruct_PhysConn(pMemberDescription, _
                                                    pResourceManager, "LapWeld", _
                                                    oPenetratedPort, oCollarPort)
Cleanup:
    Set oObject = Nothing
    Set oPenetratedPart = Nothing
    Set oPenetratedPort = Nothing
    Set pCollar = Nothing
    Set oCollarPort = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
        
End Function


Public Sub CMCreateBaseRightPC2(ByVal oMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
                                 
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    ' determine Physical Connection between Collar Bottom Edge and Base Plate
    Dim bIsTrue As Boolean
    bIsTrue = False
    
    CMCreateBaseEdgePC2 JXSEC_BOTTOM_RIGHT, oMD, bIsTrue
    
    If bIsTrue Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
        
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCreateBaseRightPC2").Number
End Sub


Public Sub CMCreateBaseLeftPC2(ByVal oMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
                                  
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(oMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
        
    ' Create Physical Connection between Collar Bottom Edge and Base Plate
    Dim bIsTrue As Boolean
    bIsTrue = False
        
    CMCreateBaseEdgePC2 JXSEC_BOTTOM_LEFT, oMD, bIsTrue
    
    If bIsTrue Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMCreateBaseLeftPC2").Number
End Sub


'***********************************************************************
' METHOD:  CMCreateClipNormalSide
'
' DESCRIPTION:  Determines if an extra base pc is needed where a collar overlaps with more than one leaf plate part of Base Plate
'                           (stiffener base plate)
'***********************************************************************
Public Sub CMCreateBaseEdgePC2(ByVal eBaseEdgeId As GSCADSDCreateModifyUtilities.JXSEC_CODE, ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    
    On Error GoTo ErrorHandler
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
       
    bIsNeeded = False
    Dim oCollarWrapper As New StructDetailObjects.Collar
    Set oCollarWrapper.object = pMD.CAO
    
    Dim pSecondBasePCPort As IJPort
    
    bIsNeeded = IsSecondBasePCValid(oCollarWrapper.object, pSecondBasePCPort)
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Sub

Public Sub CMConstructNormalSideLapPC2(ByVal oMD As IJDMemberDescription, _
                                      ByVal oResourceManager As IUnknown, _
                                      ByRef oObject As Object)
    On Error GoTo ErrorHandler
    
    ' Create normal side Lap PC
    Set oObject = CreateLapPC2forCollar(oMD, oResourceManager)
    
    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
  
End Sub

Public Sub CMConstructOppositeSideLapPC2(ByVal oMD As IJDMemberDescription, _
                                      ByVal oResourceManager As IUnknown, _
                                      ByRef oObject As Object)
    On Error GoTo ErrorHandler

    ' Create normal side Lap PC
    Set oObject = CreateLapPC2forCollarOnOtherSide(oMD, oResourceManager)

    Exit Sub
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
 
End Sub

Public Sub CMConstructPC2Lap2(ByVal oMD As IJDMemberDescription, _
                                           ByVal oResourceManager As IUnknown, _
                                           ByRef oObject As Object)
    Dim oCollarCM As New CollarDefCM
  
    Set oObject = oCollarCM.CreateLapPC2forCollar(oMD, oResourceManager)
    Set oCollarCM = Nothing
    
End Sub

Public Sub CMConstructPC2bLap2b(ByVal oMD As IJDMemberDescription, _
                                             ByVal oResourceManager As IUnknown, _
                                             ByRef oObject As Object)
  Dim oCollarCM As New CollarDefCM
  
  Set oObject = oCollarCM.CreateLapPC2forCollarOnOtherSide(oMD, oResourceManager)
  Set oCollarCM = Nothing
  
End Sub


Public Sub CMConstructBaseRightPC2(ByVal oMD As IJDMemberDescription, _
                                  ByVal oResourceManager As IUnknown, _
                                  ByRef oObject As Object)
                                  
    '
    ' Create Physical Connection between Collar Bottom Edge and Base Plate
    ConstructBaseEdgePC2 JXSEC_BOTTOM_RIGHT, oMD, oResourceManager, oObject
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructBaseRightPC").Number
End Sub

Public Sub CMConstructBaseLeftPC2(ByVal oMD As IJDMemberDescription, _
                                 ByVal oResourceManager As IUnknown, _
                                 ByRef oObject As Object)
    
    ' Create Physical Connection between Collar Bottom Left Edge and Base Plate
    ConstructBaseEdgePC2 JXSEC_BOTTOM_LEFT, oMD, oResourceManager, oObject
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CMConstructBaseLeftPC").Number
End Sub

Public Function IsSecondLapPCValid(pCollar As Object, ByRef pPenetratedPart As Object) As Boolean
    
    On Error GoTo ErrorHandler
    
    IsSecondLapPCValid = False
    'get the penetrated plate part associated with collar and then loop through all its connected leaf plates
    'see if any of the leaf plate other than oAC.connectedObject1 overlaps collar
    'if yes, create PC with it
    
    Dim oPlatePart As New StructDetailObjects.PlatePart
    Dim oAssyConn As New StructDetailObjects.AssemblyConn
    
    Dim oCollar As New StructDetailObjects.Collar
    Set oCollar.object = pCollar
    
    Dim oChild As IJDesignChild
    Set oChild = oCollar.object
    
    Dim oParent As Object
    Set oParent = oChild.GetParent
    
    If TypeOf oParent Is IJAssemblyConnection Then
        Set oAssyConn.object = oParent
    Else
        Exit Function
    End If
    
    If Not TypeOf oAssyConn.ConnectedObject1 Is IJPlatePart Then
        Set oPlatePart = Nothing
        Set oAssyConn = Nothing
        Set oCollar = Nothing
        Set oChild = Nothing
        Set oParent = Nothing
        IsSecondLapPCValid = False
        Exit Function
    End If
    
    Set oPlatePart.object = oAssyConn.ConnectedObject1
    
    Dim oRootPenetratedPart As IJPlate
    Dim ppLeafParts As Collection
    Dim oSDObjectsHelper As New StructDetailObjects.Helper
        
    'get the root parent system of penetrated part
    Set oRootPenetratedPart = oSDObjectsHelper.Object_RootParentSystem(oPlatePart.object)
    If oRootPenetratedPart Is Nothing Then
        Set oRootPenetratedPart = oPlatePart.ParentSystem
    End If
    
    'get lateral connected parts of the first plate part
    GetLateralConnectedParts oPlatePart.object, ppLeafParts
            
    Dim oWireBodyUtil As New GSCADShipGeomOps.SGOWireBodyUtilities
    Dim oSurfaceBodyUtil As New GSCADShipGeomOps.SGOSurfaceBodyUtilities
    
    Dim oLeafPart As New StructDetailObjects.PlatePart
    
    Dim oCollarWire As IJWireBody
    Dim oSlotWire As IJWireBody
    
    Dim oPlATtr As GSCADCreateModifyUtilities.IJPlateAttributes
    Set oPlATtr = New GSCADCreateModifyUtilities.PlateUtils
    
    Dim oCollarPart As IJCollarPart
    Set oCollarPart = oCollar.object
    
    On Error Resume Next
    
    'get collar and slot outline
    Set oCollarWire = oCollarPart.GetOutline
    oPlATtr.GetWireBodyFromSymbol oCollar.Slot, oSlotWire
    
    If oSlotWire Is Nothing Or oCollarWire Is Nothing Then
        GoTo Cleanup
    End If
    Dim dToler As Double
    
    If oCollarPart.SideOfPlate = PLATE_MOLDED Then
        If oPlatePart.MoldedSide = "Base" Then
            dToler = 0.01
        Else
            dToler = oPlatePart.PlateThickness + 0.01
        End If
    Else
        If oPlatePart.MoldedSide = "Offset" Then
            dToler = 0.01
        Else
            dToler = oPlatePart.PlateThickness + 0.01
        End If
    End If
    
    Dim oSlotSurface As IJSurfaceBody
    Dim oOverLapGeom As IJWireBody
        'get the surface body from slot wire
    'get the overlapping geometry of collar wire with slot surface body
    If Not oSlotWire Is Nothing Then
        Set oSlotSurface = oWireBodyUtil.GetBoundedSurfaceFromPlanarWire(oSlotWire)
        If Not oSlotSurface Is Nothing Then
            oSurfaceBodyUtil.FindWireOverlapWithTolerance oSlotSurface, oCollarWire, dToler, Nothing, Nothing, oOverLapGeom
        End If
    End If
    
                       
    'get the non-verlapping geometry of collar wire with slot surface body
    Dim oNonOverLapGeom As IJWireBody
    If Not oOverLapGeom Is Nothing Then
        oWireBodyUtil.FindOverlapOrNonOverlapGeometryBetweenWires oCollarWire, oOverLapGeom, False, oNonOverLapGeom, dToler
    End If
    
    If oNonOverLapGeom Is Nothing Then
        GoTo Cleanup
    End If
    
    Dim oOverLapWithPenetrated As IJWireBody
    Dim n As Long
    'loop through all the leaf systems
    
    For n = 1 To ppLeafParts.Count
        If TypeOf ppLeafParts.Item(n) Is IJPlatePart Then
            Set oLeafPart.object = ppLeafParts.Item(n)
            If oCollarPart.SideOfPlate = PLATE_MOLDED Then
                If oPlatePart.MoldedSide = "Base" Then
                    oSurfaceBodyUtil.FindWireOverlapWithTolerance oLeafPart.BasePort(BPT_Base).Geometry, oNonOverLapGeom, dToler, Nothing, Nothing, oOverLapWithPenetrated
                Else
                    oSurfaceBodyUtil.FindWireOverlapWithTolerance oLeafPart.BasePort(BPT_Offset).Geometry, oNonOverLapGeom, dToler, Nothing, Nothing, oOverLapWithPenetrated
                End If
            Else
                If oPlatePart.AlternateMoldedSide = "Base" Then
                    oSurfaceBodyUtil.FindWireOverlapWithTolerance oLeafPart.BasePort(BPT_Offset).Geometry, oNonOverLapGeom, dToler, Nothing, Nothing, oOverLapWithPenetrated
                Else
                    oSurfaceBodyUtil.FindWireOverlapWithTolerance oLeafPart.BasePort(BPT_Base).Geometry, oNonOverLapGeom, dToler, Nothing, Nothing, oOverLapWithPenetrated
                End If
            End If
                       
            If Not oOverLapWithPenetrated Is Nothing Then
                IsSecondLapPCValid = True
                Set pPenetratedPart = oLeafPart.object
                Set oOverLapWithPenetrated = Nothing
                GoTo Cleanup
            End If
            Set oLeafPart.object = Nothing
        End If
    Next
    
Cleanup:
    Set oPlatePart = Nothing
    Set oAssyConn = Nothing
    Set oChild = Nothing
    Set oParent = Nothing
    Set oRootPenetratedPart = Nothing
    Set oSDObjectsHelper = Nothing
    Set oLeafPart = Nothing
    Set oWireBodyUtil = Nothing
    Set oSurfaceBodyUtil = Nothing
    Set oCollarPart = Nothing
    Set oPlATtr = Nothing
    Set oSlotWire = Nothing
    Set oCollarWire = Nothing
    Set oSlotSurface = Nothing
    Set oOverLapGeom = Nothing
    Set oNonOverLapGeom = Nothing
        
    Exit Function

ErrorHandler:
    Err.Raise LogError(Err, MODULE, "IsSecondLapPCValid").Number
End Function
'Method to determine if additional base PC is valid for a collar
Public Function IsSecondBasePCValid(pCollar As Object, ByRef plastSecondBasePCPort As IJPort) As Boolean
    On Error GoTo ErrorHandler
    
    IsSecondBasePCValid = False
    Dim oCollarWrapper As New StructDetailObjects.Collar
    Set oCollarWrapper.object = pCollar
    
    'get the slot
    Dim oSlot As New StructDetailObjects.Slot
    Set oSlot.object = oCollarWrapper.Slot
    
    Dim oCollarPart As IJCollarPart
    Set oCollarPart = oCollarWrapper.object
    
    Dim oCollarWire As IJWireBody
    On Error Resume Next
    Set oCollarWire = oCollarPart.GetOutline
    
    If oCollarWire Is Nothing Then
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    'get the penetration location of slot
    Dim oPenetrationLoc As IJDPosition, oPenetrating As Object
    Set oPenetrationLoc = oSlot.PenetrationLocation
    Set oPenetrating = oCollarWrapper.Penetrating
    
    Dim oSDObjectsHelper As New StructDetailObjects.Helper
    
    Dim nConnData1 As Long, nConnData2 As Long
    Dim zConnData1() As ConnectionData, zConnData2() As ConnectionData
    Dim ncount As Long
    
    Dim oRootBaseSystem As Object, oRootSystem As Object
    Set oRootBaseSystem = oCollarWrapper.BasePlate
    
    'Get the Root plate system of the base plate part of the Collar
    If Not oSDObjectsHelper.Object_RootParentSystem(oRootBaseSystem) Is Nothing Then
        Set oRootBaseSystem = oSDObjectsHelper.Object_RootParentSystem(oRootBaseSystem)
    End If
    
    Dim oBasePlatePart As IJPlatePart
    Dim oBasePlatePort As IJPort
    
    'Get the first Base PC collar holds with base plate and corresponding base plate part.
    
    oSDObjectsHelper.Object_AppConnections oCollarWrapper.object, AppConnectionType_Physical, nConnData1, zConnData1()
    For ncount = 1 To nConnData1
        Set oRootSystem = oSDObjectsHelper.Object_RootParentSystem(zConnData1(ncount).ToConnectable)
        If oRootSystem Is oRootBaseSystem Then
            If TypeOf zConnData1(ncount).ToConnectable Is IJPlatePart Then
                Set oBasePlatePart = zConnData1(ncount).ToConnectable
                Set oBasePlatePort = zConnData1(ncount).ToConnectedPort
                Exit For
            End If
        End If
    Next
    
    Erase zConnData1
    
    'incase collar dont have any base PCs created, take the LeafPlate at AC location
    'inorder to get the context ID of the base plate part the Collar is welded to, use AppConnections on Profilepart at
    'the AC position and get the base plate part the stiffener is welded to.
    
    If oBasePlatePart Is Nothing Then 'get the plate part at the AC location
        'Get the corresponding base plate port
        oSDObjectsHelper.Object_AppConnections oPenetrating, AppConnectionType_Assembly, nConnData2, zConnData2()
        For ncount = 1 To nConnData2
            Set oRootSystem = oSDObjectsHelper.Object_RootParentSystem(zConnData2(ncount).ToConnectable)
            If oRootSystem Is oRootBaseSystem Then
                If TypeOf zConnData2(ncount).ToConnectable Is IJPlatePart Then
                    Set oBasePlatePart = zConnData2(ncount).ToConnectable
                    Set oBasePlatePort = zConnData2(ncount).ToConnectedPort
                    Exit For
                End If
            End If
        Next
    End If
    Erase zConnData2
    
    If oBasePlatePart Is Nothing Or oBasePlatePort Is Nothing Then
        Exit Function
    End If
    
    Dim oUtil As New GSCADStructDetailUtil.StructDetailHelper
    Dim ppLeafParts As Collection
    Set ppLeafParts = New Collection
    
    'get lateral connected parts of the first plate part; these are the leaf parts need to be queried if
    'any of them overlap with collar
    GetLateralConnectedParts oBasePlatePart, ppLeafParts
    
    Dim dTolerance As Double
    
    Dim oSurface As IJSurfaceBody
    Dim oSUtilites As New GSCADShipGeomOps.SGOSurfaceBodyUtilities
    
    Dim oOverLapGeom As IJWireBody
    Dim objPlatePart As New StructDetailObjects.PlatePart
    Dim oSecondBasePCPort As IJPort
    
    Dim sBaseContext As String
    
    'Get the Context ID of the base plate port welded to Collar/Stiffener
    'oCollar.BasePlatePort is not recommended as it is giving asserts when the penetrating is split; observed in WeldSymbol_lap ATP
    If Not ppLeafParts Is Nothing Then
        
        Dim oBasePlateStructPort As IJStructPort
        Set oBasePlateStructPort = oBasePlatePort
                
        If oBasePlateStructPort Is Nothing Then
            IsSecondBasePCValid = False
            Exit Function
        End If
        
        sBaseContext = ""
        
        If oBasePlateStructPort.ContextID = CTX_BASE Then
            sBaseContext = "Base"
        ElseIf oBasePlateStructPort.ContextID = CTX_OFFSET Then
            sBaseContext = "Offset"
        End If
    End If
    
    dTolerance = 0.004
        
    For ncount = 1 To ppLeafParts.Count  'loop through all leaf plates of base plate which are connected to first leaf plate
        If TypeOf ppLeafParts.Item(ncount) Is IJPlatePart Then
            Set objPlatePart.object = ppLeafParts.Item(ncount)
            If sBaseContext = "Base" Then
                Set oSurface = objPlatePart.BasePort(BPT_Base).Geometry
            Else
                Set oSurface = objPlatePart.BasePort(BPT_Offset).Geometry
            End If
            
            On Error Resume Next
            oSUtilites.FindWireOverlapWithTolerance oSurface, oCollarWire, dTolerance, Nothing, Nothing, oOverLapGeom
            
            If Not oOverLapGeom Is Nothing Then
                IsSecondBasePCValid = True
                
                If sBaseContext = "Base" Then
                    Set oSecondBasePCPort = objPlatePart.BasePort(BPT_Base)
                Else
                    Set oSecondBasePCPort = objPlatePart.BasePort(BPT_Offset)
                End If
                
                If Not oSDObjectsHelper.GetEquivalentLastPort(oSecondBasePCPort) Is Nothing Then
                    Set plastSecondBasePCPort = oSDObjectsHelper.GetEquivalentLastPort(oSecondBasePCPort)   'This will be the port with which the additional
                    'PC is created for the Collar.
                Else
                    Set plastSecondBasePCPort = oSecondBasePCPort
                End If

                GoTo Cleanup
            End If
    
            Set objPlatePart.object = Nothing
        End If
    Next
    
Cleanup:
    Set oSurface = Nothing
    Set oSUtilites = Nothing
    Set oOverLapGeom = Nothing
    Set oCollarWrapper = Nothing
    Set oSlot = Nothing
    Set oPenetrationLoc = Nothing
    
    Set objPlatePart.object = Nothing
    Set oBasePlatePart = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
    
End Function


