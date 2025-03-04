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

Const MODULE = "S:\StructDetail\Data\SmartOccurrence\CollarRules\CollarClipParm.cls"

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
    Dim oProfilePart As New StructDetailObjects.ProfilePart
    sERROR = "Setting Profile Part wrapper object to penetrating object"
    Set oProfilePart.object = oPenetratingPart
    Dim oPenetratingPort As IJPort
    sERROR = "Getting Penetrating object's port"
    Set oPenetratingPort = oProfilePart.SubPort(ePenetratingSubPort)
   
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

    sMETHOD = "CMCreateClipNormalSide"
    sERROR = "Setting Conditional value"
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

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

    sMETHOD = "CMCreateClipNormalSide"
    sERROR = "Setting Conditional value"
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

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
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim oPenetrating As New StructDetailObjects.ProfilePart
    Set oPenetrating.object = pCollar.Penetrating
    
      
    Select Case oPenetrating.SectionType
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
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim oPenetrating As New StructDetailObjects.ProfilePart
    Set oPenetrating.object = pCollar.Penetrating
    
      
    Select Case oPenetrating.SectionType
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
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim oPenetrating As New StructDetailObjects.ProfilePart
    Set oPenetrating.object = pCollar.Penetrating
    
      
    Select Case oPenetrating.SectionType
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
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim oPenetrating As New StructDetailObjects.ProfilePart
    Set oPenetrating.object = pCollar.Penetrating
    
      
    Select Case oPenetrating.SectionType
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
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim oPenetrating As New StructDetailObjects.ProfilePart
    Set oPenetrating.object = pCollar.Penetrating
    
      
    Select Case oPenetrating.SectionType
        Case "B", "EA", "UA", "CSType", "C_SS", "I", "ISType", "TSType", "T_XType"
            bIsNeeded = True
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
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim oPenetrating As New StructDetailObjects.ProfilePart
    Set oPenetrating.object = pCollar.Penetrating
    
      
    Select Case oPenetrating.SectionType
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
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim oPenetrating As New StructDetailObjects.ProfilePart
    Set oPenetrating.object = pCollar.Penetrating
    
      
    Select Case oPenetrating.SectionType
        Case "B", "EA", "UA", "CSType", "C_SS", "I", "ISType", "TSType", "T_XType"
            bIsNeeded = True
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
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim oPenetrating As New StructDetailObjects.ProfilePart
    Set oPenetrating.object = pCollar.Penetrating
    
      
    Select Case oPenetrating.SectionType
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
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim oPenetrating As New StructDetailObjects.ProfilePart
    Set oPenetrating.object = pCollar.Penetrating
    
      
    Select Case oPenetrating.SectionType
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
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim oPenetrating As New StructDetailObjects.ProfilePart
    Set oPenetrating.object = pCollar.Penetrating
    
      
    Select Case oPenetrating.SectionType
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
    bIsNeeded = False
    
    Dim pCollar As New StructDetailObjects.Collar
    Set pCollar.object = pMD.CAO

    Dim oPenetrating As New StructDetailObjects.ProfilePart
    Set oPenetrating.object = pCollar.Penetrating
    
      
    Select Case oPenetrating.SectionType
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



