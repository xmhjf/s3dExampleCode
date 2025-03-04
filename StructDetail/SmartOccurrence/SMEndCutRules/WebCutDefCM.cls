VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WebCutDefCM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const MODULE = "S:\StructDetail\Data\SmartOccurrence\" + CUSTOMERID + "EndCutRules\WebCutDefCM.cls"

Private sError As String
Private sMETHOD As String

Public Function GetBoundingProfilePort(ByVal pMemberDescription As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown) As Long
                                    
    sMETHOD = "GetBoundingProfilePort"
    Dim dPortValue As Long
    
    ' Get Class
    Dim pWebCut As New StructDetailObjects.WebCut ' create new web cut wrapper class
    Dim oFlangeCut As New StructDetailObjects.FlangeCut
    sError = "Setting web cut Inputs"
    
    Dim oFeature As GSCADSDCreateModifyUtilities.IJStructFeature
    Set oFeature = pMemberDescription.CAO
    
    ' Get Bounding Part's Port
    Dim oBoundingPart As Object ' This is the Bounding part
    sError = "Getting Bounding object"
    Set oBoundingPart = pWebCut.Bounding ' Get Bounding part from wrapper
    
    Dim oPortBounded As IJPort
    Dim oPortBounding As IJPort
    
    If oFeature.get_StructFeatureType = SF_WebCut Then
        Set pWebCut.object = oFeature
        Set oBoundingPart = pWebCut.Bounding
    Set oPortBounding = pWebCut.BoundingPort
    Set oPortBounded = pWebCut.BoundedPort
    ElseIf oFeature.get_StructFeatureType = SF_FlangeCut Then
        Set oFlangeCut.object = oFeature
        Set oBoundingPart = oFlangeCut.Bounding
        Set oPortBounding = oFlangeCut.BoundingPort
        Set oPortBounded = oFlangeCut.BoundedPort
    End If
    
    If TypeOf oBoundingPart Is IJPlate Then
        dPortValue = JXSEC_IDEALIZED_BOUNDARY
    ElseIf TypeOf oBoundingPart Is IJProfilePart Then
        'Get the name of the port, used to determine side for the symbol
        Dim strPortName As String
        Dim oPartInfo As New StructDetailObjects.Helper
        dPortValue = oPartInfo.GetBoundingProfileFace(oPortBounded, oPortBounding)
    End If
                             
    GetBoundingProfilePort = dPortValue
                             
End Function

'***********************************************************************
' METHOD:  GetEndCutCornerPorts
'
' DESCRIPTION:
'
'***********************************************************************
Private Sub GetEndCutCornerPorts(oEndCut As Object, _
                                oBoundedPart As Object, _
                                lEndCutEdgeId As Long, _
                                oFacePort As IJPort, _
                                oEdgePort1 As IJPort, _
                                oEdgePort2 As IJPort)
sMETHOD = "GetEndCutCornerPorts"
    
    On Error GoTo ErrorHandler
    
    Dim sError As String
    
    Dim lFeatureCtx As Long
    Dim lFeatureOptId As Long
    Dim lFeatureOprId As Long
    
    Dim lStructPortCtx As Long
    Dim lStructPortXid As Long
    Dim lStructPortOptId As Long
    Dim lStructPortOprId As Long
    
    Dim oStructPort As IJStructPort
    
    Dim z_oFacePort As IJPort
    Dim z_oEdgePort1 As IJPort
    Dim z_oEdgePort2 As IJPort
    
    ' save Local copy of Ports passed in
    Set z_oFacePort = oFacePort
    Set z_oEdgePort1 = oEdgePort1
    Set z_oEdgePort2 = oEdgePort2
    
    ' Get EndCut Feature Context, Operation, Operator data
    GetLateBindFeatureData oEndCut, oBoundedPart, _
                           lFeatureCtx, lFeatureOptId, lFeatureOprId
    
    ' Replace Face Port with Port After EndCut is applied (Trim Operation)
    Set oStructPort = z_oFacePort
    lStructPortCtx = oStructPort.ContextID
    lStructPortXid = oStructPort.SectionID
    lStructPortOptId = oStructPort.OperationID
    lStructPortOprId = oStructPort.OperatorID
    Set oFacePort = GetLateBindFeaturePort(oEndCut, oBoundedPart, lFeatureOptId, _
                                           lStructPortCtx, _
                                           lStructPortOptId, _
                                           lStructPortOprId, _
                                           lStructPortXid)
                                                 
    ' Replace Edge Port with Port After EndCut is applied (Trim Operation)
    Set oStructPort = z_oEdgePort1
    lStructPortCtx = oStructPort.ContextID
    lStructPortXid = oStructPort.SectionID
    lStructPortOptId = oStructPort.OperationID
    lStructPortOprId = oStructPort.OperatorID
    Set oEdgePort1 = GetLateBindFeaturePort(oEndCut, oBoundedPart, lFeatureOptId, _
                                            lStructPortCtx, _
                                            lStructPortOptId, _
                                            lStructPortOprId, _
                                            lStructPortXid)
                                                          
    ' Replace Edge Port with Port After EndCut is applied (Trim Operation)
    Set oStructPort = z_oEdgePort2
    lStructPortCtx = oStructPort.ContextID
    lStructPortXid = oStructPort.SectionID
    lStructPortOptId = oStructPort.OperationID
    lStructPortOprId = oStructPort.OperatorID
    Set oEdgePort2 = GetLateBindFeaturePort(oEndCut, oBoundedPart, lFeatureOptId, _
                                            lFeatureCtx, _
                                            lFeatureOptId, _
                                            lFeatureOprId, _
                                            lEndCutEdgeId)

Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number

End Sub

    
'***********************************************************************
' METHOD:  GetLateBindFeatureData
'
' DESCRIPTION:
'
'***********************************************************************
Private Sub GetLateBindFeatureData(oFeatureObject As Object, oPartObject As Object, _
                                  lFeatureCtx As Long, _
                                  lFeatureOptId As Long, _
                                  lFeatureOprId As Long)
sMETHOD = "GetLateBindFeatureData"
    
    On Error GoTo ErrorHandler
    Dim sError As String

    Dim oFeatureUtils As GSCADSDCreateModifyUtilities.IJSDFeatureAttributes
    Set oFeatureUtils = New GSCADSDCreateModifyUtilities.SDFeatureUtils
    
    lFeatureCtx = oFeatureUtils.DetermineFeatureContext(oFeatureObject, oPartObject)
    lFeatureOptId = oFeatureUtils.GetFeatureOperationID(oFeatureObject)
    lFeatureOprId = oFeatureUtils.FeatureOperand(oFeatureObject, oPartObject)
    
    Exit Sub
    
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number

End Sub

'***********************************************************************
' METHOD:  GetLateBindFeaturePort
'
' DESCRIPTION:  Get Ports required for Corner Feature
'
'
'***********************************************************************
Private Function GetLateBindFeaturePort(oFeatureObject As Object, _
                                       oPartObject As Object, _
                                       lFeatureOptId As Long, _
                                       lCtx As Long, _
                                       lOptId As Long, _
                                       lOprId As Long, _
                                       lXid As Long) As Object
sMETHOD = "GetLateBindFeaturePort"
    
    On Error GoTo ErrorHandler
    Dim sError As String

    Dim oMoniker As IMoniker
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
    
    oTopologyLocate.GetLateBindingPortMonikerAfterCutout lOptId, _
                                                         lOprId, _
                                                         lCtx, _
                                                         lXid, _
                                                         oMoniker
        
    Dim oLateBindFeaturePort As IJPort
    Set oLateBindFeaturePort = oTopologyLocate.BindMonikerToPort(oPartObject, _
                                                                 oMoniker, _
                                                                 lFeatureOptId)
        
    Set GetLateBindFeaturePort = oLateBindFeaturePort
    Exit Function
    
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number

End Function

    
Public Function CAConstruct_PhysConn(ByVal pMemberDescription As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown, _
                                    strStartClass As String, _
                                    ByVal eBoundingSubPort As JXSEC_CODE, _
                                    ByVal eBoundedSubPort As JXSEC_CODE) As Object
zMsgBox "WebCutDefCM::CAConstruct_PhysConn" & vbCrLf & _
        "eBoundingSubPort" & Str(eBoundingSubPort) & vbCrLf & _
        "eBoundedSubPort" & Str(eBoundedSubPort)
                                    
    sMETHOD = "CAConstruct_PhysConn"
    Dim phelper As New StructDetailObjects.Helper
    
    ' Get Class
    sError = "Setting Webcut Inputs"
    Dim pWebCut As StructDetailObjects.WebCut
    Set pWebCut = New StructDetailObjects.WebCut
    Set pWebCut.object = pMemberDescription.CAO
    
    ' Get bounding Part's Port
    Dim oBoundingPort As IJPort
    sError = "Getting bounding object"
    
    Dim oBoundedPort As IJPort
    
    Dim oAC As New StructDetailObjects.AssemblyConn
    Dim oChild As IJDesignChild
    Set oChild = pWebCut.object
    
    Set oAC.object = oChild.GetParent
    
    If TypeOf pWebCut.BoundingPort Is IJPort Then
        Set oBoundingPort = phelper.GetEquivalentLastPort(pWebCut.BoundingPort)
    Else
        If phelper.ObjectType(oAC.ConnectedObject2) = SDOBJECT_STIFFENER Then
            Dim oBoundingPart As New StructDetailObjects.ProfilePart
            Set oBoundingPart.object = oAC.ConnectedObject2
            GetBaseOrOffsetPortsAtAC oAC, oBoundingPort, oBoundedPort
        ElseIf phelper.ObjectType(oAC.ConnectedObject2) = SDOBJECT_BEAM Then
            Dim oBoundingBeamPart As New StructDetailObjects.BeamPart
            Set oBoundingBeamPart.object = oAC.ConnectedObject2
            GetBaseOrOffsetPortsAtAC oAC, oBoundingPort, oBoundedPort
        End If
        Set oBoundingPort = phelper.GetEquivalentLastPort(oBoundingPort)
    End If
zMsgBox "   ...GetEquivalentLastPort(pWebCut.BoundingPort)"
  
    ' Get bounded Part's Port
    sError = "Setting Profile Part to bounded object"
    Set oBoundedPort = pWebCut.BoundedPort
    
    If phelper.ObjectType(pWebCut.Bounded) = SDOBJECT_STIFFENER Then
        Dim oBoundedPart As New StructDetailObjects.ProfilePart
        Set oBoundedPart.object = pWebCut.Bounded
        Set oBoundedPort = oBoundedPart.CutoutSubPort(pWebCut.object, eBoundedSubPort)
zMsgBox "   ...ProfilePart::CutoutSubPort(pWebCut.object)"
    
    ElseIf phelper.ObjectType(pWebCut.Bounded) = SDOBJECT_BEAM Then
        Dim oBoundedBeamPart As New StructDetailObjects.BeamPart
        Set oBoundedBeamPart.object = pWebCut.Bounded
        Set oBoundedPort = oBoundedBeamPart.CutoutSubPort(pWebCut.object, eBoundedSubPort)
zMsgBox "   ...BeamPart::CutoutSubPort(pWebCut.object)"
    End If
    
    ' Get the profile part, since it is the parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the webcut's ca
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    ' Create physical connection
    Dim oPhysicalConnection As New PhysicalConn
    sError = "Creating Physical Connection"
    oPhysicalConnection.Create pResourceManager, _
                                            oBoundedPort, _
                                            oBoundingPort, _
                                            strStartClass, _
                                            pSystemParent, ConnectionStandard
zMsgBox "   ...oPhysicalConnection.Create"
                               
    sError = "Setting Physical Connection to private variable"
    Set CAConstruct_PhysConn = oPhysicalConnection.object
    
    Set oBoundingPort = Nothing
    Set oBoundedPort = Nothing
    Set oPhysicalConnection = Nothing
    Set pSystemParent = Nothing
    Set pWebCut = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Public Sub WebCut_FinalConstruct(ByVal pAggregatorDescription As IJDAggregatorDescription)

    On Error GoTo ErrorHandler
    
    sMETHOD = "WebCut_FinalConstruct"
    If (pAggregatorDescription Is Nothing) Then Exit Sub
    ' Get slot to add to cut
    Dim pWebCut As IJStructFeature
    Set pWebCut = pAggregatorDescription.CAO
      
    ' Get Bounded object to add webcut to
    Dim pBoundedPart As Object
    Dim oWebCut As New StructDetailObjects.WebCut
    Set oWebCut.object = pWebCut
    Set pBoundedPart = oWebCut.Bounded
    
    ' Create the SDCreateModifyUtilities
    Dim pIJFeatureDefinition As IJSDFeatureDefinition
    Set pIJFeatureDefinition = New SDFeatureUtils
    pIJFeatureDefinition.PlaceFeature pWebCut, pBoundedPart
    
    Set pIJFeatureDefinition = Nothing
    Set pWebCut = Nothing
    Set pBoundedPart = Nothing
    Set oWebCut = Nothing
    
    Exit Sub
      
ErrorHandler:
      Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

'***********************************************************************
' METHOD:  CMWWebLeftCon
'
' DESCRIPTION:  Conditional for physical connection
'
'***********************************************************************
Public Sub CMWWebLeftCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMWWebLeftCon"
    sError = "Setting Conditional value"

    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
    End Sub
'***********************************************************************
' METHOD:  CMWWebLeft
'
' DESCRIPTION:  Create physical connection for JXSEC_WEB_LEFT
'
'
'***********************************************************************
Public Sub CMWWebLeftPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_WEB_LEFT, JXSEC_WEB_LEFT)
End Sub

'***********************************************************************
' METHOD:  CMIdealizedCon
'
' DESCRIPTION:  Conditional for physical connection
'
'***********************************************************************
Public Sub CMIdealizedCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMIdealizedCon"
    sError = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    If TypeOf oSDO_WebCut.BoundingPort Is IJPort Then
        bIsNeeded = True
    Else
        bIsNeeded = False
    End If
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
    End Sub
    '***********************************************************************
' METHOD:  CMWWebRightCon
'
' DESCRIPTION:  Conditional for physical connection
'
'***********************************************************************
Public Sub CMWWebRightCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMWWebRightCon"
    sError = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
    End Sub

'***********************************************************************
' METHOD:  CMWWebRight
'
' DESCRIPTION:  Create physical connection for JXSEC_WEB_RIGHT
'
'
'***********************************************************************
Public Sub CMWWebRightPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_WEB_RIGHT, JXSEC_WEB_RIGHT)
End Sub

' *****************************************
' *****************************************
' This following are temporary copies of Structdetailobjects methods
' for use while TR??? is being resolved
' *****************************************
' *****************************************
Private Sub CornerFeatureData(bAtTopCorner As Boolean, _
                             oCornerPosition As IJDPosition, _
                             oFacePort As IJPort, _
                             oEdgePort1 As IJPort, _
                             oEdgePort2 As IJPort, _
                             ByVal pMD As IJDMemberDescription)
sMETHOD = "CornerFeatureData"
    
    On Error GoTo ErrorHandler
    
    Dim oBoundedLocation As IJDPosition
    Dim eFaceEdgePort As JXSEC_CODE
    
' ---------------------------
' This is customized code
' ---------------------------
Dim oSmartOccurrence As DEFINITIONHELPERSINTFLib.IJSmartOccurrence
Set oSmartOccurrence = pMD.CAO

Dim oSDOWebCut As New StructDetailObjects.WebCut
Set oSDOWebCut.object = oSmartOccurrence

Set oBoundedLocation = oSDOWebCut.BoundedLocation

    ' set initial location for Corner Feature as the Bounded Location
'    Set oBoundedLocation = BoundedLocation

' ---------------------------
' end customized code
' ---------------------------
    
    ' Using Web_left as the Corner Face Port
    ' Use Top Face Port if Corner Feature is to be at Top Corner of Web
    ' Else Use Bottom Face Port if Corner Feature is NOT at Top Corner of Web
    If bAtTopCorner Then
        eFaceEdgePort = JXSEC_TOP
    Else
        eFaceEdgePort = JXSEC_BOTTOM
    End If
    
    ' determine the Ports required for the Web Cut Corner Feature
    EndCutCornerFeatureData oSDOWebCut.Bounded, oBoundedLocation, _
                            JXSEC_WEB_LEFT, eFaceEdgePort, _
                            oCornerPosition, oFacePort, oEdgePort1, oEdgePort2
    
' ---------------------------
' This is more customized code
' ---------------------------
Dim sItemName As String
sItemName = oSDOWebCut.ItemName

Dim edgeID As Long

Select Case sItemName
    Case "WebCut_C2Left", _
         "WebCut_C3Left", _
         "WebCut_W1Left", _
         "WebCutI_C3Left", _
         "WebCutT_W1Left"
        edgeID = JXSEC_WEB_LEFT
    Case "WebCut_C2Right", _
         "WebCut_C3Right", _
         "WebCut_C3RBulb", _
         "WebCut_W1Right", _
         "WebCut_W1RBulb", _
         "WebCutI_C3Right", _
         "WebCutT_W1Right"
        edgeID = JXSEC_WEB_RIGHT
    Case Else
        edgeID = JXSEC_IDEALIZED_BOUNDARY
End Select
' ---------------------------
' end customized code
' ---------------------------

    ' need to get the Port Monikers for After "Trim" Operation
    GetEndCutCornerPorts oSDOWebCut.object, oSDOWebCut.Bounded, edgeID, _
                         oFacePort, oEdgePort1, oEdgePort2
    
Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD).Number

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
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)

    Select Case sSectionType
        Case "BUT", "BUTL2", "I", "ISType", "TSType", "T_XType", "H"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeLeftBottom
'
' DESCRIPTION:  Create physical connection for JXSEC_TOP_FLANGE_LEFT_BOTTOM
'
'
'***********************************************************************
Public Sub CMTopFlangeLeftBottomPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_TOP_FLANGE_LEFT_BOTTOM, JXSEC_TOP_FLANGE_LEFT_BOTTOM)
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
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "BUT", "BUTL2", "I", "ISType", "TSType", "T_XType", gsB, "H"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeRightBottom
'
' DESCRIPTION:  Create physical connection for JXSEC_TOP_FLANGE_RIGHT_BOTTOM
'
'
'***********************************************************************
Public Sub CMTopFlangeRightBottomPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_TOP_FLANGE_RIGHT_BOTTOM, JXSEC_TOP_FLANGE_RIGHT_BOTTOM)
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
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "BUT", "BUTL2", "I", "ISType", "TSType", "T_XType", "H"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
    End Sub
'***********************************************************************
' METHOD:  CMTopFlangeLeft
'
' DESCRIPTION:  Create physical connection for JXSEC_TOP_FLANGE_LEFT
'
'
'***********************************************************************
Public Sub CMTopFlangeLeftPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_TOP_FLANGE_LEFT, JXSEC_TOP_FLANGE_LEFT)
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
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "BUT", "BUTL2", "I", "ISType", "TSType", "T_XType", "H"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
    End Sub
'***********************************************************************
' METHOD:  CMTopFlangeRight
'
' DESCRIPTION:  Create physical connection for JXSEC_TOP_FLANGE_RIGHT
'
'
'***********************************************************************
Public Sub CMTopFlangeRightPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_TOP_FLANGE_RIGHT, JXSEC_TOP_FLANGE_RIGHT)
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
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case gsB, "I", "ISType", "TSType", "T_XType", "H"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMWebLeftTopCorner
'
' DESCRIPTION:  Create physical connection for JXSEC_WEB_LEFT_TOP_CORNER
'
'
'***********************************************************************
Public Sub CMWebLeftTopCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_WEB_LEFT_TOP_CORNER, JXSEC_WEB_LEFT_TOP_CORNER)
End Sub

'***********************************************************************
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
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
    
    Select Case sSectionType
        Case gsB, "I", "ISType", "TSType", "T_XType", "H"
           If oSDO_WebCut.ItemName = "Cope_AngleRight" Then
                bIsNeeded = False
            Else
                bIsNeeded = True
            End If
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMWebRightTopCorner
'
' DESCRIPTION:  Create physical connection for JXSEC_WEB_RIGHT_TOP_CORNER
'
'
'***********************************************************************
Public Sub CMWebRightTopCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_WEB_RIGHT_TOP_CORNER, JXSEC_WEB_RIGHT_TOP_CORNER)
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
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "I", "ISType", "TSType", "T_XType"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeLeftBottomCorner
'
' DESCRIPTION:  Create physical connection for JXSEC_TOP_FLANGE_LEFT_BOTTOM_CORNER
'
'
'***********************************************************************
Public Sub CMTopFlangeLeftBottomCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_TOP_FLANGE_LEFT_BOTTOM_CORNER, JXSEC_TOP_FLANGE_LEFT_BOTTOM_CORNER)
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
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "I", "ISType", "TSType", "T_XType"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMTopFlangeRightBottomCorner
'
' DESCRIPTION:  Create physical connection for JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER
'
'
'***********************************************************************
Public Sub CMTopFlangeRightBottomCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER, JXSEC_TOP_FLANGE_RIGHT_BOTTOM_CORNER)
End Sub

'***********************************************************************
' METHOD:  CMWebLeftBottomCornerCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMWebLeftBottomCornerCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMWebLeftBottomCornerCon"
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "I", "ISType", "H"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMWebLeftBottomCorner
'
' DESCRIPTION:  Create physical connection for JXSEC_WEB_LEFT_BOTTOM_CORNER
'
'
'***********************************************************************
Public Sub CMWebLeftBottomCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_WEB_LEFT_BOTTOM_CORNER, JXSEC_WEB_LEFT_BOTTOM_CORNER)
End Sub

'***********************************************************************
' METHOD:  CMWebRightBottomCornerCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMWebRightBottomCornerCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMWebRightBottomCornerCon"
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "I", "ISType", "H"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMWebRightBottomCorner
'
' DESCRIPTION:  Create physical connection for JXSEC_WEB_RIGHT_BOTTOM_CORNER
'
'
'***********************************************************************
Public Sub CMWebRightBottomCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_WEB_RIGHT_BOTTOM_CORNER, JXSEC_WEB_RIGHT_BOTTOM_CORNER)
End Sub


'***********************************************************************
' METHOD:  CMBottomFlangeLeftTopCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMBottomFlangeLeftTopCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMBottomFlangeLeftTopCon"
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "I", "ISType", "H"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMBottomFlangeLeftTop
'
' DESCRIPTION:  Create physical connection for JXSEC_BOTTOM_FLANGE_LEFT_TOP
'
'
'***********************************************************************
Public Sub CMBottomFlangeLeftTopPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_BOTTOM_FLANGE_LEFT_TOP, JXSEC_BOTTOM_FLANGE_LEFT_TOP)
End Sub

'***********************************************************************
' METHOD:  CMBottomFlangeRightTopCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMBottomFlangeRightTopCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMBottomFlangeRightTopCon"
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "I", "ISType", "H"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMBottomFlangeRightTop
'
' DESCRIPTION:  Create physical connection for JXSEC_BOTTOM_FLANGE_RIGHT_TOP
'
'
'***********************************************************************
Public Sub CMBottomFlangeRightTopPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_BOTTOM_FLANGE_RIGHT_TOP, JXSEC_BOTTOM_FLANGE_RIGHT_TOP)
End Sub

'***********************************************************************
' METHOD:  CMBottomFlangeLeftTopCornerCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMBottomFlangeLeftTopCornerCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMBottomFlangeLeftTopCornerCon"
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If

    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "I", "ISType"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMBottomFlangeLeftTopCorner
'
' DESCRIPTION:  Create physical connection for JXSEC_BOTTOM_FLANGE_LEFT_TOP_CORNER
'
'
'***********************************************************************
Public Sub CMBottomFlangeLeftTopCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_BOTTOM_FLANGE_LEFT_TOP_CORNER, JXSEC_BOTTOM_FLANGE_LEFT_TOP_CORNER)
End Sub

'***********************************************************************
' METHOD:  CMBottomFlangeRightTopCornerCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMBottomFlangeRightTopCornerCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMBottomFlangeRightTopCornerCon"
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "I", "ISType"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMBottomFlangeRightTopCorner
'
' DESCRIPTION:  Create physical connection for JXSEC_BOTTOM_FLANGE_RIGHT_TOP_CORNER
'
'
'***********************************************************************
Public Sub CMBottomFlangeRightTopCornerPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_BOTTOM_FLANGE_RIGHT_TOP_CORNER, JXSEC_BOTTOM_FLANGE_RIGHT_TOP_CORNER)
End Sub


'***********************************************************************
' METHOD:  CMBottomFlangeLeftCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMBottomFlangeLeftCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMBottomFlangeLeftCon"
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "I", "ISType", "H"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMBottomFlangeLeft
'
' DESCRIPTION:  Create physical connection for JXSEC_BOTTOM_FLANGE_LEFT
'
'
'***********************************************************************
Public Sub CMBottomFlangeLeftPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_BOTTOM_FLANGE_LEFT, JXSEC_BOTTOM_FLANGE_LEFT)
End Sub

'***********************************************************************
' METHOD:  CMBottomFlangeRightCon
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub CMBottomFlangeRightCon(ByRef pMD As IJDMemberDescription, _
                                  ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMBottomFlangeRightCon"
    sError = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "I", "ISType", "H"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub
'***********************************************************************
' METHOD:  CMBottomFlangeRight
'
' DESCRIPTION:  Create physical connection for JXSEC_BOTTOM_FLANGE_RIGHT
'
'
'***********************************************************************
Public Sub CMBottomFlangeRightPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, "TeeWeld", _
                             JXSEC_BOTTOM_FLANGE_RIGHT, JXSEC_BOTTOM_FLANGE_RIGHT)
End Sub



Public Function CAConstruct_PhysConnNames(ByVal pMemberDescription As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown, _
                                    strStartClass As String, _
                                    ByVal eBoundingSubPort As JXSEC_CODE, _
                                    ByVal eBoundedSubPort As JXSEC_CODE) As Object

    sMETHOD = "CAConstruct_PhysConnNames"
    
    ' Get Class
    Dim pWebCut As New StructDetailObjects.WebCut
    Dim oFlangCut As New StructDetailObjects.FlangeCut
    sError = "Setting web cut Inputs"
    
    
    ' Get Bounding Part's Port
    Dim oBoundingPart As Object ' This is the Bounding part
    Dim oBoundingPort As IJPort
    sError = "Getting Bounding object"
    ' Get Bounding part from wrapper
    Dim phelper As StructDetailObjects.Helper
    Set phelper = New StructDetailObjects.Helper
    
    Dim oBounding As Object

    Dim oFeature As GSCADSDCreateModifyUtilities.IJStructFeature
    Set oFeature = pMemberDescription.CAO
    
    Dim oSmOcc As IJSmartOccurrence
    Dim oSmItem As IJSmartItem
    Dim sItemName As String
    
    If oFeature.get_StructFeatureType = SF_WebCut Then
        Set pWebCut.object = pMemberDescription.CAO
        Set oBoundingPart = pWebCut.Bounding
        Set oSmOcc = pWebCut.object
        Set oSmItem = oSmOcc.SmartItemObject
        sItemName = oSmItem.Name
    ElseIf oFeature.get_StructFeatureType = SF_FlangeCut Then
        Set oFlangCut.object = pMemberDescription.CAO
        Set oBoundingPart = oFlangCut.Bounding
        Set oSmOcc = oFlangCut.object
        Set oSmItem = oSmOcc.SmartItemObject
        sItemName = oSmItem.Name
    End If
    
    If TypeOf oBoundingPart Is IJPlate Then
        Dim pSDOHelper As New StructDetailObjects.Helper
        Dim oLastPlatePort As IJPort
        If oFeature.get_StructFeatureType = SF_WebCut Then
        Set oLastPlatePort = pWebCut.BoundingPort
        Set oBoundingPort = pSDOHelper.GetEquivalentLastPort(oLastPlatePort)
        ElseIf oFeature.get_StructFeatureType = SF_FlangeCut Then
            Set oLastPlatePort = oFlangCut.BoundingPort
            Set oBoundingPort = pSDOHelper.GetEquivalentLastPort(oLastPlatePort)
        End If
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        sError = "Setting Profile Part wrapper object to penetrating object"
        Set oProfilePart.object = oBoundingPart
        sError = "Getting Bounding object's port"
        If oFeature.get_StructFeatureType = SF_WebCut Then
            Set oBoundingPort = oProfilePart.SubPort(eBoundingSubPort)
        ElseIf oFeature.get_StructFeatureType = SF_FlangeCut Then
            If sItemName = "FlangeCut_W4" Or sItemName = "FlangeCut_W1" Then
                If eBoundingSubPort = JXSEC_WEB_LEFT Or eBoundingSubPort = JXSEC_TOP_FLANGE_LEFT Then
                    Set oBoundingPort = oProfilePart.SubPort(JXSEC_TOP_FLANGE_LEFT)
                ElseIf eBoundingSubPort = JXSEC_WEB_RIGHT Or eBoundingSubPort = JXSEC_TOP_FLANGE_RIGHT Then
                    Set oBoundingPort = oProfilePart.SubPort(JXSEC_TOP_FLANGE_RIGHT)
                End If
            Else
        Set oBoundingPort = oProfilePart.SubPort(eBoundingSubPort)
            End If
        End If
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_BEAM Then
        Dim oBeamPart As New StructDetailObjects.BeamPart
        sError = "Setting Beam Part wrapper object to penetrating object"
        Set oBeamPart.object = oBoundingPart
        sError = "Getting Bounding object's port"
        If oFeature.get_StructFeatureType = SF_WebCut Then
        Set oBoundingPort = oBeamPart.SubPort(eBoundingSubPort)
        ElseIf oFeature.get_StructFeatureType = SF_FlangeCut Then
            If sItemName = "FlangeCut_W4" Or sItemName = "FlangeCut_W1" Then
                If eBoundingSubPort = JXSEC_WEB_LEFT Or eBoundingSubPort = JXSEC_TOP_FLANGE_LEFT Then
                    Set oBoundingPort = oBeamPart.SubPort(JXSEC_TOP_FLANGE_LEFT)
                ElseIf eBoundingSubPort = JXSEC_WEB_RIGHT Or eBoundingSubPort = JXSEC_TOP_FLANGE_RIGHT Then
                    Set oBoundingPort = oBeamPart.SubPort(JXSEC_TOP_FLANGE_RIGHT)
                End If
            Else
                Set oBoundingPort = oBeamPart.SubPort(eBoundingSubPort)
            End If
        End If
    End If
   
    ' Get bounded profile Port
        Dim oBoundedPort As IJPort
        Dim oBoundedPart As New StructDetailObjects.ProfilePart
        Dim oBoundedBeamPart As New StructDetailObjects.BeamPart
        If oFeature.get_StructFeatureType = SF_WebCut Then
    If phelper.ObjectType(pWebCut.Bounded) = SDOBJECT_STIFFENER Then
        Set oBoundedPart = New StructDetailObjects.ProfilePart
        Set oBoundedPart.object = pWebCut.Bounded
        Set oBoundedPort = oBoundedPart.CutoutSubPort(pWebCut.object, eBoundedSubPort)
    ElseIf phelper.ObjectType(pWebCut.Bounded) = SDOBJECT_BEAM Then
        Set oBoundedBeamPart = New StructDetailObjects.BeamPart
        Set oBoundedBeamPart.object = pWebCut.Bounded
        Set oBoundedPort = oBoundedBeamPart.CutoutSubPort(pWebCut.object, eBoundedSubPort)
    End If
        ElseIf oFeature.get_StructFeatureType = SF_FlangeCut Then
            If phelper.ObjectType(oFlangCut.Bounded) = SDOBJECT_STIFFENER Then
                Set oBoundedPart.object = oFlangCut.Bounded
                Set oBoundedPort = oBoundedPart.CutoutSubPort(oFlangCut.object, eBoundedSubPort)
            ElseIf phelper.ObjectType(oFlangCut.Bounded) = SDOBJECT_BEAM Then
                Set oBoundedBeamPart.object = oFlangCut.Bounded
                Set oBoundedPort = oBoundedBeamPart.CutoutSubPort(oFlangCut.object, eBoundedSubPort)
            End If
        End If
        
    ' Get the web cut, make it a parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the slot's ca
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    ' Create physical connection
    Dim oPhysicalConnection As New PhysicalConn
    sError = "Creating Physical Connection"
    oPhysicalConnection.Create pResourceManager, _
                               oBoundedPort, _
                               oBoundingPort, _
                               strStartClass, _
                               pSystemParent, ConnectionStandard
                               
    sError = "Setting Physical Connection to private variable"
    Set CAConstruct_PhysConnNames = oPhysicalConnection.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
        
End Function

'
'***********************************************************************
' METHOD:  CMWebLeftPC
'
' DESCRIPTION:  Create physical connection for WebLeft  # 257
'
'
'***********************************************************************
Public Sub CMWebLeftPC(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, _
                                          "TeeWeld", _
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

  Set pObject = CAConstruct_PhysConnNames(pMemberDescription, pResourceManager, _
                                          "TeeWeld", _
                                          JXSEC_WEB_RIGHT, JXSEC_WEB_RIGHT)
End Sub

'***********************************************************************
' METHOD:  ConstructPC1_Con
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub ConstructPC1_Con(ByRef pMD As IJDMemberDescription, _
                            ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "ConstructPC1_Con"
    sError = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True
    
zMsgBox "WebCutDefCM::ConstructPC1_Con ...bIsNeeded:" & bIsNeeded
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

'***********************************************************************
' METHOD:  ConstructPC2_Con
'
' DESCRIPTION:  Determines if a physical connection should be placed
'
'
'***********************************************************************
Public Sub ConstructPC2_Con(ByRef pMD As IJDMemberDescription, _
                            ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "ConstructPC2_Con"
    sError = "Setting Conditional value"
    
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If

    Dim sSectionType As String
    Dim sSplitEndToEndCase As String
    
    sSplitEndToEndCase = GetAnswer_SplitEndToEndCase(pMD)
    
    ' Initialize wrapper class and get the bounded profile
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
        
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounded)
    Set oSDO_WebCut = Nothing
    
    If (sSectionType = "FB") Then
        bIsNeeded = False
    
    ElseIf LCase(Trim(sSplitEndToEndCase)) = LCase("AngleWebSquareFlange") Or _
       LCase(Trim(sSplitEndToEndCase)) = LCase("AngleWebBevelFlange") Or _
       LCase(Trim(sSplitEndToEndCase)) = LCase("AngleWebAngleFlange") Or _
       LCase(Trim(sSplitEndToEndCase)) = LCase("DistanceWebDistanceFlange") Or _
       LCase(Trim(sSplitEndToEndCase)) = LCase("OffsetWebOffsetFlange") Then
        bIsNeeded = True
    End If
    
    
zMsgBox "WebCutDefCM::ConstructPC2_Con ...bIsNeeded:" & bIsNeeded
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

Public Function GetAnswer_SplitEndToEndCase(ByRef pMD As IJDMemberDescription) As String
sMETHOD = "GetAnswer_SplitEndToEndCase"
    
    On Error GoTo ErrorHandler
    
    Dim vAnswer As Variant
    Dim sSplitEndToEndCase As String
    
    Dim oObject As IJDObject
    Dim oResourceManagerUnknown As IUnknown
    
    Dim oDefinitions As IMSSymbolEntities.IJDDefinitions
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim oParentSmartOccurrence As DEFINITIONHELPERSINTFLib.IJSmartOccurrence
    Dim oSymbolEntitiesFactory As IMSSymbolEntities.IJDSymbolEntitiesFactory
    
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    sSplitEndToEndCase = ""
    
    Set oParentSmartOccurrence = pMD.CAO
    Set oObject = oParentSmartOccurrence.ItemObject
    Set oResourceManagerUnknown = oObject.ResourceManager
 
    
    On Error GoTo ErrorHandler

    Set oSymbolEntitiesFactory = New IMSSymbolEntities.DSymbolEntitiesFactory
    Set oDefinitions = oSymbolEntitiesFactory.DefinitionCollection(oResourceManagerUnknown)
    Set oSymbolDefinition = oDefinitions.Item(CUSTOMERID + "EndCutRules.EndToEndWebCutSel")
    
    Set oCommonHelper = New DefinitionHlprs.CommonHelper
    vAnswer = oCommonHelper.GetAnswer(oParentSmartOccurrence, _
                                      oSymbolDefinition, _
                                      "SplitEndToEndCase")
    sSplitEndToEndCase = vAnswer
    On Error Resume Next

    GetAnswer_SplitEndToEndCase = sSplitEndToEndCase
zMsgBox "WebCutDefCM::GetAnswer_SplitEndToEndCase ...:" & sSplitEndToEndCase

    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Function

Private Function GetAnswer_AddCornerFeature(ByRef pMD As IJDMemberDescription) As String
sMETHOD = "GetAnswer_AddCornerFeature"
    
    On Error GoTo ErrorHandler
    
    Dim oSmartOccurrence As DEFINITIONHELPERSINTFLib.IJSmartOccurrence
    Set oSmartOccurrence = pMD.CAO
    
    ' Since we know the FlangeCutSelector does not have the question, we don't want to ask
    ' Set answer to "No" and exit early
    Dim oFeature As IJStructFeature
    If TypeOf oSmartOccurrence Is IJStructFeature Then
        Set oFeature = oSmartOccurrence
        If Not oFeature.get_StructFeatureType = SF_WebCut Then
            GetAnswer_AddCornerFeature = "No"
            Exit Function
        End If
    End If
    
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Set oSymbolDefinition = oSmartOccurrence.RootSelectionObject.SelectionRuleDef
    
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    Set oCommonHelper = New DefinitionHlprs.CommonHelper

    Dim vAnswer As Variant
    GetAnswer_AddCornerFeature = ""
    vAnswer = oCommonHelper.GetAnswer(oSmartOccurrence, _
                                      oSymbolDefinition, _
                                      gsAddCornerFeature)
    GetAnswer_AddCornerFeature = vAnswer
    On Error Resume Next

    zMsgBox "WebCutDefCM::GetAnswer_AddCornerFeature ...:" & GetAnswer_AddCornerFeature

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

Public Sub CMConstructFET(ByVal pMemberDescription As IJDMemberDescription, ByVal pResourceManager As IUnknown, ByRef pObject As Object)
    
    On Error GoTo ErrorHandler
    Set pObject = CMConstructFETForFeature(pMemberDescription, pResourceManager)
    
    Exit Sub
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMConstructFET1").Number
End Sub
Public Sub CMGetEdgeTreatment(ByRef pMD As IJDMemberDescription, _
                          ByRef bIsNeeded As Boolean)
 On Error GoTo ErrorHandler
    bIsNeeded = False
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_FreeEdgeTreatment) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    
    Dim vAnswer As Variant
    Dim sApplyTreatment As String
    
    Dim oObject As IJDObject
    Dim oResourceManagerUnknown As IUnknown
    
    Dim oDefinitions As IMSSymbolEntities.IJDDefinitions
    Dim oSymbolDefinition As IMSSymbolEntities.IJDSymbolDefinition
    Dim oParentSmartOccurrence As DEFINITIONHELPERSINTFLib.IJSmartOccurrence
    Dim oSymbolEntitiesFactory As IMSSymbolEntities.IJDSymbolEntitiesFactory
    
    Dim oCommonHelper As DefinitionHlprs.CommonHelper
    
    sApplyTreatment = ""
    
    Set oParentSmartOccurrence = pMD.CAO
 
    Dim oSmartItem As IJSmartItem
    Dim oParentSmartClass As IJSmartClass
    
    Set oSmartItem = oParentSmartOccurrence.ItemObject
    Set oParentSmartClass = oSmartItem.Parent
    
    Set oSymbolDefinition = oParentSmartClass.SelectionRuleDef
        
    On Error GoTo ErrorHandler

    Set oCommonHelper = New DefinitionHlprs.CommonHelper
    vAnswer = oCommonHelper.GetAnswer(oParentSmartOccurrence, _
                                      oSymbolDefinition, _
                                      "ApplyTreatment")
    sApplyTreatment = vAnswer
    
    If LCase(sApplyTreatment) = LCase("No") Then
        bIsNeeded = False
    ElseIf LCase(sApplyTreatment) = LCase("Yes") Then
        bIsNeeded = True
    End If
    
    Exit Sub
  
ErrorHandler:
  Err.Raise LogError(Err, MODULE, "CMGetEdgeTreatment").Number
End Sub


Public Sub CMConstructLapPC(ByVal pMemDesc As IJDMemberDescription, _
                            ByVal oResourceMgr As IUnknown, _
                            ByRef pObject As Object)
                            
    Set pObject = CAConstruct_LappedPhysConn(pMemDesc, oResourceMgr)
End Sub

Private Function CAConstruct_LappedPhysConn(ByVal pMemberDescription As IJDMemberDescription, _
                                           ByVal pResourceManager As IUnknown) As Object

    sMETHOD = "CAConstruct_LappedPhysConn"
            
    ' ------------------------------------
    ' Get the web cut object
    ' ------------------------------------
    Dim oWebCut As StructDetailObjects.WebCut
    Set oWebCut = New StructDetailObjects.WebCut
    
    With oWebCut
    
    Set .object = pMemberDescription.CAO
       
    ' -------------------------------------
    ' Get primary ports on bounded part
    ' -------------------------------------
    ' Get bounded position
    Dim oBoundedPos As IJDPosition
    Set oBoundedPos = .BoundedLocation
    
    ' Presume that it is lapped to WebLeft, WebRight, Top, or Bottom
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
        
    Dim oBoundedPort(4) As IJPort
    
    Set oBoundedPort(1) = oTopologyLocate.GetSubPort(.Bounded, JXSEC_WEB_LEFT)
    Set oBoundedPort(2) = oTopologyLocate.GetSubPort(.Bounded, JXSEC_WEB_RIGHT)
    Set oBoundedPort(3) = oTopologyLocate.GetSubPort(.Bounded, JXSEC_TOP)
    Set oBoundedPort(4) = oTopologyLocate.GetSubPort(.Bounded, JXSEC_BOTTOM)
    
    Dim oBoundedGeom(4) As IJSurfaceBody
    
    Set oBoundedGeom(1) = oBoundedPort(1).Geometry
    Set oBoundedGeom(2) = oBoundedPort(2).Geometry
    Set oBoundedGeom(3) = oBoundedPort(3).Geometry
    Set oBoundedGeom(4) = oBoundedPort(4).Geometry
    
    Dim oBoundedProjPos(4) As IJDPosition
    
    Dim oModelUtil As IJSGOModelBodyUtilities
    Set oModelUtil = New SGOModelBodyUtilities
    
    Dim dist As Double
    
    oModelUtil.GetClosestPointOnBody oBoundedGeom(1), oBoundedPos, oBoundedProjPos(1), dist
    oModelUtil.GetClosestPointOnBody oBoundedGeom(2), oBoundedPos, oBoundedProjPos(2), dist
    oModelUtil.GetClosestPointOnBody oBoundedGeom(3), oBoundedPos, oBoundedProjPos(3), dist
    oModelUtil.GetClosestPointOnBody oBoundedGeom(4), oBoundedPos, oBoundedProjPos(4), dist
    
    Dim oBoundedNorm(4) As IJDVector
    
    oBoundedGeom(1).GetNormalFromPosition oBoundedProjPos(1), oBoundedNorm(1)
    oBoundedGeom(2).GetNormalFromPosition oBoundedProjPos(2), oBoundedNorm(2)
    oBoundedGeom(3).GetNormalFromPosition oBoundedProjPos(3), oBoundedNorm(3)
    oBoundedGeom(4).GetNormalFromPosition oBoundedProjPos(4), oBoundedNorm(4)
    
    ' -------------------------------------
    ' Get primary ports on bounding part
    ' -------------------------------------
    ' See notes above for bounded part
    Dim oBoundingPort(4) As IJPort
    
    Set oBoundingPort(1) = oTopologyLocate.GetSubPort(.Bounding, JXSEC_WEB_LEFT)
    Set oBoundingPort(2) = oTopologyLocate.GetSubPort(.Bounding, JXSEC_WEB_RIGHT)
    Set oBoundingPort(3) = oTopologyLocate.GetSubPort(.Bounding, JXSEC_TOP)
    Set oBoundingPort(4) = oTopologyLocate.GetSubPort(.Bounding, JXSEC_BOTTOM)
    
    Dim oBoundingGeom(4) As IJSurfaceBody
    
    Set oBoundingGeom(1) = oBoundingPort(1).Geometry
    Set oBoundingGeom(2) = oBoundingPort(2).Geometry
    Set oBoundingGeom(3) = oBoundingPort(3).Geometry
    Set oBoundingGeom(4) = oBoundingPort(4).Geometry
    
    Dim oBoundingProjPos(4) As IJDPosition
    
    oModelUtil.GetClosestPointOnBody oBoundingGeom(1), oBoundedPos, oBoundingProjPos(1), dist
    oModelUtil.GetClosestPointOnBody oBoundingGeom(2), oBoundedPos, oBoundingProjPos(2), dist
    oModelUtil.GetClosestPointOnBody oBoundingGeom(3), oBoundedPos, oBoundingProjPos(3), dist
    oModelUtil.GetClosestPointOnBody oBoundingGeom(4), oBoundedPos, oBoundingProjPos(4), dist
    
    Dim oBoundingNorm(4) As IJDVector
    
    oBoundingGeom(1).GetNormalFromPosition oBoundingProjPos(1), oBoundingNorm(1)
    oBoundingGeom(2).GetNormalFromPosition oBoundingProjPos(2), oBoundingNorm(2)
    oBoundingGeom(3).GetNormalFromPosition oBoundingProjPos(3), oBoundingNorm(3)
    oBoundingGeom(4).GetNormalFromPosition oBoundingProjPos(4), oBoundingNorm(4)
    
    ' ----------------------------------
    ' Find matching ports for connection
    ' ----------------------------------
    ' All ports are within 1mm of the bounded location
    ' The good ports should have normals closest to 180° apart (dot product = -1)
    Dim i As Long
    Dim j As Long
    Dim dDot As Double
    
    Dim oBoundedToBoundingPos As IJDPosition
    Dim oBoundingToBoundedPos As IJDPosition
    
    Dim oBoundedConnPort As IJPort
    Dim oBoundingConnPort As IJPort
    
    ' Loop through bounded ports
    For i = 1 To 4
        ' Loop through bounding ports
        For j = 1 To 4
            ' If port normals are opposite
            dDot = oBoundedNorm(i).Dot(oBoundingNorm(j))
            If dDot < 0 And (1 - (Abs(dDot)) < 0.001) Then
                ' See if they touch each other.  The projected positions may
                ' not be in the common area, so we need to do more work
                
                ' Project position on bounded to bounding geom
                oModelUtil.GetClosestPointOnBody oBoundingGeom(j), _
                                                 oBoundedProjPos(i), _
                                                 oBoundedToBoundingPos, _
                                                 dist
                
                ' If touching, should project back onto bounded
                oModelUtil.GetClosestPointOnBody oBoundedGeom(i), _
                                                 oBoundedToBoundingPos, _
                                                 oBoundingToBoundedPos, _
                                                 dist
                
                If dist < 0.001 Then
                    Set oBoundedConnPort = oBoundedPort(i)
                    Set oBoundingConnPort = oBoundingPort(j)
                    Exit For
                End If
            End If
        Next j
        
        If Not oBoundedConnPort Is Nothing Then
            Exit For
        End If
    Next i
    
    If oBoundedConnPort Is Nothing Or oBoundingConnPort Is Nothing Then
        ' May be in the process of changing,exit now.
        ' This method will be called later after port is update
        Exit Function
    End If
    ' -----------------------------
    ' Set the web cut as the parent
    ' -----------------------------
    Dim pSystemParent As IJSystemChild
    sError = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    ' --------------------------
    ' Create physical connection
    ' --------------------------
    Dim oPhysicalConnection As New PhysicalConn
    sError = "Creating Physical Connection"
    oPhysicalConnection.Create pResourceManager, _
                               oBoundedConnPort, _
                               oBoundingConnPort, _
                               "LapWeld", _
                               pSystemParent, _
                               ConnectionStandard
                               
    sError = "Setting Physical Connection to private variable"
    Set CAConstruct_LappedPhysConn = oPhysicalConnection.object
    
    Set pSystemParent = Nothing
    
    End With
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
        
End Function

'***********************************************************************
' METHOD:  CAConstruct_CornerFeature
'
' DESCRIPTION:  Add the root corner selector as a child item
'***********************************************************************
Public Function CAConstruct_CornerFeature(ByVal pMemberDescription As IJDMemberDescription, _
                                          ByVal pResourceManager As IUnknown) As Object

    sMETHOD = "CAConstruct_CornerFeature"
    
    Dim sFeatureSelection As String
    sFeatureSelection = GetAnswer_AddCornerFeature(pMemberDescription)
    
    ' Get the web cut
    Dim oEndCutObject As Object
    Set oEndCutObject = pMemberDescription.CAO
 
    ' Initialize WebCut wrapper class
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = oEndCutObject

    'Retrieve Ports to be used to create Corner Feature
    sError = "Retrieve Ports to be used to create WebCut Corner Feature"
    
    Dim bIsTopCorner As Boolean
    Dim oProfFacePort As IJPort
    Dim oProfEdgePort1 As IJPort
    Dim oProfEdgePort2 As IJPort
    Dim oCornerPosition As IJDPosition
    bIsTopCorner = False
'    oSDO_WebCut.CornerFeatureData bIsTopCorner, oCornerPosition, _
'                                  oProfFacePort, oProfEdgePort1, oProfEdgePort2
    CornerFeatureData bIsTopCorner, oCornerPosition, _
                      oProfFacePort, oProfEdgePort1, oProfEdgePort2, ByVal pMemberDescription

    ' Get the End cut, make it a parent of the Corner Feature
    sError = "Setting system parent to Member Description Custom Assembly"
    Dim oSystemParent As IJSystemChild
    Set oSystemParent = oEndCutObject
    
    ' Create Corner Feature
    sError = "Creating Corner Feature"
    Dim oSDO_CornerFeature As StructDetailObjects.CornerFeature
    Set oSDO_CornerFeature = New StructDetailObjects.CornerFeature
    oSDO_CornerFeature.Create pResourceManager, _
                               oProfFacePort, _
                               oProfEdgePort1, _
                               oProfEdgePort2, _
                               "RootCorner", _
                               oSystemParent
                               
    sError = "Returning CornerFeature just created"
    Set CAConstruct_CornerFeature = oSDO_CornerFeature.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
        
End Function

'***********************************************************************
' METHOD:  CMCornerFeature
'
' DESCRIPTION:  Add the root corner selector as a child item
'***********************************************************************
Public Sub CMCornerFeature(ByVal pMemberDescription As IJDMemberDescription, _
                           ByVal pResourceManager As IUnknown, _
                           ByRef pObject As Object)

  Set pObject = CAConstruct_CornerFeature(pMemberDescription, pResourceManager)

End Sub


'***********************************************************************
' METHOD:  CMCornerFeatureCon
'
' DESCRIPTION:  Determine whether or not to add a corner feature
'***********************************************************************
Public Sub CMCornerFeatureCon(ByRef pMD As IJDMemberDescription, _
                              ByRef bIsNeeded As Boolean)
    On Error GoTo ErrorHandler

    sMETHOD = "CMCornerFeatureCon"
    
    
    If ExcludeObjectBasedOnDetailedState(pMD.CAO, eObjectType.e_StiffenerEndCutCornerFeature) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    
    Dim sAddCorner As String
    sAddCorner = GetAnswer_AddCornerFeature(pMD)

    bIsNeeded = False
    
    If LCase(sAddCorner) = LCase("Yes") Then
        bIsNeeded = True
    End If

    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
    End Sub
    
 Public Sub CMLapCond(ByRef pMemberDescription As IJDMemberDescription, _
                     ByRef bIsNeeded As Boolean)
    
    bIsNeeded = False
    
    On Error GoTo ErrorHandler

    sMETHOD = "CMLapCond"
    sError = "Setting Conditional value"
    
    If ExcludeObjectBasedOnDetailedState(pMemberDescription.CAO, eObjectType.e_PhysicalConnection) Then
        bIsNeeded = False
        Exit Sub
    End If
    
    bIsNeeded = True
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
End Sub

'------------------------------------------------------------------------------------------------------------
' Procedure:
'     EndCutCornerFeatureData
'
' Description:
'   returns the following data required to create a Corner Feature:
'       Face Port that the Corner Feature is to be Located On
'       Edge Port that spefices the first side of the Corner Feature
'       Edge Port that spefices the other side of the Corner Feature
'       X,Y,Z postion of the Corner Feature
'
' Arguments:
'   oBoundedPart        - Profile Part Conrner Feature is applied to
'   oBoundedLocation    - approxiamate location of Corner Feature (EndCut Location)
'   eCornerFacePort     - Face Port Corner Feature is to be Located On
'   eFaceEdgePort       - Face Port to locate Corner Feature on Corner Face Port
'
' Usage Example:
'   for Corner Feature on WebCut at top corner
'       oBoundedPart is the Bounded Profile with the Web Cut
'       oBoundedLocation is the location of the Web Cut
'       eCornerFacePort is JXSEC_LEFT
'       eFaceEdgePort is JXSEC_TOP
'           retrieve the JXSEC_TOP Face Port from the Bounded Profile Part
'           project oBoundedLocation onto the JXSEC_TOP Face Port
'           retrieve the JXSEC_LEFT Face Port from the Bounded Profile Part
'           project JXSEC_TOP projected point onto JXSEC_LEFT face Port
'           call utility method ::FindClosestCorner
'               with following inputs:
'                   the IJModelBody from the oBoundedPart
'                   JXSEC_LEFT projected point as the oBoreLineOrigin argument
'                   JXSEC_LEFT projected normal vector as the oBoreLineDirection
'               this should return the data required to create the Corner Feature
'------------------------------------------------------------------------------------------------------------
Private Sub EndCutCornerFeatureData(oBoundedPart As Object, _
                                   oBoundedLocation As IJDPosition, _
                                   eCornerFacePort As JXSEC_CODE, _
                                   eFaceEdgePort As JXSEC_CODE, _
                                   oCornerPosition As IJDPosition, _
                                   oFacePort As IJPort, _
                                   oEdgePort1 As IJPort, _
                                   oEdgePort2 As IJPort)
sMETHOD = "EndCutCornerFeatureData"
    
    On Error GoTo ErrorHandler
    
    Dim sError As String
    
    Dim iIndex As Long
    Dim eFaceId As JXSEC_CODE
    
    Dim oBoreLineOrigin As IJDPosition
    Dim oBoreLineDirection As IJDVector
    
    Dim oFaceEdgeOrigin As IJDPosition
    Dim oFaceEdgeNormal As IJDVector
    
    Dim oPort As IJPort
    Dim oModelBody As IJModelBody
    
    Dim oTopologyLocate As GSCADStructGeomUtilities.TopologyLocate
    
    ' Verify given Bounded Part is valid object
    sError = "EndCut Bounded Part not initialized"
    If oBoundedPart Is Nothing Then
        LogError Err, MODULE, sMETHOD, sError
        Exit Sub
    End If
    
    ' get TopologyLocate helper class
    Set oTopologyLocate = New GSCADStructGeomUtilities.TopologyLocate
        
    ' Project the Bounded Location onto the given Face Edge Port
    sError = "EndCut Bounded FaceEdgePort not valid"
    Set oPort = oTopologyLocate.GetSubPort(oBoundedPart, eFaceEdgePort)
    If oPort Is Nothing Then
        LogError Err, MODULE, sMETHOD, sError
        Exit Sub
    End If
    
    oTopologyLocate.GetProjectedPointOnModelBody oPort.Geometry, _
                                                 oBoundedLocation, _
                                                 oFaceEdgeOrigin, _
                                                 oFaceEdgeNormal
    Set oPort = Nothing
    
    ' adjust the projected point by minimal amount in opposite direction
    ' so that the point will actually on the Corner Face Port
    oFaceEdgeNormal.Length = 1#
    oFaceEdgeOrigin.x = oFaceEdgeOrigin.x - (oFaceEdgeNormal.x * 0.0005)
    oFaceEdgeOrigin.y = oFaceEdgeOrigin.y - (oFaceEdgeNormal.y * 0.0005)
    oFaceEdgeOrigin.z = oFaceEdgeOrigin.z - (oFaceEdgeNormal.z * 0.0005)
    
    ' Project the oFaceEdgeOrigin onto the given Corner Face Port
    sError = "EndCut Bounded eCornerFacePort not valid"
    Set oPort = oTopologyLocate.GetSubPort(oBoundedPart, eCornerFacePort)
    If oPort Is Nothing Then
        LogError Err, MODULE, sMETHOD, sError
        Exit Sub
    End If
    
    oTopologyLocate.GetProjectedPointOnModelBody oPort.Geometry, _
                                                 oFaceEdgeOrigin, _
                                                 oBoreLineOrigin, _
                                                 oBoreLineDirection
    Set oPort = Nothing
    
    ' Reverse the Projection direction vector from the Corner Face Port
    ' for the BoreLineDirection
    oBoreLineDirection.x = -oBoreLineDirection.x
    oBoreLineDirection.y = -oBoreLineDirection.y
    oBoreLineDirection.z = -oBoreLineDirection.z
    
    ' Retrieve Corner Feature Data for this EndCut
    sError = "Error returned from::GetIJModelBodyFromObject"
    oTopologyLocate.GetIJModelBodyFromObject oBoundedPart, oModelBody
    
    sError = "Error returned from::FindClosestCorner"
    oTopologyLocate.FindClosestCorner oModelBody, _
                                      oBoreLineOrigin, oBoreLineDirection, _
                                      oCornerPosition, _
                                      oFacePort, oEdgePort1, oEdgePort2


Exit Sub
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number

End Sub


Public Function IsChamferExpected(pMD As IJDMemberDescription) As Boolean
    
sMETHOD = "IsChamferExpected"
    
    On Error GoTo ErrorHandler
    
    IsChamferExpected = False
    
    Dim sChamferType As String
    Dim oSmartParent As Object
    
    GetSmartOccurrenceParent pMD.CAO, oSmartParent
    'Retrieved the answer for Chamfer type question with out hard coding.
    GetSelectorAnswer oSmartParent, "ChamferType", sChamferType
    
    If Len(sChamferType) = 0 Or (sChamferType = "None") Then Exit Function
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    
    Dim oBoundedPart As Object
    
    Set oSDO_WebCut.object = pMD.CAO
    Set oBoundedPart = oSDO_WebCut.Bounded
    
    Dim oBounded As New StructDetailObjects.ProfilePart
    Set oBounded.object = oSDO_WebCut.Bounded
    
    If TypeOf oSmartParent Is IJFreeEndCut Then
        
        Dim oFreeEndCut_Object As Object
        Set oFreeEndCut_Object = oSmartParent
        
        Dim oFreeEndCut As GSCADFreeEndCut.IJFreeEndCut
        Set oFreeEndCut = oFreeEndCut_Object
        
        Dim oBoundingObject As Object
        Dim oBoundedObject As Object
        oFreeEndCut.get_FreeEndCutInputs oBoundedObject, oBoundingObject
        
        If TypeOf oBoundedObject Is IJPort Then
            Dim oPort As IJPort
            Set oPort = oBoundedObject
        Else
            Exit Function
        End If
        
        Dim oSDO_ProfilePart As New StructDetailObjects.ProfilePart
        Set oSDO_ProfilePart.object = oPort.Connectable
        
        If oBoundedPart Is oSDO_ProfilePart.object Then
            IsChamferExpected = True
        End If
        
    ElseIf TypeOf oSmartParent Is IJAssemblyConnection Then
        Dim oSDO_AssemblyConn As StructDetailObjects.AssemblyConn
        Set oSDO_AssemblyConn = New StructDetailObjects.AssemblyConn
        Set oSDO_AssemblyConn.object = oSmartParent
        
        Dim oPort1 As IJPort
        Dim oPort2 As IJPort
        
        Set oPort1 = oSDO_AssemblyConn.Port1
        Set oPort2 = oSDO_AssemblyConn.Port2
        
        Dim oProfilePart1 As StructDetailObjects.ProfilePart
        Set oProfilePart1 = New StructDetailObjects.ProfilePart
        
        Dim oProfilePart2 As StructDetailObjects.ProfilePart
        Set oProfilePart2 = New StructDetailObjects.ProfilePart
        
        If TypeOf oPort1.Connectable Is IJProfilePart And TypeOf oPort2.Connectable Is IJProfilePart Then
            Set oProfilePart1.object = oPort1.Connectable
            Set oProfilePart2.object = oPort2.Connectable
        Else
            Exit Function
        End If
        
        If oProfilePart1.WebThickness > oProfilePart2.WebThickness Then
            If oBoundedPart Is oProfilePart1.object Then
                IsChamferExpected = True
            End If
        ElseIf oProfilePart2.WebThickness > oProfilePart1.WebThickness Then
            If oBoundedPart Is oProfilePart2.object Then
                IsChamferExpected = True
            End If
        End If
    End If
       
    Exit Function
ErrorHandler:
  Err.Raise LogError(Err, MODULE, sMETHOD, sError).Number
  
End Function
