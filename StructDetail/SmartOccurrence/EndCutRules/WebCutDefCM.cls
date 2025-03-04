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

Const MODULE = "S:\StructDetail\Data\SmartOccurrence\EndCutRules\WebCutDefCM.cls"

Private sERROR As String
Private sMETHOD As String

Public Function GetBoundingProfilePort(ByVal pMemberDescription As IJDMemberDescription, _
                                    ByVal pResourceManager As IUnknown) As Long
                                    
    sMETHOD = "GetBoundingProfilePort"
    Dim dPortValue As Long
    
    ' Get Class
    Dim pWebCut As New StructDetailObjects.WebCut ' create new web cut wrapper class
    sERROR = "Setting web cut Inputs"
    Set pWebCut.object = pMemberDescription.CAO
    
    ' Get Bounding Part's Port
    Dim oBoundingPart As Object ' This is the Bounding part
    sERROR = "Getting Bounding object"
    Set oBoundingPart = pWebCut.Bounding ' Get Bounding part from wrapper
    
    Dim oPortBounded As IJPort
    Dim oPortBounding As IJPort
    Set oPortBounding = pWebCut.BoundingPort
    Set oPortBounded = pWebCut.BoundedPort
    
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
    sERROR = "Setting Webcut Inputs"
    Dim pWebCut As StructDetailObjects.WebCut
    Set pWebCut = New StructDetailObjects.WebCut
    Set pWebCut.object = pMemberDescription.CAO
    
    ' Get bounding Part's Port
    Dim oBoundingPort As IJPort
    sERROR = "Getting bounding object"
    Set oBoundingPort = phelper.GetEquivalentLastPort(pWebCut.BoundingPort)
zMsgBox "   ...GetEquivalentLastPort(pWebCut.BoundingPort)"
  
    ' Get bounded Part's Port
    Dim oBoundedPort As IJPort
    sERROR = "Setting Profile Part to bounded object"
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
    sERROR = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    ' Create physical connection
    Dim oPhysicalConnection As New PhysicalConn
    sERROR = "Creating Physical Connection"
    oPhysicalConnection.Create pResourceManager, _
                                            oBoundedPort, _
                                            oBoundingPort, _
                                            strStartClass, _
                                            pSystemParent, ConnectionStandard
zMsgBox "   ...oPhysicalConnection.Create"
                               
    sERROR = "Setting Physical Connection to private variable"
    Set CAConstruct_PhysConn = oPhysicalConnection.object
    
    Set oBoundingPort = Nothing
    Set oBoundedPort = Nothing
    Set oPhysicalConnection = Nothing
    Set pSystemParent = Nothing
    Set pWebCut = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
      Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    sERROR = "Setting Conditional value"
             
    bIsNeeded = True
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    sERROR = "Setting Conditional value"
         
    bIsNeeded = True
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    sERROR = "Setting Conditional value"
         
    bIsNeeded = True
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)

    Select Case sSectionType
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
    sERROR = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
        Case "B", "BUT", "BUTL2", "I", "ISType", "TSType", "T_XType"
            bIsNeeded = True
        Case Else
            bIsNeeded = False
    End Select
            
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    sERROR = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
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
    sERROR = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
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
    sERROR = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
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
    sERROR = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
    
    Select Case sSectionType
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
    sERROR = "Setting Conditional value"
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
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    sERROR = "Setting Conditional value"
    bIsNeeded = False
    
    Dim oSDO_WebCut As StructDetailObjects.WebCut
    Set oSDO_WebCut = New StructDetailObjects.WebCut
    Set oSDO_WebCut.object = pMD.CAO
    
    Dim sSectionType As String
    sSectionType = GetCrossSectionType(oSDO_WebCut.Bounding)
      
    Select Case sSectionType
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
    sERROR = "Setting Conditional value"
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
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    sERROR = "Setting Conditional value"
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
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    sERROR = "Setting Conditional value"
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
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    sERROR = "Setting Conditional value"
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
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    sERROR = "Setting Conditional value"
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
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    sERROR = "Setting Conditional value"
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
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    sERROR = "Setting Conditional value"
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
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    sERROR = "Setting Conditional value"
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
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    Dim pWebCut As StructDetailObjects.WebCut
    ' create new web cut wrapper class
    Set pWebCut = New StructDetailObjects.WebCut
    sERROR = "Setting web cut Inputs"
    Set pWebCut.object = pMemberDescription.CAO
    
    ' Get Bounding Part's Port
    Dim oBoundingPart As Object ' This is the Bounding part
    Dim oBoundingPort As IJPort
    sERROR = "Getting Bounding object"
    Set oBoundingPart = pWebCut.Bounding ' Get Bounding part from wrapper
    Dim phelper As StructDetailObjects.Helper
    Set phelper = New StructDetailObjects.Helper
    Dim oBounding As Object

    
    If TypeOf oBoundingPart Is IJPlate Then
        Dim pSDOHelper As New StructDetailObjects.Helper
        Dim oLastPlatePort As IJPort
        Set oLastPlatePort = pWebCut.BoundingPort
        Set oBoundingPort = pSDOHelper.GetEquivalentLastPort(oLastPlatePort)
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_STIFFENER Then
        Dim oProfilePart As New StructDetailObjects.ProfilePart
        sERROR = "Setting Profile Part wrapper object to penetrating object"
        Set oProfilePart.object = oBoundingPart
        sERROR = "Getting Bounding object's port"
        Set oBoundingPort = oProfilePart.SubPort(eBoundingSubPort)
    ElseIf phelper.ObjectType(oBoundingPart) = SDOBJECT_BEAM Then
        Dim oBeamPart As New StructDetailObjects.BeamPart
        sERROR = "Setting Beam Part wrapper object to penetrating object"
        Set oBeamPart.object = oBoundingPart
        sERROR = "Getting Bounding object's port"
        Set oBoundingPort = oBeamPart.SubPort(eBoundingSubPort)
    End If
   
    ' Get bounded profile Port
        Dim oBoundedPort As IJPort
    If phelper.ObjectType(pWebCut.Bounded) = SDOBJECT_STIFFENER Then
        Dim oBoundedPart As StructDetailObjects.ProfilePart
        Set oBoundedPart = New StructDetailObjects.ProfilePart
        Set oBoundedPart.object = pWebCut.Bounded
        Set oBoundedPort = oBoundedPart.CutoutSubPort(pWebCut.object, eBoundedSubPort)
    ElseIf phelper.ObjectType(pWebCut.Bounded) = SDOBJECT_BEAM Then
        Dim oBoundedBeamPart As StructDetailObjects.BeamPart
        Set oBoundedBeamPart = New StructDetailObjects.BeamPart
        Set oBoundedBeamPart.object = pWebCut.Bounded
        Set oBoundedPort = oBoundedBeamPart.CutoutSubPort(pWebCut.object, eBoundedSubPort)
    End If
        
    ' Get the web cut, make it a parent of the PC
    Dim pSystemParent As IJSystemChild ' this is the slot's ca
    sERROR = "Setting system parent to Member Description Custom Assembly"
    Set pSystemParent = pMemberDescription.CAO
       
    ' Create physical connection
    Dim oPhysicalConnection As New PhysicalConn
    sERROR = "Creating Physical Connection"
    oPhysicalConnection.Create pResourceManager, _
                               oBoundedPort, _
                               oBoundingPort, _
                               strStartClass, _
                               pSystemParent, ConnectionStandard
                               
    sERROR = "Setting Physical Connection to private variable"
    Set CAConstruct_PhysConnNames = oPhysicalConnection.object
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
        
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
    sERROR = "Setting Conditional value"
    bIsNeeded = True
    
zMsgBox "WebCutDefCM::ConstructPC1_Con ...bIsNeeded:" & bIsNeeded
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    sERROR = "Setting Conditional value"
    bIsNeeded = False
    
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
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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
    Set oSymbolDefinition = oDefinitions.Item("EndCutRules.EndToEndWebCutSel")
    
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
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
End Function

Public Function GetAnswer_EndToEndCase(ByRef pMD As IJDMemberDescription) As String
sMETHOD = "GetAnswer_EndToEndCase"
    GetAnswer_EndToEndCase = ""
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, sMETHOD, sERROR).Number
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


