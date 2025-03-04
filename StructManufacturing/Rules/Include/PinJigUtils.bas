Attribute VB_Name = "PinJigUtils"
'*******************************************************************
'  Copyright (C) 2006 Intergraph.  All rights reserved.
'
'  Project:
'
'  Abstract:    PinJigUtils.bas
'
'  History:
'     Anand     July 31, 2007  Creation
'******************************************************************

Option Explicit
Private Const MODULE As String = "PinJigUtils::"

Public m_dRemarkingSurfaceOffset  As Double

Public Function GetSystemParent(oDetailedPart As Object) As IJSystem
    Const METHOD As String = "GetSystemParent"
    On Error GoTo ErrorHandler
    
    Dim oPartSupport As IJPartSupport
    Set oPartSupport = New PartSupport
    
    Set oPartSupport.Part = oDetailedPart
    
    Dim oReturnObject As IJSystem
    oPartSupport.IsSystemDerivedPart oReturnObject
    
    Set GetSystemParent = oReturnObject
    
    Set oPartSupport = Nothing
    Set oReturnObject = Nothing
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, "Failed to get system parent")
End Function

Public Function GetPinJigSupportedPlateSystems(oPinJig As IJPinJig) As IJElements
    Const METHOD As String = "GetPinJigSupportedPlateSystems"
    On Error GoTo ErrorHandler

    Dim oRetColl As IJElements
    Set oRetColl = New JObjectCollection
    
    Dim oPlateObj As Object
    For Each oPlateObj In oPinJig.SupportedPlates
        oRetColl.Add GetSystemParent(oPlateObj)
    Next

    Set GetPinJigSupportedPlateSystems = oRetColl
    Set oRetColl = Nothing
    Set oPlateObj = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, "Failed to get Pin Jig's Supported Plate Systems")
End Function

Public Function GetReferencePlanesInPinJigRange(PartialPinJig As IJPinJig, oDirectionVec As IJDVector, ErrorCodeListNum As Long) As IJElements
    Const METHOD As String = "GetReferencePlanesInPinJigRange"
    On Error GoTo ErrorHandler

    Dim BOwithFS As IJDMfgFrameSystem
    Set BOwithFS = PartialPinJig
    
    Dim oPinJigFrameSys As IHFrameSystem
    Set oPinJigFrameSys = BOwithFS.FrameSysParent
    
    ' If no frame system was set on Pin jig (directly/indirectly), then nothing to do!
    If oPinJigFrameSys Is Nothing Then
        Set GetReferencePlanesInPinJigRange = Nothing
        Exit Function
    End If
    
    Dim oFramesColl As IJElements
    Set oFramesColl = GetReferencePlanesFromAllCSinRange(PartialPinJig, oDirectionVec)
    
    If oFramesColl Is Nothing Then
        Set GetReferencePlanesInPinJigRange = Nothing
        Exit Function
    End If
    
    ' Filter this collection to return only those from Pin Jig's FS.
    Dim oFrameSet As IJElements
    Set oFrameSet = New JObjectCollection

    Dim nIndex As Long
    For nIndex = 1 To oFramesColl.Count
    
        Dim oFrame As IHFrame
        Set oFrame = oFramesColl.Item(nIndex)
        
        Dim oFrameAxis As IHFrameAxis
        Set oFrameAxis = oFrame.FrameAxis
        
        Dim oFrameSystem As IHFrameSystem
        Set oFrameSystem = oFrameAxis.FrameSystem

        If oFrameSystem Is oPinJigFrameSys Then
            oFrameSet.Add oFramesColl.Item(nIndex)
        End If
        
        Set oFrame = Nothing
        Set oFrameAxis = Nothing
        Set oFrameSystem = Nothing
        
    Next nIndex

    ' Send the data back
    Set GetReferencePlanesInPinJigRange = oFrameSet
    
    Set oFrameSet = Nothing
    Set oPinJigFrameSys = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, "Failed to get reference planes in Pinjig range", "SMCustomWarningMessages", ErrorCodeListNum, , "RULES")
End Function

Public Sub GetConnectedObjects(oPlateObj As Object, _
                               TypeOfConnectedObject As StrMfgPartType, _
                               TypeOfConnection As ConnectionType, _
                               TeeButtLapOrEnd As ContourConnectionType, _
                               SideOfInputPlate As eUSER_CTX_FLAGS, _
                               oConnectedObjects As Collection, _
                               oConnections As Collection)

    Const METHOD As String = "GetConnectedObjects"
    On Error GoTo ErrorHandler

    Dim oPartSupport As IJPartSupport
    If TypeOf oPlateObj Is IJPlatePart Then
        Set oPartSupport = New PartSupport
        Set oPartSupport.Part = oPlateObj
    ElseIf TypeOf oPlateObj Is IJPlateSystem Then
        Dim oStructConnectable As IJStructConnectable
        Set oStructConnectable = oPlateObj
    Else
        Exit Sub
    End If
    
    ' oConnectedObjects is output collection
    Set oConnectedObjects = New Collection
    ' Input filters will be applied ConnectedObjColl to fill oConnectedObjects
    Dim ConnectedObjColl As Collection
    
    ' oConnections is output collection
    Set oConnections = New Collection
    ' Input filters will be applied ConnectionsColl to fill oConnections
    Dim ConnectionsColl As Collection
    
    If TypeOfConnection = ConnectionPhysical Then
        Dim ThisPortColl As Collection
        Dim OtherPortColl As Collection
        oPartSupport.GetConnectedObjects ConnectionPhysical, _
                                         ConnectedObjColl, ConnectionsColl, _
                                         ThisPortColl, OtherPortColl
    ElseIf TypeOfConnection = ConnectionLogical Then
        If oStructConnectable Is Nothing Then
            ' Get Plate system for input plate part
            Dim oPlateSystem As IJSystem
            oPartSupport.IsSystemDerivedPart oPlateSystem
            
            Set oStructConnectable = oPlateSystem
        End If

        'Get all objects connected to plate with a logical connection
        Dim ConnectedObjIEUK As IEnumUnknown
        Dim ConnectionsIEUK As IEnumUnknown
        oStructConnectable.enumConnectedObjects2 ConnectionLogical, ConnectedObjIEUK, ConnectionsIEUK

        Dim oConvertUtils As New CONVERTUTILITIESLib.CCollectionConversions
        Dim TmpConnectedObjColl As Collection
        oConvertUtils.CreateVBCollectionFromIEnumUnknown ConnectedObjIEUK, TmpConnectedObjColl

        Set oConvertUtils = Nothing
        Set ConnectedObjIEUK = Nothing
        Set ConnectionsIEUK = Nothing

        'Get all the connections from the connected objects
        Set ConnectedObjColl = New Collection
        Set ConnectionsColl = New Collection

        Dim oConnectable As IJConnectable
        Set oConnectable = oStructConnectable

        Dim iCnt As Integer
        For iCnt = 1 To TmpConnectedObjColl.Count
            Dim bIsConnected As Boolean
            Dim TmpConnections As IJElements

            oConnectable.isConnectedTo TmpConnectedObjColl.Item(iCnt), bIsConnected, TmpConnections

            'Populate ConnectedObjColl and ConnectionsColl to have a one to one mapping
            If (Not TmpConnections Is Nothing) Then
                Dim jCnt As Integer
                For jCnt = 1 To TmpConnections.Count
                     ConnectedObjColl.Add TmpConnectedObjColl.Item(iCnt)
                     ConnectionsColl.Add TmpConnections.Item(jCnt)
                Next jCnt
            End If

            Set TmpConnections = Nothing

        Next iCnt

        Set TmpConnectedObjColl = Nothing
        Set oConnectable = Nothing

    End If
    
    Dim OtherPortInTeeConn As IJElements
    Set OtherPortInTeeConn = New JObjectCollection
    
    Dim i As Integer
    ' For each connected object ...
    For i = 1 To ConnectedObjColl.Count
        Dim oThisStructPort As IJStructPort
        Dim eThisPortContext As eUSER_CTX_FLAGS
        
        Dim oAppCon As IJAppConnection
        Set oAppCon = ConnectionsColl.Item(i)
                    
        If TypeOfConnection = ConnectionPhysical Then
            Set oThisStructPort = ThisPortColl.Item(i)
            eThisPortContext = oThisStructPort.ContextID
            Set oThisStructPort = Nothing
            
        ' ... check if type of connected object is that of requested type
            If ((TypeOfConnectedObject = PLATE_TYPE And TypeOf ConnectedObjColl.Item(i) Is IJPlatePart) Or _
               (TypeOfConnectedObject = PROFILE_TYPE And TypeOf ConnectedObjColl.Item(i) Is IJProfilePart)) And _
               (eThisPortContext And SideOfInputPlate) <> 0 _
            Then
                Dim ConnType As ContourConnectionType
                Dim ThisPartCross As Boolean
                oPartSupport.GetConnectionTypeForContour oAppCon, ConnType, ThisPartCross
                
                ' Split PC causes problems with GetConnectionContour_Tee
                ' So check if same port has participated earlier
                If ConnType = PARTSUPPORT_CONNTYPE_TEE And Not ThisPartCross And _
                   Not OtherPortInTeeConn.Contains(OtherPortColl.Item(i)) And _
                   ConnType = TeeButtLapOrEnd Then
                   
                    OtherPortInTeeConn.Add OtherPortColl.Item(i)
                    oConnectedObjects.Add ConnectedObjColl.Item(i)
                    oConnections.Add ConnectionsColl.Item(i)
                    
                ElseIf ConnType = TeeButtLapOrEnd Then
                
                    oConnectedObjects.Add ConnectedObjColl.Item(i)
                    oConnections.Add ConnectionsColl.Item(i)
                    
                End If
                
                Set oAppCon = Nothing
            End If
        ElseIf TypeOfConnection = ConnectionLogical Then
            Dim oStructPort(1 To 2) As IJStructPort
            Dim oPort(1 To 2) As IJPort
            IdentifyPortsParticipatingInConnection oAppCon, _
                                                   oStructConnectable, _
                                                   ConnectedObjColl.Item(i), _
                                                   oStructPort, _
                                                   oPort
            eThisPortContext = oStructPort(1).ContextID
            Dim j As Integer
            For j = 1 To 2
                Set oStructPort(j) = Nothing
                Set oPort(j) = Nothing
            Next

        ' ... check if type of connected object is that of requested type
            If ((TypeOfConnectedObject = PLATE_TYPE And TypeOf ConnectedObjColl.Item(i) Is IJPlate) Or _
               (TypeOfConnectedObject = PROFILE_TYPE And TypeOf ConnectedObjColl.Item(i) Is IJProfile)) _
            Then
               
               If (eThisPortContext And SideOfInputPlate) <> 0 Then 'On the requested side
                    oConnectedObjects.Add ConnectedObjColl.Item(i)
                    oConnections.Add ConnectionsColl.Item(i)
               Else
                    'Might be a manually placed LC.
                    If SideOfInputPlate = CTX_INVALID And eThisPortContext = CTX_NOP Then
                        'Add if manually placed LC
                        If IsManualConnection(ConnectionsColl.Item(i)) Then
                            oConnectedObjects.Add ConnectedObjColl.Item(i)
                            oConnections.Add ConnectionsColl.Item(i)
                        End If 'end of manual conn check
                    End If 'end of side check
               End If
               
            End If
        End If ' end check for type of connected part
    Next ' end looping around connected objects
    
    Set oPartSupport = Nothing
    Set oPlateSystem = Nothing
    Set oStructConnectable = Nothing
    Set OtherPortInTeeConn = Nothing

    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, "Failed to get connected objects")
End Sub

Public Function GetReferencePlanesFromAllCSinRange(PartialPinJig As IJPinJig, oDirectionVec As IJDVector) As IJElements
    Const METHOD As String = "GetReferencePlanesFromAllCSinRange"
    On Error GoTo ErrorHandler
    
    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper
    
    Dim oFramesColl As New Collection
    oGeomHelper.GetReferencePlanesInRange PartialPinJig.RemarkingSurface, oDirectionVec, oFramesColl

    If oFramesColl Is Nothing Then
        Set GetReferencePlanesFromAllCSinRange = Nothing
        GoTo CleanUp
    End If
    
    Dim oFrameSet As IJElements
    Set oFrameSet = New JObjectCollection

    Dim nIndex As Long
    For nIndex = 1 To oFramesColl.Count
        oFrameSet.Add oFramesColl.Item(nIndex)
    Next nIndex
    
    Set GetReferencePlanesFromAllCSinRange = oFrameSet
    
CleanUp:
    Set oGeomHelper = Nothing
    Set oFrameSet = Nothing
    Set oFramesColl = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, "Failed to get reference planes from all Coordinate Systems in range")
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetPCPortionOnSurface
' Author    : Anand Hariharan
' Purpose   : Returns the geometry portion(s) of a Physical connection that are in contact with a surface
'---------------------------------------------------------------------------------------

Public Function GetPCPortionOnSurface(ByVal Connection As Object, _
                                      ByVal SurfaceBody As IJSurfaceBody) As IJElements
    On Error GoTo ErrorHandler
    Const METHOD As String = "GetPCPortionOnSurface"
    
    Dim oGeomOpsToolBox As IJDTopologyToolBox
    Set oGeomOpsToolBox = New DGeomOpsToolBox
    
    Dim oMfgGeomHelper As New MfgGeomHelper
    
    Dim oPhyConGeomColl As IJDObjectCollection
    
    If TypeOf Connection Is IJStructPhysicalConnection Then
        Dim oSDPhysicalConn As StructDetailObjects.PhysicalConn
        Set oSDPhysicalConn = New StructDetailObjects.PhysicalConn
        Set oSDPhysicalConn.object = Connection
    
        Set oPhyConGeomColl = oSDPhysicalConn.GetConnectionGeometries
    Else 'It is an FET
        Dim oFET As IJShpStrEdgeTreatment
        Set oFET = Connection
        
        Dim oFETGeom As IJWireBody
        Set oFETGeom = oFET.GetTreatmentGeometry
                
        Dim oFETGeomColl As IJElements
        On Error Resume Next
        oGeomOpsToolBox.ExplodeModelBodyByEdges Nothing, oFETGeom, oFETGeomColl
        On Error GoTo ErrorHandler
        
        If Not oFETGeomColl Is Nothing Then
            Set oPhyConGeomColl = New JObjectCollection
            Dim idX As Integer
            For idX = 1 To oFETGeomColl.Count
                oPhyConGeomColl.Add (oFETGeomColl.Item(idX))
            Next
        Else
            Set oPhyConGeomColl = oMfgGeomHelper.OptimizedMergingOfInputCurves(oFETGeom)
        End If
        
    End If
        
    Dim oTouchingComplexStringColl As IJElements
    Set oTouchingComplexStringColl = New JObjectCollection
    
    Dim oPhyConGeom As Object
    For Each oPhyConGeom In oPhyConGeomColl
        Dim oPhyConCurve As IJCurve
        Set oPhyConCurve = oPhyConGeom
        
        Dim dStartX As Double, dStartY As Double, dStartZ As Double
        Dim dEndX As Double, dEndY As Double, dEndZ As Double
        oPhyConCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
        
        Dim oPhyConPos As New DPosition
        oPhyConPos.Set dStartX, dStartY, dStartZ
        
        Dim oNormal As IJDVector
        Dim oProjectedPhyConPos As IJDPosition
        On Error Resume Next
        oGeomOpsToolBox.ProjectPointOnSurfaceBody SurfaceBody, _
                                                  oPhyConPos, _
                                                  oProjectedPhyConPos, _
                                                  oNormal, False
        Err.Clear
        On Error GoTo ErrorHandler
        
        Dim dDist As Double
        dDist = 1000
        If Not oProjectedPhyConPos Is Nothing Then dDist = oProjectedPhyConPos.DistPt(oPhyConPos)
        
        If dDist < 0.001 Then
            oPhyConPos.Set dEndX, dEndY, dEndZ
            On Error Resume Next
            oGeomOpsToolBox.ProjectPointOnSurfaceBody SurfaceBody, _
                                                      oPhyConPos, _
                                                      oProjectedPhyConPos, _
                                                      oNormal, False
            Err.Clear
            On Error GoTo ErrorHandler
            
            dDist = 1000
            If Not oProjectedPhyConPos Is Nothing Then dDist = oProjectedPhyConPos.DistPt(oPhyConPos)
        
            If dDist < 0.001 Then
                oTouchingComplexStringColl.Add oPhyConCurve
            End If
        End If
        
        Set oNormal = Nothing
        Set oProjectedPhyConPos = Nothing
        Set oPhyConPos = Nothing
        Set oPhyConCurve = Nothing
    Next
    
    Set GetPCPortionOnSurface = oMfgGeomHelper.OptimizedMergingOfInputCurves(oTouchingComplexStringColl)
    
CleanUp:
    Set oMfgGeomHelper = Nothing
    Set oTouchingComplexStringColl = Nothing
    Set oGeomOpsToolBox = Nothing
    Set oSDPhysicalConn = Nothing
    Set oPhyConGeom = Nothing
    Set oPhyConGeomColl = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
    GoTo CleanUp
End Function

Public Function LookUpUniqList(OprIdForConn As Long, OpnIdForConn As Long, UniqOprIds() As Long, UniqOpnIds() As Long) As Long
    Dim i As Long
    For i = LBound(UniqOpnIds) To UBound(UniqOpnIds)
        If OprIdForConn = UniqOprIds(i) And OpnIdForConn = UniqOpnIds(i) Then
            LookUpUniqList = i
            Exit Function
        End If
    Next
    LookUpUniqList = -1
End Function
Private Function IsManualConnection(oAppConn As IJAppConnection) As Boolean
        
    If TypeOf oAppConn Is IJStructLogicalConnection Then
        Dim oConnUtils As IJConnectionUtils
        Set oConnUtils = New CConnectionUtils
        IsManualConnection = oConnUtils.IsManualConnection(oAppConn)
    Else 'Physical Conn
        Dim oStructDetailConnUtils As New StructDetailConnectionUtil
        IsManualConnection = oStructDetailConnUtils.IsManualPhysicalConnection(oAppConn)
    End If
    
End Function
Private Function CheckIfFETIsInner(outerPars As Collection, oFETEdge As IJStructPort) As Boolean
    
    CheckIfFETIsInner = True
    
    Dim outerEntities As IJElements
    Set outerEntities = New JObjectCollection

    Dim idX As Integer
    If Not outerPars Is Nothing Then
        For idX = 1 To outerPars.Count
            outerEntities.Add outerPars.Item(idX)
        Next
    End If
    
    For idX = 1 To outerEntities.Count
        Dim parentObj As Object
        Set parentObj = outerEntities.Item(idX)
        
        If TypeOf parentObj Is IJPort Then 'Natural Edge
            Dim oPort As IJStructPort
            Set oPort = parentObj
            
            Dim oSysPort As IJPort
            Dim SDhelper As New StructDetailHelper
            Set oSysPort = SDhelper.GetEquivalentSystemPortEx(oFETEdge, JS_TOPOLOGY_FILTER_LCONNECT_SYS_LEDGES)
                        
            'Check if ports are equal
            Dim oSysStructPort As IJStructPort
            Set oSysStructPort = oSysPort
            If (oPort.ContextID = oSysStructPort.ContextID) And (oPort.OperatorID = oSysStructPort.OperatorID) And (oPort.SectionID = oSysStructPort.SectionID) Then
                CheckIfFETIsInner = False
            End If
        End If
    Next

End Function
'---------------------------------------------------------------------------------------
' Procedure : GetOperatorsFromConnections
' Author    : Anand Hariharan
'---------------------------------------------------------------------------------------

Public Function GetOperatorsFromConnections(oConnectionsColl As IJElements, OprIdsForConn() As Long, OpnIdsForConn() As Long, _
                                            UniqOprIds() As Long, UniqOpnIds() As Long) As IJElements
    Const METHOD As String = "GetOperatorsFromConnections "
    On Error GoTo ErrorHandler

    Dim NumConns As Long
    NumConns = oConnectionsColl.Count
    
    ReDim OprIdsForConn(1 To NumConns) As Long
    ReDim OpnIdsForConn(1 To NumConns) As Long
    
    ReDim UniqOprIds(1 To NumConns) As Long
    ReDim UniqOpnIds(1 To NumConns) As Long
    
    Dim ReturnOperatorCollection As IJElements
    Set ReturnOperatorCollection = New JObjectCollection
    
    Dim SDhelper As New StructDetailHelper
    Dim NumUniqOp As Long
    
    Dim i As Long
    For i = 1 To oConnectionsColl.Count
        If Not TypeOf oConnectionsColl.Item(i) Is IJAppConnection Then
        
            'The object can be a FET
            If TypeOf oConnectionsColl.Item(i) Is IJShpStrEdgeTreatment Then
                
                Dim oFET As IJShpStrEdgeTreatment
                Set oFET = oConnectionsColl.Item(i)
                
                Dim oFETPort As IJStructPort
                Set oFETPort = oFET.GetEdge()
                
                Dim FETTopoType As JS_TOPOLOGY_PROXY_TYPE
                Dim FETCtxFlag As eUSER_CTX_FLAGS
                Dim FETSecId As Long
                Dim FETOprId As Long
                Dim FETOpnId As Long
                
                oFETPort.GetAttributes FETTopoType, FETCtxFlag, FETOpnId, FETOprId, FETSecId
                
                OprIdsForConn(i) = FETOprId
                OpnIdsForConn(i) = FETOpnId
                
                If LookUpUniqList(OprIdsForConn(i), OpnIdsForConn(i), UniqOprIds, UniqOpnIds) <> -1 Then
                    GoTo NextConn
                End If
                
                ReturnOperatorCollection.Add oFET 'Add the FET itself as the remarking entity
                NumUniqOp = ReturnOperatorCollection.Count
                UniqOpnIds(NumUniqOp) = OpnIdsForConn(i)
                UniqOprIds(NumUniqOp) = OprIdsForConn(i)
             
            Else
                MsgBox MODULE & METHOD & "Item no. " & i & " of input collection of connections NOT a connection!"
            End If
            
            GoTo NextConn
        End If
        
        Dim oAppConn As IJAppConnection
        Set oAppConn = oConnectionsColl.Item(i)
        
        Dim PortsInConn As IJElements
        oAppConn.enumPorts PortsInConn
        
        Dim oDbgNI As IJNamedItem
        Set oDbgNI = oAppConn
        
        If PortsInConn.Count <> 2 Then
            MsgBox MODULE & METHOD & "Bad connection: '" & oDbgNI.Name & "' has " & PortsInConn.Count & " ports participating (expected 2)"
            GoTo NextConn
        End If
        
        Dim oStructPort(1 To 2) As IJStructPort
        Dim oPort(1 To 2) As IJPort
        Dim TopoType(1 To 2) As JS_TOPOLOGY_PROXY_TYPE
        Dim CtxFlag(1 To 2) As eUSER_CTX_FLAGS
        Dim SecId(1 To 2) As Long
        Dim OprId(1 To 2) As Long
        Dim OpnId(1 To 2) As Long
        
        Dim j As Integer
        For j = 1 To 2
            Set oStructPort(j) = PortsInConn.Item(j)
            oStructPort(j).GetAttributes TopoType(j), CtxFlag(j), OpnId(j), OprId(j), SecId(j)
            Set oPort(j) = oStructPort(j)
        Next
        
        'Check if the connection is manually placed.
        If IsManualConnection(oAppConn) Or (OpnId(1) <> OpnId(2) Or OprId(1) <> OprId(2)) Then
             OprIdsForConn(i) = OprId(1)
             OpnIdsForConn(i) = OpnId(1)
             
             If LookUpUniqList(OprIdsForConn(i), OpnIdsForConn(i), UniqOprIds, UniqOpnIds) <> -1 Then
                 GoTo NextConn
             End If
             
             ReturnOperatorCollection.Add oAppConn 'Add the manual conn itself as the remarking entity
             NumUniqOp = ReturnOperatorCollection.Count
             UniqOpnIds(NumUniqOp) = OpnIdsForConn(i)
             UniqOprIds(NumUniqOp) = OprIdsForConn(i)
             
             GoTo NextConn
        End If
        
'       If OpnId(1) <> OpnId(2) Or OprId(1) <> OprId(2) Then
'            MsgBox MODULE & METHOD & "Bad connection: '" & oDbgNI.Name & "' has different operators on participating ports!"'
'            GoTo NextConn
'        End If
        
        OprIdsForConn(i) = OprId(1)
        OpnIdsForConn(i) = OpnId(1)
        
        If LookUpUniqList(OprIdsForConn(i), OpnIdsForConn(i), UniqOprIds, UniqOpnIds) <> -1 Then
            GoTo NextConn
        End If

        If oPort(1).Connectable Is oPort(2).Connectable Then
           MsgBox MODULE & METHOD & "Bad connection: '" & oDbgNI.Name & "' has same entity on either side of connection!"
           GoTo NextConn
        End If
        
        Dim oOperation As IJStructOperation
        Dim oOperator As Object
        SDhelper.FindOperatorForOperationInGraphByID oPort(2).Connectable, OpnId(2), OprId(2), oOperation, oOperator
        
        ReturnOperatorCollection.Add oOperator
        NumUniqOp = ReturnOperatorCollection.Count
        UniqOpnIds(NumUniqOp) = OpnIdsForConn(i)
        UniqOprIds(NumUniqOp) = OprIdsForConn(i)

NextConn:
        Set oOperation = Nothing
        Set oOperator = Nothing
        For j = 1 To 2
            Set oStructPort(j) = Nothing
            Set oPort(j) = Nothing
        Next
        Set PortsInConn = Nothing
        Set oAppConn = Nothing
    Next
    
    NumUniqOp = ReturnOperatorCollection.Count
    ReDim Preserve UniqOpnIds(1 To NumUniqOp) As Long
    ReDim Preserve UniqOprIds(1 To NumUniqOp) As Long
    
    Set GetOperatorsFromConnections = ReturnOperatorCollection

CleanUp:
    Set oAppConn = Nothing
    Set SDhelper = Nothing
    Set ReturnOperatorCollection = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
    GoTo CleanUp
End Function

Public Function GetConnectionsBetweenPinJigPlates(LogOrPhy As ConnectionType, oPinJig As IJPinJig) As IJElements
    Const METHOD As String = "GetLogicalConnectionsBetweenPlates"
    On Error GoTo ErrorHandler
    
    Dim oPlateColl As IJElements
    
    If LogOrPhy = ConnectionPhysical Then
        Set oPlateColl = oPinJig.SupportedPlates
    ElseIf LogOrPhy = ConnectionLogical Then
        Set oPlateColl = GetPinJigSupportedPlateSystems(oPinJig)
    Else
        Exit Function
    End If
    
    Dim oUniqueConns As IJElements
    Set oUniqueConns = New JObjectCollection
    
    ' For each supported plate ...
    Dim Iter As Long
    For Iter = 1 To oPlateColl.Count
        Dim oConnectedObjects As Collection
        Dim oConnections As Collection
        GetConnectedObjects oPlateColl.Item(Iter), _
                            PLATE_TYPE, _
                            LogOrPhy, _
                            PARTSUPPORT_CONNTYPE_BUTT, _
                            CTX_INVALID, _
                            oConnectedObjects, oConnections
        
        Dim i As Long
        ' For each connected object ...
        For i = 1 To oConnectedObjects.Count
            ' ... check if this connected object is a supported plate (system)
            If oPlateColl.Contains(oConnectedObjects.Item(i)) Then
                oUniqueConns.Add oConnections.Item(i)
            End If
        Next
    Next
    
    Dim oPlatePartColl As IJElements
    Set oPlatePartColl = oPinJig.SupportedPlates
    
    For Iter = 1 To oPlatePartColl.Count
        Dim oPlatePartSupport As IJPlatePartSupport
        Set oPlatePartSupport = New PlatePartSupport
        
        Dim oPartSupport As IJPartSupport
        Set oPartSupport = oPlatePartSupport
        
        Set oPartSupport.Part = oPlatePartColl.Item(Iter)
        
        Dim oFETColl As Collection
        oPartSupport.GetFET Nothing, oFETColl
        
        If Not oFETColl Is Nothing Then
            If oFETColl.Count > 0 Then
                Dim oMfgGeomUtil As IJMfgGeomUtil
                Set oMfgGeomUtil = New MfgUtilSurface
                
                Dim outerPars As Collection, outerWires As Collection
                Dim outerNames() As String
                oMfgGeomUtil.GetEntitiesCreatingOuterEdgesInPlateCollection oPlatePartColl, oPinJig.RemarkingSurface, outerPars, outerNames, outerWires
            
                Dim idX As Integer
                For idX = 1 To oFETColl.Count
                    Dim oFET As IJShpStrEdgeTreatment
                    Set oFET = oFETColl.Item(idX)
                    
                    If CheckIfFETIsInner(outerPars, oFET.GetEdge()) Then
                        oUniqueConns.Add oFET
                    End If
                Next idX
            End If
        End If
    Next
        
    Set GetConnectionsBetweenPinJigPlates = oUniqueConns
    Set oPlateColl = Nothing
    Set oUniqueConns = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
End Function
                                       
Public Sub IdentifyPortsParticipatingInConnection(oConnection As IJAppConnection, _
                                                  oConnectedObject1 As Object, _
                                                  oConnectedObject2 As Object, _
                                                  oConnStructPort() As IJStructPort, _
                                                  oConnIJPort() As IJPort)
                                                  
    Const METHOD As String = "IdentifyPortsParticipatingInConnection "
    On Error GoTo ErrorHandler
        
    Dim oPortsInConn As IJElements
    oConnection.enumPorts oPortsInConn
    
    Dim oDbgNI As IJNamedItem
    Set oDbgNI = oConnection

    If oPortsInConn.Count <> 2 Then
        MsgBox MODULE & METHOD & "Bad connection: '" & oDbgNI.Name & "' has " & oPortsInConn.Count & " ports participating (expected 2)"
        Exit Sub
    End If
    
    Dim oStructPort(1 To 2) As IJStructPort
    Dim oPort(1 To 2) As IJPort

    Dim j As Integer
    For j = 1 To 2
        Set oStructPort(j) = oPortsInConn.Item(j)
        Set oPort(j) = oStructPort(j)
    Next
        
    If oPort(1).Connectable Is oConnectedObject1 And _
       oPort(2).Connectable Is oConnectedObject2 Then
       
       Set oConnStructPort(1) = oStructPort(1)
       Set oConnStructPort(2) = oStructPort(2)
       Set oConnIJPort(1) = oPort(1)
       Set oConnIJPort(2) = oPort(2)
       
    ElseIf oPort(1).Connectable Is oConnectedObject2 And _
           oPort(2).Connectable Is oConnectedObject1 Then
           
       Set oConnStructPort(1) = oStructPort(2)
       Set oConnStructPort(2) = oStructPort(1)
       Set oConnIJPort(1) = oPort(2)
       Set oConnIJPort(2) = oPort(1)
    
    Else
        MsgBox MODULE & METHOD & "Connection '" & oDbgNI.Name & "' is between unexpected entities"
    End If
    
    Set oDbgNI = Nothing
    
    For j = 1 To 2
        Set oStructPort(j) = Nothing
        Set oPort(j) = Nothing
    Next
    Set oPortsInConn = Nothing
    
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
End Sub
                                       
Public Function LogConnIsTeeTypeConn(oLogicalConnection As IJAppConnection, _
                                     oConnectedObject As Object, _
                                     oSupportedPlateSystem As IJPlateSystem) As Boolean
    Const METHOD As String = "LogConnIsTeeTypeConn "
    On Error GoTo ErrorHandler
        
    LogConnIsTeeTypeConn = False
    
    Dim oStructPort(1 To 2) As IJStructPort
    Dim oPort(1 To 2) As IJPort
    IdentifyPortsParticipatingInConnection oLogicalConnection, _
                                           oConnectedObject, oSupportedPlateSystem, _
                                           oStructPort, oPort
        
    If oStructPort(1).ProxyType = JS_TOPOLOGY_PROXY_LEDGE And _
       oStructPort(2).ProxyType = JS_TOPOLOGY_PROXY_LFACE Then
        LogConnIsTeeTypeConn = True
    End If
    
    Dim j As Integer
    For j = 1 To 2
        Set oStructPort(j) = Nothing
        Set oPort(j) = Nothing
    Next
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
End Function

Public Function GetSidesOfPlatesFacingPlane(oPinJig As IJPinJig, _
                                            ConnType As ConnectionType, _
                                            Optional ByVal oPlateColl As IJElements) As eUSER_CTX_FLAGS()

    Const METHOD As String = "GetSidesOfPlatesFacingPlane "
    On Error GoTo ErrorHandler

    If oPlateColl Is Nothing Then Set oPlateColl = oPinJig.SupportedPlates
    
    Dim RootX As Double, RootY As Double, RootZ As Double
    Dim NormX As Double, NormY As Double, NormZ As Double
    oPinJig.GetBasePlane NormX, NormY, NormZ, RootX, RootY, RootZ
    
    Dim oPlane As IJPlane
    Set oPlane = New Plane3d
    
    oPlane.DefineByPointNormal RootX, RootY, RootZ, NormX, NormY, NormZ
    
    Dim oSurfaceUtil As IJMfgUtilSurface
    Set oSurfaceUtil = New MfgUtilSurface
    
    Dim PlateSide() As Long
    PlateSide = oSurfaceUtil.GetRemarkingSidesOfPlates(oPlateColl, oPlane)
    
    ReDim RetArray(1 To UBound(PlateSide) - LBound(PlateSide) + 1) As eUSER_CTX_FLAGS
    
    Dim i As Long
    For i = LBound(PlateSide) To UBound(PlateSide)
        If ConnType = ConnectionLogical Then
            If PlateSide(i) = CTX_BASE Then
                RetArray(i) = CTX_NMINUS
            ElseIf PlateSide(i) = CTX_OFFSET Then
                RetArray(i) = CTX_NPLUS
            Else
                MsgBox MODULE & METHOD & "Encountered neither base nor offset!"
                RetArray(i) = CTX_INVALID
            End If
        ElseIf ConnType = ConnectionPhysical Then
            If PlateSide(i) = CTX_BASE Or PlateSide(i) = CTX_OFFSET Then
                RetArray(i) = PlateSide(i)
            Else
                MsgBox MODULE & METHOD & "Encountered neither base nor offset!"
                RetArray(i) = CTX_INVALID
            End If
        Else
            MsgBox MODULE & METHOD & "Wrong input sent to RetArray!"
            RetArray(i) = CTX_INVALID
        End If
    Next
    
    GetSidesOfPlatesFacingPlane = RetArray
    
    Set oPlane = Nothing
    Set oSurfaceUtil = Nothing
    Set oPlateColl = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
End Function

Public Function GetCustomAttributeOfInputObject(oInputObject As Object, CustomInterfaceName As String, AttributeName As String) As IJDAttribute
    Const METHOD As String = "GetMarkingLineCustomAttribute"
    On Error GoTo ErrorHandler

    Dim oMetaDataHelp As IJDAttributeMetaData
    Set oMetaDataHelp = oInputObject
    
    If oMetaDataHelp Is Nothing Then Exit Function
    
    Dim CustomInterface As Variant
    CustomInterface = oMetaDataHelp.IID(CustomInterfaceName)

    Set oMetaDataHelp = Nothing
    If vbEmpty = VarType(CustomInterface) Or vbNull = VarType(CustomInterface) Then Exit Function

    Dim oAttributes As IJDAttributes
    Set oAttributes = oInputObject
    
    If oAttributes Is Nothing Then Exit Function
    
    Dim oAttributesCol As IJDAttributesCol
    Set oAttributesCol = oAttributes.CollectionOfAttributes(CustomInterface)

    Set oAttributes = Nothing
    If oAttributesCol Is Nothing Then Exit Function
    
    Dim oAttribute As IJDAttribute
    Set oAttribute = oAttributesCol.Item(AttributeName)
    
    Set oAttributesCol = Nothing
    Set GetCustomAttributeOfInputObject = oAttribute
    Set oAttribute = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
End Function


Public Function GetMarkingLineCustomAttribute(oMarkingLineAE As IJMfgMarkingLines_AE, AttributeName As String) As Variant
    Const METHOD As String = "GetMarkingLineCustomAttribute"
    On Error GoTo ErrorHandler
    
    Dim oAttribute As IJDAttribute
    Set oAttribute = GetCustomAttributeOfInputObject(oMarkingLineAE, _
                                                     "IJMfgSketchLocation", _
                                                     AttributeName)

    If Not oAttribute Is Nothing Then GetMarkingLineCustomAttribute = oAttribute.Value
    Set oAttribute = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")
End Function
Public Sub GetAPSMarkingLines(oPinJig As IJPinJig, lMarkingType As Long, ByRef oMarkingLines As IJElements)
    Const METHOD As String = "GetAPSMarkingLines"
    On Error GoTo ErrorHandler
    
    Dim oAPSMarkingLines As IJElements
    Set oAPSMarkingLines = oPinJig.MarkingLinesOnSupportedPlates(False, lMarkingType, PinJigRemarkingSide)

    If Not oAPSMarkingLines Is Nothing Then
        If oAPSMarkingLines.Count > 0 Then
            oMarkingLines.AddElements oAPSMarkingLines
        End If
        Set oAPSMarkingLines = Nothing
    End If
         
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")

End Sub

Public Sub CreateGeom3dFromAPSMarkingLines(oPinJig As IJPinJig, ElemsToRemark As IJElements, oMarkingLines As IJElements, lGeomSubType As Long, ByRef oGeomCol3d As MfgGeomCol3d)
    Const METHOD As String = "CreateGeom3dFromAPSMarkingLines"
    On Error GoTo ErrorHandler
    
    Dim oRemarkingSurface As IJSurfaceBody
    If (m_dRemarkingSurfaceOffset > 0.001) Then
        Set oRemarkingSurface = oPinJig.RemarkingSurface
    End If
  
    Dim oMfgMGHelper As New MfgMGHelper
    
    Dim oMarkLine As IJMfgMarkingLines_AE
    For Each oMarkLine In oMarkingLines
        If ElemsToRemark.Contains(oMarkLine) Then
            Dim ConnectedPartName As String
            ConnectedPartName = GetMarkingLineCustomAttribute(oMarkLine, "RelatedPartName")
           
            Dim oCSColl     As IJDObjectCollection
            Dim oProjCS     As IJComplexString
            Dim oCS         As IJComplexString
            
            Set oCSColl = oMarkLine.GeometryAsComplexStrings
            
            For Each oCS In oCSColl
        
                ' remarking surface is obtained only when the offset is > 0.001
                If (Not oRemarkingSurface Is Nothing) And m_dRemarkingSurfaceOffset > 0.001 Then
                    
                    On Error Resume Next
                    
                    'Check if the related part is a plate/profile.If so use its orientation for projection
                    Dim oRelatedObj As Object
                    Set oRelatedObj = oMarkLine.GetMfgMarkingRelatedObject
                    
                    Dim oProjVec As IJDVector
                    
                    If Not oRelatedObj Is Nothing Then
                        If TypeOf oRelatedObj Is IJPlate Then
                            Set oProjVec = GetPlateOrientationVector(oCS, oRelatedObj)
                        ElseIf TypeOf oRelatedObj Is IJProfile Then
                            Set oProjVec = GetProfileOrientationVector(oCS, oRelatedObj)
                        End If
                    End If
                    
                    oMfgMGHelper.ProjectComplexStringToSurface oCS, oRemarkingSurface, oProjVec, oProjCS
                    On Error GoTo ErrorHandler
                    
                    If oProjCS Is Nothing Then
                        StrMfgLogError Err, MODULE, METHOD, "Failed to project marking line onto remarking surface", , , , "RULES"
                        
                        Set oProjCS = oCS
                    End If
                    
                Else
                    Set oProjCS = oCS
                End If
                
                CreateMfgGeom3dObject oProjCS, STRMFG_PinJigRemarkingLine3D, _
                                  oMarkLine, _
                                  oGeomCol3d, ConnectedPartName, _
                                  lGeomSubType
                                  
                Set oCS = Nothing
                Set oProjCS = Nothing
                
            Next ' Next oCS
                        
        End If
        
    Next ' Next oMarkLine
    
    Set oMarkLine = Nothing
           
    Exit Sub

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , , , , "RULES")

End Sub

'---------------------------------------------------------------------------------------
' Procedure : CreateMfgGeom3d
' Purpose   :
' Inputs    :
    ' Geometry as IJComplexString
    ' Geometry Type as Long (ideally StrMfgGeometryType)
    ' Original Object as Object (will be stored as Moniker)
    ' Geometry Collection Holder as IJMfgGeomCol3d
    ' Name of this object as String (if empty, will be name of parent object)
    ' Geometry sub-type (not normally required)
'---------------------------------------------------------------------------------------

Public Function CreateMfgGeom3dObject(oCS As IJComplexString, _
                                      lGeomType As Long, _
                                      oParentObject As Object, _
                                      Optional oGeomCol3d As IJMfgGeomCol3d, _
                                      Optional bstrName As String, _
                                      Optional lGeomSubType As Long) As IJMfgGeom3d
    
    Const METHOD As String = "CreateMfgGeom3d"
    On Error GoTo ErrorHandler

    Dim oGeom3dFactory As GSCADMfgGeometry.MfgGeom3dFactory
    Set oGeom3dFactory = New GSCADMfgGeometry.MfgGeom3dFactory
    
    ' Create the object from factory
    Dim oGeom3d As IJMfgGeom3d
    Set oGeom3d = oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    ' Stuff it with attributes
    oGeom3d.PutGeometry oCS
    oGeom3d.PutGeometrytype lGeomType
    oGeom3d.PutSubGeometryType lGeomSubType
    
    If Not oGeomCol3d Is Nothing Then
        ' Make a relation between newly created object and input parent object
        oGeomCol3d.AddGeometry oGeomCol3d.Getcount + 1, oGeom3d
    End If
    
    ' Set its name
    If TypeOf oGeom3d Is IJNamedItem Then
        Dim oName As IJNamedItem
        Set oName = oGeom3d
        oName.Name = bstrName
        Set oName = Nothing
    End If
    
    Dim oObjUtil As IJDMfgGeomUtilWrapper
    Set oObjUtil = New MfgGeomUtilWrapper
    
    If Not oParentObject Is Nothing Then
        If TypeOf oParentObject Is IJDObject Then
            oGeom3d.PutMoniker oObjUtil.GetMoniker(oParentObject)
        End If
    End If
    
    Set CreateMfgGeom3dObject = oGeom3d
    
    Set oGeom3dFactory = Nothing
    Set oGeom3d = Nothing
    Set oObjUtil = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5005, , "RULES")
End Function



Public Function CreateMfgGeomCol3dObject() As IJMfgGeomCol3d
    Const METHOD As String = "CreateMfgGeomCol3dObject"
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3dFactory    As GSCADMfgGeometry.MfgGeomCol3dFactory
    Set oGeomCol3dFactory = New GSCADMfgGeometry.MfgGeomCol3dFactory
    
    Set CreateMfgGeomCol3dObject = oGeomCol3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    Set oGeomCol3dFactory = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub CreateFrameRemarkingLines(PartialPinJig As Object, ElemsToRemark As IJElements, eSubGeomType As StrMfgGeometryType, ReturnColl As IJMfgGeomCol3d)
    Const METHOD As String = "CreateFrameRemarkingLines"
    On Error GoTo ErrorHandler
    
    Set ReturnColl = CreateMfgGeomCol3dObject
        
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig
    
    Dim oFramecoll As IJElements
    Set oFramecoll = New JObjectCollection
    
    Dim i As Long
    For i = 1 To ElemsToRemark.Count
        If TypeOf ElemsToRemark.Item(i) Is IHFrame Then oFramecoll.Add ElemsToRemark.Item(i)
    Next
    
    If oFramecoll Is Nothing Then
        Exit Sub
    End If
    
    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New MfgMGHelper
    
    Dim oRemarkingSurface As IJSurfaceBody
    Set oRemarkingSurface = oPinJig.RemarkingSurface
    
    Dim oFrame As IHFrame
    For Each oFrame In oFramecoll
    
        Dim RefPlane_RootPoint As IJDPosition
        Dim RefPlane_Normal As IJDVector
        oFrame.GetPositionVector RefPlane_RootPoint, RefPlane_Normal
        
        Dim oCSelems As IJElements
        Set oCSelems = oMfgMGHelper.IntersectSurfaceBodyWithInfinitePlane(oRemarkingSurface, RefPlane_RootPoint, RefPlane_Normal)
        
        Set RefPlane_Normal = Nothing
        Set RefPlane_RootPoint = Nothing
        
        Dim oNI As IJNamedItem
        Set oNI = oFrame
     
        If oCSelems.Count > 1 Then
            
            Dim lCodelistNum As Long
            Select Case eSubGeomType
                Case STRMFG_PinJig_Remark_FrameX
                    lCodelistNum = 5013
                Case STRMFG_PinJig_Remark_FrameY
                    lCodelistNum = 5015
                Case STRMFG_PinJig_Remark_FrameZ
                    lCodelistNum = 5007
            End Select
        
            LogMessage Err, MODULE, METHOD, oNI.Name & _
                       " & Remarking surface intersection returned multiple outputs!", _
                       "SMCustomWarningMessages", lCodelistNum, oNI.Name
        End If
        
        Dim oCS As IJComplexString
        For Each oCS In oCSelems

            CreateMfgGeom3dObject oCS, STRMFG_PinJigRemarkingLine3D, oFrame, _
                                  ReturnColl, oNI.Name, eSubGeomType

            Set oCS = Nothing

        Next
            
        Set oNI = Nothing
        Set oCSelems = Nothing
        Set oFrame = Nothing
    Next
    
    Set oRemarkingSurface = Nothing
    Set oFramecoll = Nothing
    Set oMfgMGHelper = Nothing
    Set oPinJig = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Function IsRefCurveUsedForSplit(oPlateSystem As IJSystem, oRefCurve As IJRefCurveOnSurface) As Boolean
    Const METHOD As String = "IsRefCurveUsedForSplit"
    On Error GoTo ErrorHandler
    
    IsRefCurveUsedForSplit = False
    
    Dim oSDhelper As New StructDetailHelper
                        
    Dim oRootSys   As IUnknown
    Dim oSplitters As IEnumUnknown
    oSDhelper.IsResultOfSplitWithOpr oPlateSystem, oRootSys, oSplitters
    
    If oSplitters Is Nothing Then
        Exit Function
    End If
    
    Dim oCollectionOfSplitters  As Collection
    Dim ConvertUtils            As CCollectionConversions
    Dim SplitterColl            As IJElements
    
    Set ConvertUtils = New CCollectionConversions
    ConvertUtils.CreateVBCollectionFromIEnumUnknown oSplitters, oCollectionOfSplitters
    Set SplitterColl = ConvertUtils.CreateIJElementsCollectionFromVBCollection(oCollectionOfSplitters)
    
    If Not SplitterColl Is Nothing Then
         If SplitterColl.Contains(oRefCurve) Then 'Reference curve has been used for split
           IsRefCurveUsedForSplit = True
        End If
    End If
   
 Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetRefCurveGeomDataFromPinJigSupportedPlates(oPinJig As IJPinJig) As Collection
    Const METHOD = "GetRefCurveGeomDataFromPinJigSupportedPlates"
    On Error GoTo ErrorHandler
    
    Dim ReturnColl As New Collection

    Dim oPlatePartSupport As IJPlatePartSupport
    Set oPlatePartSupport = New PlatePartSupport
    
    Dim PartSupport As IJPartSupport
    Set PartSupport = oPlatePartSupport
    
    Dim oPlateColl As IJElements
    Set oPlateColl = oPinJig.SupportedPlates

    Dim RemarkingSideOfPlate() As eUSER_CTX_FLAGS
    RemarkingSideOfPlate = GetSidesOfPlatesFacingPlane(oPinJig, ConnectionPhysical, oPlateColl)
    
    Dim i As Long
    For i = 1 To oPlateColl.Count
        Dim oPlate As IJPlate
        Set oPlate = oPlateColl.Item(i)
        
        If oPlate.plateType = Hull Then
            Set PartSupport.Part = oPlate
            
            Dim oPlateSystem As IJSystem
            PartSupport.IsSystemDerivedPart oPlateSystem
            
            Dim WhichSide As PlateThicknessSide
            If RemarkingSideOfPlate(i) = CTX_BASE Then
                WhichSide = PlateBaseSide
            ElseIf RemarkingSideOfPlate(i) = CTX_OFFSET Then
                WhichSide = PlateOffsetSide
            Else
                WhichSide = PlateSideUnspecified
            End If
            
            Dim RefCurveColl As Collection
            On Error Resume Next
            oPlatePartSupport.GetReferenceCurvesOnSurface WhichSide, RefCurveColl
            If Err.Number <> 0 Then
                StrMfgLogError Err, MODULE, METHOD, "Failed to get reference curves on surface", "SMCustomWarningMessages", 5017, , "RULES"
            End If
            On Error GoTo ErrorHandler
            
            If Not RefCurveColl Is Nothing Then
                Dim oRefCurveData As IJRefCurveData
                For Each oRefCurveData In RefCurveColl
                    'Need to restrict the types supported(We do not want to support GRID and MARK referenc curves in MFG)
                    If oRefCurveData.Type = JSRCOS_REFERENCE Or _
                       oRefCurveData.Type = JSRCOS_TANGENT Or _
                       oRefCurveData.Type = JSRCOS_KNUCKLE Or _
                       oRefCurveData.Type = JSRCOS_UNKNOWN Then
                       
                        'Check if the reference curve has been used for split.
                        'If so, do not add the curve.
                        If Not IsRefCurveUsedForSplit(oPlateSystem, oRefCurveData.ParentReferenceCurve) Then
                            ReturnColl.Add oRefCurveData
                        End If
                                                
                    End If
                Next
            End If
        End If
    Next
    
    Set oPlatePartSupport = Nothing
    Set PartSupport = Nothing
    Set oPlateColl = Nothing
    
    Set GetRefCurveGeomDataFromPinJigSupportedPlates = ReturnColl
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5017, , "RULES")
End Function

Public Sub CreateNavalArchMarks(PartialPinJig As Object, ElemsToRemark As IJElements, ReturnColl As IJMfgGeomCol3d)
    Const METHOD = "CreateNavalArchMarks"
    On Error GoTo ErrorHandler
    
    Set ReturnColl = CreateMfgGeomCol3dObject
    
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig
    
    Dim oMfgMGHelper As New MfgMGHelper

    Dim oRefCurveColl As Collection
    Set oRefCurveColl = GetRefCurveGeomDataFromPinJigSupportedPlates(oPinJig)
    
    Dim RefCurve As IJRefCurveData
    For Each RefCurve In oRefCurveColl
        Dim oActualReferenceCurve As IJRefCurveOnSurface
        Set oActualReferenceCurve = RefCurve.ParentReferenceCurve
        
        If ElemsToRemark.Contains(oActualReferenceCurve) Then
            Dim RemarkLineName As String
            RemarkLineName = RefCurve.Name
            
            Dim MLcoll As Collection
            Set MLcoll = RefCurve.GetMarkingLineCollection
            
            Dim i As Long
            Dim MarkLine As IJWireBody
            'For Each MarkLine In MLcoll
            For i = 1 To MLcoll.Count
                Set MarkLine = MLcoll.Item(i)
                Dim oCSelems As IJElements
                oMfgMGHelper.WireBodyToComplexStrings MarkLine, oCSelems
                
                If oCSelems.Count > 1 Then
                    StrMfgLogError Err, MODULE, METHOD, _
                               "Multiple broken segments for " & RemarkLineName, _
                               "SMCustomWarningMessages", 5017, RemarkLineName
                End If
                
                Dim oCS As IJComplexString
                For Each oCS In oCSelems
                    CreateMfgGeom3dObject oCS, STRMFG_PinJigRemarkingLine3D, _
                                          oActualReferenceCurve, _
                                          ReturnColl, RemarkLineName, _
                                          STRMFG_PinJig_Remarking_NavalArch
    
                    Set oCS = Nothing
                Next
                    
                Set oCSelems = Nothing
                Set MarkLine = Nothing
            Next
            
            Set MLcoll = Nothing
            Set RefCurve = Nothing
        End If
    Next
    
    Dim oNavalArchMarks As IJElements
    Set oNavalArchMarks = New JObjectCollection
     
    GetAPSMarkingLines oPinJig, STRMFG_NAVALARCHLINE, oNavalArchMarks

    If oNavalArchMarks.Count > 0 Then
        CreateGeom3dFromAPSMarkingLines oPinJig, ElemsToRemark, oNavalArchMarks, STRMFG_PinJig_Remarking_NavalArch, ReturnColl
    End If
    Set oNavalArchMarks = Nothing
    
    Exit Sub

ErrorHandler:
     Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5017, , "RULES")
End Sub
Public Function GetPlateOrientationVector(oCS As IJComplexString, oPlate As IJPlate) As IJDVector
    Const METHOD = "GetPlateOrientationVector"
    On Error GoTo ErrorHandler
        
    Dim oIntPos As IJDPosition
    Set oIntPos = GetMidPointOfCurve(oCS)
    
    Dim oDirOnToPlate As IJDVector
    Dim oThicknessDirVec As IJDVector
                   
    Dim oPlateUtil As IJPlateAttributes
    Set oPlateUtil = New PlateUtils
    oPlateUtil.GetPlateOrientation oPlate, oIntPos, oDirOnToPlate, oThicknessDirVec
    
    Set GetPlateOrientationVector = oDirOnToPlate
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetProfileOrientationVector(oCS As IJComplexString, oProfile As IJProfile) As IJDVector
    Const METHOD = "GetProfileOrientationVector"
    On Error GoTo ErrorHandler

    Dim oIntPos As IJDPosition
    Set oIntPos = GetMidPointOfCurve(oCS)
    
    Dim oProfileVVec As IJDVector
    Dim oProfileUVec As IJDVector
                   
    Dim oProfileUtil As IJProfileAttributes
    Set oProfileUtil = New ProfileUtils
    Dim oOriginPos As IJDPosition
    
    oProfileUtil.GetProfileOrientationAndLocation oProfile, oIntPos, oProfileUVec, oProfileVVec, oOriginPos
    
    Set GetProfileOrientationVector = oProfileVVec

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Private Function GetMidPointOfCurve(ByVal oCS As IJComplexString) As IJDPosition
    Const METHOD = "GetMidPointofCurve"
    On Error GoTo ErrorHandler
    
    Dim oRuleHelper As New MfgRuleHelpers.Helper
    Dim oWireBody   As IJWireBody
    Set oWireBody = oRuleHelper.ComplexStringToWireBody(oCS)
    
    Dim oMidPos    As IJDPosition
    Set oMidPos = oRuleHelper.GetMiddlePoint(oWireBody)
    
    Set GetMidPointOfCurve = oMidPos
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function AreThePlatePartSurfacesHavingGaps(ByVal PartialPinJig As Object) As Boolean
    Const METHOD = "AreThePlatePartSurfacesHavingGaps"
    On Error GoTo ErrorHandler
    
    Const GAP_TOLERANCE = 0.005
    AreThePlatePartSurfacesHavingGaps = False
    
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig
    
    Dim RootX As Double, RootY As Double, RootZ As Double
    Dim NormX As Double, NormY As Double, NormZ As Double
    oPinJig.GetBasePlane NormX, NormY, NormZ, RootX, RootY, RootZ

    Dim oPlane As IJPlane
    Set oPlane = New Plane3d

    oPlane.DefineByPointNormal RootX, RootY, RootZ, NormX, NormY, NormZ
    
    Dim oPlatePartColl As IJElements
    Set oPlatePartColl = oPinJig.SupportedPlates
    
    Dim oSurfaceUtil As IJMfgUtilSurface
    Set oSurfaceUtil = New MfgUtilSurface
    
    Dim RemarkingSides() As Long
    RemarkingSides = oSurfaceUtil.GetRemarkingSidesOfPlates(oPlatePartColl, oPlane)
    
    Dim oPlateSysColl As IJElements
    Set oPlateSysColl = GetPinJigSupportedPlateSystems(oPinJig)
    
    Dim iCnt As Long, jCnt As Long
    For iCnt = 1 To oPlateSysColl.Count - 1
        Dim oStructConnectable As IJStructConnectable
        Set oStructConnectable = oPlateSysColl.Item(iCnt)
        
        Dim oConnectable As IJConnectable
        Set oConnectable = oStructConnectable
        
        'Check with every connected supported plate
        For jCnt = (iCnt + 1) To oPlateSysColl.Count
            
             Dim bIsConnected As Boolean
             Dim TmpConnections As IJElements
             oConnectable.isConnectedTo oPlateSysColl.Item(jCnt), bIsConnected, TmpConnections
             
             If bIsConnected = True Then
                
                Dim oIthSurface As IJDModelBody
                Dim oJthSurface As IJDModelBody
                
                Set oIthSurface = GetPlatePartSurfaceOnRemarkingSide(oPlatePartColl.Item(iCnt), RemarkingSides(iCnt))
                Set oJthSurface = GetPlatePartSurfaceOnRemarkingSide(oPlatePartColl.Item(jCnt), RemarkingSides(jCnt))
                
                Dim oClosestPos1 As IJDPosition, oClosestPos2 As IJDPosition
                Dim dMinDist As Double
                
                oIthSurface.GetMinimumDistance oJthSurface, oClosestPos1, oClosestPos2, dMinDist
                                
                If dMinDist > GAP_TOLERANCE Then
                    AreThePlatePartSurfacesHavingGaps = True
                    Exit Function
                End If
                
                Set oIthSurface = Nothing
                Set oJthSurface = Nothing
                
             End If
             
             Set TmpConnections = Nothing
        
        Next jCnt
        
        Set oStructConnectable = Nothing
        Set oConnectable = Nothing
        
    Next iCnt
    
    Set oPlateSysColl = Nothing
    Set oPlatePartColl = Nothing
    Set oSurfaceUtil = Nothing
    Set oPlane = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Private Function GetPlatePartSurfaceOnRemarkingSide(ByVal oPlatePart As Object, ByVal RemarkingSide As Long) As Object
    Const METHOD = "GetPlatePartSurfaceOnRemarkingSide"
    On Error GoTo ErrorHandler
    
    Dim oElements As IJElements
    Dim pBasePort As IJPort, pOffsetPort As IJPort
        
    Dim oConnectable As IJStructConnectable
    Set oConnectable = oPlatePart
    oConnectable.GetBaseOffsetLateralPorts vbNullString, False, pBasePort, pOffsetPort, oElements
    
    If RemarkingSide = CTX_BASE Then
        Set GetPlatePartSurfaceOnRemarkingSide = pBasePort.Geometry
    Else
        Set GetPlatePartSurfaceOnRemarkingSide = pOffsetPort.Geometry
    End If
    
    Set oElements = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ArePlatesHavingCenterLineIntersection(ByVal oSupportedplates As IJElements) As Boolean
    Const METHOD = "ArePlatesHavingCenterLineIntersection"
    On Error GoTo ErrorHandler
    
    ArePlatesHavingCenterLineIntersection = False
    
    Dim lCnt As Long
    For lCnt = 1 To oSupportedplates.Count
        Dim oSDPlateWrapper   As StructDetailObjects.PlatePart
        Set oSDPlateWrapper = New StructDetailObjects.PlatePart
        Set oSDPlateWrapper.object = oSupportedplates.Item(lCnt)
        
        If (oSDPlateWrapper.HasCenterLineIntersection = True) And (oSDPlateWrapper.plateType = Hull Or oSDPlateWrapper.plateType = DeckPlate) Then
            ArePlatesHavingCenterLineIntersection = True
            Exit Function
        End If
        
    Next
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function GetPinJigContourLines(ByVal PartialPinJig As Object, Optional ByVal eDirectionType As JigRemarkingDirectionType = 2) As IJElements
    Const METHOD = "GetPinJig3DContourLines"
    On Error GoTo ErrorHandler
    
    Dim oPinJig As IJPinJig
    Set oPinJig = PartialPinJig
        
    If eDirectionType = JigSurfaceRemarking Then
        Dim oJigPart3d As IJJigPart3D
        Set oJigPart3d = oPinJig.GetJigPart3D
        
        Dim oProcessData As IJJigProcessData
        Set oProcessData = oJigPart3d.GetJigProcessData
        
        Set GetPinJigContourLines = oProcessData.GetEdgesByType(STRMFG_PinJigContourLine3D)
    
    ElseIf eDirectionType = ProjectedJigRemarking Then
            
        Dim oJigOutput As IJJigOutput
        Set oJigOutput = oPinJig.GetJigOutput
        
        Dim oProjectedData As IJJigProjectedData
        Set oProjectedData = oJigOutput.GetJigProjectedData
        
        Set GetPinJigContourLines = oProjectedData.GetEdgesByType(STRMFG_PinJigContourLine2D)
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetOffsetCurve(ByVal oSurfaceBody As IJSurfaceBody, ByVal oWireToOffset As IJWireBody, ByVal dOffsetValue As Double, ByVal oRemarkingSurface As IJSurfaceBody) As IJComplexString
    Const METHOD = "GetOffsetCurve"
    On Error GoTo ErrorHandler
    
    Dim oOffsetWire     As IJWireBody
    Dim oOffsetCS       As IJComplexString
    Dim oMfgUtilWrapper As New GSCADMathGeom.MfgGeomUtilWrapper
    Dim oMfgMGHelper    As New MfgMGHelper
                        
     On Error Resume Next
     Set oOffsetWire = oMfgUtilWrapper.OffsetCurveByNormal(oSurfaceBody, oWireToOffset, dOffsetValue, oRemarkingSurface)
     If Err.Number <> 0 Then
         StrMfgLogError Err, MODULE, METHOD, "Failed to get offset curve", , , , "RULES"
         Set oOffsetWire = Nothing
     End If
     On Error GoTo ErrorHandler
     
    If Not oOffsetWire Is Nothing Then
        oMfgMGHelper.WireBodyToComplexString oOffsetWire, oOffsetCS
    End If
    
    Set GetOffsetCurve = oOffsetCS
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function GetOffsetValue(ByVal oPlatePart As IJStructConnectable, ByVal oLCWire As IJWireBody, ByVal sSide As String) As Double
    Const METHOD = "GetOffsetValue"
    On Error GoTo ErrorHandler
    
    Dim oStructConnectable  As IJStructConnectable
    Dim oBasePort           As IJPort
    Dim oOffsetPort         As IJPort
    Dim oLateralPortsCol    As IJElements
    
    Dim oLCMB               As IJDModelBody
    Dim oSurfaceMB          As IJDModelBody
    Dim oClosestPos1        As IJDPosition
    Dim oClosestPos2        As IJDPosition
    Dim dMinDist            As Double

    Set oStructConnectable = oPlatePart
    
    'Get the Base and Offset Ports
    On Error Resume Next
    oStructConnectable.GetBaseOffsetLateralPorts vbNullString, False, oBasePort, oOffsetPort, oLateralPortsCol
    If Err.Number <> 0 Then
         StrMfgLogError Err, MODULE, METHOD, "Failed to get base and offset ports", , , , "RULES"
         Exit Function
     End If
     On Error GoTo ErrorHandler
    
    If sSide = "Base" Then
        Set oSurfaceMB = oBasePort.Geometry
    ElseIf sSide = "Offset" Then
        Set oSurfaceMB = oOffsetPort.Geometry
    End If
    
    Set oLCMB = oLCWire
    oLCMB.GetMinimumDistance oSurfaceMB, oClosestPos1, oClosestPos2, dMinDist
    
    Dim oVec As IJDVector
    Set oVec = New DVector
    oVec.Set oClosestPos2.X - oClosestPos1.X, oClosestPos2.Y - oClosestPos1.Y, oClosestPos2.Z - oClosestPos1.Z
    
    If oVec.X < 0 Or oVec.Y Or oVec.Z < 0 Then
        dMinDist = -1 * dMinDist
    End If
    
    GetOffsetValue = Round(dMinDist, 4)
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
