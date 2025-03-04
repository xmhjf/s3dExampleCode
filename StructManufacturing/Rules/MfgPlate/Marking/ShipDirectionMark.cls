VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShipDirectionMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Ship Direction Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      TBH      feb. 6. 2002    created
'      KONI     july 3 2002     implemented/tested
'               (no specific implementation for brackets,collars and edgereenforcements)
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Private Const MODULE = "MfgPlateMarking.ShipDirectionMark"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub


'********************************************************************************
'*   Internal check to make sure we only create 2 DirectionMarks at the most
'********************************************************************************

Private Function CountCheck(ByRef Count As Integer) As Boolean
    Count = Count + 1
    If Count >= 2 Then
        CountCheck = False
    Else
        CountCheck = True
    End If
End Function
Private Sub GetEdgeDirections(ByVal oPosColl As VBA.Collection, ByVal oDirColl As VBA.Collection, ByRef strDirection() As String)
    Const METHOD As String = "GetEdgeDirections"
    On Error GoTo ErrorHandler
    
    If oDirColl Is Nothing Then
        'Since one of the inputs is nothing we can immediately leave
        Exit Sub
    End If
    
    Dim ind As Long
    Dim pDirection As IJDVector
    Dim pPosition As IJDPosition
    Dim x As Double, y As Double, z As Double, pos_x As Double, pos_y As Double, pos_z As Double, prev_val As Double
    Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
    Dim ind_fore As Long, ind_port As Long, ind_upper As Long, ind_aft As Long, ind_starboard As Long, ind_lower As Long
    Dim bSkip As Boolean
    
    ind_fore = 0
    ind_port = 0
    ind_upper = 0
    ind_aft = 0
    ind_starboard = 0
    ind_lower = 0
    
start_again:

    For ind = 1 To oDirColl.Count
        bSkip = True
        Set pDirection = oDirColl.Item(ind)
        Set pPosition = oPosColl.Item(ind)
        
        pDirection.Get x, y, z
        pPosition.Get pos_x, pos_y, pos_z
        
        If x >= 0.5 And x > Abs(y) And x > Abs(z) Then
            If ind_fore = 0 Then
                strDirection(ind) = "fore"
                ind_fore = ind
            ElseIf (ind_fore <> ind) Then
                Set pPosition = oPosColl.Item(ind_fore)
                pPosition.Get prev_val, tmp_y, tmp_z
                If pos_x > prev_val Then
                    strDirection(ind_fore) = ""
                    strDirection(ind) = "fore"
                    ind_fore = ind
                    GoTo start_again
                Else
                    bSkip = False
                End If
            End If
            If (bSkip) Then GoTo Next_Ind
        End If
        
        bSkip = True
        
        If y >= 0.5 And y > Abs(z) Then
            If ind_port = 0 Then
                strDirection(ind) = "port"
                ind_port = ind
            ElseIf (ind_port <> ind) Then
                Set pPosition = oPosColl.Item(ind_port)
                pPosition.Get tmp_x, prev_val, tmp_z
                If pos_y > prev_val Then
                    strDirection(ind_port) = ""
                    strDirection(ind) = "port"
                    ind_port = ind
                    GoTo start_again
                Else
                    bSkip = False
                End If
            End If
            If (bSkip) Then GoTo Next_Ind
        End If
        
        bSkip = True
        
        If z >= 0.5 And z > Abs(x) And z > Abs(y) Then
            If ind_upper = 0 Then
                strDirection(ind) = "upper"
                ind_upper = ind
            ElseIf (ind_upper <> ind) Then
                Set pPosition = oPosColl.Item(ind_upper)
                pPosition.Get tmp_x, tmp_y, prev_val
                If pos_z > prev_val Then
                    strDirection(ind_upper) = ""
                    strDirection(ind) = "upper"
                    ind_upper = ind
                    GoTo start_again
                Else
                    bSkip = False
                End If
            End If
            If (bSkip) Then GoTo Next_Ind
        End If
        
        bSkip = True
        
        If x < -0.5 And x < y And x < z Then
            If ind_aft = 0 Then
                strDirection(ind) = "aft"
                ind_aft = ind
            ElseIf (ind_aft <> ind) Then
                Set pPosition = oPosColl.Item(ind_aft)
                pPosition.Get prev_val, tmp_y, tmp_z
                If pos_x < prev_val Then
                    strDirection(ind_aft) = ""
                    strDirection(ind) = "aft"
                    ind_aft = ind
                    GoTo start_again
                Else
                    bSkip = False
                End If
            End If
            If (bSkip) Then GoTo Next_Ind
        End If
        
        bSkip = True
        
        If y < -0.5 And y < z Then
            If ind_starboard = 0 Then
                strDirection(ind) = "starboard"
                ind_starboard = ind
            ElseIf (ind_starboard <> ind) Then
                Set pPosition = oPosColl.Item(ind_starboard)
                pPosition.Get tmp_x, prev_val, tmp_z
                If pos_y < prev_val Then
                    strDirection(ind_starboard) = ""
                    strDirection(ind) = "starboard"
                    ind_starboard = ind
                    GoTo start_again
                Else
                    bSkip = False
                End If
            End If
            If (bSkip) Then GoTo Next_Ind
        End If
        
        bSkip = True
        
        If z < -0.5 And z < x And z < y Then
            If ind_lower = 0 Then
                strDirection(ind) = "lower"
                ind_lower = ind
            ElseIf (ind_lower <> ind) Then
                Set pPosition = oPosColl.Item(ind_lower)
                pPosition.Get tmp_x, tmp_y, prev_val
                If pos_z < prev_val Then
                    strDirection(ind_lower) = ""
                    strDirection(ind) = "lower"
                    ind_lower = ind
                    GoTo start_again
                Else
                    bSkip = False
                End If
            End If
            If (bSkip) Then GoTo Next_Ind
        End If
Next_Ind:
    Next
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
'********************************************************************************
'*  Gets the mark created in the before unfold rule and clips it so that it is
'*   only 10 mm long from an intersection with the outer contour
'********************************************************************************
Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d
    Dim METHOD As String
    METHOD = "IJDMfgSystemMarkingRule_CreateAfterUnfold"
    On Error GoTo ErrorHandler
    
    Dim oMfgGeomHelper As MfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    Dim oMfgMGHelper As IJMfgMGHelper
    Set oMfgMGHelper = New GSCADMathGeom.MfgMGHelper
    
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Dim oMfgPart As IJMfgPlatePart

    If Not oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Exit Function
    End If
    
    Dim oGeomCol2d As IJMfgGeomCol2d
    Set oGeomCol2d = oMfgPart.Geometries
    
    Dim oOuterContourCol As Collection
    Set oOuterContourCol = New Collection
    
    Dim oMfgGeom As MfgGeom2d
    Dim eGeomType As StrMfgGeometryType
    Dim lIdx As Long
    
    For lIdx = 1 To oGeomCol2d.GetCount
        Set oMfgGeom = oGeomCol2d.GetGeometry(lIdx)
        eGeomType = oMfgGeom.GetGeometryType
        If eGeomType = STRMFG_OUTER_CONTOUR Then
            oOuterContourCol.Add oMfgGeom.GetGeometry
        End If
        Set oMfgGeom = Nothing
    Next
    
    Dim dDirCount As Double
    dDirCount = 0
    
    For lIdx = 1 To oGeomCol2d.GetCount
    
        Set oMfgGeom = oGeomCol2d.GetGeometry(lIdx)
        eGeomType = oMfgGeom.GetGeometryType
        
        If eGeomType = STRMFG_DIRECTION Then
            Dim oCurve As IJCurve
            Dim oOuterCurve As IJCurve
            Set oCurve = oMfgGeom.GetGeometry
            
            Dim oStartPos As IJDPosition
            Dim oEndPos As IJDPosition
            Dim oTempPos As IJDPosition
            Dim oTempPos2 As IJDPosition
            Dim oLine As Line3d
            Dim oCS As IJComplexString
            
            Dim dSX As Double
            Dim dSY As Double
            Dim dSZ As Double
            Dim dummy As Double
            
            oCurve.EndPoints dSX, dSY, dSZ, dummy, dummy, dummy
            Set oStartPos = New DPosition
            oStartPos.Set dSX, dSY, dSZ
            
            Dim lJdx As Long
            For lJdx = 1 To oOuterContourCol.Count
                Set oOuterCurve = oOuterContourCol.Item(lJdx)
                
                Dim dMinDist As Double
                Dim dX As Double
                Dim dY As Double
                Dim dZ As Double
                
                oCurve.DistanceBetween oOuterCurve, dMinDist, dX, dY, dZ, dummy, dummy, dummy
                
                If dMinDist <= 0.0001 Then
                    If oTempPos Is Nothing Then
                        Set oTempPos = New DPosition
                        oTempPos.Set dX, dY, dZ
                    Else
                        Set oTempPos2 = New DPosition
                        oTempPos2.Set dX, dY, dZ
                        
                        Dim dDist1 As Double
                        Dim dDist2 As Double
                        
                        oMfgGeomHelper.GetLengthBet2Points oCurve, oStartPos, oTempPos, dDist1
                        oMfgGeomHelper.GetLengthBet2Points oCurve, oStartPos, oTempPos2, dDist2
                        
                        If dDist2 < dDist1 Then
                            Set oTempPos = Nothing
                            Set oTempPos = oTempPos2
                        End If
                        
                        Set oOuterCurve = Nothing
                        Exit For
                    End If
                End If
                
                Set oOuterCurve = Nothing
                Set oTempPos2 = Nothing
            Next
            
            If Not oTempPos Is Nothing Then
                Set oStartPos = oTempPos
            Else
                Set oStartPos = Nothing
            End If
                                    
            If oStartPos Is Nothing Then
                Dim sX As Double
                Dim sY As Double
                Dim sZ As Double
                Dim dDummy As Double
                
                Set oStartPos = New DPosition
                oCurve.EndPoints sX, sY, sZ, dDummy, dDummy, dDummy
                oStartPos.Set sX, sY, sZ
                
                Set oEndPos = oMfgGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, 0.01)
                
                Set oLine = New Line3d
                oLine.DefineBy2Points sX, sY, sZ, oEndPos.x, oEndPos.y, oEndPos.z
                Set oCS = New ComplexString3d
                oCS.AddCurve oLine, True
                
                oMfgGeom.PutGeometry oCS
                
                Set oCurve = Nothing
                Set oCS = Nothing
                Set oLine = Nothing
                Set oStartPos = Nothing
                Set oEndPos = Nothing
            Else
                Set oEndPos = oMfgGeomHelper.GetPointAtDistAlongCurve(oCurve, oStartPos, 0.01)
                
                Set oLine = New Line3d
                oLine.DefineBy2Points oStartPos.x, oStartPos.y, oStartPos.z, oEndPos.x, oEndPos.y, oEndPos.z
                
                Set oCS = New ComplexString3d
                oCS.AddCurve oLine, True
                
                oMfgGeom.PutGeometry oCS
                
                Set oCurve = Nothing
                Set oCS = Nothing
                Set oLine = Nothing
                Set oStartPos = Nothing
                Set oEndPos = Nothing
            End If
            
            Set oCurve = Nothing
            Set oOuterCurve = Nothing
            
            Set oStartPos = Nothing
            Set oEndPos = Nothing
            Set oTempPos = Nothing
            Set oTempPos2 = Nothing
            Set oLine = Nothing
            Set oCS = Nothing
            
            dSX = 0
            dSY = 0
            dSZ = 0
            dummy = 0
        End If
        
        Set oMfgGeom = Nothing
        
    Next
    
CleanUp:
    Set oMfgGeomHelper = Nothing
    Set oPlateWrapper = Nothing
    Set oSDPlateWrapper = Nothing
    Set oGeomCol2d = Nothing
    
ErrorHandler:
    LogError Err, Err.Source, METHOD, Err.Description
    Resume Next

End Function
'********************************************************************************
'*  Creates an extremely long making line from the center of an outer contour
'*   outwards. This will be trimmed back after unfold so that the mark will not
'*   overlap with features.
'********************************************************************************
Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
    Const METHOD = "SHipDirectionMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
    On Error Resume Next
    
    Dim ogeomcol3dout As IJMfgGeomCol3d
    Set ogeomcol3dout = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    CreateAPSMarkings STRMFG_DIRECTION, ReferenceObjColl, ogeomcol3dout
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = ogeomcol3dout
        
    If bSelectiveRecompute Then
        Exit Function
    End If
    
    'Create the SD plate Wrapper and initialize it
    Dim oSDPlateWrapper As StructDetailObjects.PlatePart
    Set oSDPlateWrapper = New StructDetailObjects.PlatePart
    Set oSDPlateWrapper.object = Part

    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part

    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Dim oMfgPart As IJMfgPlatePart

    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
        Set oMfgPlateWrapper.object = oMfgPart
    Else
        Exit Function
    End If
    
    Dim oPlateCreation_AE As IJMfgPlateCreation_AE
    Dim oNeutralSurface As IJSurfaceBody
    Set oPlateCreation_AE = oMfgPart.ActiveEntity
    
    Set oNeutralSurface = oPlateCreation_AE.NeutralSurface
    
    Dim pResMgr As IUnknown
    Set pResMgr = GetActiveConnection.GetResourceManager(GetActiveConnectionName)

    ' Get the goemetries before unfold so that we can examine the edges
    Dim oGeomCol3d As IJMfgGeomCol3d
    
    Set oGeomCol3d = m_oGeom3dColFactory.Create(pResMgr)
    
    Set oGeomCol3d = oPlateCreation_AE.GeometriesBeforeUnfold

    If oGeomCol3d Is Nothing Then
        'Since we do not find any geometries to be marked we can exit
        Exit Function
    End If
    
    Dim oMoniker As IMoniker
    
    ' determine centerline-crossings
    Dim bCenterLineX As Boolean
    Dim bCenterLineY As Boolean

    bCenterLineX = oPlateWrapper.CrossesMidShip
    bCenterLineY = oSDPlateWrapper.HasCenterLineIntersection

    Dim oMidPosCol As Collection
    Dim oDirCol As Collection
    Dim oPortCol As Collection

    Set oMidPosCol = New Collection
    Set oDirCol = New Collection
    Set oPortCol = New Collection

    Dim eGeomType As StrMfgGeometryType

    ' First determine the positions of the outer contours
    ' in order to figure out which ones to mark.
    ' And while we're working with them get the directions right away.
    Dim i As Long
    Dim j As Long
    
    For i = 1 To oGeomCol3d.GetCount
        Dim oGeom3d As IJMfgGeom3D
        Set oGeom3d = oGeomCol3d.GetGeometry(i)
        eGeomType = oGeom3d.GetGeometryType

        If eGeomType = STRMFG_OUTER_CONTOUR Then
            Set oMoniker = oGeom3d.GetMoniker
            If oMoniker Is Nothing Then GoTo NextContourEdge
            Dim oObject As IUnknown
            Set oObject = m_oMfgRuleHelper.BindToObject(pResMgr, oMoniker)
            Dim oFacePort As IJPort
            Set oFacePort = oObject ' lateral faceport
            Dim oTmp1MidPos As IJDPosition
           
            Set oTmp1MidPos = m_oMfgRuleHelper.GetMiddlePoint(m_oMfgRuleHelper.ComplexStringToWireBody(oGeom3d.GetGeometry))
            
            If Not (oTmp1MidPos Is Nothing) Then
                Dim oUnkSurface As IUnknown
                Set oUnkSurface = oFacePort.Geometry
                Dim oSurface As IJSurfaceBody
                Set oSurface = oUnkSurface
                    
                Dim oDirection As IJDVector
                oSurface.GetNormalFromPosition oTmp1MidPos, oDirection
                                
                If Not oDirection Is Nothing Then
                    oMidPosCol.Add oTmp1MidPos
                    oDirCol.Add oDirection
                    oPortCol.Add oFacePort
                End If
                
            End If
        End If

NextContourEdge:
        Set oGeom3d = Nothing
        Set oMoniker = Nothing
        Set oObject = Nothing
        Set oFacePort = Nothing
        Set oTmp1MidPos = Nothing
        Set oUnkSurface = Nothing
        Set oSurface = Nothing
        Set oDirection = Nothing
    Next i
    
    If oDirCol.Count = 0 Then
        ' This might be a case where we are in Partial Recompute, but the outer contour has not changed.
        Exit Function
    End If
    
    Dim indStr As Long
    Dim strEdgeDirs() As String
    
    ReDim strEdgeDirs(1 To oDirCol.Count)
    
    GetEdgeDirections oMidPosCol, oDirCol, strEdgeDirs()
    
    ' Determine the type of plate
    Dim oPlateSystem As IJPlate
    Set oPlateSystem = Part

    Dim ePlateType As StructPlateType
    ePlateType = oPlateSystem.plateType

    Dim oDirString As String

    Dim x As Double, y As Double, z As Double
    ' Attributes to be persisted
    Dim oComplexStringCol As Collection
    Dim oDirStringCol As Collection
    Dim oPortsMarkedCol As Collection
    Set oComplexStringCol = New Collection
    Set oDirStringCol = New Collection
    Set oPortsMarkedCol = New Collection
    Dim oCS As IJComplexString
    Dim MarkCount As Integer
    Dim bPortExists As Boolean, bStarboardExists As Boolean
    
    MarkCount = 0
    Select Case ePlateType
    ' currently :
    ' DeckPlate,Hull,LBulkheadPlate,TBulkheadPlate,UntypedPlate
    ' Missing from commonschema :
    ' Collar,Bracket,EdgeReinforcement


        '***********************************************************
        '*  CASE DECK
        '*
        '*  Mark Primary fore,port or out(if it does not cross centerline)
        '*  Secondary aft,starboard
        '***********************************************************

        Case DeckPlate
            ' First find longitudinal direction (fore/aft) if it excists
            ' then transversal directions (port/starboard/out) if it excists

            For i = 1 To oMidPosCol.Count
                oDirString = strEdgeDirs(i)
                ' Mark fore if it excists
                If oDirString = "fore" Then
                    If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                        oComplexStringCol.Add oCS
                        oDirStringCol.Add "fore"
                        oPortsMarkedCol.Add oPortCol.Item(i)
                        ' No need to look at that direction/edge anymore
                        
                        strEdgeDirs(i) = vbNullString

                        If Not CountCheck(MarkCount) Then GoTo DoneSearching
                        For j = 1 To oMidPosCol.Count
                            oDirString = strEdgeDirs(j)
                            If oDirString = "aft" Then
                                strEdgeDirs(j) = vbNullString
                                Exit For
                            End If
                        Next j
                        
                        GoTo deck2 ' no need to mark aft
                    End If
                End If
            Next i

            For i = 1 To oMidPosCol.Count
                oDirString = strEdgeDirs(i)
                ' if fore does not excist mark aft if it excists
                If oDirString = "aft" Then
                    If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                        oComplexStringCol.Add oCS
                        oDirStringCol.Add "aft"
                        oPortsMarkedCol.Add oPortCol.Item(i)
                        ' No need to look at that direction/edge anymore
                        strEdgeDirs(i) = vbNullString
                        
                        If Not CountCheck(MarkCount) Then GoTo DoneSearching
                        For j = 1 To oMidPosCol.Count
                            oDirString = strEdgeDirs(j)
                            If oDirString = "fore" Then
                                strEdgeDirs(j) = vbNullString
                                Exit For
                            End If
                        Next j
                        GoTo deck2
                    End If
                End If
            Next i

deck2:
            If bCenterLineY Then
                ' if it crosses the centerline we mark transversal direction as port/starboard
                For i = 1 To oMidPosCol.Count
                    oDirString = strEdgeDirs(i)
                    'the edge is pointed towards portside and it is locates left of the centerline
                    If oDirString = "port" Then
                       If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                            oComplexStringCol.Add oCS
                            oDirStringCol.Add "port"
                            oPortsMarkedCol.Add oPortCol.Item(i)
                            ' No need to look at that direction/edge anymore
                            strEdgeDirs(i) = vbNullString

                            If Not CountCheck(MarkCount) Then GoTo DoneSearching
                            
                            For j = 1 To oMidPosCol.Count
                                oDirString = strEdgeDirs(j)
                                If oDirString = "starboard" Then
                                    strEdgeDirs(j) = vbNullString
                                    Exit For
                                End If
                            Next j
                            
                            GoTo deck3
                        End If
                    End If
                Next i
                 ' if no port was found look for starboard
                For i = 1 To oMidPosCol.Count
                    oDirString = strEdgeDirs(i)
                    'the edge is pointed towards portside and it is locates left of the centerline
                    If oDirString = "starboard" Then
                       If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                            oComplexStringCol.Add oCS
                            oDirStringCol.Add "starboard"
                            oPortsMarkedCol.Add oPortCol.Item(i)
                            ' No need to look at that direction/edge anymore
                            strEdgeDirs(i) = vbNullString

                            If Not CountCheck(MarkCount) Then GoTo DoneSearching
                            For j = 1 To oMidPosCol.Count
                                oDirString = strEdgeDirs(j)
                                If oDirString = "port" Then
                                    strEdgeDirs(j) = vbNullString
                                    Exit For
                                End If
                            Next j
                            GoTo tBulk3
                            
                        End If
                    End If
                Next i
            End If

deck3:
            If Not bCenterLineY Then
            
                bPortExists = False
                bStarboardExists = False
                For i = 1 To oMidPosCol.Count
                    oDirString = strEdgeDirs(i)
                    If oDirString = "starboard" Then
                        bStarboardExists = True
                    ElseIf oDirString = "port" Then
                        bPortExists = True
                    End If
                Next
                
                ' if it does not cross the centerline we mark transversal direction as out (not in)
                For i = 1 To oMidPosCol.Count
                    oDirString = strEdgeDirs(i)
                    'the edge is pointed towards portside and it is locates left of the centerline
                    oMidPosCol.Item(i).Get x, y, z
                    If oDirString = "port" And (bStarboardExists = False Or y >= 0) Then
                        If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                            oComplexStringCol.Add oCS
                            If y >= 0 Then
                                oDirStringCol.Add "out"
                            Else
                                oDirStringCol.Add "in"
                            End If
                            oPortsMarkedCol.Add oPortCol.Item(i)
                            ' No need to look at that direction/edge anymore
                            strEdgeDirs(i) = vbNullString

                            If Not CountCheck(MarkCount) Then GoTo DoneSearching
                            
                            For j = 1 To oMidPosCol.Count
                                oDirString = strEdgeDirs(j)
                                If oDirString = "starboard" Then
                                    strEdgeDirs(j) = vbNullString
                                    Exit For
                                End If
                            Next j
                            GoTo deckend
                        End If
                    End If
                    If oDirString = "starboard" And (bPortExists = False Or y < 0) Then
                        If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                            oComplexStringCol.Add oCS
                            If y < 0 Then
                                oDirStringCol.Add "out"
                            Else
                                oDirStringCol.Add "in"
                            End If
                            
                            oPortsMarkedCol.Add oPortCol.Item(i)
                            ' No need to look at that direction/edge anymore
                            strEdgeDirs(i) = vbNullString

                            If Not CountCheck(MarkCount) Then GoTo DoneSearching
                            
                            For j = 1 To oMidPosCol.Count
                                oDirString = strEdgeDirs(j)
                                If oDirString = "port" Then
                                    strEdgeDirs(j) = vbNullString
                                    Exit For
                                End If
                            Next j
                            GoTo deckend
                        End If
                    End If
                Next i
            End If
deckend:

        '***********************************************************
        '*  CASE HULL
        '*
        '*  Mark Primary fore,upper,port
        '*  Seconday aft,lower,starboard
        '***********************************************************

        Case Hull

            StandardMark oMidPosCol, oPortCol, strEdgeDirs, oDirCol, oNeutralSurface, oComplexStringCol, oDirStringCol, oPortsMarkedCol, MarkCount

        '***********************************************************
        '*  CASE LONGDITUDINAL BULKHEAD
        '*
        '*  Mark Primary fore,upper
        '*  Seconday aft,lower
        '***********************************************************

        Case LBulkheadPlate

            For i = 1 To oMidPosCol.Count
                oDirString = strEdgeDirs(i)
                ' Mark fore if it excists
                If oDirString = "fore" Then
                    If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                        oComplexStringCol.Add oCS
                        oDirStringCol.Add oDirString
                        oPortsMarkedCol.Add oPortCol.Item(i)
                        ' No need to look at that direction/edge anymore
                        strEdgeDirs(i) = vbNullString
                        
                        If Not CountCheck(MarkCount) Then GoTo DoneSearching
                        ' remove aft
                        For j = 1 To oMidPosCol.Count
                            oDirString = strEdgeDirs(j)
                            If oDirString = "aft" Then
                                strEdgeDirs(j) = vbNullString
                                Exit For
                            End If
                        Next j
                        GoTo lBulk2 ' no need to mark aft
                    End If
                End If
            Next i

            For i = 1 To oMidPosCol.Count
                oDirString = strEdgeDirs(i)
                ' if fore does not excist mark aft if it excists
                If oDirString = "aft" Then
                    If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                        oComplexStringCol.Add oCS
                        oDirStringCol.Add oDirString
                        oPortsMarkedCol.Add oPortCol.Item(i)
                        ' No need to look at that direction/edge anymore
                        strEdgeDirs(i) = vbNullString
                        
                        ' remove aft
                        For j = 1 To oMidPosCol.Count
                            oDirString = strEdgeDirs(j)
                            If oDirString = "fore" Then
                                strEdgeDirs(j) = vbNullString
                                Exit For
                            End If
                        Next j
                        GoTo lBulk2 'no need to look for more 'aft' directions
                        If Not CountCheck(MarkCount) Then GoTo DoneSearching
                    End If
                End If
            Next i

lBulk2:
            For i = 1 To oMidPosCol.Count
                oDirString = strEdgeDirs(i)
                ' mark 'upper' if it excists
                If oDirString = "upper" Then
                    If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                        oComplexStringCol.Add oCS
                        oDirStringCol.Add oDirString
                        oPortsMarkedCol.Add oPortCol.Item(i)
                        ' No need to look at that direction/edge anymore
                        strEdgeDirs(i) = vbNullString
                        
                        If Not CountCheck(MarkCount) Then GoTo DoneSearching
                        
                        For j = 1 To oMidPosCol.Count
                            oDirString = strEdgeDirs(j)
                            If oDirString = "lower" Then
                                strEdgeDirs(j) = vbNullString
                                Exit For
                            End If
                        Next j
                        GoTo lBulk1End
                    End If
                End If
            Next i

lBulk3:
            For i = 1 To oMidPosCol.Count
                oDirString = strEdgeDirs(i)
                ' if upper does not excist mark lower if it excists
                If oDirString = "lower" Then
                    If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                        oComplexStringCol.Add oCS
                        oDirStringCol.Add oDirString
                        oPortsMarkedCol.Add oPortCol.Item(i)
                        ' No need to look at that direction/edge anymore
                        strEdgeDirs(i) = vbNullString
                        
                        If Not CountCheck(MarkCount) Then GoTo DoneSearching
                        For j = 1 To oMidPosCol.Count
                            oDirString = strEdgeDirs(j)
                            If oDirString = "upper" Then
                                strEdgeDirs(j) = vbNullString
                                Exit For
                            End If
                        Next j
                        GoTo lBulk1End
                    End If
                End If
            Next i

lBulk1End:

        '***********************************************************
        '*  CASE TRANSVERSAL BULKHEAD
        '*
        '*  Mark Primary upper,port or in(if it does not cross centerline)
        '*  Secondary lower,starboard
        '***********************************************************

        Case TBulkheadPlate

            For i = 1 To oMidPosCol.Count
                oDirString = strEdgeDirs(i)
                ' Mark upper if it excists
                If oDirString = "upper" Then
                    If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                        oComplexStringCol.Add oCS
                        oDirStringCol.Add "upper"
                        oPortsMarkedCol.Add oPortCol.Item(i)
                        ' No need to look at that direction/edge anymore
                        strEdgeDirs(i) = vbNullString
                        
                        If Not CountCheck(MarkCount) Then GoTo DoneSearching
                        For j = 1 To oMidPosCol.Count
                            oDirString = strEdgeDirs(j)
                            If oDirString = "lower" Then
                                strEdgeDirs(j) = vbNullString
                                Exit For
                            End If
                        Next j
                        GoTo tBulk2 ' no need to mark lower
                    End If
                End If
            Next i

tBulk1:
            For i = 1 To oMidPosCol.Count
                oDirString = strEdgeDirs(i)
                ' if upper does not excist mark lower if it excists
                If oDirString = "lower" Then
                    If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                        oComplexStringCol.Add oCS
                        oDirStringCol.Add "lower"
                        oPortsMarkedCol.Add oPortCol.Item(i)
                        ' No need to look at that direction/edge anymore
                        strEdgeDirs(i) = vbNullString

                        If Not CountCheck(MarkCount) Then GoTo DoneSearching
                        For j = 1 To oMidPosCol.Count
                            oDirString = strEdgeDirs(j)
                            If oDirString = "upper" Then
                                strEdgeDirs(j) = vbNullString
                                Exit For
                            End If
                        Next j
                        GoTo tBulk2
                    End If
                End If
            Next i

tBulk2:
            If bCenterLineY Then
                ' if it crosses the centerline we mark transversal direction as port/starboard
                For i = 1 To oMidPosCol.Count
                    oDirString = strEdgeDirs(i)
                    'the edge is pointed towards portside and it is locates left of the centerline
                    If oDirString = "port" Then
                        If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                            oComplexStringCol.Add oCS
                            oDirStringCol.Add "port"
                            oPortsMarkedCol.Add oPortCol.Item(i)
                            ' No need to look at that direction/edge anymore
                            strEdgeDirs(i) = vbNullString

                            If Not CountCheck(MarkCount) Then GoTo DoneSearching
                            For j = 1 To oMidPosCol.Count
                                oDirString = strEdgeDirs(j)
                                If oDirString = "starboard" Then
                                    strEdgeDirs(j) = vbNullString
                                    Exit For
                                End If
                            Next j
                            GoTo tBulk3
                        End If
                    End If
                Next i
                 ' if no port was found look for starboard
                For i = 1 To oMidPosCol.Count
                    oDirString = strEdgeDirs(i)
                    'the edge is pointed towards portside and it is locates left of the centerline
                    If oDirString = "starboard" Then
                        If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                            oComplexStringCol.Add oCS
                            oDirStringCol.Add "starboard"
                            oPortsMarkedCol.Add oPortCol.Item(i)
                            ' No need to look at that direction/edge anymore
                            strEdgeDirs(i) = vbNullString

                            If Not CountCheck(MarkCount) Then GoTo DoneSearching
                            For j = 1 To oMidPosCol.Count
                                oDirString = strEdgeDirs(j)
                                If oDirString = "port" Then
                                    strEdgeDirs(j) = vbNullString
                                    Exit For
                                End If
                            Next j
                            GoTo tBulk3
                        End If
                    End If
                Next i
           End If

tBulk3:
            If Not bCenterLineY Then
                
                bPortExists = False
                bStarboardExists = False
                For i = 1 To oMidPosCol.Count
                    oDirString = strEdgeDirs(i)
                    If oDirString = "starboard" Then
                        bStarboardExists = True
                    ElseIf oDirString = "port" Then
                        bPortExists = True
                    End If
                Next
                
                ' if it does not cross the centerline we mark transversal direction as out (not in)
                For i = 1 To oMidPosCol.Count
                    oDirString = strEdgeDirs(i)
                    'the edge is pointed towards portside and it is locates left of the centerline
                    oMidPosCol.Item(i).Get x, y, z
                    If oDirString = "starboard" And (bPortExists = False Or y >= 0) Then
                        If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                            oComplexStringCol.Add oCS
                            If y >= 0 Then
                                oDirStringCol.Add "in"
                            Else
                                oDirStringCol.Add "out"
                            End If
                            oPortsMarkedCol.Add oPortCol.Item(i)
                            ' No need to look at that direction/edge anymore
                            strEdgeDirs(i) = vbNullString

                            If Not CountCheck(MarkCount) Then GoTo DoneSearching
                            For j = 1 To oMidPosCol.Count
                                oDirString = strEdgeDirs(j)
                                If oDirString = "port" Then
                                    strEdgeDirs(j) = vbNullString
                                    Exit For
                                End If
                            Next j
                            
                            GoTo tBulk1End
                        End If
                    End If
                    If oDirString = "port" And (bStarboardExists = False Or y < 0) Then
                        If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                            oComplexStringCol.Add oCS
                            If y < 0 Then
                                oDirStringCol.Add "in"
                            Else
                                oDirStringCol.Add "out"
                            End If
                            oPortsMarkedCol.Add oPortCol.Item(i)
                            ' No need to look at that direction/edge anymore
                            strEdgeDirs(i) = vbNullString

                            If Not CountCheck(MarkCount) Then GoTo DoneSearching
                            For j = 1 To oMidPosCol.Count
                                oDirString = strEdgeDirs(j)
                                If oDirString = "starboard" Then
                                    strEdgeDirs(j) = vbNullString
                                    Exit For
                                End If
                            Next j
                            GoTo tBulk1End
                        End If
                    End If
                Next i
            End If
tBulk1End:

        '***********************************************************
        '*  CASE UNKNOWN (STANDALONE PLATES)
        '***********************************************************

        Case Else

            StandardMark oMidPosCol, oPortCol, strEdgeDirs, oDirCol, oNeutralSurface, oComplexStringCol, oDirStringCol, oPortsMarkedCol, MarkCount

'           Case Collar
'
'           Case Bracket
'
'           Case EdgeReinforcement

    End Select

DoneSearching:

    ' If only one ship direction mark is created above
    If MarkCount < 2 Then
        For i = 1 To oMidPosCol.Count
            oDirString = strEdgeDirs(i)
           
            If oDirString <> vbNullString Then
                If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                    
                    If oDirString = "starboard" Then
                        oMidPosCol.Item(i).Get x, y, z
                        If y >= 0 Then
                            oDirStringCol.Add "in"
                        Else
                            oDirStringCol.Add "out"
                        End If
                    ElseIf oDirString = "port" Then
                        oMidPosCol.Item(i).Get x, y, z
                        If y < 0 Then
                            oDirStringCol.Add "in"
                        Else
                            oDirStringCol.Add "out"
                        End If
                    Else
                        oDirStringCol.Add oDirString
                    End If
                        
                    oComplexStringCol.Add oCS
                    oPortsMarkedCol.Add oPortCol.Item(i)
                    Exit For
                End If
            End If
        Next i
    End If

    Dim oSystemMark As IJMfgSystemMark
    Dim oMarkingInfo As MarkingInfo
    Dim oNewGeom3d As IJMfgGeom3D
            
    For i = 1 To oComplexStringCol.Count
                        
        'Create a SystemMark object to store additional information
        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))

        'Set the marking side
        oSystemMark.SetMarkingSide UpSide

        'QI for the MarkingInfo object on the SystemMark
        Set oMarkingInfo = oSystemMark
        oMarkingInfo.Direction = oDirStringCol.Item(i)
        oMarkingInfo.Name = oDirStringCol.Item(i)

        Set oNewGeom3d = m_oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
        oNewGeom3d.PutGeometry oComplexStringCol.Item(i)
        oNewGeom3d.PutGeometrytype STRMFG_DIRECTION
        
        Set oMoniker = m_oMfgRuleHelper.GetMoniker(oPortsMarkedCol.Item(i))
        oNewGeom3d.PutMoniker oMoniker
        
        oSystemMark.Set3dGeometry oNewGeom3d

        ogeomcol3dout.AddGeometry 1, oNewGeom3d

        oNewGeom3d.TrimToBoundaries = True
        
        Set oSystemMark = Nothing
        Set oMarkingInfo = Nothing
        Set oNewGeom3d = Nothing
    Next i

    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = ogeomcol3dout
    
CleanUp:
    Set oGeomCol3d = Nothing
    Set oSDPlateWrapper = Nothing
    Set oPlateWrapper = Nothing
    Set oMfgPlateWrapper = Nothing
    Set oMidPosCol = Nothing
    Set oDirCol = Nothing
    Set oPortCol = Nothing
    Set oPlateSystem = Nothing
    Set oComplexStringCol = Nothing
    Set oDirStringCol = Nothing
    Set oCS = Nothing
    ReDim strEdgeDirs(0)
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1029, , "RULES")
    GoTo CleanUp
End Function


Private Function StandardMark(oMidPosCol As Collection, oPortCol As Collection, strEdgeDirs() As String, oDirCol As Collection, ByRef oNeutralSurface As IJSurfaceBody, ByRef oComplexStringCol As Collection, ByRef oDirStringCol As Collection, ByRef oPortsMarkedCol As Collection, ByRef MarkCount As Integer)
    Const METHOD = "StandardMark"
    On Error GoTo ErrorHandler

    Dim i As Long, indStr As Long
    Dim oDirString As String
    Dim oCS As IJComplexString
    Dim MaxMarks As Integer
    Dim bIsUpperLowerPlaced As Boolean, bIsForeOrAftExists As Boolean
    
    MarkCount = 0
    bIsUpperLowerPlaced = False
    bIsForeOrAftExists = False
    
    For i = 1 To oMidPosCol.Count
        oDirString = strEdgeDirs(i)
        ' Mark upper if it excists
        If oDirString = "upper" Then
            If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                oComplexStringCol.Add oCS
                oDirStringCol.Add "upper"
                oPortsMarkedCol.Add oPortCol.Item(i)
                ' No need to look at that direction/edge anymore
                For indStr = i To (oMidPosCol.Count - 1)
                    strEdgeDirs(indStr) = strEdgeDirs(indStr + 1)
                Next
                oMidPosCol.Remove i
                oPortCol.Remove i
                oDirCol.Remove i
                If Not CountCheck(MarkCount) Then GoTo PlateEnd
                bIsUpperLowerPlaced = True
                GoTo Plate1 ' no need to mark lower
            End If
        End If
    Next i
            
    For i = 1 To oMidPosCol.Count
        oDirString = strEdgeDirs(i)
        ' if upper does not excist mark lower if it excists
        If oDirString = "lower" Then
            If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                oComplexStringCol.Add oCS
                oDirStringCol.Add "lower"
                oPortsMarkedCol.Add oPortCol.Item(i)
                ' No need to look at that direction/edge anymore
                For indStr = i To (oMidPosCol.Count - 1)
                    strEdgeDirs(indStr) = strEdgeDirs(indStr + 1)
                Next
                oMidPosCol.Remove i
                oPortCol.Remove i
                oDirCol.Remove i
                If Not CountCheck(MarkCount) Then GoTo PlateEnd
                bIsUpperLowerPlaced = True
                GoTo Plate1
            End If
        End If
    Next i

Plate1:
    For i = 1 To oMidPosCol.Count
        oDirString = strEdgeDirs(i)
        If oDirString = "fore" Or oDirString = "aft" Then
            bIsForeOrAftExists = True
        End If
    Next i
    
    If bIsUpperLowerPlaced = True And bIsForeOrAftExists = True Then
        GoTo Plate2
    End If
    
    For i = 1 To oMidPosCol.Count
        oDirString = strEdgeDirs(i)
        ' Mark port if it excists
        If oDirString = "port" Then
            If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                oComplexStringCol.Add oCS
                oDirStringCol.Add "out"
                oPortsMarkedCol.Add oPortCol.Item(i)
                ' No need to look at that direction/edge anymore
                For indStr = i To (oMidPosCol.Count - 1)
                    strEdgeDirs(indStr) = strEdgeDirs(indStr + 1)
                Next
                oMidPosCol.Remove i
                oPortCol.Remove i
                oDirCol.Remove i
                If Not CountCheck(MarkCount) Then GoTo PlateEnd
                GoTo Plate2
            End If
        End If
    Next i
            
    For i = 1 To oMidPosCol.Count
        oDirString = strEdgeDirs(i)
        ' if port does not excist mark starboard if it excists
        If oDirString = "starboard" Then
            If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                oComplexStringCol.Add oCS
                oDirStringCol.Add "in"
                oPortsMarkedCol.Add oPortCol.Item(i)
                ' No need to look at that direction/edge anymore
                For indStr = i To (oMidPosCol.Count - 1)
                    strEdgeDirs(indStr) = strEdgeDirs(indStr + 1)
                Next
                oMidPosCol.Remove i
                oPortCol.Remove i
                oDirCol.Remove i
                If Not CountCheck(MarkCount) Then GoTo PlateEnd
                GoTo Plate2
            End If
        End If
    Next i
                
Plate2:

    For i = 1 To oMidPosCol.Count
        oDirString = strEdgeDirs(i)
        ' Mark fore if it excists
        If oDirString = "fore" Then
            If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                oComplexStringCol.Add oCS
                oDirStringCol.Add "fore"
                oPortsMarkedCol.Add oPortCol.Item(i)
                ' No need to look at that direction/edge anymore
                For indStr = i To (oMidPosCol.Count - 1)
                    strEdgeDirs(indStr) = strEdgeDirs(indStr + 1)
                Next
                oMidPosCol.Remove i
                oPortCol.Remove i
                oDirCol.Remove i
                If Not CountCheck(MarkCount) Then GoTo PlateEnd
                GoTo PlateEnd
            End If
        End If
    Next i
            
    For i = 1 To oMidPosCol.Count
        oDirString = strEdgeDirs(i)
        ' if fore does not excist mark aft if it excists
        If oDirString = "aft" Then
            If CreateShipDirectionMarkLine(oMidPosCol.Item(i), oPortCol.Item(i), oDirCol.Item(i), oNeutralSurface, oCS) Then
                oComplexStringCol.Add oCS
                oDirStringCol.Add "aft"
                oPortsMarkedCol.Add oPortCol.Item(i)
                ' No need to look at that direction/edge anymore
                For indStr = i To (oMidPosCol.Count - 1)
                    strEdgeDirs(indStr) = strEdgeDirs(indStr + 1)
                Next
                oMidPosCol.Remove i
                oPortCol.Remove i
                oDirCol.Remove i
                If Not CountCheck(MarkCount) Then GoTo PlateEnd
                GoTo PlateEnd
            End If
        End If
    Next i
    
PlateEnd:
    Set oCS = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

 
