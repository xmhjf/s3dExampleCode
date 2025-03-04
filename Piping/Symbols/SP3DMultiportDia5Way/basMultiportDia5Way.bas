Attribute VB_Name = "basMultiportDia5Way"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   basMultiportDia5Way.bas
'   Author:         RUK
'   Creation Date:  Thursday Oct 18 2007
'   Description:
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'       Site address: www.saundersvalves.com, File is "Saunders Multiport Diverter Valve – 5 way.pdf"
'       CR-127644  Provide 2-way, 3-way, 4-way, and 5-way diverter valve body & operator symbols
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------    -----    ------------------
'   27.Sep.07       RUK     Created
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "basMultiportDia5Way:" 'Used for error messages
Public Const MULTI_PORT_OPTIONS_5WAY = 463
Public Const STRAIGHT_INLET = 1
Public Const INLET_WITH_90DEG_ELBOW = 2
Public Const STRAIGHT_OUTLET = 1
Public Const OUTLET_WITH_90DEG_ELBOW = 2
Public Const OUTLET_WITH_OFFSET = 3

Public Type InputType
    name        As String
    description As String
    properties  As IMSDescriptionProperties
    uomValue    As Double
End Type

Public Type OutputType
    name            As String
    description     As String
    properties      As IMSDescriptionProperties
    Aspect          As SymbolRepIds
End Type

Public Type AspectType
    name                As String
    description         As String
    properties          As IMSDescriptionProperties
    AspectId            As SymbolRepIds
End Type

Public Sub ReportUnanticipatedError(InModule As String, InMethod As String)
    Const E_FAIL = -2147467259
    Err.Raise E_FAIL
End Sub

Public Function ReturnMax5(A#, B#, C#, D#, E#) As Double
    Dim MaxValue As Double

    MaxValue = A
    If MaxValue < B Then MaxValue = B
    If MaxValue < C Then MaxValue = C
    If MaxValue < D Then MaxValue = D
    If MaxValue < E Then MaxValue = E
    ReturnMax5 = MaxValue
End Function

Public Function RotateObject(Obj As Object, vec As IJDVector, Angle As Double) As Object
    Dim oTransMat As AutoMath.DT4x4
    Set oTransMat = New DT4x4
    
    oTransMat.LoadIdentity
    oTransMat.Rotate Angle, vec
    Obj.Transform oTransMat
    Set RotateObject = Obj
    
    Set oTransMat = Nothing
End Function

Public Function CreatePortGeometry(OutputColl As Object, ByVal PortGeom As Integer, ByVal oStartPoint As IJDPosition, _
            ByVal dDiamter As Double, ByVal dStartToEnd As Double, ByVal dHeight As Double, _
            Optional dRotAbtX As Double, Optional dRotAbtY As Double, Optional dRotAbtZ As Double, _
            Optional transVec As IJDVector) As Object
    Const METHOD = "CreatePortGeometry"
    On Error GoTo ErrorHandler

    Dim objPort As Object
    Dim oGeomFact As IngrGeom3D.GeometryFactory
    Dim oCenter As AutoMath.DPosition
    Dim oNormal As AutoMath.DVector
    Dim oTransMat As AutoMath.DT4x4
    Dim oCircle As IngrGeom3D.Circle3d
    
    Set oGeomFact = New GeometryFactory
    Set oCenter = New DPosition
    Set oNormal = New DVector
    Set oTransMat = New DT4x4
    
    Dim Surfset As IngrGeom3D.IJElements
    Dim stnorm() As Double
    Dim ednorm() As Double
    Dim iCount As Integer
    
    If PortGeom = INLET_WITH_90DEG_ELBOW Or PortGeom = OUTLET_WITH_90DEG_ELBOW Then
        Dim oStPoint As AutoMath.DPosition
        Dim oEnPoint As AutoMath.DPosition
        Dim oTraceStr As IngrGeom3D.ComplexString3d
        Dim oCollection As Collection
        Dim oLine As IngrGeom3D.Line3d
        Dim oArc As IngrGeom3D.Arc3d
        
        Set oStPoint = New DPosition
        Set oEnPoint = New DPosition
        Set oTraceStr = New ComplexString3d
        Set oCollection = New Collection
        Set oLine = New Line3d
        Set oArc = New Arc3d
        
        oCenter.Set oStartPoint.x, oStartPoint.y, oStartPoint.z
        oNormal.Set 1, 0, 0
        Set oCircle = oGeomFact.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                 oCenter.x, oCenter.y, oCenter.z, _
                                                oNormal.x, oNormal.y, oNormal.z, dDiamter / 2)
        
        
        oStPoint.Set oStartPoint.x, oStartPoint.y, oStartPoint.z
        oEnPoint.Set oStPoint.x + dHeight - 0.2 * dStartToEnd, oStPoint.y, oStPoint.z
        Set oLine = PlaceTrLine(oStPoint, oEnPoint)
        oCollection.Add oLine
        
        oCenter.Set oEnPoint.x, oStartPoint.y + 0.2 * dStartToEnd, oStartPoint.z
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set oStartPoint.x + dHeight, oCenter.y, oCenter.z
        
        Set oArc = oGeomFact.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                                oCenter.x, oCenter.y, oCenter.z, _
                                                oStPoint.x, oStPoint.y, oStPoint.z, _
                                                oEnPoint.x, oEnPoint.y, oEnPoint.z)
        oCollection.Add oArc
        
        oStPoint.Set oEnPoint.x, oEnPoint.y, oEnPoint.z
        oEnPoint.Set oStPoint.x, oStartPoint.y + dStartToEnd, oEnPoint.z
        Set oLine = PlaceTrLine(oStPoint, oEnPoint)
        oCollection.Add oLine
        
        oStPoint.Set oStartPoint.x, oStartPoint.y, oStartPoint.z
        Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
        
        Set Surfset = oGeomFact.GeometryServices.CreateBySingleSweep( _
                                OutputColl.ResourceManager, oTraceStr, oCircle, _
                                CircularCorner, 0, stnorm, ednorm, False)
        For iCount = 1 To oCollection.Count
            oCollection.Remove 1
        Next iCount
        Set oCollection = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing
        Set oArc = Nothing
        Set oLine = Nothing
        Set oTraceStr = Nothing
    ElseIf PortGeom = OUTLET_WITH_OFFSET Then
        oCenter.Set oStartPoint.x, oStartPoint.y, oStartPoint.z
        oNormal.Set -1, 0, 0
        Set oCircle = oGeomFact.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                 oCenter.x, oCenter.y, oCenter.z, _
                                                oNormal.x, oNormal.y, oNormal.z, dDiamter / 2)
        Dim oLineStr As IngrGeom3D.LineString3d
        Dim dPoints(0 To 11) As Double
        
        dPoints(0) = oStartPoint.x
        dPoints(1) = oStartPoint.y
        dPoints(2) = oStartPoint.z
        
        dPoints(3) = oStartPoint.x - dHeight / 3
        dPoints(4) = dPoints(1)
        dPoints(5) = dPoints(2)
        
        dPoints(6) = dPoints(3) - dHeight / 3
        dPoints(7) = oStartPoint.y + dStartToEnd
        dPoints(8) = dPoints(2)
        
        dPoints(9) = dPoints(6) - dHeight / 3
        dPoints(10) = oStartPoint.y + dStartToEnd
        dPoints(11) = dPoints(2)
        
        Set oLineStr = oGeomFact.LineStrings3d.CreateByPoints(Nothing, 4, dPoints)
        
        Set Surfset = oGeomFact.GeometryServices.CreateBySingleSweep( _
                                OutputColl.ResourceManager, oLineStr, oCircle, _
                                CircularCorner, 0, stnorm, ednorm, False)
        Set oLineStr = Nothing
    End If
    For Each objPort In Surfset
        If Not objPort Is Nothing Then
            Exit For
        End If
    Next objPort
    
    oTransMat.LoadIdentity
    If Not CmpDblEqual(dRotAbtX, LINEAR_TOLERANCE) Then
        oNormal.Set 1, 0, 0
        oTransMat.Rotate dRotAbtX, oNormal
    End If
    If Not CmpDblEqual(dRotAbtY, LINEAR_TOLERANCE) Then
        oNormal.Set 0, 1, 0
        oTransMat.Rotate dRotAbtY, oNormal
    End If
    If Not CmpDblEqual(dRotAbtZ, LINEAR_TOLERANCE) Then
        oNormal.Set 0, 0, 1
        oTransMat.Rotate dRotAbtZ, oNormal
    End If
    objPort.Transform oTransMat
    
    If Not transVec Is Nothing Then
        oTransMat.LoadIdentity
        oTransMat.Translate transVec
        objPort.Transform oTransMat
    End If
    
    Set CreatePortGeometry = objPort

    'Remove the References
    For iCount = 1 To Surfset.Count
        Surfset.Remove 1
    Next iCount
    Set Surfset = Nothing
    Set oCenter = Nothing
    Set oNormal = Nothing
    Set oCircle = Nothing
    Set oTransMat = Nothing
    Set objPort = Nothing
    Set oGeomFact = Nothing
    
    Exit Function
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Function

