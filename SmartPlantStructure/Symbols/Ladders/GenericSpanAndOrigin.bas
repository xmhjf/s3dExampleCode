Attribute VB_Name = "SpanAndOrigin"
 
Option Explicit
'*******************************************************************
'  Copyright (C) 2006, Intergraph Corporation.  All rights reserved.
'
'  File:  GenericSpanAndOrigin.bas
'
'  Description: The file contains the functions used for finding and setting the Origin and the Height of the
'               ladder.
'
'  History:
'   03/08/2000 - JM - Creation
'   04/18/2000 - JM - Get and Set of Attributes has changed.  Given the collection
'                                        now, so no need to have the IID to find it.
'   06/02/2000 - JM - Changed the Matrix and Offset Directions so that the positive
'                     offset direction is correct and the lighting is correct.
'   Jan 2002    AP    V3 Additions new Functionality
'   09-May-06   JMS DI#97751 - All lines are constructed through the
'                   geometry factory instead of being "new"ed
'  June 12, 2006 AS TR#88968: Added code to normalize and make the vectors in the xform matrix orthogonal
'  May 05, 2007  AS CR#116177 Added new functions used for mirorr copy support
'  01/10/2008 C.C.P. TR130780, TR124571, TR133869, TR133944 Problems with horizontal offset for Stairs/ladders with Curved Member top support.
'
'  Jul-11-2008  WR  TR-CP-144430 - Changing vertical offset of ladder results in incorrect geometry.
'                                  Span was not taking into account the top extension thus Span was computed incorrectly.
'
'******************************************************************
    
Private Const MODULE = "SPSLadderMacros.SetSpanAndOrigin"
Private m_oErrors As IJEditErrors
Private Const TOL = 0.0000001
Public Const E_FAIL = -2147467259
Private oLocalizer As IJLocalizer

'This function will be used to set the Position and Matrix of the Ladder and
'determine the height of the ladder.
'This is just a generic function that can be rewriten or changed for
'each symbol.LadderBO, Top, Bottom, Ref, OccAttributeCol)
Public Function SetOriginAndSpan(LadderBo As ISPSLadder, ByRef Top As Object, _
                                 ByRef bottom As Object, ByRef Ref As Object, _
                                 PartOccInfoCol As IJDInfosCol) As Variant
Set oLocalizer = New IMSLocalizer.Localizer
oLocalizer.Initialize App.Path & "\" & "SPSLadderMacros"
On Error GoTo ErrorHandler
Const METHOD = "SetOriginAndSpan"
    
    Dim Flip As Boolean
    Dim oBottomSurf As IJSurface
    Dim oBottomPlane As IJPlane
    Dim BottomNormal As DVector
    Set BottomNormal = New DVector
    Dim dHoriz As Double
    Dim dVert As Double
    Dim TopExt As Double
    Dim BottomExt As Double
    Dim gPosOrig As DPosition
    Dim oMountingWall As IJPlane
    Dim oTopSurf As IJSurface
    Dim oMountingNormal As DVector
    Set oMountingNormal = New DVector
    Dim oTopArc As IJArc
    Dim oTopEdge As IJLine
    Dim oTopCurve As IJCurve
    Dim oRefEdge As IJLine
    Dim oRefCurve As IJCurve
    Dim oRefPlane As IJPlane
    Dim Matrix As DT4x4
    Set Matrix = New DT4x4
    Dim startpt As DPosition
    Set startpt = New DPosition
    Dim endpt As DPosition
    Set endpt = New DPosition
    Dim SParamU As Double, SParamV As Double, EParamU As Double, EParamV As Double
    Dim xSt As Double, ySt As Double, zSt As Double
    Dim xEn As Double, yEn As Double, zEn As Double
    Dim x As Double, y As Double, z As Double
    Dim Rootx As Double, Rooty As Double, Rootz As Double
    
    Dim BoolArc As Boolean
    BoolArc = False
    Dim OccAttrs As IJDAttributes
    Set OccAttrs = LadderBo
    
    'Get the Normal of the Bottom Plane.
    On Error Resume Next
    Set oBottomPlane = bottom
    LadderBo.GetMatrix Matrix
        
    On Error Resume Next
    Set oBottomSurf = bottom
    If oBottomPlane Is Nothing Then
        If oBottomSurf Is Nothing Then
            Set oBottomSurf = oBottomPlane
            oBottomSurf.ParamRange SParamU, EParamU, SParamV, EParamV
            oBottomSurf.Normal SParamU, SParamV, x, y, z
            oBottomSurf.position SParamU, SParamV, Rootx, Rooty, Rootz
        End If
    Else
        oBottomPlane.GetNormal x, y, z
        Set oBottomSurf = oBottomPlane
    End If
     
    BottomNormal.Set 0, 0, 1 ' x, y, z
    BottomNormal.Length = 1
    Matrix.IndexValue(8) = 0
    Matrix.IndexValue(9) = 0
    Matrix.IndexValue(10) = 1
    
    'set Origin of Ladder to 0,0,0
    Set gPosOrig = New DPosition
    gPosOrig.Set 0, 0, 0
    
    On Error Resume Next
    Set oTopEdge = Top
    On Error Resume Next
    Set oTopCurve = Top
    If oTopCurve Is Nothing Then
        Set oTopCurve = oTopEdge
    End If
    On Error Resume Next
    Set oTopArc = Top
    
    If oTopEdge Is Nothing And Not oTopCurve Is Nothing Then
        BoolArc = True
    End If
    
    If Not oTopCurve Is Nothing Then
        If oTopCurve.Form = CURVE_FORM_CLOSED_WITH_CURVATURE Then
            BoolArc = True
        End If
    End If
    
    On Error Resume Next
    Set oTopSurf = Top
    On Error Resume Next
    Set oMountingWall = Top
    If oTopSurf Is Nothing Then
        Set oTopSurf = oMountingWall
    End If
    
    ' The following check is to make sure that we have proper reference object
    ' if the top edge is a surface or a plane. Other wise, raise error.
    If Not oTopSurf Is Nothing Or Not oMountingWall Is Nothing Then
        If Ref Is Nothing Then
            GoTo RefEdgeMissingHandler
        End If
    End If
    
    On Error Resume Next
    Set oRefEdge = Ref
    On Error Resume Next
    Set oRefCurve = Ref
    If oRefCurve Is Nothing Then
        Set oRefCurve = oRefEdge
    End If
    On Error Resume Next
    Set oRefPlane = Ref
    
    Dim oTopDir As DVector
    Dim oTopAway As DVector
    Set oTopDir = New DVector
    Set oTopAway = New DVector
    
    If oTopCurve Is Nothing And oTopEdge Is Nothing Then
        If oTopSurf Is Nothing Then
            Err.Clear
            GoTo ErrorHandler
        End If
    End If
    
    If oBottomSurf Is Nothing And oBottomPlane Is Nothing Then
        Err.Clear
        GoTo ErrorHandler
    End If
    
    Dim Direc As Integer
    Direc = 1
    Flip = GetAttribute(OccAttrs, "TopSupportSide", PartOccInfoCol)
      
    If Flip Then
        Direc = -1
    End If
    
    Dim proj As IJProjection
    Set proj = Top
    
    '01/10/2008 C.C.P. TR130780, TR124571, TR133869, TR133944 Problems with horizontal offset for Stairs/ladders with Curved Member top support.
    'If Not BoolArc Then    'First step toward a fix for TR130780, TR124571, TR133869, TR133944 is to retrieve horizontal offset for curved cases.
        dHoriz = LadderBo.HorizVertOffset.HorizontalOffset
    'End If
    dVert = LadderBo.HorizVertOffset.VerticalOffset
    TopExt = GetAttribute(OccAttrs, "TopExtension", PartOccInfoCol)
    Dim Maxdist As Double, Mindist As Double, position As DPosition
    Set position = New DPosition
    
    If Not proj Is Nothing Or oTopSurf Is Nothing Then
        'Begin 01/10/2008 C.C.P. TR130780, TR124571, TR133869, TR133944 Problems with horizontal offset for Stairs/ladders with Curved Member top support.
        'Call MaxOffset(dHoriz, Maxdist, Mindist, position, Top, Ref)
        Call MaxOffset(BoolArc, dHoriz, Maxdist, Mindist, position, Top, Ref)
        
        'If Not proj Is Nothing Then
        If Not proj Is Nothing Or BoolArc Then
        'End 01/10/2008 C.C.P. TR130780, TR124571, TR133869, TR133944 Problems with horizontal offset for Stairs/ladders with Curved Member top support.
            LadderBo.GetPosition x, y, z
            If x <> position.x Or y <> position.y Or z <> position.z Then
                LadderBo.SetPosition position.x, position.y, position.z
            End If
        End If
        
        If dHoriz <= Maxdist And dHoriz >= Mindist Then
        Else
            dHoriz = Mindist
            LadderBo.HorizVertOffset.HorizontalOffset = dHoriz
            LadderBo.SetPosition position.x, position.y, position.z
        End If
    End If
      
    If Not oTopCurve Is Nothing And Not BoolArc Then
        oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        oTopDir.Set xEn - xSt, yEn - ySt, 0 ' zEn - zSt
        oTopDir.Length = 1
        Set oTopAway = oTopDir.Cross(BottomNormal)
        Matrix.IndexValue(4) = oTopAway.x * Direc
        Matrix.IndexValue(5) = oTopAway.y * Direc
        Matrix.IndexValue(6) = oTopAway.z * Direc
    ElseIf Not oTopSurf Is Nothing And Not oRefCurve Is Nothing Then
        LadderBo.GetPosition x, y, z
        oTopSurf.Parameter x, y, z, SParamU, SParamV
        oTopSurf.Normal SParamU, SParamV, x, y, z

'************************************************************************************************
        
        Dim x1 As Double, y1 As Double, z1 As Double  'Temp Vars
        If Not proj Is Nothing Then
            If proj.IsOrthogonal Then
'                x = x * -1
'                y = y * -1
'                z = z * -1
            End If
        End If
'************************************************************************************************
        oMountingNormal.Set x, y, z
        Matrix.IndexValue(4) = x
        Matrix.IndexValue(5) = y
        Matrix.IndexValue(6) = z
    ElseIf BoolArc Then  ' for ARC
        oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        LadderBo.GetPosition x, y, z
        Dim param As Double
        Dim xTan1 As Double, yTan1 As Double, zTan1 As Double, xTan2 As Double, yTan2 As Double, zTan2 As Double
        oTopCurve.Parameter x, y, z, param
        oTopCurve.Evaluate param, x, y, z, xTan1, yTan1, zTan1, xTan2, yTan2, zTan2
        oTopDir.Set xTan1, yTan1, zTan1
        oTopDir.Length = 1
        Set oTopAway = oTopDir.Cross(BottomNormal)
        
        'Re-obtain the topDir so that it is orthogonal to TopAway vector. Otherwise GTransf can not be obtained for the matrix
        ' in SetMatrix
        oTopAway.Length = 1
        'get the oTopAway again using the cross so that it is orhtogonal
        Set oTopDir = BottomNormal.Cross(oTopAway)
        oTopDir.Length = 1
        
        Matrix.IndexValue(4) = oTopAway.x * Direc
        Matrix.IndexValue(5) = oTopAway.y * Direc
        Matrix.IndexValue(6) = oTopAway.z * Direc
    End If
    
    startpt.Set xSt, ySt, zSt
    endpt.Set xEn, yEn, zEn
       
    Dim IntersectPt As DPosition
    Set IntersectPt = New DPosition
    Set IntersectPt = GetIntersectPoint(Top, Ref)
    If oRefCurve Is Nothing And BoolArc Or Not proj Is Nothing Then
        LadderBo.GetPosition x, y, z
        gPosOrig.Set x, y, IntersectPt.z
    Else
        'TR-198568
        gPosOrig.Set IntersectPt.x, IntersectPt.y, IntersectPt.z
        'Set gPosOrig = IntersectPt.Clone 'GetIntersectPoint(Top, Ref) 'IntersectPt
    End If
  
    'Apply the Vertical Offset to the origin.
    gPosOrig.x = gPosOrig.x + ((dVert + TopExt) * BottomNormal.x)
    gPosOrig.y = gPosOrig.y + ((dVert + TopExt) * BottomNormal.y)
    gPosOrig.z = gPosOrig.z + ((dVert + TopExt) * BottomNormal.z)

    'Horizontal Offset Direction
    Dim OffsetDir As DVector
    Set OffsetDir = New DVector
    Dim iDirMod As Integer
    iDirMod = 1
    Dim dot As Double
    
    If Not oTopCurve Is Nothing And Not BoolArc Then
        Dim Direc1 As DVector
        Set Direc1 = New DVector
        Dim Direc2 As DVector
        Set Direc2 = New DVector
        
        Direc1.Set endpt.x - IntersectPt.x, endpt.y - IntersectPt.y, endpt.z - IntersectPt.z
        Direc2.Set startpt.x - IntersectPt.x, startpt.y - IntersectPt.y, startpt.z - IntersectPt.z
        Direc1.Length = 1
        Direc2.Length = 1
        OffsetDir.Set endpt.x - startpt.x, endpt.y - startpt.y, endpt.z - startpt.z
        dot = Direc1.dot(Direc2)
            If (dot < 0.000001) And (dot > -0.0000001) Then
                dot = 0#
            End If

        If dot >= 0# Then
            Set OffsetDir = Direc2
            If IntersectPt.DistPt(startpt) <= TOL Then
                Set OffsetDir = Direc1
            ElseIf IntersectPt.DistPt(endpt) <= TOL Then
                Set OffsetDir = Direc2
            End If
            iDirMod = 1

        Else
            OffsetDir.Set endpt.x - startpt.x, endpt.y - startpt.y, endpt.z - startpt.z
        End If
    
    ElseIf Not oTopSurf Is Nothing And Not BoolArc Then
        Set OffsetDir = oMountingNormal.Cross(BottomNormal)
        OffsetDir.Length = 1
        
        
    ElseIf BoolArc Then

    End If
    
    OffsetDir.Length = 1 'Unit Directional vector
    If Not BoolArc And proj Is Nothing Then
        gPosOrig.x = gPosOrig.x + (dHoriz * OffsetDir.x) * iDirMod
        gPosOrig.y = gPosOrig.y + (dHoriz * OffsetDir.y) * iDirMod
        gPosOrig.z = gPosOrig.z + (dHoriz * OffsetDir.z) * iDirMod
    End If

    'The offset direction is not the same as the matrices x-direction (Matrix 0 1 2).  Incorrect values in here
    'result in the lighting for the object to be incorrect.
 
   If Not oTopCurve Is Nothing Or BoolArc Then
        Matrix.IndexValue(0) = oTopDir.x * -1 * Direc
        Matrix.IndexValue(1) = oTopDir.y * -1 * Direc
        Matrix.IndexValue(2) = oTopDir.z * -1 * Direc
    ElseIf Not oTopSurf Is Nothing Then
        Dim oTemp As DVector
        Set oTemp = New DVector
        Set oTemp = oMountingNormal.Cross(BottomNormal)
        Matrix.IndexValue(0) = oTemp.x
        Matrix.IndexValue(1) = oTemp.y
        Matrix.IndexValue(2) = oTemp.z
    End If
    
    Matrix.IndexValue(12) = gPosOrig.x
    Matrix.IndexValue(13) = gPosOrig.y
    Matrix.IndexValue(14) = gPosOrig.z
    Matrix.IndexValue(3) = 0
    Matrix.IndexValue(7) = 0
    Matrix.IndexValue(11) = 0
    Matrix.IndexValue(15) = 1

    Dim Span As Double
    Dim Height As Double

    LadderBo.SetMatrix Matrix

     'Code for sloped surface as bottom input
    Dim Angle As Double
    Angle = GetAttribute(OccAttrs, "Angle", PartOccInfoCol)

    Dim IdMatrix As DT4x4
    Set IdMatrix = New DT4x4
    IdMatrix.LoadIdentity
    
    Dim Vec1 As DVector
    Dim Vec2 As DVector
    Dim Vec3 As DVector
    Dim Vec4 As DVector
    Set Vec1 = New DVector
    Set Vec2 = New DVector
    Set Vec3 = New DVector
    Set Vec4 = New DVector
    Vec1.Set Matrix.IndexValue(4), Matrix.IndexValue(5), Matrix.IndexValue(6)
    Vec2.Set 0, 0, 1  'gPosOrig.x, gPosOrig.y, gPosOrig.z - 0.1
'    Vec2.Length = 1#
    Set Vec3 = Vec1.Cross(Vec2) ' Normal Vector to V1 and V2
    IdMatrix.Rotate Angle, Vec3
    Set Vec4 = IdMatrix.TransformVector(Vec1)
    Vec4.Get x, y, z
    
    Dim GeomFactory As New IngrGeom3D.GeometryFactory
    Dim line As Line3d
    Set line = GeomFactory.Lines3d.CreateByPtVectLength(Nothing, gPosOrig.x, gPosOrig.y, gPosOrig.z, x, y, z, 1)
    line.Infinite = True
    '0, 1 * Cos(Angle), -1 * Sin(Angle)
    'line.SetDirection 0, 0, -1
    Dim temp As IJElements
    Set temp = New JObjectCollection ' elements

    'oBottomSurf should be made infinite so that we calculate the height
    'always even though the top intersect pt is not exactly above the bottom
    oBottomPlane.GetRootPoint x, y, z
    oBottomPlane.GetNormal Rootx, Rooty, Rootz

    Dim DummyFace As New Plane3d
    Dim oNewBottomSurf As IJSurface
    Dim code As Geom3dIntersectConstants
    Set DummyFace = GeomFactory.Planes3d.CreateByPointNormal(Nothing, x, y, z, Rootx, Rooty, Rootz)
    Set oNewBottomSurf = DummyFace

    oNewBottomSurf.Intersect line, temp, code
      
     If temp.Count <> 0 Then
          Dim pt1 As Double, pt2 As Double, pt3 As Double
          Dim point As IJPoint
          Set point = New Point3d
          Set point = temp.Item(1)
          Dim dist As Double
          point.GetPoint pt1, pt2, pt3
          Dim NewPos As DPosition
          Set NewPos = New DPosition
          NewPos.Set pt1, pt2, pt3
         
          dist = NewPos.DistPt(gPosOrig)
     End If
     
    If pt3 >= gPosOrig.z Then
          SetOriginAndSpan = E_FAIL
    End If
    
    Height = dist * Sin(Angle)
    Span = Height - dVert - TopExt '[TR-CP-144430]
    Call SetAttribute(OccAttrs, Height, "Height", PartOccInfoCol)
    Call SetAttribute(OccAttrs, Span, "Span", PartOccInfoCol)

    Set oBottomPlane = Nothing
    Set BottomNormal = Nothing
    Set gPosOrig = Nothing
    Set oMountingWall = Nothing
    Set oTopEdge = Nothing
    Set oRefEdge = Nothing
    Set oRefPlane = Nothing
    Set Matrix = Nothing
    Set startpt = Nothing
    Set endpt = Nothing
    Set oLocalizer = Nothing
    Exit Function
    
RefEdgeMissingHandler:
    Set oLocalizer = Nothing
    SetOriginAndSpan = LADDER_E_SIDE_REF_MISSING
    Err.Raise LADDER_E_SIDE_REF_MISSING
    Exit Function
    
ErrorHandler:
'    Set m_oErrors = New IMSErrorLog.JServerErrors
'    m_oErrors.Add Err.Number, METHOD, Err.Description
    SetOriginAndSpan = E_FAIL
    Err.Description = oLocalizer.GetString(IDS_LADDER_MISSING_INPUTS, "Ladder has missing inputs")
    Err.Raise E_FAIL
    Set oLocalizer = Nothing
End Function

Public Sub SetAttribute(OccAttrs As IJDAttributes, Value As Variant, strAttribute As String, InfosColl As IJDInfosCol)
On Error GoTo ErrHandler
    
    Dim oAttr As IJDAttribute
    Dim Attrcol As IJDAttributesCol
    Dim AttrInfo As IJDAttributeInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
    
    
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim jContext As IJContext
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
     
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
    Set oAttributeMetaData = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Dim icounter As Integer
    Dim iInterfaceCount As Integer
        
    If Not InfosColl Is Nothing Then
        iInterfaceCount = InfosColl.Count
        If iInterfaceCount > 0 Then
        For icounter = 1 To iInterfaceCount
            On Error Resume Next
            Set AttrInfo = oAttributeMetaData.AttributeInfo(InfosColl.Item(icounter).Type, strAttribute)
            If Not AttrInfo Is Nothing And Err.Number = 0 Then
                Set Attrcol = OccAttrs.CollectionOfAttributes(InfosColl.Item(icounter).Type)
                Set oAttr = Attrcol.Item(strAttribute)
                oAttr.Value = Value
                Exit Sub
            End If
            Set AttrInfo = Nothing
        Next icounter
        End If
    End If
     
    Exit Sub

ErrHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, "SetAttribute", Err.Description
    Err.Raise E_FAIL
End Sub

Public Function GetAttribute(OccAttrs As IJDAttributes, strAttribute As String, InfosColl As IJDInfosCol)    'this is collection of all interfaces
On Error GoTo ErrHandler

    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim jContext As IJContext
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
     
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
     
    Dim oAttr As IJDAttribute
    Dim Attrcol As IJDAttributesCol
    Dim AttrInfo As IJDAttributeInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
    
    Set oAttributeMetaData = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Dim icounter As Integer
    Dim iInterfaceCount As Integer
        
        If Not InfosColl Is Nothing Then
            iInterfaceCount = InfosColl.Count
            If iInterfaceCount > 0 Then
            For icounter = 1 To iInterfaceCount
                On Error Resume Next
                Set AttrInfo = oAttributeMetaData.AttributeInfo(InfosColl.Item(icounter).Type, strAttribute)
                If Not AttrInfo Is Nothing And Err.Number = 0 Then
                    Set Attrcol = OccAttrs.CollectionOfAttributes(InfosColl.Item(icounter).Type)
                    Set oAttr = Attrcol.Item(strAttribute)
                    GetAttribute = oAttr.Value
                    Exit Function
                End If
                Set AttrInfo = Nothing
            Next icounter
            End If
        End If
        
    Exit Function
    
ErrHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, "GetAttribute", Err.Description
    Err.Raise E_FAIL
End Function

Public Function GetAttribute1(OccAttrs As IJDAttributes, strAttribute As String, InfosColl As IJDInfosCol)        'this is collection of all interfaces
On Error GoTo ErrHandler
    Dim oAttr As IJDAttribute
    Dim Attrcol As IJDAttributesCol
    Dim AttrInfo As IJDAttributeInfo
    Dim oAttributeMetaData As IJDAttributeMetaData
      
    Dim oDBTypeConfig As IJDBTypeConfiguration
    Dim jContext As IJContext
    Dim oConnectMiddle As IJDAccessMiddle
    Dim strModelDBID As String
     
    Set jContext = GetJContext()
    Set oDBTypeConfig = jContext.GetService("DBTypeConfiguration")
    Set oConnectMiddle = jContext.GetService("ConnectMiddle")
    strModelDBID = oDBTypeConfig.get_DataBaseFromDBType("Model")
     
    Set oAttributeMetaData = oConnectMiddle.GetResourceManager(strModelDBID)
    
    Dim icounter As Integer
    Dim iInterfaceCount As Integer
        
        If Not InfosColl Is Nothing Then
            iInterfaceCount = InfosColl.Count
            If iInterfaceCount > 0 Then
            For icounter = 1 To iInterfaceCount
                On Error Resume Next
                Set AttrInfo = oAttributeMetaData.AttributeInfo(InfosColl.Item(icounter).Type, strAttribute)
                If Not AttrInfo Is Nothing And Err.Number = 0 Then
                    Set Attrcol = OccAttrs.CollectionOfAttributes(InfosColl.Item(icounter).Type)
                    Set oAttr = Attrcol.Item(strAttribute)
                    GetAttribute1 = oAttr.Value
                    Exit Function
                End If
                Set AttrInfo = Nothing
            Next icounter
            End If
        End If
        
    Exit Function
ErrHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, "GetAttribute1", Err.Description
    Err.Raise E_FAIL
End Function

Private Function GetIntersectPoint(m_pTopEdge As Object, m_pRefEdge As Object) As DPosition
Const METHOD = "GetIntersectPoint"
On Error GoTo ErrorHandler

    Set GetIntersectPoint = New DPosition
    Dim x2 As Double, y2 As Double, z2 As Double
    Dim xSt As Double, ySt As Double, zSt As Double
    Dim xEn As Double, yEn As Double, zEn As Double
    Dim ClosestParam As Double
    Dim ParmU As Double, ParmV As Double
    Dim oTopface As IJPlane
    Dim oTopSurf As IJSurface
    Dim oTopCurve As IJCurve
    Dim oTopEdge As IJLine
    Dim oRefCurve As IJCurve
    Dim oRefEdge As IJLine
    Dim oRefFace As IJPlane
    Dim oRefSurf As IJSurface
    
    On Error Resume Next
    Set oTopEdge = m_pTopEdge
    On Error Resume Next
    Set oTopCurve = m_pTopEdge
    If oTopCurve Is Nothing Then
        Set oTopCurve = oTopEdge
    End If
    On Error Resume Next
    Set oTopface = m_pTopEdge
    On Error Resume Next
    Set oTopSurf = m_pTopEdge
    If oTopSurf Is Nothing Then
        Set oTopSurf = oTopface
    End If
    
    On Error Resume Next
    Set oRefEdge = m_pRefEdge
    On Error Resume Next
    Set oRefCurve = m_pRefEdge
    If oRefCurve Is Nothing Then
        Set oRefCurve = oRefEdge
    End If
    On Error Resume Next
    Set oRefFace = m_pRefEdge
    On Error Resume Next
    Set oRefSurf = m_pRefEdge
    If oRefSurf Is Nothing Then
        Set oRefSurf = oRefFace
    End If
        
    If Not oTopSurf Is Nothing And Not oRefSurf Is Nothing Then
        GoTo ErrorHandler
    End If
    
    If m_pRefEdge Is Nothing Then
        If Not oTopCurve Is Nothing Then
            oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
            GetIntersectPoint.Set xSt, ySt, zSt
            Exit Function
        End If
    Else
        If oRefSurf Is Nothing And oRefCurve Is Nothing Then   ' i know Top will be curve ' kludge as gridplanes dont support IJsurface
            Set GetIntersectPoint = FindIntersection(oRefFace, oTopCurve)
        Exit Function
        End If
    End If
    
    Dim x1 As Double, y1 As Double, z1 As Double, Mindist As Double
    If Not oRefCurve Is Nothing Or oTopSurf Is Nothing And oRefSurf Is Nothing Then
        oRefCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        If Not oTopCurve Is Nothing Then
            oTopCurve.EndPoints x1, y1, z1, x2, y2, z2  'TR#110689  Null pointer used in GenericSpanAndOrigin.bas, variable oTopCurve
            If Abs(z2 - z1) <= TOL Then
                oTopCurve.DistanceBetween oRefCurve, Mindist, x1, y1, z1, x2, y2, z2
                oTopCurve.Parameter x2, y2, z2, ClosestParam
                oTopCurve.position ClosestParam, x2, y2, z2
            Else
                Dim GeomFactory As New IngrGeom3D.GeometryFactory
                Dim DummyFace As New Plane3d
                Dim oNewBottomSurf As IJSurface
                Dim code As Geom3dIntersectConstants
                Dim temp As IJElements
                Dim dotprod As Double
                Dim NormVec As DVector
                Dim RefVec As DVector
                Dim zVec As DVector
                Set RefVec = New DVector
                Set zVec = New DVector
                Set NormVec = New DVector
                RefVec.Set xEn - xSt, yEn - ySt, zEn - zSt
                RefVec.Length = 1
                zVec.Set 0, 0, 1
                dotprod = RefVec.dot(zVec)
                If (dotprod < 0.000001) And (dotprod > -0.0000001) Then
                    Set NormVec = RefVec.Cross(zVec)
                Else
                    NormVec.Set x2 - x1, y2 - y1, 0
                End If
                Set temp = New JObjectCollection ' elements
                Set DummyFace = GeomFactory.Planes3d.CreateByPointNormal(Nothing, xSt, ySt, zSt, NormVec.x, NormVec.y, NormVec.z)
                Set oNewBottomSurf = DummyFace
                oNewBottomSurf.Intersect oTopCurve, temp, code
                If temp.Count <> 0 Then
                    Dim point As IJPoint
                    Set point = New Point3d
                    Set point = temp.Item(1)
                    point.GetPoint x2, y2, z2
                End If
            End If
        ElseIf Not oTopSurf Is Nothing Then
            oTopSurf.Parameter xEn, yEn, zEn, ParmU, ParmV
            oTopSurf.position ParmU, ParmV, x2, y2, z2
        End If
    ElseIf Not oRefSurf Is Nothing And Not oTopCurve Is Nothing Then 'TR#110689  Null pointer used in GenericSpanAndOrigin.bas, variable oTopCurve
        oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        oRefSurf.Parameter xSt, ySt, zSt, ParmU, ParmV
        oRefSurf.position ParmU, ParmV, x2, y2, z2
    End If
    GetIntersectPoint.Set x2, y2, z2
    
    Exit Function
    
ErrorHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Public Function FindIntersection(oPlane As IJPlane, oLine As IJCurve) As DPosition
    On Error GoTo ErrHandler
    Dim v1 As DVector
    Dim v2 As DVector
    Dim StartPoint As DPosition
    Dim EndPoint As DPosition
    Dim PlanePoint As DPosition
    Dim Dot_prod As Double


    Set v1 = New DVector
    Set v2 = New DVector
    Set StartPoint = New DPosition
    Set EndPoint = New DPosition
    Set PlanePoint = New DPosition
    Dim x As Double, y As Double, z As Double
    Dim x1 As Double, y1 As Double, z1 As Double
    
    oLine.EndPoints x, y, z, x1, y1, z1
    StartPoint.Set x, y, z
    EndPoint.Set x1, y1, z1
    oPlane.GetRootPoint x, y, z
    PlanePoint.Set x, y, z
    
    v1.Set EndPoint.x - StartPoint.x, EndPoint.y - StartPoint.y, EndPoint.z - StartPoint.z
    v2.Set PlanePoint.x - StartPoint.x, PlanePoint.y - StartPoint.y, PlanePoint.z - StartPoint.z

    v1.Length = 1
    Dot_prod = v1.dot(v2)

    Dim Intersection As DPosition
    Set Intersection = New DPosition

    Intersection.x = StartPoint.x + (Dot_prod * v1.x)
    Intersection.y = StartPoint.y + (Dot_prod * v1.y)
    Intersection.z = StartPoint.z + (Dot_prod * v1.z)
    Set FindIntersection = Intersection


    Set v1 = Nothing
    Set v2 = Nothing
    Set StartPoint = Nothing
    Set EndPoint = Nothing
    Set PlanePoint = Nothing
    Set Intersection = Nothing

    Exit Function
ErrHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, "FindIntersection", Err.Description
    Err.Raise E_FAIL
End Function

        
'01/10/2008 C.C.P. TR130780, TR124571, TR133869, TR133944 Problems with horizontal offset for Stairs/ladders with Curved Member top support.
'Private Function MaxOffset(dist As Double, ByRef Maxdist As Double, ByRef Mindist As Double, _
'                            ByRef position As DPosition, ByVal m_ptop As Object, m_pref As Object)
Private Function MaxOffset(BoolArc As Boolean, dist As Double, ByRef Maxdist As Double, ByRef Mindist As Double, _
                            ByRef position As DPosition, ByVal m_ptop As Object, m_pref As Object)
Const METHOD = "MaxOffset"
On Error GoTo ErrorHandler
    Dim SParamU As Double, SParamV As Double, EParamU As Double, EParamV As Double
    Dim SParam As Double, EParam As Double
    Dim ClosestParam As Double
    Dim xSt As Double, ySt As Double, zSt As Double
    Dim xEn As Double, yEn As Double, zEn As Double
    Dim startpt As DPosition
    Set startpt = New DPosition
    Dim endpt As DPosition
    Set endpt = New DPosition
    Dim IntersectPt As DPosition
    Set IntersectPt = New DPosition
    Set IntersectPt = GetIntersectPoint(m_ptop, m_pref)

    Dim oTopEdge As IJLine
    Dim oTopCurve As IJCurve
    Dim oTopface As IJPlane
    Dim oTopSurf As IJSurface
    Dim oProj As IJProjection
    Dim TempArc As IJArc
    Dim Length As Double
    On Error Resume Next
    Set TempArc = m_ptop
    If Not TempArc Is Nothing Then
        Set oTopCurve = TempArc
        Length = oTopCurve.Length
        Mindist = 0#
        Maxdist = Length
        TempArc.GetStartPoint xSt, ySt, zSt
        If dist <> 0# Then
            Set oTopCurve = TempArc
            oTopCurve.ParamRange SParam, EParam
            ClosestParam = ((dist * (EParam - SParam)) / Length) + SParam
            oTopCurve.position ClosestParam, xSt, ySt, zSt
        End If
        position.Set xSt, ySt, zSt
        Exit Function
    End If
    On Error Resume Next
    Set oTopEdge = m_ptop
    On Error Resume Next
    Set oTopCurve = m_ptop
    If oTopCurve Is Nothing Then
        Set oTopCurve = oTopEdge
    End If
    
    On Error Resume Next
    Set oTopface = m_ptop
    On Error Resume Next
    Set oTopSurf = m_ptop
    If oTopSurf Is Nothing Then
        Set oTopSurf = oTopface
    End If
    
    On Error Resume Next
    Set oProj = m_ptop
    If Not oProj Is Nothing Then
        Dim oProjCurve As IJCurve
        Dim IntersectParam As Double
        Dim PtonTopParam  As Double
        Set oProjCurve = oProj.Curve
        oProjCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        oProjCurve.ParamRange SParam, EParam
        oProjCurve.Parameter IntersectPt.x, IntersectPt.y, IntersectPt.z, IntersectParam
        Length = oProjCurve.Length
        Mindist = (Abs(IntersectParam - SParam) / (EParam - SParam)) * Length
        Maxdist = (Abs(IntersectParam - EParam) / (EParam - SParam)) * Length
        Dim Tmpdist As Double
        If dist = 0# Then
            oProjCurve.position IntersectParam, xSt, ySt, zSt
        ElseIf dist < 0 Then
            Tmpdist = Mindist - Abs(dist)
            ClosestParam = ((Tmpdist * (EParam - SParam)) / Length) + SParam
            oProjCurve.position ClosestParam, xSt, ySt, zSt
        ElseIf dist > 0 Then
            Tmpdist = Mindist + dist
            ClosestParam = ((Tmpdist * (EParam - SParam)) / Length) + SParam
            oProjCurve.position ClosestParam, xSt, ySt, zSt
        End If
        Mindist = Mindist * -1
        position.Set xSt, ySt, zSt
        Exit Function
    End If
    
    If Not oTopSurf Is Nothing Then
        position.Set 0, 0, 0
        Exit Function
    End If
    
    Dim Direc1 As DVector
    Set Direc1 = New DVector
    Dim Direc2 As DVector
    Set Direc2 = New DVector
    Dim OffsetDir As DVector
    Set OffsetDir = New DVector
    If Not oTopCurve Is Nothing Then
        oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
        position.Set xSt, ySt, zSt
        startpt.Set xSt, ySt, zSt
        endpt.Set xEn, yEn, zEn
        If m_pref Is Nothing Then
            Mindist = 0#
            Maxdist = oTopCurve.Length
        End If
        
        '01/10/2008 C.C.P. TR130780, TR124571, TR133869, TR133944 Problems with horizontal offset for Stairs/ladders with Curved Member top support.
        'If oTopCurve.Form = CURVE_FORM_CLOSED_WITH_CURVATURE Or oTopCurve.Form = CURVE_FORM_CLOSED_WITH_TANGENT Then
        If oTopCurve.Form = CURVE_FORM_CLOSED_WITH_CURVATURE Or oTopCurve.Form = CURVE_FORM_CLOSED_WITH_TANGENT Or BoolArc Then
            oTopCurve.ParamRange SParam, EParam
            Length = oTopCurve.Length
            ClosestParam = ((dist * (EParam - SParam)) / Length) + SParam
            oTopCurve.position ClosestParam, xSt, ySt, zSt
            position.Set xSt, ySt, zSt
            Exit Function
        End If
        Dim dot As Double
        Direc1.Set endpt.x - IntersectPt.x, endpt.y - IntersectPt.y, endpt.z - IntersectPt.z
        Direc2.Set startpt.x - IntersectPt.x, startpt.y - IntersectPt.y, startpt.z - IntersectPt.z
        dot = Direc1.dot(Direc2)
        If dot >= 0# Then
            If IntersectPt.DistPt(startpt) > IntersectPt.DistPt(endpt) Then
                Mindist = IntersectPt.DistPt(endpt)
                Maxdist = IntersectPt.DistPt(startpt)
            Else
                Mindist = IntersectPt.DistPt(startpt)
                Maxdist = IntersectPt.DistPt(endpt)
            End If
        Else
            OffsetDir.Set endpt.x - startpt.x, endpt.y - startpt.y, endpt.z - startpt.z
            OffsetDir.Length = 1 'Unit Directional vector
            Mindist = -1 * (IntersectPt.DistPt(startpt))
            Maxdist = IntersectPt.DistPt(endpt)
        End If
    End If
        
    Exit Function

ErrorHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function
 
Public Function GetIntersectPointOnSideRef(oLadder As ISPSLadder) As IJDPosition
Const METHOD = "GetIntersectPointOnSideRef"
On Error GoTo ErrorHandler
    Dim oPointOnRefPlane As IJDPosition
    Dim oTopPlane As IJPlane
    Dim oTopSurface As IJSurface
    Dim oMatrix As IJDT4x4
    Dim posX As Double
    Dim posY As Double
    Dim posZ As Double
    Dim tempX As Double
    Dim tempY As Double
    Dim tempZ As Double
    Dim normX As Double
    Dim normY As Double
    Dim normZ As Double
    Dim oGeomFactory As New GeometryFactory
    Dim oLines3D As ILines3d
    Dim oPlanes3D As IPlanes3d
    
    'Initialize vars
    Set oLines3D = oGeomFactory.Lines3d
    Set oPlanes3D = oGeomFactory.Planes3d
    Set oPointOnRefPlane = New DPosition
    
    'Now get the position of the ladder and the matrix
    oLadder.GetPosition posX, posY, posZ
    oLadder.GetMatrix oMatrix
    
    'First lets create a plane from the top ref
    'create a plane with out vector and the ladder position
    Set oTopPlane = oPlanes3D.CreateByPointNormal(Nothing, posX, posY, posZ, oMatrix.IndexValue(4), oMatrix.IndexValue(5), oMatrix.IndexValue(6))
    Set oTopSurface = oTopPlane
    
    Dim oRef As Object
    Dim oRefPlane As IJPlane
    Dim oRefSurf As IJSurface
    Dim oTopCurve As IJCurve
    
    Set oRef = oLadder.ReferenceEdge
    
    'Now get the intersection point on the sideref plane for different cases
    If oRef Is Nothing Then
        'remember side ref is optional, in this case just take the starting point of the top curve
        Set oTopCurve = oLadder.TopEdge
        
        If Not oTopCurve Is Nothing Then
            Dim xSt As Double
            Dim ySt As Double
            Dim zSt As Double
            Dim xEn As Double
            Dim yEn As Double
            Dim zEn As Double
            oTopCurve.EndPoints xSt, ySt, zSt, xEn, yEn, zEn
            oPointOnRefPlane.Set xSt, ySt, zSt
        End If
    Else
        If TypeOf oRef Is IJLine Then
            Dim oRefLine As IJLine
            Set oRefLine = oRef
            Set oPointOnRefPlane = FindIntersection(oTopPlane, oRefLine)
        Else
            'For this case, we get a curve from the top support and get the min dist/pt from this curve
            If TypeOf oRef Is IJCurve Then
                Dim oRefCurve As IJCurve
                Set oRefCurve = oRef
                
                Dim oTopLine As IJLine
                Dim Mindist As Double
        
                Set oTopLine = oLines3D.CreateByPtVectLength(Nothing, posX, posY, posZ, oMatrix.IndexValue(0), oMatrix.IndexValue(1), oMatrix.IndexValue(2), 1000#)
                Set oTopCurve = oTopLine
                oTopCurve.DistanceBetween oRefCurve, Mindist, tempX, tempY, tempZ, oPointOnRefPlane.x, oPointOnRefPlane.y, oPointOnRefPlane.z
                
            Else
            'if oRef is any other type, create a surface first which will be sued for intersection
                If TypeOf oRef Is IJPlane Then
                    Dim oTempPlane As IJPlane
                    Set oTempPlane = oRef
                    oTempPlane.GetNormal normX, normY, normZ
                    oTempPlane.GetRootPoint tempX, tempY, tempZ
                    Set oRefPlane = oPlanes3D.CreateByPointNormal(Nothing, tempX, tempY, tempZ, normX, normY, normZ)
                    Set oRefSurf = oRefPlane
                
                ElseIf TypeOf oRef Is IJSurface Then
                    Set oRefSurf = oRef
                End If
                
                If Not oRefSurf Is Nothing Then
                    Dim numPts As Long
                    Dim pnts1() As Double
                    Dim pnts2() As Double
                    Dim pars1() As Double
                    Dim pars2() As Double
                    oRefSurf.DistanceBetween oTopSurface, Mindist, oPointOnRefPlane.x, oPointOnRefPlane.y, oPointOnRefPlane.z, tempX, tempY, tempZ, numPts, pnts1, pnts2, pars1, pars2
                End If
            End If
        End If
    End If

    Set GetIntersectPointOnSideRef = oPointOnRefPlane
    Set oGeomFactory = Nothing
    Exit Function
    
ErrorHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function

Public Function GetDirectionOfTopSupport(oLadder As ISPSLadder) As IJDVector
Const METHOD = "GetDirectionOfTopSupport"
On Error GoTo ErrorHandler
    'Returns direction of the top support.
    ' for line the vector from start to end
    ' for surface, the surface normal at the position of ladder
    ' for curve, the normal at the position of the ladder
    Dim oTop As Object
    Dim oTopCurve As IJCurve
    Dim oTopLine As IJLine
    Dim oTopSurface As IJSurface
    Dim oVecTop As IJDVector
    Dim tempX As Double
    Dim tempY As Double
    Dim tempZ As Double
    
    Set oVecTop = New DVector
    
    Set oTop = oLadder.TopEdge
    
    If TypeOf oTop Is IJLine Then
        Set oTopLine = oTop
        oTopLine.GetDirection tempX, tempY, tempZ
        oVecTop.Set tempX, tempY, tempZ
        
        Dim tempX1 As Double
        Dim tempY1 As Double
        Dim tempZ1 As Double
       oTopLine.GetStartPoint tempX, tempY, tempZ
       oTopLine.GetEndPoint tempX1, tempY1, tempZ1
        oVecTop.Set tempX1 - tempX, tempY1 - tempY, tempZ1 - tempZ
    ElseIf TypeOf oTop Is IJSurface Then
        Dim paramU As Double
        Dim paramV As Double
    
        Set oTopSurface = oTop
        oLadder.GetPosition tempX, tempY, tempZ
        oTopSurface.Parameter tempX, tempY, tempZ, paramU, paramV
        oTopSurface.Normal paramU, paramV, tempX, tempY, tempZ
        oVecTop.Set tempX, tempY, tempZ
    
    ElseIf TypeOf oTop Is IJCurve Then
        Dim param As Double
        Dim xTan1 As Double, yTan1 As Double, zTan1 As Double, xTan2 As Double, yTan2 As Double, zTan2 As Double
                
        Set oTopCurve = oTop
        oLadder.GetPosition tempX, tempY, tempZ
        
        oTopCurve.Parameter tempX, tempY, tempZ, param
        oTopCurve.Evaluate param, tempX, tempY, tempZ, xTan1, yTan1, zTan1, xTan2, yTan2, zTan2
        oVecTop.Set xTan1, yTan1, zTan1
    End If
    
    Set GetDirectionOfTopSupport = oVecTop
    
    Exit Function
ErrorHandler:
    Set m_oErrors = New IMSErrorLog.JServerErrors
    m_oErrors.Add Err.Number, METHOD, Err.Description
    Err.Raise E_FAIL
End Function


