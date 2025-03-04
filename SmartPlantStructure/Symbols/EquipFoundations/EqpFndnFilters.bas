Attribute VB_Name = "EqpFndnFilters"
Option Explicit

'******************************************************************
' Copyright (C) 2003, Intergraph Corporation. All rights reserved.
'
'File
'    EqpFndnFilters.bas
'
'Author
'       15-Oct-03        Sudha Srikakolapu
'
'Description
'
'Notes
'
'History:
'   10-Jan-08   SS  TR#122653 - user should not be able to select bounding surfaces
'                   that are above the equipment foundation port plane
'
'   21-Mar-08   SS  TR#138674 - Filter function HasFoundationPort() is exiting early
'                   with out checking all foundation ports, overlooked in if then else...
'
'*******************************************************************

Private Const MODULE = "EqpFndnFilters :: "

Public Function HasFoundationPort(oInputObj As Object) As Integer
Const METHOD = "HasFoundationPort"
On Error GoTo ErrHandler
    Dim oConnectable As IJConnectable
    Dim oPort As IJPort
    Dim oenumPorts As IJElements
    Dim oFndPort As IJEqpFoundationPort
    Dim index As Integer, ii As Integer
    
    HasFoundationPort = 0
    
    If oInputObj Is Nothing Then Exit Function
    
    If TypeOf oInputObj Is IJEquipment Or TypeOf oInputObj Is IJEqpFoundationPort Then Else Exit Function
    
    If TypeOf oInputObj Is IJEquipment Then
        Set oConnectable = oInputObj
        If Not oConnectable Is Nothing Then
            Call oConnectable.enumConnectablePorts(oenumPorts, 2)
            If Not oenumPorts Is Nothing Then    ' tr 77196
                For Each oPort In oenumPorts
                    If TypeOf oPort Is IJEqpFoundationPort Then     'tr 70225
                        Set oFndPort = oPort
                    End If
                    If Not oFndPort Is Nothing Then
                        If oFndPort.NumberOfHoles >= 2 Then
                            HasFoundationPort = 1
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If
    ElseIf TypeOf oInputObj Is IJEqpFoundationPort Then
        Set oFndPort = oInputObj
        If oFndPort.NumberOfHoles >= 2 Then
            HasFoundationPort = 1
            Exit Function
        End If
        Set oFndPort = Nothing
    End If

    Set oConnectable = Nothing
    Set oenumPorts = Nothing
    Set oFndPort = Nothing
    Set oPort = Nothing
    Exit Function
    
ErrHandler:
    HandleError MODULE, METHOD
    
End Function

'TR#71850- Last argument added. Which contains object in select set . this is used to check that plane
' doesn't belong to same foundation which is being modified.
Public Function IsValidPlane(oPlane As Object, UserArg As IJElements, ObjsinSelectSet As IJElements) As Integer
Const METHOD = "IsValidPlane"
On Error GoTo ErrHandler
    IsValidPlane = 0

    If Not TypeOf oPlane Is IJPlane Then

    Exit Function
    End If
    Dim oPortGeom As IJLineString
    
    Dim oFndnPorts As IJElements
    Set oFndnPorts = UserArg
    Dim oPortGeomEles As IJElements 'DynElements
    Set oPortGeomEles = New JObjectCollection ' DynElements
    Dim oFndPort As IJEqpFoundationPort
    If oFndnPorts.count > 0 Then
        If TypeOf oFndnPorts.Item(1) Is IJEqpFoundationPort Then
            Set oFndPort = oFndnPorts.Item(1)
            
            Dim oX As Double, oY As Double, oZ As Double
            Dim xX As Double, xY As Double, xZ As Double
            Dim ZX As Double, ZY As Double, zZ As Double
        
        
            oFndPort.GetCS oX, oY, oZ, xX, xY, xZ, ZX, ZY, zZ
            
            Dim iTotalPortCnt As Integer
            
            If oFndPort.NumberOfHoles >= 1 Then
                Call GetFoundationPorts(oFndnPorts, oPortGeomEles, iTotalPortCnt)
            End If
     
            Set oPortGeom = oPortGeomEles.Item(1)
            If oPortGeom Is Nothing Then Exit Function
        End If
    End If
    'TR#71850
     Dim EqpFdnObj As ISPSEquipFoundation
     If ObjsinSelectSet Is Nothing Then Exit Function
     If ObjsinSelectSet.count > 0 Then
        If TypeOf ObjsinSelectSet.Item(1) Is ISPSEquipFoundation Then
            
            Set EqpFdnObj = ObjsinSelectSet.Item(1)
            
            If TypeOf oPlane Is IJDReferenceProxy Then
                Dim oRefProxy As IJDReferenceProxy
                Dim EqpFdnComp As ISPSFoundationComponent
                Dim systemChild As IJDesignChild
                Dim desParent As IJDesignParent
    
                Set oRefProxy = oPlane
                If TypeOf oRefProxy.Reference Is ISPSFoundationComponent Then
                    Set EqpFdnComp = oRefProxy.Reference
                    Set systemChild = EqpFdnComp
                    Set desParent = systemChild.GetParent
                ElseIf TypeOf oRefProxy.Reference Is ISPSEquipFoundation Then
    '                Dim Footing As ISPSFooting
    '                Set Footing = oRefProxy.Reference
                    Set desParent = oRefProxy.Reference 'footing
                End If
                
                If desParent Is EqpFdnObj Then
                    Set oRefProxy = Nothing
                    'Set EqpFdnObj = Nothing
                    Set EqpFdnComp = Nothing
                    Set systemChild = Nothing
                    Set desParent = Nothing
                    IsValidPlane = 0
                    Exit Function
                End If
      
            End If
            
            Set oRefProxy = Nothing
'            Set EqpFdnObj = Nothing
            Set EqpFdnComp = Nothing
            Set systemChild = Nothing
            Set desParent = Nothing
    '        Set Footing =nothing
        End If
    End If
    'TR#71850
    
    Dim dMinDist As Double
    Dim dsrcX As Double, dsrcY As Double, dsrcZ As Double
    Dim dinX As Double, dinY As Double, dinZ As Double
    Dim pts1() As Double, pts2() As Double
    Dim pars1() As Double, pars2() As Double
    Dim iNumPts As Long
    
    Dim i As Integer
    Dim oTempEles As IJElements
    Dim oLine As Line3d
    Dim portHoleX As Double, portHoleY As Double, portHoleZ As Double
    Dim maxDiff As Double, dHt As Double
    
    maxDiff = dHt = 0#
    
    Dim odummyPlane As IJPlane
    Dim dox As Double, doy As Double, doz As Double
    Dim nox As Double, noy As Double, noz As Double
        
    If TypeOf oPlane Is IJPlane Then    'tr 70225
        Set odummyPlane = oPlane
        odummyPlane.GetRootPoint dox, doy, doz
        odummyPlane.GetNormal nox, noy, noz
    End If
    
    ' check the lowest point when in modify and there is no points and no members
'    If Not EqpFdnObj Is Nothing And elems.count = 0 Then
'        EqpFdnObj.GetPosition Stx, Sty, Stz
'        If (Rtz < Stz) Then
'            IsValidPlane = 1
'        Else
'            IsValidPlane = 0
'        End If
'        Exit Function
'    End If
    
    ' get/check the lowest points when points have been picked
    Dim oPoint As IJPoint
    If oFndnPorts.count > 0 Then
        If TypeOf oFndnPorts.Item(1) Is IJPoint Or TypeOf oFndnPorts.Item(1) Is IJDPosition Then
            Dim tempX As Double, tempY As Double, tempZ As Double
            Dim firstPointZ As Double
            Dim oPosition As IJDPosition
            On Error Resume Next
            Set oPoint = oFndnPorts.Item(1)
            On Error GoTo ErrHandler
            If oPoint Is Nothing Then
                Set oPosition = oFndnPorts.Item(1)
                oPosition.Get tempX, tempY, firstPointZ
            Else
            oPoint.GetPoint tempX, tempY, firstPointZ
            End If
            For i = 2 To oFndnPorts.count
                If TypeOf oFndnPorts.Item(i) Is IJPoint Then
                    Set oPoint = oFndnPorts.Item(i)
                    oPoint.GetPoint tempX, tempY, tempZ
                ElseIf TypeOf oFndnPorts.Item(i) Is IJDPosition Then
                    
                    Set oPosition = oFndnPorts.Item(i)
                    oPosition.Get tempX, tempY, tempZ
                End If
                If tempZ < firstPointZ Then
                    firstPointZ = tempZ
                End If
            Next i
            If (doz < firstPointZ) Then
                If noz >= 0.7071 And noz <= 1 Then
                    IsValidPlane = 1
                ElseIf noz <= -0.7071 And noz >= -1 Then
                    IsValidPlane = 1
                End If
            Else
                IsValidPlane = 0
            End If

            Exit Function
        End If
    Else
        Dim olocal As IJLocalCoordinateSystem
        Set olocal = EqpFdnObj
        Dim oPos As DPosition
        Set oPos = olocal.Position
        If (doz < oPos.z) Then
            If noz >= 0.7071 And noz <= 1 Then
                IsValidPlane = 1
            ElseIf noz <= -0.7071 And noz >= -1 Then
                IsValidPlane = 1
            End If
        Else
            IsValidPlane = 0
        End If
        Set EqpFdnObj = Nothing
        Exit Function
    End If
    Set EqpFdnObj = Nothing
    If Not oPortGeom Is Nothing And Not odummyPlane Is Nothing Then     ' oPlaneSurface
        
        Dim code As Geom3dIntersectConstants
        Dim oGeomFactory As New IngrGeom3D.GeometryFactory
        ' construct an infinite plane with inputPlane
        Dim oDummySuppInfiPlane As IJPlane
        Set oDummySuppInfiPlane = oGeomFactory.Planes3d.CreateByPointNormal(Nothing, _
                                                                        dox, doy, doz, _
                                                                        nox, noy, noz)
                                                                        
        Dim oSuppPlaneSurface As IJSurface
        Set oSuppPlaneSurface = oDummySuppInfiPlane
        
        For i = 1 To oPortGeom.PointCount
    
            oPortGeom.GetPoint i, portHoleX, portHoleY, portHoleZ
            Set oLine = oGeomFactory.Lines3d.CreateByPtVectLength(Nothing, portHoleX, portHoleY, portHoleZ, ZX, ZY, zZ, INFINITY)
                      
            oSuppPlaneSurface.Intersect oLine, oTempEles, code
           
            If Not oTempEles Is Nothing Then        'tr 70225
                If code <> ISECT_NOSOLUTION And oTempEles.count <> 0 Then
            
                    Dim pt1 As Double, pt2 As Double, pt3 As Double
                    'Dim oPoint As IJPoint
                    Dim dist As Double
                    
                    Set oPoint = New Point3d
                    Set oPoint = oTempEles.Item(1)
                    oPoint.GetPoint pt1, pt2, pt3
                    
                    ' measure distance between interesect pt and port hole
                    dHt = Sqr((pt1 - portHoleX) * (pt1 - portHoleX) + (pt2 - portHoleY) * (pt2 - portHoleY) + (pt3 - portHoleZ) * (pt3 - portHoleZ))
                    
                    If dHt > maxDiff Then maxDiff = dHt
                
                End If
            End If
            
            Set oLine = Nothing
            If Not oTempEles Is Nothing Then
                oTempEles.Clear
                Set oTempEles = Nothing
            End If
        Next i
        
        
        
        If maxDiff >= 0.0001 Then IsValidPlane = 1
        Set oGeomFactory = Nothing
        Set oDummySuppInfiPlane = Nothing
        Set oSuppPlaneSurface = Nothing
        Set oLine = Nothing
    End If
    
    Set oPortGeomEles = Nothing
    Set oPortGeom = Nothing
    Set odummyPlane = Nothing
    Set oFndnPorts = Nothing
    Set oFndPort = Nothing
    
    Exit Function
    
ErrHandler:
    HandleError MODULE, METHOD
End Function


Public Sub GetFoundationPorts(oFndPorts As IJElements, _
                              PortHiliter As IJElements, _
                              iPortCount As Integer)
Const METHOD = "GetFoundationPorts"
On Error GoTo ErrHandler

    Dim geomFactory As GeometryFactory
    Dim OriginX As Double, OriginY As Double, OriginZ As Double
    Dim XAxisX As Double, XAxisY As Double, XAxisZ As Double
    Dim ZAxisX As Double, ZAxisY As Double, ZAxisZ As Double
    Dim dim1Count As Long, dim2Count As Long, dim1Start As Long, dim2Start As Long
    Dim holeParam() As Double
    Dim holes() As Variant
    Dim oLine As Line3d
    Dim index As String
    Dim xAxis As DVector, yAxis As DVector, zaxis As DVector
    Dim oLinestring As LineString3d
    Dim PortPoints() As Double
    Dim otrans As IJDT4x4
    Dim tempPos As IJDPosition, oNewPos As IJDPosition
    Dim ii As Integer, i As Integer, j As Integer
    
    Dim oFndPort As IJEqpFoundationPort
    Dim ofndportid As IJEqpFoundationPortID
    
    Set geomFactory = New GeometryFactory
    
    index = 0
    For ii = 1 To oFndPorts.count
        
        If TypeOf oFndPorts.Item(ii) Is IJEqpFoundationPort Then    'tr 70225
            Set oFndPort = oFndPorts.Item(ii)
        End If
        
        If Not oFndPort Is Nothing Then
            
            ' not all ports are foundation ports.
            Set ofndportid = oFndPort
            index = ofndportid.ID
        
            Set xAxis = New DVector
            Set yAxis = New DVector
            Set zaxis = New DVector
            Set oLinestring = New LineString3d
            Set otrans = New DT4x4
        
            Call oFndPort.GetCS(OriginX, OriginY, OriginZ, XAxisX, XAxisY, XAxisZ, ZAxisX, ZAxisY, ZAxisZ)
            xAxis.Set XAxisX, XAxisY, XAxisZ
            zaxis.Set ZAxisX, ZAxisY, ZAxisZ
            Set yAxis = zaxis.Cross(xAxis)
                    
            Call oFndPort.GetHoles(holes())
        
            dim1Count = UBound(holes, 1)
            dim2Count = UBound(holes, 2)
            
            dim1Start = LBound(holes, 1)
            dim2Start = LBound(holes, 2)
    
            ReDim PortPoints((dim1Count + 2) * 3) As Double
            
            otrans.IndexValue(0) = xAxis.x
            otrans.IndexValue(1) = xAxis.y
            otrans.IndexValue(2) = xAxis.z
            otrans.IndexValue(3) = 0
            otrans.IndexValue(4) = yAxis.x
            otrans.IndexValue(5) = yAxis.y
            otrans.IndexValue(6) = yAxis.z
            otrans.IndexValue(7) = 0
            otrans.IndexValue(8) = zaxis.x
            otrans.IndexValue(9) = zaxis.y
            otrans.IndexValue(10) = zaxis.z
            otrans.IndexValue(11) = 0
            otrans.IndexValue(12) = OriginX
            otrans.IndexValue(13) = OriginY
            otrans.IndexValue(14) = OriginZ
            otrans.IndexValue(15) = 1
    
            ReDim holeParam((dim2Count + 1) * (dim1Count + 1))
            For i = dim1Start To dim1Count
                For j = dim2Start To dim2Count
                    holeParam((3 * i) + j + 1) = holes(i, j)
                Next j
            
                Set tempPos = New DPosition
                Set oNewPos = New DPosition
                tempPos.Set holeParam((i * 3) + 2), holeParam((i * 3) + 3), 0
                Set oNewPos = otrans.TransformPosition(tempPos)
            
                PortPoints(3 * i) = oNewPos.x
                PortPoints(3 * i + 1) = oNewPos.y
                PortPoints(3 * i + 2) = oNewPos.z
            Next i
    
            PortPoints(3 * (dim1Count + 1)) = PortPoints(0)
            PortPoints(3 * (dim1Count + 1) + 1) = PortPoints(1)
            PortPoints(3 * (dim1Count + 1) + 2) = PortPoints(2)
     
            Set oLinestring = geomFactory.LineStrings3d.CreateByPoints(Nothing, dim1Count - dim1Start + 2, PortPoints)
            PortHiliter.Add oLinestring, CStr(ii)
            Set oLinestring = Nothing
        End If
    Next
    Exit Sub
    
ErrHandler:
    HandleError MODULE, METHOD
End Sub



