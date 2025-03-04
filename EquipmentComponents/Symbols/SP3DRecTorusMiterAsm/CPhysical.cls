VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         KKC
'   Creation Date:  Tuesday, Aug 28 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'    06.Nov.2006         KKC          TR-106572  All of the new Mitered components should be capped.
'    06.Nov.2006         KKC          TR-106569  New Mitered components do not modify correctly
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

    Private m_oSymGeomHelper As IJSymbolGeometryHelper
    Private Const MODULE = "SimplePhysical:" 'Used for error messageDim
    Private PI As Double
    
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
      PI = 4 * Atn(1)
On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
End Sub
    Private Sub Class_Terminate()
        Set m_oSymGeomHelper = Nothing
    End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    
    Dim iOutput     As Double
    Dim parwidth As Double
    Dim pardepth As Double
    Dim parBendRadius As Double
    Dim parNoOfSegments As Long
    Dim parBendAngle As Double
    Dim parInsulationThickness As Double
    Dim lNoOfMiterCuts As Long
    Dim LineStrPoints() As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBendRadius = arrayOfInputs(2)
    parNoOfSegments = arrayOfInputs(3)
    parBendAngle = arrayOfInputs(4)
    parwidth = arrayOfInputs(5)
    pardepth = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    iOutput = 0
    
    If CmpDblLessThanOrEqualTo(parBendRadius, 0) Or CmpDblLessThanOrEqualTo(parNoOfSegments, 0) Then
        GoTo ErrorLabel
    End If
    
    lNoOfMiterCuts = parNoOfSegments - 1
        
'   Use parameter to store elbow angle
    Dim dElbowAng As Double
    dElbowAng = parBendAngle
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim objrectangle   As IngrGeom3D.LineString3d
    Dim objrectangle1   As IngrGeom3D.LineString3d
    Dim dElbowRadius As Double
    Dim LineStrPts(0 To 14)  As Double
    Dim recmidX As Double
    Dim recmidY As Double
    Dim recmidZ As Double
    
    Dim ObjMiter1 As Object
    Dim ObjMiter2 As Object
    Dim MajorX As Double, MajorY As Double, MajorZ As Double
    Dim RefPlane As IngrGeom3D.Plane3d
    Dim RefPlane1 As IngrGeom3D.Plane3d
    
    If lNoOfMiterCuts = 0 And CmpDblEqual(parBendAngle, PI / 2) Then
            GoTo ErrorLabel
    End If
    Dim dBRadius As Double ' This is the value of the radius considered for creating the Miter.
    If lNoOfMiterCuts = 0 Then
        ' Note:parBendRadius is considered as the segment length when No of Miter cuts is 1.
        dBRadius = parBendRadius / Tan(parBendAngle)
    Else
        dBRadius = parBendRadius
    End If
                
   If lNoOfMiterCuts = 0 Then
            'Point 1 (Top Right)
            LineStrPts(0) = -parBendRadius
            LineStrPts(1) = pardepth / 2
            LineStrPts(2) = parwidth / 2
        
        'Point 2 (Top Left)
            LineStrPts(3) = LineStrPts(0)
            LineStrPts(4) = pardepth / 2
            LineStrPts(5) = -parwidth / 2
            
        'Point 3 (Bottom Left)
            LineStrPts(6) = LineStrPts(0)
            LineStrPts(7) = -pardepth / 2
            LineStrPts(8) = -parwidth / 2
            
        'Point 4 (Bottom Right)
            LineStrPts(9) = LineStrPts(0)
            LineStrPts(10) = -pardepth / 2
            LineStrPts(11) = parwidth / 2
            
        'Point 5 (Initial Point (Top Right))
            LineStrPts(12) = LineStrPts(0)
            LineStrPts(13) = pardepth / 2
            LineStrPts(14) = parwidth / 2
   
            Set objrectangle = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPts)
            Set RefPlane = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objrectangle)
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), RefPlane
            recmidX = -(parwidth / 2) * Tan(parBendAngle)
            recmidY = 0
            recmidZ = (parwidth / 2)
            
        'If the inclined end face of the segment intersects with the axis of the first face (The vertical face)
        'then limit the inclination of the inclined face to prevent failing of CreateByCurves function.
            If CmpDblGreaterthan(recmidX, dBRadius) Then recmidX = dBRadius
            
        'Point 1 (Top Right)
            LineStrPts(0) = recmidX
            LineStrPts(1) = pardepth / 2
            LineStrPts(2) = recmidZ
        
        'Point 2 (Top Left)
            LineStrPts(3) = recmidX
            LineStrPts(4) = -pardepth / 2
            LineStrPts(5) = recmidZ
            
        'Point 3 (Bottom Left)
            LineStrPts(6) = -recmidX
            LineStrPts(7) = -pardepth / 2
            LineStrPts(8) = -recmidZ
            
        'Point 4 (Bottom Right)
            LineStrPts(9) = -recmidX
            LineStrPts(10) = pardepth / 2
            LineStrPts(11) = -recmidZ
            
        'Point 5 (Initial Point (Top Right))
            LineStrPts(12) = recmidX
            LineStrPts(13) = pardepth / 2
            LineStrPts(14) = recmidZ
        
            Set objrectangle1 = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPts)
            Set RefPlane1 = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objrectangle1)
           
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), RefPlane1
            
            '   Created Ruled surface
            Set ObjMiter1 = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                                    objrectangle, objrectangle1, False)
        '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMiter1
            Set ObjMiter1 = Nothing
    Else
         'Point 1 (Top Right)
            LineStrPts(0) = -parBendRadius * Tan(parBendAngle / 2)
            LineStrPts(1) = pardepth / 2
            LineStrPts(2) = parwidth / 2
        
        'Point 2 (Top Left)
            LineStrPts(3) = LineStrPts(0)
            LineStrPts(4) = pardepth / 2
            LineStrPts(5) = -parwidth / 2
            
        'Point 3 (Bottom Left)
            LineStrPts(6) = LineStrPts(0)
            LineStrPts(7) = -pardepth / 2
            LineStrPts(8) = -parwidth / 2
            
        'Point 4 (Bottom Right)
            LineStrPts(9) = LineStrPts(0)
            LineStrPts(10) = -pardepth / 2
            LineStrPts(11) = parwidth / 2
            
        'Point 5 (Initial Point (Top Right))
            LineStrPts(12) = LineStrPts(0)
            LineStrPts(13) = pardepth / 2
            LineStrPts(14) = parwidth / 2
   
            Set objrectangle = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPts)
            Set RefPlane = geomFactory.Planes3d.CreateByOuterBdry _
                               (m_OutputColl.ResourceManager, objrectangle)
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), RefPlane
            Dim trObj As IJElements
            Dim csObj As IJElements
            Set trObj = New JObjectCollection
            Set csObj = New JObjectCollection
            
            csObj.Add objrectangle
            Dim iCount As Double
            Dim nCount As Double
            Dim noOfPoints As Double
            Dim stPoint As AutoMath.DPosition
            Set stPoint = New AutoMath.DPosition
            stPoint.Set -parBendRadius * Tan(parBendAngle / 2), 0, 0    ' The starting point
            
            noOfPoints = lNoOfMiterCuts + 2 'Total No of points in the miter =No of Mitre cuts+ St and en points
            nCount = noOfPoints * 3 'Total No of points required in the linestring
            ReDim LineStrPoints(0 To nCount - 1) As Double
            Dim dSegAngle As Double
            
            ' The starting point-
            LineStrPoints(0) = stPoint.x
            LineStrPoints(1) = stPoint.y
            LineStrPoints(2) = stPoint.z
            dSegAngle = parBendAngle / (lNoOfMiterCuts * 2)
           'The second point-
            LineStrPoints(3) = stPoint.x + dBRadius * Tan(dSegAngle)
            LineStrPoints(4) = stPoint.y
            LineStrPoints(5) = stPoint.z
            Dim dCurSegCumAngle As Double
            dCurSegCumAngle = dSegAngle
            'First point in the for loop is the third point in the line string.
            'The last count till which for loop runs (Stops just before the enpoint), hence till noOfPoints - 1
            Dim dMiterRadius As Double
            dMiterRadius = dBRadius / Cos(dSegAngle)
            For iCount = 3 To noOfPoints - 1
                dCurSegCumAngle = dCurSegCumAngle + dSegAngle * 2
                LineStrPoints(3 * iCount - 3) = stPoint.x + dMiterRadius * Sin(dCurSegCumAngle)
                LineStrPoints(3 * iCount - 2) = stPoint.y
                LineStrPoints(3 * iCount - 1) = stPoint.z + dBRadius - dMiterRadius * Cos(dCurSegCumAngle)
             Next iCount
            'The end point
            LineStrPoints(nCount - 3) = stPoint.x + dBRadius * Sin(parBendAngle)
            LineStrPoints(nCount - 2) = stPoint.y
            LineStrPoints(nCount - 1) = stPoint.z + dBRadius * (1 - Cos(parBendAngle))
            
            recmidX = parBendRadius * Tan(parBendAngle / 2) * Cos(parBendAngle)
            recmidY = 0
            recmidZ = parBendRadius * Tan(parBendAngle / 2) * Sin(parBendAngle)
            Dim oLineString As IngrGeom3D.LineString3d
            Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, lNoOfMiterCuts + 2, LineStrPoints)
            
            trObj.Add oLineString
            
            Dim Surfset   As IngrGeom3D.IJDObjectCollection
            Dim WCaps As Long
            Dim stnorm() As Double
            Dim ednorm() As Double
            Set Surfset = geomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
            oLineString, objrectangle, CircularCorner, 3, stnorm, ednorm, False)
        'Point 1 (Top Right)
            LineStrPts(0) = recmidX - (parwidth / 2) * Sin(parBendAngle)
            LineStrPts(1) = pardepth / 2
            LineStrPts(2) = recmidZ + (parwidth / 2) * Cos(parBendAngle)
        
        'Point 2 (Top Left)
            LineStrPts(3) = recmidX - (parwidth / 2) * Sin(parBendAngle)
            LineStrPts(4) = -pardepth / 2
            LineStrPts(5) = recmidZ + (parwidth / 2) * Cos(parBendAngle)
            
        'Point 3 (Bottom Left)
            LineStrPts(6) = recmidX + (parwidth / 2) * Sin(parBendAngle)
            LineStrPts(7) = -pardepth / 2
            LineStrPts(8) = recmidZ - (parwidth / 2) * Cos(parBendAngle)
            
        'Point 4 (Bottom Right)
            LineStrPts(9) = recmidX + (parwidth / 2) * Sin(parBendAngle)
            LineStrPts(10) = pardepth / 2
            LineStrPts(11) = recmidZ - (parwidth / 2) * Cos(parBendAngle)
            
        'Point 5 (Initial Point (Top Right))
            LineStrPts(12) = recmidX - (parwidth / 2) * Sin(parBendAngle)
            LineStrPts(13) = pardepth / 2
            LineStrPts(14) = recmidZ + (parwidth / 2) * Cos(parBendAngle)
            
            Set objrectangle1 = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPts)
            Set RefPlane1 = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objrectangle1)
           
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), RefPlane1
            Dim ObjSurface As Object
            For Each ObjSurface In Surfset
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface
            Next ObjSurface
          
    End If
    
   
' Set the output
    Set RefPlane = Nothing
    Set RefPlane1 = Nothing
    Set objrectangle = Nothing
    Set objrectangle1 = Nothing
    Set ObjSurface = Nothing
    Set Surfset = Nothing
    Set stPoint = Nothing
    Set oLineString = Nothing
    Set trObj = Nothing
    Set csObj = Nothing
 Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
        
 
