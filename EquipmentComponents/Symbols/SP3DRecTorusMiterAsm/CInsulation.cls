VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         KKC
'   Creation Date:  Tuesday, Aug 28 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'           Source: PDS Equipment Modeling Document(PD_EQP), Pg-342, Rectangular Torus Miter(U861)
'   Change History:
'   dd.mmm.yyyy          Who                 Change description
'   -----------         -----           ------------------
'    06.Nov.2006         KKC          TR-106572  All of the new Mitered components should be capped.
'    06.Nov.2006         KKC          TR-106569  New Mitered components do not modify correctly
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
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
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBendRadius = arrayOfInputs(2)
    parNoOfSegments = arrayOfInputs(3)
    parBendAngle = arrayOfInputs(4)
    parwidth = arrayOfInputs(5)
    pardepth = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    lNoOfMiterCuts = parNoOfSegments - 1
    
    iOutput = 0

'   Use parameter to store elbow angle
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim objrectangle   As IngrGeom3D.LineString3d
    Dim objrectangle1   As IngrGeom3D.LineString3d
    Dim ObjMiter1 As Object
    Dim recCenterX As Double
    Dim recCenterY As Double
    Dim recCenterZ As Double
    Dim LineStrPts(0 To 14)  As Double
    Dim recmidX As Double
    Dim recmidY As Double
    Dim recmidZ As Double
    Dim dBRadius As Double  'The vertical distance between the two faces.
    If lNoOfMiterCuts = 0 Then
        ' Note:parBendRadius is considered as the segment length when No of Miter cuts is 1.
        dBRadius = parBendRadius / Tan(parBendAngle)
    Else
        dBRadius = parBendRadius
    End If
    If lNoOfMiterCuts = 0 Then
            'Point 1 (Top Right)
            LineStrPts(0) = -parBendRadius
            LineStrPts(1) = (pardepth / 2 + parInsulationThickness)
            LineStrPts(2) = (parwidth / 2 + parInsulationThickness)
        
        'Point 2 (Top Left)
            LineStrPts(3) = LineStrPts(0)
            LineStrPts(4) = -(pardepth / 2 + parInsulationThickness)
            LineStrPts(5) = (parwidth / 2 + parInsulationThickness)
            
        'Point 3 (Bottom Left)
            LineStrPts(6) = LineStrPts(0)
            LineStrPts(7) = -(pardepth / 2 + parInsulationThickness)
            LineStrPts(8) = -(parwidth / 2 + parInsulationThickness)
            
        'Point 4 (Bottom Right)
            LineStrPts(9) = LineStrPts(0)
            LineStrPts(10) = (pardepth / 2 + parInsulationThickness)
            LineStrPts(11) = -(parwidth / 2 + parInsulationThickness)
            
        'Point 5 (Initial Point (Top Right))
            LineStrPts(12) = LineStrPts(0)
            LineStrPts(13) = (pardepth / 2 + parInsulationThickness)
            LineStrPts(14) = (parwidth / 2 + parInsulationThickness)
   
            Set objrectangle = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPts)
            recmidX = -(parwidth / 2 + parInsulationThickness) * Tan(parBendAngle)
            recmidY = 0
            recmidZ = parwidth / 2 + parInsulationThickness
        
        'If the inclined end face of the segment intersects with the axis of the first face (The vertical face)
        'then limit the inclination of the inclined face to prevent failing of CreateByCurves function.
            If CmpDblGreaterthan(recmidX, dBRadius) Then recmidX = dBRadius
        
        'Point 1 (Top Right)
            LineStrPts(0) = recmidX
            LineStrPts(1) = (pardepth / 2 + parInsulationThickness)
            LineStrPts(2) = recmidZ
        
        'Point 2 (Top Left)
            LineStrPts(3) = recmidX
            LineStrPts(4) = -(pardepth / 2 + parInsulationThickness)
            LineStrPts(5) = recmidZ
            
        'Point 3 (Bottom Left)
            LineStrPts(6) = -recmidX
            LineStrPts(7) = -(pardepth / 2 + parInsulationThickness)
            LineStrPts(8) = -recmidZ
            
        'Point 4 (Bottom Right)
            LineStrPts(9) = -recmidX
            LineStrPts(10) = (pardepth / 2 + parInsulationThickness)
            LineStrPts(11) = -recmidZ
            
        'Point 5 (Initial Point (Top Right))
            LineStrPts(12) = recmidX
            LineStrPts(13) = (pardepth / 2 + parInsulationThickness)
            LineStrPts(14) = recmidZ
        
            Set objrectangle1 = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPts)
            '   Created Ruled surface
            
            Set ObjMiter1 = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                                    objrectangle, objrectangle1, False)
        '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjMiter1
            Set ObjMiter1 = Nothing
    ElseIf lNoOfMiterCuts > 0 Then
      'Point 1 (Top Right)
            LineStrPts(0) = -parBendRadius * Tan(parBendAngle / 2)
            LineStrPts(1) = pardepth / 2 + parInsulationThickness
            LineStrPts(2) = parwidth / 2 + parInsulationThickness
        
        'Point 2 (Top Left)
            LineStrPts(3) = LineStrPts(0)
            LineStrPts(4) = -pardepth / 2 - parInsulationThickness
            LineStrPts(5) = parwidth / 2 + parInsulationThickness
            
        'Point 3 (Bottom Left)
            LineStrPts(6) = LineStrPts(0)
            LineStrPts(7) = -pardepth / 2 - parInsulationThickness
            LineStrPts(8) = -parwidth / 2 - parInsulationThickness
            
        'Point 4 (Bottom Right)
            LineStrPts(9) = LineStrPts(0)
            LineStrPts(10) = pardepth / 2 + parInsulationThickness
            LineStrPts(11) = -parwidth / 2 - parInsulationThickness
            
        'Point 5 (Initial Point (Top Right))
            LineStrPts(12) = LineStrPts(0)
            LineStrPts(13) = pardepth / 2 + parInsulationThickness
            LineStrPts(14) = parwidth / 2 + parInsulationThickness
   
            Set objrectangle = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, LineStrPts)
                                
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
            Dim LineStrPoints() As Double
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
            Dim oLineString As IngrGeom3D.LineString3d
            Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, lNoOfMiterCuts + 2, LineStrPoints)

            trObj.Add oLineString
            
            Dim Surfset   As IngrGeom3D.IJElements
            Set Surfset = geomFactory.GeometryServices.CreateBySkinning(m_OutputColl.ResourceManager, _
                trObj, csObj, 0)
            
            Dim ObjTopSupport As Object
            For Each ObjTopSupport In Surfset
                iOutput = iOutput + 1
                m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopSupport
            Next ObjTopSupport
            
    End If
    Set ObjTopSupport = Nothing
    Set Surfset = Nothing
    Set stPoint = Nothing
    Set oLineString = Nothing
    Set objrectangle = Nothing
    Set trObj = Nothing
    Set csObj = Nothing
 
 Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
        
 


