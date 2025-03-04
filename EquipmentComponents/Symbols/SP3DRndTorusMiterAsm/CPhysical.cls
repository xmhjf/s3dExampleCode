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
'    Source: PDS Equipment Modeling Document(PD_EQP), Pg-340, Round Torus Miter (U860)
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
    Dim parcylinderdia     As Double
    Dim parBendRadius As Double
    Dim parNoOfSegments As Long
    Dim parBendAngle As Double
    Dim lNoOfMiterCuts As Long
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBendRadius = arrayOfInputs(2)
    parNoOfSegments = arrayOfInputs(3)
    parBendAngle = arrayOfInputs(4)
    parcylinderdia = arrayOfInputs(5)
    parInsulationThickness = arrayOfInputs(6)
    lNoOfMiterCuts = parNoOfSegments - 1
    iOutput = 0
'   Use parameter to store elbow angle
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim objCircle   As IngrGeom3D.Circle3d
    Dim objCircle1   As IngrGeom3D.Circle3d
    Dim dElbowRadius As Double
    Dim ellNormalX As Double
    Dim ellNormalY As Double
    Dim ellNormalZ As Double
    Dim ell1CenterX As Double
    Dim ell1CenterY As Double
    Dim ell1CenterZ As Double
    Dim ObjEllipse As Object
    Dim Miter1 As Object
    Dim iCount As Double
    Dim RefPlane As IngrGeom3D.Plane3d
    Dim RefPlane1 As IngrGeom3D.Plane3d
    Dim MajorX As Double, MajorY As Double, MajorZ As Double
    Dim dBRadius As Double  'The vertical distance between the two faces.
    If lNoOfMiterCuts = 0 Then
        ' Note:parBendRadius is considered as the segment length when No of Miter cuts is 1.
        dBRadius = parBendRadius / Tan(parBendAngle)
    Else
        dBRadius = parBendRadius
    End If
    
    If lNoOfMiterCuts = 0 Then
        If CmpDblEqual(parBendAngle, PI / 2) Then
            GoTo ErrorLabel
        End If
        
        Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, -parBendRadius, _
                                                    0, 0, -1, 0, 0, parcylinderdia / 2)
        Set RefPlane = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), RefPlane
         
        ell1CenterX = 0
        ell1CenterY = 0
        ell1CenterZ = 0

        ellNormalX = Cos(parBendAngle)
        ellNormalY = 0
        ellNormalZ = Sin(parBendAngle)

        MajorX = -Tan(parBendAngle) * parcylinderdia / 2
        MajorY = 0
        MajorZ = parcylinderdia / 2

        'If the inclined end face of the segment intersects with the axis of the first face (The vertical face)
        'then limit the inclination of the inclined face to prevent failing of CreateByCurves function.
        If CmpDblGreaterthan(MajorX, dBRadius) Then MajorX = dBRadius

        Set ObjEllipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                            ell1CenterX, ell1CenterY, ell1CenterZ, _
                                            ellNormalX, ellNormalY, ellNormalZ, _
                                             MajorX, MajorY, MajorZ, Cos(parBendAngle))
         Set RefPlane1 = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, ObjEllipse)
        
        iOutput = iOutput + 1
         m_OutputColl.AddOutput arrayOfOutputs(iOutput), RefPlane1
    ' Created Ruled surface
        Set Miter1 = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                               objCircle, ObjEllipse, False)
    '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), Miter1
        Set Miter1 = Nothing
       
    ElseIf lNoOfMiterCuts > 0 Then
        Set objCircle = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, -parBendRadius * Tan(parBendAngle / 2), _
                                                    0, 0, -1, 0, 0, parcylinderdia / 2)
         Set RefPlane = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objCircle)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), RefPlane
        Dim nCount As Double
        Dim noOfPoints As Double
        Dim stPoint As AutoMath.DPosition
        Set stPoint = New AutoMath.DPosition
        stPoint.Set -parBendRadius * Tan(parBendAngle / 2), 0, 0  ' The starting point
        
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
        
        Dim Surfset   As IngrGeom3D.IJDObjectCollection
        Dim stnorm() As Double
        Dim ednorm() As Double
        Set Surfset = geomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                                        oLineString, objCircle, CircularCorner, 3, stnorm, ednorm, False)
        
        Set objCircle1 = geomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, parBendRadius * Tan(parBendAngle / 2) * Cos(parBendAngle), _
                            0, parBendRadius * Tan(parBendAngle / 2) * Sin(parBendAngle), _
                            Cos(parBendAngle), 0, Sin(parBendAngle), _
                            parcylinderdia / 2)
        Set RefPlane1 = geomFactory.Planes3d.CreateByOuterBdry _
                                           (m_OutputColl.ResourceManager, objCircle1)
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
    Set objCircle = Nothing
    Set objCircle1 = Nothing
    Set ObjEllipse = Nothing
    Set ObjSurface = Nothing
    Set Surfset = Nothing
    Set stPoint = Nothing
    Set oLineString = Nothing
 Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
    Err.HelpFile, Err.HelpContext
    
End Sub



