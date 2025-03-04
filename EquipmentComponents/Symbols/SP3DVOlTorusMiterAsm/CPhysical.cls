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
'   Author: KKC
'   Creation Date:  Tuesday, Aug 28 2006
'   Description:
'       TODO - fill in header description information
'        Source: PDS Equipment Modeling Document(PD_EQP), Pg-344, Vertical Oval Torus Miter (U862)
'   Change History:
'   dd.mmm.yyyy     who                     change description
'   -----------     ---                     ------------------
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
    If lNoOfMiterCuts = 0 And CmpDblEqual(parBendAngle, PI / 2) Then
        GoTo ErrorLabel
    End If
    
'   Use parameter to store elbow angle
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim dAngle As Double
    Dim objBRA  As IngrGeom3D.Arc3d
    Dim OCSP   As New AutoMath.DPosition
    Dim OCEP   As New AutoMath.DPosition
    Dim OCMP   As New AutoMath.DPosition
    
    Dim OCL1(0 To 2) As Double
    Dim OCL2(0 To 2) As Double
    Dim OCL3(0 To 2) As Double
    Dim OCL4(0 To 2) As Double
    Dim OCM1(0 To 2) As Double
    Dim OCM2(0 To 2) As Double
    Dim ellipse1 As Object
    Dim ellipse2 As Object
    Dim ellNormalX As Double
    Dim ellNormalY As Double
    Dim ellNormalZ As Double
    Dim objLineP1 As IngrGeom3D.Line3d
    Dim objLineP2 As IngrGeom3D.Line3d
    Dim objTorussta  As IngrGeom3D.ComplexString3d
    Dim StartTorus   As New AutoMath.DPosition
    Dim RefPlane As IngrGeom3D.Plane3d
    Dim RefPlane1 As IngrGeom3D.Plane3d
    Dim obottomcurves  As Collection
    Set obottomcurves = New Collection
    
    Dim ell1CenterX As Double
    Dim ell1CenterY As Double
    Dim ell1CenterZ As Double
    Dim dBRadius As Double  'The vertical distance between the two faces.
    If lNoOfMiterCuts = 0 Then
        ' Note:parBendRadius is considered as the segment length when No of Miter cuts is 1.
        dBRadius = parBendRadius / Tan(parBendAngle)
    Else
        dBRadius = parBendRadius
    End If
    If CmpDblGreaterthan(parwidth / 2, dBRadius) Then 'To avoid intersection of the faces in case_
        OCM1(2) = dBRadius                   ' their cross sections ar large and Bend Radius is small
    End If
    
    Set obottomcurves = New Collection
    If lNoOfMiterCuts = 0 Then
        OCL1(0) = -parBendRadius
        OCL1(1) = -pardepth / 2
        OCL1(2) = parwidth / 2 - pardepth / 2
        
        OCL2(0) = -parBendRadius
        OCL2(1) = pardepth / 2
        OCL2(2) = parwidth / 2 - pardepth / 2
        
        OCL3(0) = -parBendRadius
        OCL3(1) = pardepth / 2
        OCL3(2) = pardepth / 2 - parwidth / 2
        
        OCL4(0) = -parBendRadius
        OCL4(1) = -pardepth / 2
        OCL4(2) = pardepth / 2 - parwidth / 2
        
        OCM1(0) = -parBendRadius
        OCM1(1) = 0
        OCM1(2) = parwidth / 2
        
        OCM2(0) = -parBendRadius
        OCM2(1) = 0
        OCM2(2) = -parwidth / 2
        If CmpDblGreaterthan(OCL2(2), dBRadius) Then 'To avoid intersection of the faces in case_
            OCL2(2) = dBRadius                 ' their cross sections ar large and Bend Radius is small
            OCL3(2) = dBRadius
        End If
        Set objLineP1 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        OCL4(0), OCL4(1), OCL4(2), OCL1(0), OCL1(1), OCL1(2))
        obottomcurves.Add objLineP1
        OCSP.Set OCL1(0), OCL1(1), OCL1(2)
        OCEP.Set OCL2(0), OCL2(1), OCL2(2)
        OCMP.Set OCM1(0), OCM1(1), OCM1(2)
        Set objBRA = PlaceTrArcBy3Pts(OCSP, OCEP, OCMP)
        
        obottomcurves.Add objBRA
        Set objLineP2 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                         OCL2(0), OCL2(1), OCL2(2), OCL3(0), OCL3(1), OCL3(2))
        obottomcurves.Add objLineP2
        OCSP.Set OCL3(0), OCL3(1), OCL3(2)
        OCEP.Set OCL4(0), OCL4(1), OCL4(2)
        OCMP.Set OCM2(0), OCM2(1), OCM2(2)
        Set objBRA = PlaceTrArcBy3Pts(OCSP, OCEP, OCMP)
        
        obottomcurves.Add objBRA
        StartTorus.Set OCL4(0), OCL4(1), OCL4(2)
        Set objTorussta = PlaceTrCString(StartTorus, obottomcurves)
        
        OCL1(0) = -(parwidth / 2 - pardepth / 2) * Tan(parBendAngle)
        OCL1(1) = pardepth / 2
        OCL1(2) = parwidth / 2 - pardepth / 2
        
        OCL2(0) = -(parwidth / 2 - pardepth / 2) * Tan(parBendAngle)
        OCL2(1) = -pardepth / 2
        OCL2(2) = parwidth / 2 - pardepth / 2
        
        OCL3(0) = (parwidth / 2 - pardepth / 2) * Tan(parBendAngle)
        OCL3(1) = -pardepth / 2
        OCL3(2) = pardepth / 2 - parwidth / 2
        
        OCL4(0) = (parwidth / 2 - pardepth / 2) * Tan(parBendAngle)
        OCL4(1) = pardepth / 2
        OCL4(2) = pardepth / 2 - parwidth / 2
        
        OCM1(0) = -(parwidth / 2 - pardepth / 2) * Tan(parBendAngle)
        OCM1(1) = 0
        OCM1(2) = parwidth / 2 - pardepth / 2
        
        OCM2(0) = (parwidth / 2 - pardepth / 2) * Tan(parBendAngle)
        OCM2(1) = 0
        OCM2(2) = pardepth / 2 - parwidth / 2
        Set obottomcurves = New Collection
        Set objLineP1 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        OCL4(0), OCL4(1), OCL4(2), OCL1(0), OCL1(1), OCL1(2))
        obottomcurves.Add objLineP1
        
        Dim ObjEllipse As Object
        Set ObjEllipse = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                     OCM1(0), OCM1(1), OCM1(2), _
                                    Cos(parBendAngle), 0, Sin(parBendAngle), _
                                    -pardepth / 2 * Tan(parBendAngle), 0, pardepth / 2, _
                                    Cos(parBendAngle), 1.5 * PI, PI)
        obottomcurves.Add ObjEllipse
        Set objLineP2 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                         OCL2(0), OCL2(1), OCL2(2), OCL3(0), OCL3(1), OCL3(2))
       
        obottomcurves.Add objLineP2
        Set ObjEllipse = geomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                     OCM2(0), OCM2(1), OCM2(2), _
                                     Cos(parBendAngle), 0, Sin(parBendAngle), _
                                    pardepth / 2 * Tan(parBendAngle), 0, -pardepth / 2, _
                                    Cos(parBendAngle), 1.5 * PI, PI)
         
        obottomcurves.Add ObjEllipse
        StartTorus.Set OCL4(0), OCL4(1), OCL4(2)
        Dim objTorussta1  As IngrGeom3D.ComplexString3d
        Set objTorussta1 = PlaceTrCString(StartTorus, obottomcurves)
        Set RefPlane = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objTorussta)
        
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), RefPlane
    
        Set RefPlane1 = geomFactory.Planes3d.CreateByOuterBdry _
                                       (m_OutputColl.ResourceManager, objTorussta1)
                                          
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), RefPlane1
        
        
        Dim Miter1 As Object
        Set Miter1 = geomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                   objTorussta, objTorussta1, False)
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), Miter1
        Set Miter1 = Nothing
         
    Else
        Set obottomcurves = New Collection
        OCL1(0) = -parBendRadius * Tan(parBendAngle / 2)
        OCL1(1) = -pardepth / 2
        OCL1(2) = parwidth / 2 - pardepth / 2
        
        OCL2(0) = -parBendRadius * Tan(parBendAngle / 2)
        OCL2(1) = pardepth / 2
        OCL2(2) = parwidth / 2 - pardepth / 2
        
        OCL3(0) = -parBendRadius * Tan(parBendAngle / 2)
        OCL3(1) = pardepth / 2
        OCL3(2) = pardepth / 2 - parwidth / 2
        
        OCL4(0) = -parBendRadius * Tan(parBendAngle / 2)
        OCL4(1) = -pardepth / 2
        OCL4(2) = pardepth / 2 - parwidth / 2
        
        OCM1(0) = -parBendRadius * Tan(parBendAngle / 2)
        OCM1(1) = 0
        OCM1(2) = parwidth / 2
        
        OCM2(0) = -parBendRadius * Tan(parBendAngle / 2)
        OCM2(1) = 0
        OCM2(2) = -parwidth / 2
   
        Set objLineP1 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                        OCL4(0), OCL4(1), OCL4(2), OCL1(0), OCL1(1), OCL1(2))
                                         
        obottomcurves.Add objLineP1
        
        OCSP.Set OCL1(0), OCL1(1), OCL1(2)
        OCEP.Set OCL2(0), OCL2(1), OCL2(2)
        OCMP.Set OCM1(0), OCM1(1), OCM1(2)
        Set objBRA = PlaceTrArcBy3Pts(OCSP, OCEP, OCMP)
        
        obottomcurves.Add objBRA
        Set objLineP2 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                         OCL2(0), OCL2(1), OCL2(2), OCL3(0), OCL3(1), OCL3(2))
                                         
        obottomcurves.Add objLineP2
        
        OCSP.Set OCL3(0), OCL3(1), OCL3(2)
        OCEP.Set OCL4(0), OCL4(1), OCL4(2)
        OCMP.Set OCM2(0), OCM2(1), OCM2(2)
        Set objBRA = PlaceTrArcBy3Pts(OCSP, OCEP, OCMP)
        
        obottomcurves.Add objBRA
        StartTorus.Set OCL4(0), OCL4(1), OCL4(2)
        Set objTorussta = PlaceTrCString(StartTorus, obottomcurves)
        Set RefPlane = geomFactory.Planes3d.CreateByOuterBdry _
                                               (m_OutputColl.ResourceManager, objTorussta)
              
        iOutput = iOutput + 1
         m_OutputColl.AddOutput arrayOfOutputs(iOutput), RefPlane
        Dim trObj As IJElements
        Dim csObj As IJElements
        Set trObj = New JObjectCollection
        Set csObj = New JObjectCollection
        
        csObj.Add objTorussta
        Dim iCount As Double
        Dim nCount As Double
        Dim noOfPoints As Double
        Dim stPoint As AutoMath.DPosition
        Set stPoint = New AutoMath.DPosition
        If lNoOfMiterCuts = 0 Then
        stPoint.Set -parBendRadius, 0, 0
        Else
        stPoint.Set -parBendRadius * Tan(parBendAngle / 2), 0, 0 ' The starting point
        End If
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
        'The last count till which the for loop runs (Stops just before the enpoint), is noOfPoints - 1
        Dim dMiterRadius As Double
        dMiterRadius = dBRadius / Cos(dSegAngle)
        For iCount = 3 To noOfPoints - 1
            dCurSegCumAngle = dCurSegCumAngle + dSegAngle * 2
            LineStrPoints(3 * iCount - 3) = stPoint.x + dMiterRadius * Sin(dCurSegCumAngle)
            LineStrPoints(3 * iCount - 2) = stPoint.y
            LineStrPoints(3 * iCount - 1) = stPoint.z + dBRadius - dMiterRadius * Cos(dCurSegCumAngle)
        Next iCount
        LineStrPoints(nCount - 3) = stPoint.x + dBRadius * Sin(parBendAngle)
        LineStrPoints(nCount - 2) = stPoint.y
        LineStrPoints(nCount - 1) = stPoint.z + dBRadius * (1 - Cos(parBendAngle))
       
        Dim oLineString As IngrGeom3D.LineString3d
        Set geomFactory = New IngrGeom3D.GeometryFactory
        Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, lNoOfMiterCuts + 2, LineStrPoints)
    
        trObj.Add oLineString
        
        Dim Surfset   As IngrGeom3D.IJElements
        Dim stnorm() As Double
        Dim ednorm() As Double
        Set Surfset = geomFactory.GeometryServices.CreateBySingleSweep(m_OutputColl.ResourceManager, _
                      oLineString, objTorussta, CircularCorner, 3, stnorm, ednorm, False)
        
        Set obottomcurves = New Collection
        OCL1(0) = parBendRadius * Tan(parBendAngle / 2) * Cos(parBendAngle) - (parwidth / 2 - pardepth / 2) * Sin(parBendAngle)
        OCL1(1) = pardepth / 2
        OCL1(2) = parBendRadius * Tan(parBendAngle / 2) * Sin(parBendAngle) + (parwidth / 2 - pardepth / 2) * Cos(parBendAngle)
        
        OCL2(0) = OCL1(0)
        OCL2(1) = OCL1(1) - pardepth
        OCL2(2) = OCL1(2)
        
        OCL3(0) = OCL2(0) + (parwidth - pardepth) * Sin(parBendAngle)
        OCL3(1) = OCL2(1)
        OCL3(2) = OCL2(2) - (parwidth - pardepth) * Cos(parBendAngle)
        
        OCL4(0) = OCL3(0)
        OCL4(1) = OCL1(1)
        OCL4(2) = OCL3(2)
        
        OCM1(0) = parBendRadius * Tan(parBendAngle / 2) * Cos(parBendAngle) - (parwidth / 2) * Sin(parBendAngle)
        OCM1(1) = 0
        OCM1(2) = parBendRadius * Tan(parBendAngle / 2) * Sin(parBendAngle) + (parwidth / 2) * Cos(parBendAngle)
        
        OCM2(0) = parBendRadius * Tan(parBendAngle / 2) * Cos(parBendAngle) + (parwidth / 2) * Sin(parBendAngle)
        OCM2(1) = 0
        OCM2(2) = parBendRadius * Tan(parBendAngle / 2) * Sin(parBendAngle) - (parwidth / 2) * Cos(parBendAngle)
        Set objLineP1 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                            OCL4(0), OCL4(1), OCL4(2), OCL1(0), OCL1(1), OCL1(2))
                                             
        obottomcurves.Add objLineP1
        
        OCSP.Set OCL1(0), OCL1(1), OCL1(2)
        OCEP.Set OCL2(0), OCL2(1), OCL2(2)
        OCMP.Set OCM1(0), OCM1(1), OCM1(2)
        Set objBRA = PlaceTrArcBy3Pts(OCSP, OCEP, OCMP)
        
        obottomcurves.Add objBRA
        
        Set objLineP2 = geomFactory.Lines3d.CreateBy2Points(Nothing, _
                                         OCL2(0), OCL2(1), OCL2(2), OCL3(0), OCL3(1), OCL3(2))
                                         
        obottomcurves.Add objLineP2
        
        OCSP.Set OCL3(0), OCL3(1), OCL3(2)
        OCEP.Set OCL4(0), OCL4(1), OCL4(2)
        OCMP.Set OCM2(0), OCM2(1), OCM2(2)
        Set objBRA = PlaceTrArcBy3Pts(OCSP, OCEP, OCMP)
        
        obottomcurves.Add objBRA
        
        Dim objTorus1  As IngrGeom3D.ComplexString3d
        StartTorus.Set OCL4(0), OCL4(1), OCL4(2)
        Set objTorus1 = PlaceTrCString(StartTorus, obottomcurves)
        Set RefPlane1 = geomFactory.Planes3d.CreateByOuterBdry _
                                               (m_OutputColl.ResourceManager, objTorus1)
            
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), RefPlane1
        Dim ObjSurface As Object
        For Each ObjSurface In Surfset
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSurface
        Next ObjSurface
 End If

 Set Surfset = Nothing
 Set stPoint = Nothing
 Set oLineString = Nothing
 Set objTorussta = Nothing
 Set trObj = Nothing
 Set csObj = Nothing
 Set objBRA = Nothing
 Set objLineP1 = Nothing
 Set obottomcurves = Nothing
 
  
 Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
