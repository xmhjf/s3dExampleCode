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
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Tuesday, May 29 2001
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   24.Oct.2003     SymbolTeam(India)       TR 51102 Removed duplicated Header/Copyright Information.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptoffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput     As Double
    Dim ObjFloat As Object
    Dim ObjOrifice As Object
    
    Dim parFacetoFace As Double
    Dim parBasetoInlet As Double
    Dim parBasetoOutlet As Double
    Dim parOutletOffset As Double
    Dim parSTHeight As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parBasetoInlet = arrayOfInputs(3)
    parBasetoOutlet = arrayOfInputs(4)
    parOutletOffset = arrayOfInputs(5)
    parSTHeight = arrayOfInputs(6)
    
    iOutput = 0

 ' Insert your code for output 1(Float Inlet Chamber)
''    Dim Stfloat       As New AutoMath.DPosition
''    Dim Endfloat       As New AutoMath.DPosition
''    Stfloat.Set -parFacetoFace / 2, 0, parSTHeight / 2
''    Endfloat.Set 0, 0, parSTHeight / 2
''    Set ObjFloat = PlaceCylinder(m_OutputColl, Stfloat, Endfloat, parSTHeight, True)
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim ObjFloatEllipse As Object
    Set ObjFloatEllipse = geomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
    -parFacetoFace / 2, 0, parSTHeight / 2, parFacetoFace, 0, 0, 0, 0, parSTHeight / 2, 0.65)
    
    Dim EllipseNormal       As New AutoMath.DPosition
    EllipseNormal.Set parFacetoFace / 2, 0, 0
    Set ObjFloat = geomFactory.Projections3d.CreateByCurve(m_OutputColl.ResourceManager, _
                                                        ObjFloatEllipse, _
                                                        EllipseNormal.x, EllipseNormal.y, EllipseNormal.z, _
                                                        parFacetoFace / 2, True)

    Set ObjFloatEllipse = Nothing
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFloat
    Set ObjFloat = Nothing
    Set geomFactory = Nothing
 ' Insert your code for output 2(Valve Outlet Orifice)
    Dim StOrifice       As New AutoMath.DPosition
    Dim EndOrifice       As New AutoMath.DPosition
   
    StOrifice.Set 0, 0, parSTHeight * 2 / 5
    EndOrifice.Set parFacetoFace / 2, 0, parSTHeight * 2 / 5
    Set ObjOrifice = PlaceCylinder(m_OutputColl, StOrifice, EndOrifice, parSTHeight * 4 / 5, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOrifice
    Set ObjOrifice = Nothing
    
' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim faceToFace  As Double

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    faceToFace = arrayOfInputs(2)
    oPlacePoint.Set -faceToFace / 2 - cptoffset + depth, 0, parBasetoInlet
    oDir.Set -1, 0, 0
    Set oPartFclt = arrayOfInputs(1)
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    oPlacePoint.Set faceToFace / 2 + cptoffset - depth, -parOutletOffset, parBasetoOutlet
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
    oPlacePoint.Set -faceToFace / 5, 0, parSTHeight + cptoffset - depth
    oDir.Set 0, 0, 1

    Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
