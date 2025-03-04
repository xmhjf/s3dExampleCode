VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Wednesday 19, Apr 2006
'   Description:
'    This symbol is Pipe Sleeve with Puddle Flange (Details taken from
'    product database at www.doyma.de, Curaflex Pipe Sleeve 9000 Technical Information, Doyma)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
      PI = 4 * Atn(1)
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double

    Dim ObjPenetrationSleeve As Object

    Dim iOutput     As Double

    Dim parInnerDiameter As Double
    Dim parPuddleFlangeDiameter As Double
    Dim parPipingSpecialtyLength As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInnerDiameter = arrayOfInputs(2) 'D1
    parPuddleFlangeDiameter = arrayOfInputs(3) 'D3
    parPipingSpecialtyLength = arrayOfInputs(4) 'L1

    iOutput = 0

'   Origin is assumed to be at center of coupling.

' Insert your code for output 1(Coupling)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptOffset, depth

    Dim dPoints(0 To 26) As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oCenterPoint As AutoMath.DPosition
    Set oCenterPoint = New AutoMath.DPosition
    
    oCenterPoint.Set 0, 0, 0

'   Assumptions:-
'   1. Origin is taken to be at center of parPipingSpecialtyLength value
'   2. Section of Penetration Sleeve and Puddle Flange are taken to be
'   10 mm and 15 mm respectively
    Dim dPenetrationSlvThk As Double
    Dim dPuddleFlangeThk As Double
    
    dPenetrationSlvThk = 0.01
    dPuddleFlangeThk = 0.015

    dPoints(0) = oCenterPoint.x - parPipingSpecialtyLength / 2
    dPoints(1) = oCenterPoint.y + parInnerDiameter / 2
    dPoints(2) = oCenterPoint.z
        
    dPoints(3) = oCenterPoint.x + parPipingSpecialtyLength / 2
    dPoints(4) = oCenterPoint.y + parInnerDiameter / 2
    dPoints(5) = oCenterPoint.z
        
    dPoints(6) = oCenterPoint.x + parPipingSpecialtyLength / 2
    dPoints(7) = oCenterPoint.y + parInnerDiameter / 2 + dPenetrationSlvThk
    dPoints(8) = oCenterPoint.z
        
    dPoints(9) = oCenterPoint.x + dPuddleFlangeThk / 2
    dPoints(10) = oCenterPoint.y + parInnerDiameter / 2 + dPenetrationSlvThk
    dPoints(11) = oCenterPoint.z
        
    dPoints(12) = oCenterPoint.x + dPuddleFlangeThk / 2
    dPoints(13) = oCenterPoint.y + parPuddleFlangeDiameter / 2
    dPoints(14) = oCenterPoint.z
        
    dPoints(15) = oCenterPoint.x - dPuddleFlangeThk / 2
    dPoints(16) = oCenterPoint.y + parPuddleFlangeDiameter / 2
    dPoints(17) = oCenterPoint.z
        
    dPoints(18) = oCenterPoint.x - dPuddleFlangeThk / 2
    dPoints(19) = oCenterPoint.y + parInnerDiameter / 2 + dPenetrationSlvThk
    dPoints(20) = oCenterPoint.z
        
    dPoints(21) = oCenterPoint.x - parPipingSpecialtyLength / 2
    dPoints(22) = oCenterPoint.y + parInnerDiameter / 2 + dPenetrationSlvThk
    dPoints(23) = oCenterPoint.z
        
    dPoints(24) = dPoints(0)
    dPoints(25) = dPoints(1)
    dPoints(26) = dPoints(2)

    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim axisVect    As AutoMath.DVector
    Set axisVect = New AutoMath.DVector

    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 9, dPoints)
    axisVect.Set 1, 0, 0
    Set ObjPenetrationSleeve = PlaceRevolution(m_OutputColl, oLineString, axisVect, oCenterPoint, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPenetrationSleeve
    Set ObjPenetrationSleeve = Nothing
    Set geomFactory = Nothing
    Set axisVect = Nothing
    Set oCenterPoint = Nothing
    Set oLineString = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
