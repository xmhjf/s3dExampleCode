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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:        RRK
'   Creation Date: Wednesday 5, May 2007
'   Description:
'    This symbol is created for  JIS F 7810 Pipe Sleeve as per CR-118307
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
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

    Dim ObjPipeSleeve As Object

    Dim iOutput     As Double

    Dim parInnerDiameter As Double
    Dim parPipingSpecialtyLength As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parPipingSpecialtyLength = arrayOfInputs(2)
    parInnerDiameter = arrayOfInputs(3)
    
    iOutput = 0

' Insert your code for output 1(Sleeve)

    Dim dPoints(0 To 14) As Double
    Dim oLineString As IngrGeom3D.LineString3d
    Dim oCenterPoint As AutoMath.DPosition
    Set oCenterPoint = New AutoMath.DPosition
    
'   Assumptions:-
'   1. Origin is taken to be at center of parPipingSpecialtyLength value
'   2. Section of Pipe Sleeve thickness is assumed to be 10mm
    
    oCenterPoint.Set 0, 0, 0
    
    Dim dPipeSleeveThk As Double
    dPipeSleeveThk = 0.01
    

    dPoints(0) = oCenterPoint.x - parPipingSpecialtyLength / 2
    dPoints(1) = oCenterPoint.y + parInnerDiameter / 2
    dPoints(2) = oCenterPoint.z
        
    dPoints(3) = oCenterPoint.x + parPipingSpecialtyLength / 2
    dPoints(4) = oCenterPoint.y + parInnerDiameter / 2
    dPoints(5) = oCenterPoint.z
        
    dPoints(6) = oCenterPoint.x + parPipingSpecialtyLength / 2
    dPoints(7) = oCenterPoint.y + parInnerDiameter / 2 + dPipeSleeveThk
    dPoints(8) = oCenterPoint.z
        
    dPoints(9) = oCenterPoint.x - parPipingSpecialtyLength / 2
    dPoints(10) = oCenterPoint.y + parInnerDiameter / 2 + dPipeSleeveThk
    dPoints(11) = oCenterPoint.z
        
    dPoints(12) = dPoints(0)
    dPoints(13) = dPoints(1)
    dPoints(14) = dPoints(2)
        
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Dim axisVect    As AutoMath.DVector
    Set axisVect = New AutoMath.DVector

    Set oLineString = geomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dPoints)
    axisVect.Set 1, 0, 0
    Set ObjPipeSleeve = PlaceRevolution(m_OutputColl, oLineString, axisVect, oCenterPoint, 2 * PI, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipeSleeve
    Set ObjPipeSleeve = Nothing
    Set geomFactory = Nothing
    Set axisVect = Nothing
    Set oCenterPoint = Nothing
    Set oLineString = Nothing

    Exit Sub

ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
