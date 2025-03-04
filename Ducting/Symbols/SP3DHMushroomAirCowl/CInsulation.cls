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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         pkalluri
'   Creation Date:  Monday,July 24 2007
'   Description  :  Mushroom Type Air Cowl
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    PI = 4 * Atn(1)
    
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt              As PartFacelets.IJDPart
    Dim parWidth               As Double
    Dim parHVACShape           As Double
    Dim parHeight1             As Double
    Dim parHeight2             As Double
    Dim parCowlDiameter        As Double
    Dim parHeight3             As Double
    Dim parInsulationThickness As Double
    Dim iOutput                As Integer
    Dim dDepth                 As Double
     
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)              'a
    parHVACShape = arrayOfInputs(3)          'HVACShape = 4(round)
    parHeight1 = arrayOfInputs(4)            'F
    parHeight2 = arrayOfInputs(5)            'b
    parCowlDiameter = arrayOfInputs(6)       'd1
    parHeight3 = arrayOfInputs(7)  '200mm or 300mm
    parInsulationThickness = arrayOfInputs(8)
    iOutput = 0
    
            
    'Insert your code for output 1(Insulation Cowl Body)
    Dim lines           As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim oCStr           As IngrGeom3D.ComplexString3d
    Dim oArc            As IngrGeom3D.Arc3d
    Dim oAxis           As New AutoMath.DVector
    Dim oRevolution     As Object
    Dim iCount          As Integer
    Dim oStPoint        As New AutoMath.DPosition
    Dim oCentPoint      As New AutoMath.DPosition
    Dim oFixPoint       As New AutoMath.DPosition
    Set lines = New Collection
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    oFixPoint.Set 0.1 + 0.5 * (parHeight3 + parHeight1) + (0.5 * parHeight3 + parHeight2 - 0.5 * parHeight1) * (parCowlDiameter + 2 * parInsulationThickness) / parCowlDiameter, -0.5 * parCowlDiameter - parInsulationThickness, 0

    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0.1, 0, 0, 0.1, -0.5 * parCowlDiameter - parInsulationThickness, 0)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0.1, -0.5 * parCowlDiameter - parInsulationThickness, 0, oFixPoint.x, oFixPoint.y, oFixPoint.z)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oFixPoint.x, oFixPoint.y, oFixPoint.z, 0.1 + parHeight1 + parHeight3 + parInsulationThickness, -0.0625 * (parHeight3 + parHeight1 + 2 * parInsulationThickness) / (parHeight3 + parHeight1), 0)
    lines.Add oLine
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, 0.1 + parHeight1 + parHeight3 + parInsulationThickness, -0.0625 * (parHeight3 + parHeight1 + 2 * parInsulationThickness) / (parHeight3 + parHeight1), 0, 0.1 + parHeight1 + parHeight3 + parInsulationThickness, 0, 0)
    lines.Add oLine

    oStPoint.Set 0.1, 0, 0
    Set oCStr = PlaceTrCString(oStPoint, lines)
    For iCount = 1 To lines.Count
        lines.Remove 1
    Next iCount
    oAxis.Set 1, 0, 0
    oCentPoint.Set 0.5 * parHeight2, 0, 0
    Set oRevolution = PlaceRevolution(m_OutputColl, oCStr, oAxis, oCentPoint, 2 * PI, True)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oRevolution
    
    Set oStPoint = Nothing
    Set oFixPoint = Nothing
    Set oAxis = Nothing
    Set oCentPoint = Nothing
    Set oLine = Nothing
    Set oArc = Nothing
    Set lines = Nothing
    Set oCStr = Nothing
    Set oRevolution = Nothing
    Set oGeomFactory = Nothing
    
 
    Exit Sub
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
        Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
   ''
End Sub




 


                    







