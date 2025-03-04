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
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CPhysical.cls
'   Author:         RAH
'   Creation Date:  December 2005
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       This is a Generic Olet symbol. It takes inputs from the Generic Data tables.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private PI           As Double
Private RAD          As Double

Private Const MODULE = "SimplePhysical:" 'Used for error messages

Private Sub Class_Initialize()

    PI = 4 * Atn(1)
    RAD = 180 / PI

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjOletBody As Object
    Dim ObjBody As Object
    
    Dim parA As Double
    Dim parFlangelength As Double
    Dim parNippleLength As Double
    
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parA = arrayOfInputs(2)
    parFlangelength = arrayOfInputs(3)
    parNippleLength = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0

' Insert your code for output 1(olet body)
    Dim dMajorBodyDia As Double
    Dim dHoleinHeader As Double 'Branch Outside Pipe Diameter(This is equivalent to PipeDia of Branch)
    
    'get the Header data
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, _
                        flangeThick, flangeDiam, cptOffset, depth
    
    'get branch data
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, _
                        flangeThick, flangeDiam, cptOffset, depth
                        
    dHoleinHeader = pipeDiam2
    dMajorBodyDia = 1.5 * dHoleinHeader
    
    Dim oLinesCol       As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim oComplexStr     As IngrGeom3D.ComplexString3d
    
    'do some calcs to get the point positions
    Dim oPoint1 As New AutoMath.DPosition
    Dim oPoint2 As New AutoMath.DPosition
    Dim oPoint3 As New AutoMath.DPosition
    
    Dim dLength As Double
    Dim dHeight As Double
    
    'set the start point in the center of the pipe
    Dim dVerysmalldistance As Double
    dVerysmalldistance = 0.0001
    oPoint1.Set dVerysmalldistance, 0, 0
    
    dLength = (dMajorBodyDia - dHoleinHeader) / 2#
    dHeight = dLength / Tan(40# / RAD)
    
    oPoint2.Set dMajorBodyDia / 2, dHeight + pipeDiam1 / 2, 0
    oPoint3.Set dHoleinHeader / 2, parA + pipeDiam1 / 2, 0
    
    Set oLinesCol = New Collection
    
    Set oLine = PlaceTrLine(oPoint1, oPoint2)
    oLinesCol.Add oLine
    Set oLine = PlaceTrLine(oPoint2, oPoint3)
    oLinesCol.Add oLine
    
    Set oComplexStr = PlaceTrCString(oPoint1, oLinesCol)
    
    Set oLine = Nothing
    Dim iCount As Integer
    For iCount = 1 To oLinesCol.Count
        oLinesCol.Remove 1
    Next iCount
    Set oLinesCol = Nothing
    
    Dim oCenterPoint   As New AutoMath.DPosition
    Dim oRevolutionAxis As New AutoMath.DVector
    
    oCenterPoint.Set 0, pipeDiam1 / 2, 0
    oRevolutionAxis.Set 0, 1, 0
    Set ObjOletBody = PlaceRevolution(m_OutputColl, oComplexStr, _
                                    oRevolutionAxis, oCenterPoint, 2 * PI, True)
    Set oComplexStr = Nothing
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOletBody
    Set ObjOletBody = Nothing

    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, cptOffset, depth

 ' Insert your code for output 1(Nipple Body)
    Dim Port1Position   As New AutoMath.DPosition
    Dim Port2Position   As New AutoMath.DPosition
    
    Port1Position.Set 0, parA + (pipeDiam1 / 2), 0
    Port2Position.Set 0, parA + (pipeDiam1 / 2) + parNippleLength, 0
    
    Set ObjBody = PlaceCylinder(m_OutputColl, Port1Position, Port2Position, pipeDiam2, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
 ' Insert your code for output 1(Flange Body)
    
    Port1Position.Set 0, parA + (pipeDiam1 / 2) + parNippleLength, 0
    Port2Position.Set 0, parA + (pipeDiam1 / 2) + parNippleLength + parFlangelength - flangeThick, 0
    
    Set ObjBody = PlaceCone(m_OutputColl, Port1Position, Port2Position, pipeDiam2 / 2, pipeDiam2 / 2 + flangeDiam / 8)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBody
    Set ObjBody = Nothing
    
' Place Nozzle 1
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set 0, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set 0, parA + (pipeDiam1 / 2) + parNippleLength + parFlangelength + cptOffset - depth, 0
    oDir.Set 0, 1, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set ObjOletBody = Nothing
    Set ObjBody = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    Set Port1Position = Nothing
    Set Port2Position = Nothing
    Set oCenterPoint = Nothing
    Set oRevolutionAxis = Nothing
    
    Set oLine = Nothing
    Set oGeomFactory = Nothing
    Set oComplexStr = Nothing
    
    Set oPoint1 = Nothing
    Set oPoint2 = Nothing
    Set oPoint3 = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
