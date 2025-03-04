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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         ACM
'   Creation Date:  Monday, July 12, 2004
'   Description:
'      This is a Generic Pipet symbol. Technical details are taken from the site http://www.wfi-intl.com/bwp/bwp1.html
'
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
    Dim pipeDiam1        As Double ''Header diameter
    Dim pipeDiam2        As Double ''Branch diameter
    Dim flangeThick     As Double
    Dim cptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjPipetBody As Object
    
    Dim parA As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parA = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    iOutput = 0

' Insert your code for output 1(Pipet body)
    Dim dMajorBodyDia As Double ''diameter at waist
    Dim dHoleinHeader As Double ''Branch Outside Pipe Diameter(This is equivalent to PipeDia of Branch)
    Dim dStepHeight As Double ''Height of Stepped portion at the waist.

    'get the Header(Jacketed Pipe) data
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, _
                        flangeThick, flangeDiam, cptOffset, depth
                        
    If pipeDiam1 < flangeDiam Then
       dHoleinHeader = flangeDiam
    Else
       dHoleinHeader = pipeDiam1
    End If
    
    'get branch data
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, _
                        flangeThick, flangeDiam, cptOffset, depth
                        
        
    ''Waist Dia is is assumed as 1.5 times the Branch Pipe and Step Height is 0.25 times the Branch Dia
    dMajorBodyDia = 1.5 * dHoleinHeader
    dStepHeight = 0.25 * dHoleinHeader

    Dim oLinesCol       As Collection
    Dim oLine           As IngrGeom3D.Line3d
    Dim oGeomFactory    As IngrGeom3D.GeometryFactory
    Dim oComplexStr     As IngrGeom3D.ComplexString3d
    
    'Complex curve points
    Dim oPoint1 As New AutoMath.DPosition
    Dim oPoint2 As New AutoMath.DPosition
    Dim oPoint3 As New AutoMath.DPosition
    Dim oPoint4 As New AutoMath.DPosition
    
    Dim dLength As Double
    Dim dHeight As Double
    
    'set the start point in the center of the pipe
    Dim dVerysmalldistance As Double ''which keeps the line a small distance away from th axis of revolution of complex string.
    dVerysmalldistance = 0.0001
    oPoint1.Set dVerysmalldistance, 0, 0
    
    dLength = (dMajorBodyDia - dHoleinHeader) / 2#
    dHeight = dLength / Tan(40# / RAD)
    
    oPoint2.Set dMajorBodyDia / 2, dHeight + pipeDiam1 / 2, 0
    oPoint3.Set dMajorBodyDia / 2, dHeight + dStepHeight + pipeDiam1 / 2, 0
    oPoint4.Set dHoleinHeader / 2, parA + pipeDiam1 / 2, 0
    
    Set oLinesCol = New Collection
    
    Set oLine = PlaceTrLine(oPoint1, oPoint2)
    oLinesCol.Add oLine
    Set oLine = PlaceTrLine(oPoint2, oPoint3)
    oLinesCol.Add oLine
    Set oLine = PlaceTrLine(oPoint3, oPoint4)
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
    Set ObjPipetBody = PlaceRevolution(m_OutputColl, oComplexStr, _
                                    oRevolutionAxis, oCenterPoint, 2 * PI, True)
    Set oComplexStr = Nothing
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipetBody
    Set ObjPipetBody = Nothing
  
    Set oComplexStr = Nothing
    
' Place Nozzle 1 ''Jacketed Pipe
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
     
' Place Nozzle 2 ''Branch pipe
    oPlacePoint.Set 0, parA + (pipeDiam1 / 2) + cptOffset - depth, 0
    oDir.Set 0, 1, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oPoint1 = Nothing
    Set oPoint2 = Nothing
    Set oPoint3 = Nothing
    Set oPoint4 = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

