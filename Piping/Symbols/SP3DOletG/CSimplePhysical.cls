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
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CPhysical.cls
'   Author:         ACM
'   Creation Date:  Monday, Feb 9, 2004
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'       This is a Generic Olet symbol. It takes inputs from the Generic Data tables.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC      DI-95670  Replace names with initials in all revision history sheets and symbols
'  28.SEP.2006     svsmylav TR-97093  Updated the symbol inline with the F173 (PDS EDEN).
'  18.APR.2007     KKC  CR-117893 Modified Olet symbol to be Trimmable
'  29.Jun.2007     svsmylav/dkl  Assigned parAngle = PI / 2 'For V6.1 symbol.
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
    Dim parAngle As Double
    
    Dim iOutput     As Double
    Dim ObjOletBody As Object
    
    Dim parA As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parA = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    parAngle = PI / 2 'For V6.1 symbol
    If UBound(arrayOfInputs) >= 4 Then 'Not to fail the prior to V7 instances by accessing index 31
        parAngle = arrayOfInputs(4)
    End If
        
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
                        
    If CmpDblLessThan(pipeDiam2, flangeDiam) Then
       dHoleinHeader = flangeDiam
    Else
       dHoleinHeader = pipeDiam2
    End If
    
    ''Waist Dia is is assumed as 1.5 times the Branch PipeOD as per the PDS Eden 'F-173'
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
    
    Dim dOffset As Double
    Dim dOlet_body As Double
    Dim dOlet_body_1 As Double
    Dim dOlet_body_2 As Double

    If CmpDblGreaterthan(dHoleinHeader, pipeDiam1) Then
        dOffset = 0.125 * pipeDiam1
    Else
        dOffset = 0.5 * Sqr(pipeDiam1 ^ 2 - dHoleinHeader ^ 2)
    End If
    dOlet_body = pipeDiam1 / 2 + parA - dOffset
    dOlet_body_1 = 0.4 * dOlet_body
    dOlet_body_2 = dOlet_body - dOlet_body_1

    oPoint1.Set dHoleinHeader / 2, dOffset, 0
    oPoint2.Set dMajorBodyDia / 2, oPoint1.y + dOlet_body_1, 0
    oPoint3.Set dHoleinHeader / 2, oPoint2.y + dOlet_body_2, 0

    Set oLinesCol = New Collection
    
    Set oLine = PlaceTrLine(oPoint1, oPoint2)
    oLinesCol.Add oLine
    Set oLine = PlaceTrLine(oPoint2, oPoint3)
    oLinesCol.Add oLine
    
    Set oComplexStr = PlaceTrCString(oPoint1, oLinesCol)
    
    Set oLine = Nothing
    Set oLinesCol = Nothing
    
    Dim oCenterPoint   As New AutoMath.DPosition
    Dim oRevolutionAxis As New AutoMath.DVector
    
    oCenterPoint.Set 0, dOffset + dOlet_body / 2, 0
    oRevolutionAxis.Set 0, 1, 0
    Set ObjOletBody = PlaceRevolution(m_OutputColl, oComplexStr, _
                                    oRevolutionAxis, oCenterPoint, 2 * PI, True)
                                    
    Dim oDirVector As AutoMath.DVector
    Dim oTransformationMat  As New AutoMath.DT4x4
    Dim dRotAboutZaxis As Double
    Set oDirVector = New AutoMath.DVector
    
    oTransformationMat.LoadIdentity
    dRotAboutZaxis = PI / 2 - parAngle
    oDirVector.Set 0, 0, -1
    oTransformationMat.Rotate dRotAboutZaxis, oDirVector
    ObjOletBody.Transform oTransformationMat
    
    Set oComplexStr = Nothing

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjOletBody
    Set ObjOletBody = Nothing
    Set oTransformationMat = Nothing
    Set oDirVector = Nothing
    
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
    oPlacePoint.Set (parA + (pipeDiam1 / 2) + cptOffset - depth) * Cos(parAngle), _
                    (parA + (pipeDiam1 / 2) + cptOffset - depth) * Sin(parAngle), 0
    oDir.Set Cos(parAngle), Sin(parAngle), 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

