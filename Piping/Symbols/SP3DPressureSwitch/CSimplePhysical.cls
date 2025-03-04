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
'**********************************************************************************
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   File:   SimplePhysical.cls
'   Author:         svsmylav
'   Creation Date:  Tuesday, Jul 4 2006
'   Description:
'       Pressure Switch (Source: Model 011 in Pressure Switch.pdf at www.ueonline.com)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CSimplePhysical" 'Used for error messages


Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
'
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs() As Variant, ByRef arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim ObjPressureSwitchBotBody As Object
    Dim ObjBotBodyConnector As Object
    Dim ObjPressureSwitchBody As Object
    Dim ObjPressureSwitchTopBody As Object
    Dim ObjNozzle1 As Object
    Dim ObjNozzle2 As Object
        
    Dim iOutput     As Double
    
    Dim parFacetoFace As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentHeight1 As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInstrumentHeight = arrayOfInputs(3)
    parInstrumentWidth = arrayOfInputs(4)
    parInstrumentLength = arrayOfInputs(5)
    parInstrumentHeight1 = arrayOfInputs(6)
    
    iOutput = 0

' Insert your code for output 1(Pressure Switch Bottom Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
    Dim dSqSide As Double
    Dim dBottomBodyLen As Double
    Dim dBottomBodyDia As Double

'   Assumptions:
'   1) Side of the Square (shown at bolted joints for High/Low ends) is
'   taken to be 1.5 time the pipe diameter
    dSqSide = 1.5 * pipeDiam

'   2) Length of the bottom joint is taken to be 70% of the side of the Square
    dBottomBodyLen = 0.7 * dSqSide
    
'   3) Diameter of the bottom body is taken to be 50% of the side of the Square
    dBottomBodyDia = 0.5 * dSqSide
    
'   4) Cylinder center is taken to be at 60 % of the side of the Square
    stPoint.Set 0, -dSqSide * 0.6, -dBottomBodyLen / 2
    enPoint.Set 0, -dSqSide * 0.6, dBottomBodyLen / 2
    
    Set ObjPressureSwitchBotBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, dBottomBodyDia, True)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPressureSwitchBotBody
    Set ObjPressureSwitchBotBody = Nothing
    
' Insert your code for output 2(Pressure Switch Bottom Body Connector)
    stPoint.Set -dBottomBodyDia / 2, -dSqSide * 0.6, -dBottomBodyLen / 2
    enPoint.Set dBottomBodyDia / 2, -dSqSide * 0.5, dBottomBodyLen / 2
    
    Set ObjBotBodyConnector = PlaceBox(m_OutputColl, stPoint, enPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBotBodyConnector
    Set ObjBotBodyConnector = Nothing
    
' Insert your code for output 3(Pressure Switch Body)
'   Assumption: Pressure Switch Body is taken to be equal to the side of the Square
    stPoint.Set -dSqSide * 0.5, -dSqSide * 0.5, -dSqSide * 0.5
    enPoint.Set dSqSide * 0.5, dSqSide * 0.5, dSqSide * 0.5
    
    Set ObjPressureSwitchBody = PlaceBox(m_OutputColl, stPoint, enPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPressureSwitchBody
    Set ObjPressureSwitchBody = Nothing

' Insert your code for output 4(Pressure Switch Top Body)
'   Assumption: Distance of far face (face having -Z axis as its normal) of Top body
'   is taken to be at 60% of the side of the Square
    stPoint.Set -parInstrumentLength / 2, dSqSide * 0.5, -dSqSide * 0.6
    enPoint.Set stPoint.x + parInstrumentLength, parInstrumentHeight1, _
                    stPoint.z + parInstrumentWidth
    Set ObjPressureSwitchTopBody = PlaceBox(m_OutputColl, stPoint, enPoint)
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPressureSwitchTopBody
    Set ObjPressureSwitchTopBody = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Dim dNozzleLength As Double
    dNozzleLength = (parFacetoFace - dSqSide) / 2
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, dNozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext

End Sub
