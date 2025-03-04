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
'   Author:         RRK
'   Creation Date:  Wednesday, Jul 5 2006
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      Short-stroke Rotameter (Source: Rotameter housing type 90 in Rotameters.pdf at www.yokogawa-usa.com)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
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
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    Dim ObjRotameterBody As Object
    Dim ObjInstrumentBody As Object
    
    Dim parFacetoFace As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentWidth As Double
    Dim parOffset As Double
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInstrumentLength = arrayOfInputs(3)
    parInstrumentDiameter = arrayOfInputs(4)
    parInstrumentWidth = arrayOfInputs(5)
    parOffset = arrayOfInputs(6)
    
    iOutput = 0
    
' The origin is considered to be along the flow axis and at the midpoint of the line joining two ports

' Insert your code for output 1(Rotameter Body)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0
    Set ObjRotameterBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRotameterBody
    Set ObjRotameterBody = Nothing
    
 ' Insert your code for output 2(Instrument)
'   Finding X cordinate of start point of instrument(output2)
    Dim dXcoordinate As Double 'stores X cordinate of start point of instrument
    dXcoordinate = -((parFacetoFace / 2) - parInstrumentLength - _
                        Sqr((parInstrumentDiameter / 2) ^ 2 - (parOffset ^ 2)))
    
    stPoint.Set dXcoordinate, -parOffset, pipeDiam / 2
    enPoint.Set dXcoordinate, -parOffset, (pipeDiam / 2) + parInstrumentWidth
    Set ObjInstrumentBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInstrumentBody
    Set ObjInstrumentBody = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFacetoFace / 2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFacetoFace / 2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
