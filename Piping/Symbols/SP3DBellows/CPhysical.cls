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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RRK
'   Creation Date:  Thursday, Oct 11 2007

'   Description: This symbol is Bellows of MDC Vacuum Products
'   Source: http://www.mdcvacuum.com/urd/uniface.urd/ecf0070w.display?1.8.3.3

'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   11-Oct-07               CR-127648: Provide bellows symbol for vacuum piping
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "CPhysical:" 'Used for error messages

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
    
    Dim iOutput     As Integer
    
    Dim parLength As Double
    Dim parLength2 As Double
    Dim parDiameter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parLength = arrayOfInputs(2)
    parLength2 = arrayOfInputs(3)
    parDiameter = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0
    
' Place Bellows(Output 1)
    Dim objBellow   As Object
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition

    oStPoint.Set -parLength / 2, 0, 0
    oEnPoint.Set parLength / 2, 0, 0
    
    Set objBellow = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objBellow
    Set objBellow = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
       
' Place Nozzle 1(Output 1)
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim dNozzleLength As Double
    
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
     
    oPlacePoint.Set -parLength / 2 - parLength2 - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    
    dNozzleLength = parLength2
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                            oPlacePoint, dNozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2 (Output 2)
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parLength / 2 + parLength2 + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0

    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                            oPlacePoint, dNozzleLength)
                            
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
