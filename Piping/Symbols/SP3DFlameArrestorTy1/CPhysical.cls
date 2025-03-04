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
'   Author:          Sundar
'   Creation Date:  Tuesday, Jun 4 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No.: S1AC or S1AZ, Page No. D-99 PDS Piping Component Data Reference Guide.
' SP3DFlameArrestorTy1 Symbol is created with Eight Outputs - Five are Insulation aspect outputs.
' Other three are physical aspect outputs created as follows:
'   ObjFABody - Using 'PlaceBox' function and
'   two ObjNozzle objects  by using 'CreateNozzleWithLength' function.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
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
    Dim ObjFABody As Object
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOffset1 As Double
    Dim parOffset2 As Double
    Dim parFlArrestorBodyLength1 As Double
    Dim parFlArrestorBodyLength2 As Double
    Dim parFlArrestorBodyHeight1 As Double
    Dim parFlArrestorBodyHeight2 As Double
    Dim parFlArrestorBodyWidth1 As Double
    Dim parFlArrestorBodyWidth2 As Double
    Dim parInsulationThickness As Double
    
    Dim Point1   As New AutoMath.DPosition
    Dim Point2   As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parOffset1 = arrayOfInputs(4)
    parOffset2 = arrayOfInputs(5)
    parFlArrestorBodyLength1 = arrayOfInputs(6)
    parFlArrestorBodyLength2 = arrayOfInputs(7)
    parFlArrestorBodyHeight1 = arrayOfInputs(8)
    parFlArrestorBodyHeight2 = arrayOfInputs(9)
    parFlArrestorBodyWidth1 = arrayOfInputs(10)
    parFlArrestorBodyWidth2 = arrayOfInputs(11)
    parInsulationThickness = arrayOfInputs(12)

    iOutput = 0

' Insert your code for output 6(Flame Arrestor Body)
    Point1.Set -parFlArrestorBodyLength1, -parFlArrestorBodyWidth1, -parFlArrestorBodyHeight1
    Point2.Set parFlArrestorBodyLength2, parFlArrestorBodyWidth2, parFlArrestorBodyHeight2
    Set ObjFABody = PlaceBox(m_OutputColl, Point1, Point2)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFABody
    Set ObjFABody = Nothing
    Set Point1 = Nothing
    Set Point2 = Nothing
    
' Place Nozzle 1
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    Dim NozzleLength As Double

    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, -parOffset1
    oDir.Set -1, 0, 0
    NozzleLength = parFace1toCenter - parFlArrestorBodyLength1
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, parOffset2
    oDir.Set 1, 0, 0
    NozzleLength = parFace2toCenter - parFlArrestorBodyLength2
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
