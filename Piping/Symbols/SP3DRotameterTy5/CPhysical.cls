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
'   Author:          BG
'   Creation Date:  Tuesday, Sep 17 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR5,
' PDMS Category Name=IFNK). Physical aspect consist of one vertical cylinder, One Horizontal Box
' and two nozzles, one facing -ive and the other facing +ive direction of X-axis.
' Insulation aspect consist of a cylinder along Y-axis and the Box along X-axis
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
        
    Dim ObjRotaMtrHorBox As Object
    Dim ObjRotaMtrVerCylin As Object
    
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentHeight3 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInsulationThickness As Double
    Dim iOutput     As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)   'D2
    parFace2toCenter = arrayOfInputs(3)  'D3
    parInstrumentHeight = arrayOfInputs(4)   'D4
    parInstrumentHeight1 = arrayOfInputs(5)      'D5
    parInstrumentDiameter = arrayOfInputs(6)      'D6
    parInstrumentHeight2 = arrayOfInputs(7)      'D7
    parInstrumentHeight3 = arrayOfInputs(8)      'D8
    parInstrumentWidth = arrayOfInputs(9)        'D9
    parInstrumentWidth1 = arrayOfInputs(10)      'D10
    parInsulationThickness = arrayOfInputs(11)
    
    iOutput = 0
' Insert your code for output 1(Rotameter Vertical Cylinder)
    Dim startPoint As AutoMath.DPosition
    Dim endPoint As AutoMath.DPosition
    Set startPoint = New AutoMath.DPosition
    Set endPoint = New AutoMath.DPosition
    
'Origin of the symbol is assumed at the intersection of vertical cylinder axis and nozzle  axes
    startPoint.Set 0, -parInstrumentHeight, 0
    endPoint.Set 0, parInstrumentHeight1 - parInstrumentHeight, 0
    Set ObjRotaMtrVerCylin = PlaceCylinder(m_OutputColl, startPoint, endPoint, parInstrumentDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRotaMtrVerCylin
    Set ObjRotaMtrVerCylin = Nothing

 ' Insert your code for output 2(RotaMeter Horizontal Box)
    startPoint.Set -parFace1toCenter, -parInstrumentHeight2, parInstrumentWidth
    endPoint.Set parFace2toCenter, parInstrumentHeight3, -parInstrumentWidth1
    Set ObjRotaMtrHorBox = PlaceBox(m_OutputColl, startPoint, endPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRotaMtrHorBox
    Set ObjRotaMtrHorBox = Nothing

' Place Nozzle 1 for Output 4 (Nozzle1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2 for output 5 (Nozzle2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set parFace2toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
