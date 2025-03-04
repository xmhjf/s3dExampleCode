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
'   Author:          MS
'   Creation Date:  Tuesday, Sep 17 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRVT6, PDMS Category Name=IVPS)
'   The symbol consists of nine outputs, eight physical and one insulation outputs.
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
    
    Dim iOutput     As Double
    Dim ObjValveBody As Object
    Dim ObjCyl1Body As Object
    Dim ObjCyl2Body As Object
    Dim ObjCyl3Body As Object
    Dim ObjCyl4Body As Object
    Dim ObjValveBodyIns As Object
    
    Dim parFace1toCenter   As Double
    Dim parFace3toCenter   As Double
    Dim parInstrumentHeight    As Double
    Dim parInstrumentHeight1   As Double
    Dim parInstrumentLength    As Double
    Dim parInstrumentLength1   As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1    As Double
    Dim parInstrumentRadius    As Double
    Dim parInstrumentHeight2   As Double
    Dim parInstrumentHeight3   As Double
    Dim parInstrumentRadius1   As Double
    Dim parInstrumentHeight4   As Double
    Dim parInstrumentHeight5   As Double
    Dim parInstrumentRadius2   As Double
    Dim parInstrumentHeight6   As Double
    Dim parInstrumentHeight7   As Double
    Dim parInstrumentRadius3   As Double
    Dim parInstrumentHeight8   As Double
    Dim parInstrumentHeight9   As Double
    Dim parCylOffset   As Double
    Dim parCylOffset1  As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)     'D2
    parFace3toCenter = arrayOfInputs(3)     'D3
    parInstrumentHeight = arrayOfInputs(4)      'D4
    parInstrumentHeight1 = arrayOfInputs(5)     'D5
    parInstrumentLength = arrayOfInputs(6)      'D6
    parInstrumentLength1 = arrayOfInputs(7)     'D7
    parInstrumentWidth = arrayOfInputs(8)       'D8
    parInstrumentWidth1 = arrayOfInputs(9)      'D9
    parInstrumentRadius = arrayOfInputs(10)     'D10
    parInstrumentHeight2 = arrayOfInputs(11)    'D11
    parInstrumentHeight3 = arrayOfInputs(12)    'D12
    parInstrumentRadius1 = arrayOfInputs(13)    'D13
    parInstrumentHeight4 = arrayOfInputs(14)    'D14
    parInstrumentHeight5 = arrayOfInputs(15)    'D15
    parInstrumentRadius2 = arrayOfInputs(16)    'D16
    parInstrumentHeight6 = arrayOfInputs(17)    'D17
    parInstrumentHeight7 = arrayOfInputs(18)    'D18
    parInstrumentRadius3 = arrayOfInputs(19)    'D19
    parInstrumentHeight8 = arrayOfInputs(20)    'D20
    parInstrumentHeight9 = arrayOfInputs(21)    'D21
    parCylOffset = arrayOfInputs(22)    'D22
    parCylOffset1 = arrayOfInputs(23)   'D23
    parInsulationThickness = arrayOfInputs(24)
    
    iOutput = 0
    
'   The orgin is considered to be at intersection of the nozzle axis

' Insert your code for output 1(ValveBody)
    stPoint.Set -parInstrumentLength1, -parInstrumentHeight, -parInstrumentWidth
    enPoint.Set parInstrumentLength, parInstrumentHeight1, parInstrumentWidth1
    Set ObjValveBody = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
    Set ObjValveBody = Nothing

 ' Insert your code for output 2(Cyl1Body)
    stPoint.Set parCylOffset, _
                        parInstrumentHeight1 - (parInstrumentHeight2 - parInstrumentHeight3), _
                        -(parInstrumentWidth + parInstrumentRadius)
    enPoint.Set parCylOffset, _
                        parInstrumentHeight1 + parInstrumentHeight3, _
                        -(parInstrumentWidth + parInstrumentRadius)
    Set ObjCyl1Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * parInstrumentRadius, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl1Body
    Set ObjCyl1Body = Nothing

 ' Insert your code for output 3(Cyl2Body)
    stPoint.Set 0, _
                    parInstrumentHeight1 - (parInstrumentHeight4 - parInstrumentHeight5), _
                    -(parInstrumentWidth + parInstrumentRadius1)
    enPoint.Set 0, _
                        parInstrumentHeight1 + parInstrumentHeight5, _
                        -(parInstrumentWidth + parInstrumentRadius1)
    Set ObjCyl2Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * parInstrumentRadius1, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl2Body
    Set ObjCyl2Body = Nothing

 ' Insert your code for output 4(Cyl3Body)
    stPoint.Set -parCylOffset1, _
                    parInstrumentHeight1 - (parInstrumentHeight6 - parInstrumentHeight7), _
                    -(parInstrumentWidth + parInstrumentRadius2)
    enPoint.Set -parCylOffset1, _
                        parInstrumentHeight1 + parInstrumentHeight7, _
                    -(parInstrumentWidth + parInstrumentRadius2)
    Set ObjCyl3Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * parInstrumentRadius2, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl3Body
    Set ObjCyl3Body = Nothing
  
 ' Insert your code for output 5(Cyl4Body)
    stPoint.Set 0, _
                    parInstrumentHeight1 - (parInstrumentHeight8 - parInstrumentHeight9), _
                    parInstrumentWidth1 + parInstrumentRadius3
    enPoint.Set 0, _
                    parInstrumentHeight1 + parInstrumentHeight9, _
                    parInstrumentWidth1 + parInstrumentRadius3
    Set ObjCyl4Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * parInstrumentRadius3, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl4Body
    Set ObjCyl4Body = Nothing

' Place Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
        
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    Dim NozzleLength As Double
    NozzleLength = parFace1toCenter - parInstrumentLength1
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
'   Assumption face2 to center to be equal to face1 to center
    oPlacePoint.Set parFace1toCenter + sptOffset - depth, 0, 0
    oDir.Set 1, 0, 0
    NozzleLength = parFace1toCenter - parInstrumentLength
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 3
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    oPlacePoint.Set 0, -parFace3toCenter - sptOffset + depth, 0
    oDir.Set 0, -1, 0
    NozzleLength = parFace3toCenter - parInstrumentHeight
    Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)

'   Set the output
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
    Resume Next
    
End Sub
