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
'   PDMS Category Name=IVPQ). Physical aspect consist of Relief valve body of Box shape, 3Cylinders
'   in the (-)ive side of Z axis,1 Cylinder in (+)ive side of Z-axis, Nozzle-1 along (-)ive side of X-Axis and
'   Nozzle-2 along (-)ive side of Y-axis.
'   Insulation aspect consist of a single box place covering the whole valve.
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
    
    Dim ObjRefValBody As Object
    Dim ObjVerCylin1 As Object
    Dim ObjVerCylin2 As Object
    Dim ObjVerCylin3  As Object
    Dim ObjVerCylin4  As Object
    
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parInstrumentRadius As Double
    Dim parInstrumentHeight2 As Double
    Dim parInstrumentHeight3 As Double
    Dim parInstrumentRadius1 As Double
    Dim parInstrumentHeight4 As Double
    Dim parInstrumentHeight5 As Double
    Dim parInstrumentRadius2 As Double
    Dim parInstrumentHeight6 As Double
    Dim parInstrumentHeight7 As Double
    Dim parInstrumentRadius3 As Double
    Dim parInstrumentHeight8 As Double
    Dim parInstrumentHeight9 As Double
    Dim parCylOffset As Double
    Dim parCylOffset1 As Double
    Dim parInsulationThickness As Double
    Dim iOutput     As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)             'D2
    parFace2toCenter = arrayOfInputs(3)             'D3
    parInstrumentHeight = arrayOfInputs(4)          'D4
    parInstrumentHeight1 = arrayOfInputs(5)         'D5
    parInstrumentLength = arrayOfInputs(6)          'D6
    parInstrumentLength1 = arrayOfInputs(7)         'D7
    parInstrumentWidth = arrayOfInputs(8)           'D8
    parInstrumentWidth1 = arrayOfInputs(9)          'D9
    parInstrumentRadius = arrayOfInputs(10)         'D10
    parInstrumentHeight2 = arrayOfInputs(11)        'D11
    parInstrumentHeight3 = arrayOfInputs(12)        'D12
    parInstrumentRadius1 = arrayOfInputs(13)        'D13
    parInstrumentHeight4 = arrayOfInputs(14)        'D14
    parInstrumentHeight5 = arrayOfInputs(15)        'D15
    parInstrumentRadius2 = arrayOfInputs(16)        'D16
    parInstrumentHeight6 = arrayOfInputs(17)        'D17
    parInstrumentHeight7 = arrayOfInputs(18)        'D18
    parInstrumentRadius3 = arrayOfInputs(19)        'D19
    parInstrumentHeight8 = arrayOfInputs(20)        'D20
    parInstrumentHeight9 = arrayOfInputs(21)        'D21
    parCylOffset = arrayOfInputs(22)        'D22
    parCylOffset1 = arrayOfInputs(23)        'D23
    parInsulationThickness = arrayOfInputs(24)

    iOutput = 0
    
    Dim stPoint As AutoMath.DPosition
    Dim enPoint As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition
    
'Insert your code for output 1(Relief Valve Body)
'Origin of the symbol is assumed at the intersection of Nozzle-1 and Nozzle-2 Axes
    
    ''Left Top corner point
    stPoint.Set parInstrumentLength, parInstrumentHeight1, -parInstrumentWidth
    ''Right bottom corner point
    enPoint.Set -parInstrumentLength1, -parInstrumentHeight, parInstrumentWidth1

    Set ObjRefValBody = PlaceBox(m_OutputColl, stPoint, enPoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRefValBody
    Set ObjRefValBody = Nothing
    
' Insert your code for output 2(Vertical Cylinder 1)
    
    stPoint.Set parCylOffset, parInstrumentHeight1 + parInstrumentHeight3 - parInstrumentHeight2, _
                            -(parInstrumentWidth + parInstrumentRadius)
    enPoint.Set parCylOffset, parInstrumentHeight1 + parInstrumentHeight3, _
                            -(parInstrumentWidth + parInstrumentRadius)
                            
    Set ObjVerCylin1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentRadius * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerCylin1
    Set ObjVerCylin1 = Nothing
    
' Insert your code for output 3(Vertical Cylinder 2)
    stPoint.Set 0, parInstrumentHeight1 + parInstrumentHeight5 - parInstrumentHeight4, _
                            -(parInstrumentWidth + parInstrumentRadius1)
    enPoint.Set 0, parInstrumentHeight1 + parInstrumentHeight5, _
                            -(parInstrumentWidth + parInstrumentRadius1)
                            
    Set ObjVerCylin2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentRadius1 * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerCylin2
    Set ObjVerCylin2 = Nothing

' Insert your code for output 3(Vertical Cylinder 3)
    stPoint.Set -parCylOffset1, parInstrumentHeight1 + parInstrumentHeight7 - parInstrumentHeight6, _
                            -(parInstrumentWidth + parInstrumentRadius2)
    enPoint.Set -parCylOffset1, parInstrumentHeight1 + parInstrumentHeight7, _
                            -(parInstrumentWidth + parInstrumentRadius2)
    
    Set ObjVerCylin3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, 2 * parInstrumentRadius2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerCylin3
    Set ObjVerCylin3 = Nothing

' Insert your code for output 3(Vertical Cylinder 4)
    stPoint.Set 0, parInstrumentHeight1 + parInstrumentHeight9 - parInstrumentHeight8, _
                           (parInstrumentWidth1 + parInstrumentRadius3)
    enPoint.Set 0, parInstrumentHeight1 + parInstrumentHeight9, _
                           (parInstrumentWidth1 + parInstrumentRadius3)
                           
    Set ObjVerCylin4 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInstrumentRadius3 * 2, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerCylin4
    Set ObjVerCylin4 = Nothing

' Place Nozzle 1 for Output 4 (Nozzle1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Dim NozzleLength As Double
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset + depth, 0, 0
    oDir.Set -1, 0, 0
    NozzleLength = parFace1toCenter - parInstrumentLength1         'D2-D7
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2 for output 5 (Nozzle2)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    oPlacePoint.Set 0, -parFace2toCenter - sptOffset + depth, 0                       'D3
    oDir.Set 0, -1, 0
    NozzleLength = parFace2toCenter - parInstrumentHeight         'D3-D4
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
    
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
    Resume Next
    
End Sub
