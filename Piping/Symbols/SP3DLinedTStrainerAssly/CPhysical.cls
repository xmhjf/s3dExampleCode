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
'   Author:          ACM
'   Creation Date:  Wednesday, Nov 20 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Lined Strainer Tee Assembly. This symbol is taken from www.resistoflex.com web catalog.
'   The symbol consists of 7 outputs, 5 physical and 2 insulation outputs.
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
    Dim pipeDiam1        As Double
    Dim pipeDiam2        As Double
    Dim flangeThick1     As Double
    Dim flangeThick2     As Double
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim flangeDiam1      As Double
    Dim flangeDiam2      As Double
    Dim depth1           As Double
    Dim depth2           As Double
    Dim depth3          As Double
    
    Dim iOutput     As Double
    Dim ObjVerticalCyl As Object
    Dim ObjBlindFlange1 As Object
    Dim ObjBlindFlange2 As Object
    
    Dim parFaceToCenter As Double
    Dim parStrainerHeight As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFaceToCenter = arrayOfInputs(2)
    parStrainerHeight = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)

    iOutput = 0


' Origin of the symbol is intersection of Nozzle axis and vertical axis
' Gap between two Blind Flanges is 0.1 times of Strainer Height.

' Insert your code for output 1(VerticalCyl)
' Vertical Cylinder dia is taken as Pipe dia at CP1
    
    Dim GasketLinerThk As Double
    GasketLinerThk = 0.002
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, sptOffset1, depth1

    stPoint.Set 0, (parStrainerHeight - GasketLinerThk) - (2 * flangeThick1), 0
    enPoint.Set 0, 0, 0
    
    Set ObjVerticalCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam1, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjVerticalCyl
    Set ObjVerticalCyl = Nothing
    
' Insert your code for output 2(BlindFlange1)
' Blind flange1 dia is taken as flange dia at CP1
    
    stPoint.Set 0, (parStrainerHeight - GasketLinerThk) - (2 * flangeThick1), 0
    enPoint.Set 0, (parStrainerHeight - GasketLinerThk) - flangeThick1, 0
    
    Set ObjBlindFlange1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, flangeDiam1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBlindFlange1
    Set ObjBlindFlange1 = Nothing
    
' Insert your code for output 3(BlindFlange2)
' Blind flange2 dia is taken as flange dia at CP1

    stPoint.Set 0, parStrainerHeight - flangeThick1, 0
    enPoint.Set 0, parStrainerHeight, 0
    
    Set ObjBlindFlange2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, flangeDiam1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBlindFlange2
    Set ObjBlindFlange2 = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim NozzleLength As Double
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                                                                    sptOffset2, depth2

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFaceToCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    
    NozzleLength = 2 * parFaceToCenter - flangeThick2
    
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, NozzleLength)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parFaceToCenter + sptOffset2 - depth2, 0, 0
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
    Resume Next
    
End Sub

