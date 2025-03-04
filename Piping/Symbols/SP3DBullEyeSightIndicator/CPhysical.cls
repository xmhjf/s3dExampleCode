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
'   Copyright (c) 2003-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          SSP
'   Creation Date:  Friday, Nov 15 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Bull Eye Sight Indicator.Refering www.resistoflex.com
'   The symbol consists of 18 outputs, 16 physical and 2 insulation outputs.
'
'   Change History:
'   dd.mmm.yyyy     who                   change description
'   -----------     -----                 ------------------
'   09.Jul.2003     SymbolTeam(India)     Copyright Information, Header  is added.
'   08.SEP.2006      KKC                  DI-95670  Replace names with initials in all revision history sheets and symbols
'   27.Aug.2007      VRG                  TR-124959  Insulation for ports was ensured for different end preparations
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
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
    Dim depth2          As Double
    
    Dim iOutput     As Integer
    Dim ObjCyl1 As Object
    Dim ObjCyl2 As Object
    Dim ObjFlange1 As Object
    Dim ObjFlange2 As Object
    Dim ObjFlange3 As Object
    Dim ObjFlange4 As Object
    Dim ObjBolt1 As Object
    Dim ObjBolt2 As Object
    Dim ObjBolt3 As Object
    Dim ObjBolt4 As Object
    Dim ObjBolt5 As Object
    Dim ObjBolt6 As Object
    Dim ObjBolt7 As Object
    Dim ObjBolt8 As Object
    
    Dim parFacetoCenter As Double
    Dim parIndicatorHeight As Double
    Dim parInsulationThickness As Double
    
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parIndicatorHeight = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
    
    iOutput = 0
    
'  Insert your code for output 1(Cylinder1)
'  Origin of the symbol is intersection of Nozzle axis and vertical axis
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, sptOffset2, depth2
    
'   The below four checks are kept to ensure the symbol doesn't fail for beveled ends.
    If CmpDblEqual(flangeThick1, 0) Then
            flangeThick1 = pipeDiam1 / 10
    End If
    If CmpDblEqual(flangeThick2, 0) Then
            flangeThick2 = pipeDiam2 / 10
    End If
    If CmpDblLessThanOrEqualTo(flangeDiam1, pipeDiam1) Then
            flangeDiam1 = 1.2 * pipeDiam1
    End If
    If CmpDblLessThanOrEqualTo(flangeDiam2, pipeDiam2) Then
            flangeDiam2 = 1.2 * pipeDiam2
    End If
    
    stPoint.Set -parFacetoCenter, 0, 0
    enPoint.Set parFacetoCenter, 0, 0
    
    Set ObjCyl1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl1
    Set ObjCyl1 = Nothing
    
 ' Insert your code for output 2(Cylinder2)
    'Assumption
    'Flange Thickness is FlangeThick1
    'Top Flange starts at 0.85 times parIndicatorHeight.
    'Gap between two flanges is taken to be 6 mm.
    Dim SightIndicatorThk As Double
    SightIndicatorThk = 0.006
    stPoint.Set 0, (0.85 * parIndicatorHeight - SightIndicatorThk) - 2 * flangeThick1, 0
    enPoint.Set 0, -(0.85 * parIndicatorHeight - SightIndicatorThk) + 2 * flangeThick1, 0
    
    Set ObjCyl2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, pipeDiam1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCyl2
    Set ObjCyl2 = Nothing
 
 ' Insert your code for output 3(Flange1)
    stPoint.Set 0, 0.85 * parIndicatorHeight, 0
    enPoint.Set 0, 0.85 * parIndicatorHeight - flangeThick1, 0
    
    Set ObjFlange1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, flangeDiam1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlange1
    Set ObjFlange1 = Nothing
    
 ' Insert your code for output 4(Flange2)
    stPoint.Set 0, 0.85 * parIndicatorHeight - SightIndicatorThk - flangeThick1, 0
    enPoint.Set 0, 0.85 * parIndicatorHeight - SightIndicatorThk - 2 * flangeThick1, 0
    
    Set ObjFlange2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, flangeDiam1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlange2
    Set ObjFlange2 = Nothing
    
 ' Insert your code for output 5(Flange3)
    stPoint.Set 0, -0.85 * parIndicatorHeight, 0
    enPoint.Set 0, -0.85 * parIndicatorHeight + flangeThick1, 0
    
    Set ObjFlange3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, flangeDiam1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlange3
    Set ObjFlange3 = Nothing
    
 ' Insert your code for output 6(Flange4)
    stPoint.Set 0, -0.85 * parIndicatorHeight + SightIndicatorThk + flangeThick1, 0
    enPoint.Set 0, -0.85 * parIndicatorHeight + SightIndicatorThk + 2 * flangeThick1, 0
    
    Set ObjFlange4 = PlaceCylinder(m_OutputColl, stPoint, enPoint, flangeDiam1, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFlange4
    Set ObjFlange4 = Nothing
    
' Insert your code for output 7(Bolt1)
    Dim BoltDia As Double
    Dim PlacePoint As Double   'PlacePoint is the point where Bolt is getting placed
' Assumption
' Boltdia =0.01 is taken
    BoltDia = 0.01
    
    PlacePoint = 0.8 * flangeDiam1 / 2
    
    stPoint.Set 0, parIndicatorHeight, PlacePoint
    enPoint.Set 0, 0.7 * parIndicatorHeight - SightIndicatorThk - 2 * flangeThick1, PlacePoint
    
    Set ObjBolt1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, BoltDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBolt1
    Set ObjBolt1 = Nothing

' Insert your code for output 8(Bolt2)
    stPoint.Set PlacePoint, parIndicatorHeight, 0
    enPoint.Set PlacePoint, 0.7 * parIndicatorHeight - SightIndicatorThk - 2 * flangeThick1, 0
    
    Set ObjBolt2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, BoltDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBolt2
    Set ObjBolt2 = Nothing
    
' Insert your code for output 9(Bolt3)
    stPoint.Set 0, parIndicatorHeight, -PlacePoint
    enPoint.Set 0, 0.7 * parIndicatorHeight - SightIndicatorThk - 2 * flangeThick1, -PlacePoint
    
    Set ObjBolt3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, BoltDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBolt3
    Set ObjBolt3 = Nothing
    
' Insert your code for output 10(Bolt4)
    stPoint.Set -PlacePoint, parIndicatorHeight, 0
    enPoint.Set -PlacePoint, 0.7 * parIndicatorHeight - SightIndicatorThk - 2 * flangeThick1, 0
    
    Set ObjBolt4 = PlaceCylinder(m_OutputColl, stPoint, enPoint, BoltDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBolt4
    Set ObjBolt4 = Nothing
    
' Insert your code for output 11(Bolt5)
    stPoint.Set 0, -parIndicatorHeight, PlacePoint
    enPoint.Set 0, -0.7 * parIndicatorHeight + SightIndicatorThk + 2 * flangeThick1, PlacePoint
    
    Set ObjBolt5 = PlaceCylinder(m_OutputColl, stPoint, enPoint, BoltDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBolt5
    Set ObjBolt5 = Nothing
    
' Insert your code for output 12(Bolt6)
    stPoint.Set PlacePoint, -parIndicatorHeight, 0
    enPoint.Set PlacePoint, -0.7 * parIndicatorHeight + SightIndicatorThk + 2 * flangeThick1, 0
    
    Set ObjBolt6 = PlaceCylinder(m_OutputColl, stPoint, enPoint, BoltDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBolt6
    Set ObjBolt6 = Nothing
    
' Insert your code for output 13(Bolt7)
    stPoint.Set 0, -parIndicatorHeight, -PlacePoint
    enPoint.Set 0, -0.7 * parIndicatorHeight + SightIndicatorThk + 2 * flangeThick1, -PlacePoint
    
    Set ObjBolt7 = PlaceCylinder(m_OutputColl, stPoint, enPoint, BoltDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBolt7
    Set ObjBolt7 = Nothing
    
' Insert your code for output 14(Bolt8)
    stPoint.Set -PlacePoint, -parIndicatorHeight, 0
    enPoint.Set -PlacePoint, -0.7 * parIndicatorHeight + SightIndicatorThk + 2 * flangeThick1, 0
    
    Set ObjBolt8 = PlaceCylinder(m_OutputColl, stPoint, enPoint, BoltDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBolt8
    Set ObjBolt8 = Nothing

' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFacetoCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parFacetoCenter + sptOffset2 - depth2, 0, 0
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

