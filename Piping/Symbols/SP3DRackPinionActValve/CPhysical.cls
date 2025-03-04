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
'   Author:          SSP
'   Creation Date:  Tuesday, Oct 22 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'  The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IRXPA).
'  Physical aspect consist 1 Cylinder, 2 Cones ,1 stem ,3 boxes,2 nozzles and 1 insulation aspect outputs.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'   06.Dec.2004     SymbolTeam(India)       Removed the small value (0.0000001) that was 
'                                           added to the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam1       As Double
    Dim pipeDiam2        As Double
    Dim flangeThick     As Double
    Dim sptOffset1       As Double
    Dim sptOffset2       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjValCylinder As Object
    Dim ObjConvergingCone As Object
    Dim ObjDivergingCone As Object
    Dim ObjStem As Object
    Dim ObjBox1 As Object
    Dim ObjBox2 As Object
    Dim ObjBox3 As Object
    
    Dim parFace1toCenter As Double              'D2
    Dim parFacetoFace As Double                 'D3
    Dim parCylHeight As Double               'D4
    Dim parActuatorHeight As Double             'D5
    Dim parActuatorHeight1 As Double            'D6
    Dim parActuatorHeight2 As Double            'D7
    Dim parActuatorLength As Double             'D8
    Dim parActuatorLength1 As Double            'D9
    Dim parActuatorWidth As Double              'D10
    Dim parActuatorWidth1 As Double             'D11
    Dim parActuatorLength2 As Double            'D12
    Dim parActuatorLength3 As Double            'D13
    Dim parActuatorWidth2 As Double             'D14
    Dim parActuatorWidth3 As Double             'D15
    Dim parActuatorWidth4 As Double             'D16
    Dim parActuatorLength4 As Double            'D17
    Dim parActuatorLength5 As Double            'D18
    Dim parActuatorHeight3 As Double            'D19
    Dim parActuatorHeight4 As Double            'D20
    Dim parActuatorHeight5 As Double            'D21
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFacetoFace = arrayOfInputs(3)
    parCylHeight = arrayOfInputs(4)
    parActuatorHeight = arrayOfInputs(5)
    parActuatorHeight1 = arrayOfInputs(6)
    parActuatorHeight2 = arrayOfInputs(7)
    parActuatorLength = arrayOfInputs(8)
    parActuatorLength1 = arrayOfInputs(9)
    parActuatorWidth = arrayOfInputs(10)
    parActuatorWidth1 = arrayOfInputs(11)
    parActuatorLength2 = arrayOfInputs(12)
    parActuatorLength3 = arrayOfInputs(13)
    parActuatorWidth2 = arrayOfInputs(14)
    parActuatorWidth3 = arrayOfInputs(15)
    parActuatorWidth4 = arrayOfInputs(16)
    parActuatorLength4 = arrayOfInputs(17)
    parActuatorLength5 = arrayOfInputs(18)
    parActuatorHeight3 = arrayOfInputs(19)
    parActuatorHeight4 = arrayOfInputs(20)
    parActuatorHeight5 = arrayOfInputs(21)
    parInsulationThickness = arrayOfInputs(22)
    
    iOutput = 0
    
' Insert your code for output 1(ConvergingCone)
    Dim pPos1   As New AutoMath.DPosition
    Dim pPos2   As New AutoMath.DPosition

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick, flangeDiam, sptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick, flangeDiam, sptOffset2, depth2
    pPos1.Set -parFace1toCenter, 0, 0
    pPos2.Set 0, 0, 0
    
    Set ObjConvergingCone = PlaceCone(m_OutputColl, pPos1, pPos2, pipeDiam1 / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjConvergingCone
    Set ObjConvergingCone = Nothing
    
' Insert your code for output 2(DivergingCone)
    pPos1.Set parFacetoFace - parFace1toCenter, 0, 0
    pPos2.Set 0, 0, 0

    Set ObjDivergingCone = PlaceCone(m_OutputColl, pPos1, pPos2, pipeDiam2 / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDivergingCone
    Set ObjDivergingCone = Nothing

' Insert your code for output 3(ValCylinder)
    pPos1.Set 0, 0, 0
    pPos2.Set 0, -parCylHeight, 0
    
    Dim ValCylDiam As Double
    'Assumption
'    ValCylDiam = parFace1toCenter is taken
    ValCylDiam = parFace1toCenter
    
    Set ObjValCylinder = PlaceCylinder(m_OutputColl, pPos1, pPos2, ValCylDiam, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCylinder
    Set ObjValCylinder = Nothing
    
' Insert your code for output 4(Stem)
    Dim StemDia As Double
    pPos1.Set 0, 0, 0
    pPos2.Set 0, parActuatorHeight, 0
' Assumption
' StemDia = 0.05 are taken
    StemDia = 0.05
    
    Set ObjStem = PlaceCylinder(m_OutputColl, pPos1, pPos2, StemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
' Insert your code for output 5(Box1)
    pPos1.Set -parActuatorLength, parActuatorHeight, parActuatorWidth
    pPos2.Set parActuatorLength1, parActuatorHeight + parActuatorHeight1, _
                -parActuatorWidth1
    
    Set ObjBox1 = PlaceBox(m_OutputColl, pPos1, pPos2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox1
    Set ObjBox1 = Nothing
    
' Insert your code for output 6(Box2)
    pPos1.Set -parActuatorLength2, parActuatorHeight + parActuatorHeight1, parActuatorWidth2
    pPos2.Set parActuatorLength3, _
                parActuatorHeight + parActuatorHeight1 + parActuatorHeight2, _
                -parActuatorWidth3
    
    Set ObjBox2 = PlaceBox(m_OutputColl, pPos1, pPos2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox2
    Set ObjBox2 = Nothing
    
' Insert your code for output 7(Box3)
    pPos1.Set -parActuatorLength4, _
                parActuatorHeight + parActuatorHeight5 - parActuatorHeight3, _
                parActuatorWidth + parActuatorWidth4
    pPos2.Set parActuatorLength5, _
                parActuatorHeight + parActuatorHeight5 + parActuatorHeight4, _
                parActuatorWidth
    
    Set ObjBox3 = PlaceBox(m_OutputColl, pPos1, pPos2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBox3
    Set ObjBox3 = Nothing
       
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle

    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parFacetoFace - parFace1toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set pPos1 = Nothing
    Set pPos2 = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

