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
'   Author:          svsmylav 
'   Creation Date:  Tuesday, Oct 22 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IPDA5,
'   PDMS Category Name=IVCH). The symbol consists of Eleven outputs,Ten physical outputs
'   and One Insulation Output.
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

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset1       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    
    Dim sptOffset2       As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim ObjBodyCone1 As Object
    Dim ObjBodyCone2 As Object
    Dim ObjValCyl As Object
    Dim ObjStem As Object
    Dim ObjLargeBox As Object
    Dim ObjSmallBox As Object
    Dim ObjCylinder1 As Object
    Dim ObjCylinder2 As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight  As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parActuatorLength As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorHeight2 As Double
    Dim parActuatorLength2 As Double
    Dim parActuatorLength3 As Double
    Dim parActuatorWidth2 As Double
    Dim parActuatorWidth3 As Double
    Dim parActuatorOffset As Double
    Dim parActuatorCylDiameter As Double
    Dim parActuatorCylLength As Double
    Dim parActuatorHeight3 As Double
    Dim parActuatorCyl1Diameter As Double
    Dim parActuatorCyl1Length As Double
    Dim parActuatorOffset1 As Double
    Dim parInsulationThickness As Double

    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    Dim StemDia    As Double
 
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) 'D2
    parFacetoFace = arrayOfInputs(3) 'D3
    parCylHeight = arrayOfInputs(4) 'D4
    parActuatorHeight = arrayOfInputs(5) 'D5
    parActuatorHeight1 = arrayOfInputs(6) 'D6
    parActuatorWidth = arrayOfInputs(7) 'D7
    parActuatorWidth1 = arrayOfInputs(8) 'D8
    parActuatorLength = arrayOfInputs(9) 'D9
    parActuatorLength1 = arrayOfInputs(10) 'D10
    parActuatorHeight2 = arrayOfInputs(11) 'D11
    parActuatorLength2 = arrayOfInputs(12) 'D12
    parActuatorLength3 = arrayOfInputs(13) 'D13
    parActuatorWidth2 = arrayOfInputs(14) 'D14
    parActuatorWidth3 = arrayOfInputs(15) 'D15
    parActuatorOffset = arrayOfInputs(16) 'D16
    parActuatorCylDiameter = arrayOfInputs(17) 'D17
    parActuatorCylLength = arrayOfInputs(18) 'D18
    parActuatorHeight3 = arrayOfInputs(19) 'D19
    parActuatorCyl1Diameter = arrayOfInputs(20) 'D20
    parActuatorCyl1Length = arrayOfInputs(21) 'D21
    parActuatorOffset1 = arrayOfInputs(22) 'D23
    parInsulationThickness = arrayOfInputs(23)
    
    iOutput = 0

    
'   Stem diameter is taken as 0.05
    StemDia = 0.05

' Insert your code for output 2(Valve Cone1)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
    stPoint.Set -(parFace1toCenter - flangeThick), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
 ' Insert your code for output 3(Body Cone2 of Valve)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset2, depth2
    stPoint.Set (parFacetoFace - parFace1toCenter) - flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 4(Valve Cylinder)
    Dim CylDia As Double
    stPoint.Set 0, 0, 0
    enPoint.Set 0, -parCylHeight, 0
    CylDia = parFace1toCenter
    Set ObjValCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
    
' Insert your code for output 5(Stem)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parActuatorHeight3 - parActuatorHeight, 0
    Set ObjStem = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStem
    Set ObjStem = Nothing
    
' Insert your code for output 6(LargeBox)
    stPoint.Set -parActuatorLength, parActuatorHeight3 - parActuatorHeight, -parActuatorWidth1
    enPoint.Set parActuatorLength1, parActuatorHeight3 + parActuatorHeight1, parActuatorWidth
    Set ObjLargeBox = PlaceBox(m_OutputColl, stPoint, enPoint)
     
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLargeBox
    Set ObjLargeBox = Nothing
    
' Insert your code for output 7(SmallBox)
    stPoint.Set -parActuatorLength2, parActuatorHeight3 + parActuatorHeight1, -parActuatorWidth3
    enPoint.Set parActuatorLength3, parActuatorHeight3 + parActuatorHeight1 + parActuatorHeight2, _
                    parActuatorWidth2
     Set ObjSmallBox = PlaceBox(m_OutputColl, stPoint, enPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSmallBox
    Set ObjSmallBox = Nothing
    
' Insert your code for output 8(Cylinder1)
    stPoint.Set -parActuatorLength - parActuatorCylLength, _
                        parActuatorHeight3, parActuatorOffset1
    enPoint.Set -parActuatorLength, parActuatorHeight3, parActuatorOffset1
    Set ObjCylinder1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parActuatorCylDiameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder1
    Set ObjCylinder1 = Nothing
    
' Insert your code for output 9(Cylinder2)
    stPoint.Set parActuatorLength1, parActuatorHeight3, -parActuatorOffset
    enPoint.Set parActuatorLength1 + parActuatorCyl1Length, _
                    parActuatorHeight3, -parActuatorOffset
    Set ObjCylinder2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parActuatorCyl1Diameter, True)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder2
    Set ObjCylinder2 = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oPlacePoint.Set -parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
 
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Place Nozzle 2
    oPlacePoint.Set parFacetoFace - parFace1toCenter + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
