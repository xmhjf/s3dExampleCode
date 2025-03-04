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
'   Creation Date:  Friday, Oct 4 2002
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IPDA2,
'   PDMS Category Name=IVCJ). Physical aspect consist of two cones and a cylinder as a part
'   valve body, Stem cylinder, Two boxes , two cylinders as part of operator and Two Nozzles.
'   Insulation aspect consist of single cylinder for valve body.
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
    Dim sptOffset2       As Double
    Dim flangeDiam      As Double
    Dim depth1           As Double
    Dim depth2           As Double
    
    Dim iOutput     As Double
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Dim ObjStemBody As Object
    Dim ObjLargeBoxBody As Object
    Dim ObjSmallBoxBody As Object
    Dim ObjCylinder1Body As Object
    Dim ObjCylinder2Body As Object
    Dim ObjBodyCone1  As Object
    Dim ObjBodyCone2  As Object
    Dim ObjValCyl  As Object
        
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
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
    Dim parInsulationThickness As Double

    Dim StemDia As Double
     
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)       'D2
    parFacetoFace = arrayOfInputs(3)       'D3
    parCylHeight = arrayOfInputs(4)      'D4
    parActuatorHeight = arrayOfInputs(5)       'D5
    parActuatorHeight1 = arrayOfInputs(6)      'D6
    parActuatorWidth = arrayOfInputs(7)    'D7
    parActuatorWidth1 = arrayOfInputs(8)       'D8
    parActuatorLength = arrayOfInputs(9)       'D9
    parActuatorLength1 = arrayOfInputs(10)      'D10
    parActuatorHeight2 = arrayOfInputs(11)      'D11
    parActuatorLength2 = arrayOfInputs(12)      'D12
    parActuatorLength3 = arrayOfInputs(13)     'D13
    parActuatorWidth2 = arrayOfInputs(14)      'D14
    parActuatorWidth3 = arrayOfInputs(15)      'D15
    parActuatorOffset = arrayOfInputs(16)      'D16
    parActuatorCylDiameter = arrayOfInputs(17)    'D17
    parActuatorCylLength = arrayOfInputs(18)      'D18
    parActuatorHeight3 = arrayOfInputs(19)     'D19
    parActuatorCyl1Diameter = arrayOfInputs(20)    'D20
    parActuatorCyl1Length = arrayOfInputs(21)      'D21
    parInsulationThickness = arrayOfInputs(22)
    
    iOutput = 0
    StemDia = 0.05
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
' Origin is assumed to be at the center of the valveBody along the flow axis
  ' Insert your code for output 1(Body Cone1 of Valve)
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
    stPoint.Set -(parFace1toCenter - flangeThick), 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone1
    Set ObjBodyCone1 = Nothing
    
 ' Insert your code for output 2(Body Cone2 of Valve)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset2, depth2
    stPoint.Set (parFacetoFace - parFace1toCenter) - flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    Set ObjBodyCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBodyCone2
    Set ObjBodyCone2 = Nothing
    
 ' Insert your code for output 3(Bottom Valve Cylinder)
    Dim CylDia As Double
    stPoint.Set 0, 0, 0
    enPoint.Set 0, -parCylHeight, 0
    CylDia = parFace1toCenter
    Set ObjValCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
 
 ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValCyl
    Set ObjValCyl = Nothing
        
' Insert your code for output 4(Stem Body)
    stPoint.Set 0, 0, 0
    enPoint.Set 0, parActuatorHeight3 - parActuatorHeight, 0
    Set ObjStemBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemBody
    Set ObjStemBody = Nothing
    
 ' Insert your code for output 5(Large Box Body)
    stPoint.Set -parActuatorLength, _
                        parActuatorHeight3 - parActuatorHeight, _
                        parActuatorWidth + parActuatorOffset
    enPoint.Set parActuatorLength1, _
                        parActuatorHeight3 + parActuatorHeight1, _
                        -(parActuatorWidth1 - parActuatorOffset)
    Set ObjLargeBoxBody = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLargeBoxBody
    Set ObjLargeBoxBody = Nothing
    
 ' Insert your code for output 6(Small Box Body)
    stPoint.Set -parActuatorLength2, _
                        parActuatorHeight3 + parActuatorHeight1, _
                        parActuatorWidth2
    enPoint.Set parActuatorLength3, _
                        parActuatorHeight3 + parActuatorHeight1 + parActuatorHeight2, _
                        -parActuatorWidth3
    Set ObjSmallBoxBody = PlaceBox(m_OutputColl, stPoint, enPoint)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjSmallBoxBody
    Set ObjSmallBoxBody = Nothing
    
 ' Insert your code for output 7(Cylinder1 Body)
    stPoint.Set -parActuatorLength, parActuatorHeight3, parActuatorOffset
    enPoint.Set -parActuatorLength - parActuatorCylLength, parActuatorHeight3, parActuatorOffset
    Set ObjCylinder1Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, parActuatorCylDiameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder1Body
    Set ObjCylinder1Body = Nothing
    
 ' Insert your code for output 8(Cylinder2 Body)
    stPoint.Set parActuatorLength1, parActuatorHeight3, parActuatorOffset
    enPoint.Set parActuatorLength1 + parActuatorCyl1Length, parActuatorHeight3, parActuatorOffset
    Set ObjCylinder2Body = PlaceCylinder(m_OutputColl, stPoint, enPoint, parActuatorCyl1Diameter, True)
 
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder2Body
    Set ObjCylinder2Body = Nothing
    
' Place Nozzle 1 for Output 9(Nozzle1)
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    
    oPlacePoint.Set -parFace1toCenter - sptOffset1 + depth1, 0, 0
    oDir.Set -1, 0, 0
    Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2 for output 10 (Nozzle2)
    oPlacePoint.Set (parFacetoFace - parFace1toCenter) + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0
    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
' Set the output
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
