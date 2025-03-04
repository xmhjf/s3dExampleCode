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
'   Creation Date:  Friday, Oct 25 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
''The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IPR4A),If this is
'     rotated through 180 deg about flow axis then we get (PDS Eden Name = IPL4D) symbol.
'   The symbol consists of Nine outputs,Eight physical outputs
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
    Dim sptOffset2       As Double
    Dim flangeDiam      As Double
    Dim depth1          As Double
    Dim depth2          As Double
    
    Dim iOutput     As Double
    Dim ObjValveCone1 As Object
    Dim ObjValveCone2 As Object
    Dim ObjStemRod As Object
    Dim ObjValveCylinder As Object
    Dim ObjActuatorRect As Object
    Dim ObjActuatorCyl As Object
    Dim ObjInsulationBody As Object
    
    Dim parFace1toCenter As Double
    Dim parFacetoFace As Double
    Dim parCylHeight As Double
    Dim parActuatorLength As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorHeight As Double
    Dim parActuatorHeight1 As Double
    Dim parActuatorHeight2 As Double
    Dim parActuatorWidth As Double
    Dim parActuatorWidth1 As Double
    Dim parActuatorLength1 As Double
    Dim parActuatorLength2 As Double
    Dim parActuatorWidth2 As Double
    Dim parInsulationThickness As Double
    

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2) ''D2
    parFacetoFace = arrayOfInputs(3) ''D3''
    parCylHeight = arrayOfInputs(4) ''D4
    parActuatorLength = arrayOfInputs(5) ''D5
    parActuatorDiameter = arrayOfInputs(6) ''D6
    parActuatorHeight = arrayOfInputs(7) ''D7
    parActuatorHeight1 = arrayOfInputs(8) ''D10
    parActuatorHeight2 = arrayOfInputs(9) ''D11
    parActuatorWidth = arrayOfInputs(10) ''D12
    parActuatorWidth1 = arrayOfInputs(11) ''D14
    parActuatorLength1 = arrayOfInputs(12) ''D25
    parActuatorLength2 = arrayOfInputs(13) ''D26
    parActuatorWidth2 = arrayOfInputs(14) ''D28
    parInsulationThickness = arrayOfInputs(15)
    
    iOutput = 0

 ' Insert your code for output 1(Conical Valve Body 1)
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset1, depth1
         
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
         
    stPoint.Set -parFace1toCenter + flangeThick, 0, 0
    enPoint.Set 0, 0, 0
    
    Set ObjValveCone1 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveCone1
    Set ObjValveCone1 = Nothing
    
 ' Insert your code for output 2(Conical Valve Body 2 )
   
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset2, depth2
         
    stPoint.Set parFacetoFace - parFace1toCenter - flangeThick, 0, 0
    enPoint.Set 0, 0, 0
        
    Set ObjValveCone2 = PlaceCone(m_OutputColl, stPoint, enPoint, pipeDiam / 2, 0, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveCone2
    Set ObjValveCone2 = Nothing
    
 ' Insert your code for output 3(Stem Body )
   Dim StemDia As Double
   
   StemDia = 0.05
   stPoint.Set 0, 0, 0
   enPoint.Set 0, parActuatorHeight - parActuatorHeight2, 0
 
   Set ObjStemRod = PlaceCylinder(m_OutputColl, stPoint, enPoint, StemDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjStemRod
    Set ObjStemRod = Nothing
    
 ' Insert your code for output 4(Valve Cyledrical Body)
    Dim CylDia As Double
    stPoint.Set 0, 0, 0
    enPoint.Set 0, -parCylHeight, 0
    CylDia = parFace1toCenter
    Set ObjValveCylinder = PlaceCylinder(m_OutputColl, stPoint, enPoint, CylDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveCylinder
    Set ObjValveCylinder = Nothing
    
 ' Insert your code for output 5(Actuator Rectangular box )
 
    stPoint.Set -parActuatorLength2, parActuatorHeight + parActuatorHeight1, _
                                -(parActuatorWidth1 + parActuatorWidth2)
    enPoint.Set parActuatorLength1, parActuatorHeight - parActuatorHeight2, _
                                    (parActuatorWidth - parActuatorWidth2)
    
   Set ObjActuatorRect = PlaceBox(m_OutputColl, stPoint, enPoint)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorRect
    Set ObjActuatorRect = Nothing
    
 ' Insert your code for output 6(Actuator Cylendrical Body)
   
   stPoint.Set parActuatorLength1, parActuatorHeight, -parActuatorWidth2
   enPoint.Set parActuatorLength, parActuatorHeight, -parActuatorWidth2
   
   Set ObjActuatorCyl = PlaceCylinder(m_OutputColl, stPoint, enPoint, parActuatorDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjActuatorCyl
    Set ObjActuatorCyl = Nothing
    
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
    oPlacePoint.Set (parFacetoFace - parFace1toCenter) + sptOffset2 - depth2, 0, 0
    oDir.Set 1, 0, 0

    Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

    Set stPoint = Nothing
    Set enPoint = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing

    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

