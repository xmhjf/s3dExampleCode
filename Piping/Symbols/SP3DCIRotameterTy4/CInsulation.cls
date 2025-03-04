VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:          MS
'   Creation Date:  Friday, 3rd Jan 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   The symbol is prepared based on INtools Instrument Components(PDS Eden Name=IR2, PDMS Category Name=IFNH)
'   The symbol consists of four outputs, three physical and one insulation outputs.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   16.JAN.2008     MA   CR-113430 Implemented Part Data Basis for the values 732 and 733
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Const NEGLIGIBLE_THICKNESS = 0.0001
Const NEGLIGIBLE_VALUE = 0.0001
Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub
Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 4) As Double
    Dim sptOffset(1 To 4) As Double
    Dim flangeDiam(1 To 4)    As Double
    Dim depth(1 To 4) As Double
    Dim flangeThick(1 To 4) As Double

    Dim iOutput     As Double
    Dim ObjRotameterBodyIns As Object
    Dim ObjNozz1BodyIns As Object
    Dim ObjNozz2BodyIns As Object
    Dim ObjNozz3BodyIns As Object
    Dim ObjNozz4BodyIns As Object
    
    Dim parInstrumentHeight As Double
    Dim parInstrumentHeight1 As Double
    Dim parInstrumentDiameter As Double
    Dim parInstrumentLength As Double
    Dim parInstrumentLength1 As Double
    Dim parInstrumentWidth As Double
    Dim parInstrumentWidth1 As Double
    Dim parFacetoFace As Double
    Dim parPortRotation1 As Double
    Dim parFace1toCenter As Double
    Dim parPortRotation2 As Double
    Dim parInsulationThickness As Double

    'Dynamic Port Input parameters variable declaration
    Dim parNPD As Double
    Dim parEndPreparation As Long
    Dim parScheduleThickness As Long
    Dim parEndStandard As Long
    Dim parPressureRating As Long
    Dim parFlowDirection As DistribFlow
    Dim parNPDUnitType As String

    Dim parPortIndex1 As Long
    Dim parNpd1 As Double
    Dim parEndPreparation1 As Long
    Dim parScheduleThickness1 As Long
    Dim parEndStandard1 As Long
    Dim parPressureRating1 As Long
    Dim parFlowDirection1 As DistribFlow
    Dim parId1 As String
    Dim parNPDUnitType1 As String
    
    Dim parPortIndex2 As Long
    Dim parNpd2 As Double
    Dim parEndPreparation2 As Long
    Dim parScheduleThickness2 As Long
    Dim parEndStandard2 As Long
    Dim parPressureRating2 As Long
    Dim parFlowDirection2 As DistribFlow
    Dim parId2 As String
    Dim parNPDUnitType2 As String
    
    Dim parPortIndex3 As Long
    Dim parNpd3 As Double
    Dim parEndPreparation3 As Long
    Dim parScheduleThickness3 As Long
    Dim parEndStandard3 As Long
    Dim parPressureRating3 As Long
    Dim parFlowDirection3 As DistribFlow
    Dim parId3 As String
    Dim parNPDUnitType3 As String
    
    Dim parPortIndex4 As Long
    Dim parNpd4 As Double
    Dim parEndPreparation4 As Long
    Dim parScheduleThickness4 As Long
    Dim parEndStandard4 As Long
    Dim parPressureRating4 As Long
    Dim parFlowDirection4 As DistribFlow
    Dim parId4 As String
    Dim parNPDUnitType4 As String
    Dim parFace3toCenter As Double
    Dim parFace4toCenter As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentHeight = arrayOfInputs(2)   'D2
    parInstrumentHeight1 = arrayOfInputs(3)  'D3
    parInstrumentDiameter = arrayOfInputs(4)      'D4
    parInstrumentLength = arrayOfInputs(5)        'D5
    parInstrumentLength1 = arrayOfInputs(6)       'D6
    parInstrumentWidth = arrayOfInputs(7)       'D7
    parInstrumentWidth1 = arrayOfInputs(8)       'D8
    parFacetoFace = arrayOfInputs(9)  'D9
    parPortRotation1 = arrayOfInputs(10)      'D10
    parFace1toCenter = arrayOfInputs(11)     'D11
    parPortRotation2 = arrayOfInputs(12)      'D12
    parInsulationThickness = arrayOfInputs(13)
    parNPD = arrayOfInputs(14)
    parEndPreparation = arrayOfInputs(15)
    parScheduleThickness = arrayOfInputs(16)
    parEndStandard = arrayOfInputs(17)
    parPressureRating = arrayOfInputs(18)
    parFlowDirection = arrayOfInputs(19)
    parPortIndex1 = arrayOfInputs(20)
    parNpd1 = arrayOfInputs(21)
    parEndPreparation1 = arrayOfInputs(22)
    parScheduleThickness1 = arrayOfInputs(23)
    parEndStandard1 = arrayOfInputs(24)
    parPressureRating1 = arrayOfInputs(25)
    parFlowDirection1 = arrayOfInputs(26)
    parPortIndex2 = arrayOfInputs(27)
    parNpd2 = arrayOfInputs(28)
    parEndPreparation2 = arrayOfInputs(29)
    parScheduleThickness2 = arrayOfInputs(30)
    parEndStandard2 = arrayOfInputs(31)
    parPressureRating2 = arrayOfInputs(32)
    parFlowDirection2 = arrayOfInputs(33)
    parId1 = arrayOfInputs(34)
    parId2 = arrayOfInputs(35)
    parNPDUnitType = arrayOfInputs(36)
    parNPDUnitType1 = arrayOfInputs(37)
    parNPDUnitType2 = arrayOfInputs(38)
'    parPortIndex3 = arrayOfInputs(39)
'    parNpd3 = arrayOfInputs(40)
'    parEndPreparation3 = arrayOfInputs(41)
'    parScheduleThickness3 = arrayOfInputs(42)
'    parEndStandard3 = arrayOfInputs(43)
'    parPressureRating3 = arrayOfInputs(44)
'    parFlowDirection3 = arrayOfInputs(45)
'    parPortIndex4 = arrayOfInputs(46)
'    parNpd4 = arrayOfInputs(47)
'    parEndPreparation4 = arrayOfInputs(48)
'    parScheduleThickness4 = arrayOfInputs(49)
'    parEndStandard4 = arrayOfInputs(50)
'    parPressureRating4 = arrayOfInputs(51)
'    parFlowDirection4 = arrayOfInputs(52)
'    parId3 = arrayOfInputs(53)
'    parId4 = arrayOfInputs(54)
'    parNPDUnitType3 = arrayOfInputs(55)
'    parNPDUnitType4 = arrayOfInputs(56)
'    parFace3toCenter = arrayOfInputs(57)
'    parFace4toCenter = arrayOfInputs(58)
    
    iOutput = 0
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    'The origin is considered to be mid of parInstrumentHeight1 and on the vertical axis
    'Insert your code for output 2(Rotameter Body Insulation)
     stPoint.Set -parInstrumentLength - parInsulationThickness, _
                        -parInstrumentHeight / 2 - parInsulationThickness, _
                        parInstrumentWidth + parInsulationThickness

    enPoint.Set parInstrumentLength1 + parInsulationThickness, _
                        parInstrumentHeight / 2 + parInsulationThickness, _
                        -parInstrumentWidth1 - parInsulationThickness

    Set ObjRotameterBodyIns = PlaceBox(m_OutputColl, stPoint, enPoint)
 
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRotameterBodyIns
    Set ObjRotameterBodyIns = Nothing
        
    'Insert your code for output 5(Nozzle1 Body Insulation)
     If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth
 
    Dim InsulationDiameter As Double
    InsulationDiameter = pipeDiam(parPortIndex1) + parInsulationThickness * 2
    If flangeDiam(parPortIndex1) > pipeDiam(parPortIndex1) Then InsulationDiameter = flangeDiam(parPortIndex1) + parInsulationThickness * 2
    stPoint.Set 0, -parFacetoFace / 2, 0
    enPoint.Set (parFace1toCenter) * Cos(parPortRotation2), _
                                -parFacetoFace / 2, _
                                -(parFace1toCenter) * Sin(parPortRotation2)

    Set ObjNozz1BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
  
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozz1BodyIns
    Set ObjNozz1BodyIns = Nothing
    
    'Insert your code for output 6(Nozzle2 Body Insulation)
    InsulationDiameter = pipeDiam(parPortIndex2) + parInsulationThickness * 2
    If flangeDiam(parPortIndex2) > pipeDiam(parPortIndex2) Then InsulationDiameter = flangeDiam(parPortIndex2) + parInsulationThickness * 2
    stPoint.Set 0, parFacetoFace / 2, 0
    enPoint.Set (parFace1toCenter) * Cos(parPortRotation1), _
                                parFacetoFace / 2, _
                                -(parFace1toCenter) * Sin(parPortRotation1)
                                
    Set ObjNozz2BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
 
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozz2BodyIns
    Set ObjNozz2BodyIns = Nothing

    ' This symbol is based on the following Part data basis values that govern its geometry
    '   (Part data Basis value -731): Level indicator, with inlet/outlet ports
    '   (Part data Basis value -732): Level indicator, with inlet/outlet ports and vent/drain
    '                                               ports, with straight outlet
    '   (Part data Basis value -733): Level indicator, with inlet/outlet ports and vent/drain
    '                                               ports, with 90o elbow outlet
    
    'Checking for the Part Data Basis Property
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Select Case lPartDataBasis
    
        Case Is <= 1, 731
    
        Case 732
            parPortIndex3 = arrayOfInputs(39)
            parNpd3 = arrayOfInputs(40)
            parEndPreparation3 = arrayOfInputs(41)
            parScheduleThickness3 = arrayOfInputs(42)
            parEndStandard3 = arrayOfInputs(43)
            parPressureRating3 = arrayOfInputs(44)
            parFlowDirection3 = arrayOfInputs(45)
            parPortIndex4 = arrayOfInputs(46)
            parNpd4 = arrayOfInputs(47)
            parEndPreparation4 = arrayOfInputs(48)
            parScheduleThickness4 = arrayOfInputs(49)
            parEndStandard4 = arrayOfInputs(50)
            parPressureRating4 = arrayOfInputs(51)
            parFlowDirection4 = arrayOfInputs(52)
            parId3 = arrayOfInputs(53)
            parId4 = arrayOfInputs(54)
            parNPDUnitType3 = arrayOfInputs(55)
            parNPDUnitType4 = arrayOfInputs(56)
            parFace3toCenter = arrayOfInputs(57)
            parFace4toCenter = arrayOfInputs(58)
        
            'Insert your code for Nozzle3 Body Insulation
            InsulationDiameter = pipeDiam(parPortIndex3) + parInsulationThickness * 2
            If flangeDiam(parPortIndex3) > pipeDiam(parPortIndex3) Then InsulationDiameter = flangeDiam(parPortIndex3) + parInsulationThickness * 2
            stPoint.Set 0, parInstrumentHeight / 2, 0
            enPoint.Set 0, parFace3toCenter, 0
            
            Set ObjNozz3BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
        
            'Set the output
            m_OutputColl.AddOutput "Nozz3BodyIns", ObjNozz3BodyIns
            Set ObjNozz3BodyIns = Nothing
            
            'Insert your code for Nozzle4 Body Insulation
            InsulationDiameter = pipeDiam(parPortIndex4) + parInsulationThickness * 2
            If flangeDiam(parPortIndex4) > pipeDiam(parPortIndex4) Then InsulationDiameter = flangeDiam(parPortIndex4) + parInsulationThickness * 2
            stPoint.Set 0, -parInstrumentHeight / 2, 0
            enPoint.Set 0, -parFace4toCenter, 0
            
            Set ObjNozz4BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
        
            'Set the output
            m_OutputColl.AddOutput "Nozz4BodyIns", ObjNozz4BodyIns
            Set ObjNozz4BodyIns = Nothing
            
        Case 733
            parPortIndex3 = arrayOfInputs(39)
            parNpd3 = arrayOfInputs(40)
            parEndPreparation3 = arrayOfInputs(41)
            parScheduleThickness3 = arrayOfInputs(42)
            parEndStandard3 = arrayOfInputs(43)
            parPressureRating3 = arrayOfInputs(44)
            parFlowDirection3 = arrayOfInputs(45)
            parPortIndex4 = arrayOfInputs(46)
            parNpd4 = arrayOfInputs(47)
            parEndPreparation4 = arrayOfInputs(48)
            parScheduleThickness4 = arrayOfInputs(49)
            parEndStandard4 = arrayOfInputs(50)
            parPressureRating4 = arrayOfInputs(51)
            parFlowDirection4 = arrayOfInputs(52)
            parId3 = arrayOfInputs(53)
            parId4 = arrayOfInputs(54)
            parNPDUnitType3 = arrayOfInputs(55)
            parNPDUnitType4 = arrayOfInputs(56)
            parFace3toCenter = arrayOfInputs(57)
            parFace4toCenter = arrayOfInputs(58)
        
            'Insert your code for Nozzle3 Body Insulation
            InsulationDiameter = pipeDiam(parPortIndex3) + parInsulationThickness * 2
            If flangeDiam(parPortIndex3) > pipeDiam(parPortIndex3) Then InsulationDiameter = flangeDiam(parPortIndex3) + parInsulationThickness * 2
            stPoint.Set 0, parInstrumentHeight / 2, 0
            enPoint.Set 0, parFace3toCenter, 0
            
            Set ObjNozz3BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
        
            'Set the output
            m_OutputColl.AddOutput "Nozz3BodyIns", ObjNozz3BodyIns
            Set ObjNozz3BodyIns = Nothing
            
            'Insert your code for Nozzle4 Body Insulation
            InsulationDiameter = pipeDiam(parPortIndex4) + parInsulationThickness * 2
            If flangeDiam(parPortIndex4) > pipeDiam(parPortIndex4) Then InsulationDiameter = flangeDiam(parPortIndex4) + parInsulationThickness * 2
            If CmpDblEqual(flangeThick(parPortIndex4), 0) Then flangeThick(parPortIndex4) = NEGLIGIBLE_THICKNESS
            stPoint.Set -(parFace4toCenter - parInstrumentHeight / 2), -parFace4toCenter, 0
            enPoint.Set -(parFace4toCenter - parInstrumentHeight / 2) _
                                + flangeThick(parPortIndex4) + parInsulationThickness, _
                                            -parFace4toCenter, 0
            
            Set ObjNozz4BodyIns = PlaceCylinder(m_OutputColl, stPoint, enPoint, InsulationDiameter, True)
        
            'Set the output
            m_OutputColl.AddOutput "Nozz4BodyIns", ObjNozz4BodyIns
            Set ObjNozz4BodyIns = Nothing
            
            'Place elbow
            Dim oGeomFactory As IngrGeom3D.GeometryFactory
            Set oGeomFactory = New GeometryFactory
            Dim objCircle   As IngrGeom3D.Circle3d
            Dim dInsulationRadius As Double
            Dim dElbowRadius As Double
            dInsulationRadius = pipeDiam(parPortIndex4) / 2 + parInsulationThickness
            dElbowRadius = parFace4toCenter - parInstrumentHeight / 2
            
            If CmpDblGreaterthan(dInsulationRadius, dElbowRadius) Then _
                            dInsulationRadius = dElbowRadius - NEGLIGIBLE_VALUE
            
            Set objCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                         0, -parInstrumentHeight / 2, 0, _
                         0, -1, 0, _
                         dInsulationRadius)
            
            Dim CP As New AutoMath.DPosition 'revolution center point
            Dim CV As New AutoMath.DVector 'rotation vector for rotation
            CP.Set -(parFace4toCenter - parInstrumentHeight / 2), -parInstrumentHeight / 2, 0
            CV.Set 0, 0, -1
          
            Dim ObjElbowIns As Object
            Set ObjElbowIns = PlaceRevolution(m_OutputColl, objCircle, _
                        CV, CP, PI / 2, True)
         
            'Set the output
            m_OutputColl.AddOutput "ElbowIns", ObjElbowIns
            Set ObjElbowIns = Nothing
            Set CP = Nothing
            Set CV = Nothing
            Set objCircle = Nothing
            Set oGeomFactory = Nothing
        
        Case Else
            GoTo ErrorLabel:

    End Select
    
    Exit Sub
    
    Set stPoint = Nothing
    Set enPoint = Nothing
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
