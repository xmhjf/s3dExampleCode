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
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         MP
'   Creation Date:  Thursday, Jun 26 2008
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The nozzles are fully Parametric.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   26.JUN.2008     MP     CR-144381    Created the symbol
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
      
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim dPipeDiam(1 To 3) As Double
    Dim dSptOffset(1 To 3) As Double
    Dim dFlangeDiam(1 To 3)    As Double
    Dim dDepth(1 To 3) As Double
    Dim dFlangeThick(1 To 3) As Double
    
    Dim iOutput     As Long
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter  As Double
    Dim parInsulationThickness As Double
    Dim parWyeAngle As Double
            
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
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
    Dim parInsulationDiameter As Double
    
    Dim parPortIndex3 As Long
    Dim parNpd3 As Double
    Dim parEndPreparation3 As Long
    Dim parScheduleThickness3 As Long
    Dim parEndStandard3 As Long
    Dim parPressureRating3 As Long
    Dim parFlowDirection3 As DistribFlow
    Dim parId3 As String
    Dim parNPDUnitType3 As String
    Dim ObjInsBody1 As Object
    Dim ObjInsBody2 As Object
    Dim ObjInsBody3 As Object
    Dim ObjInsPort1 As Object
    Dim ObjInsPort2 As Object
    Dim ObjInsPort3 As Object
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parFace3toCenter = arrayOfInputs(4)
    parWyeAngle = arrayOfInputs(5)
    parNPD = arrayOfInputs(6)                       'Input parameters for Dynamic Nozzles
    parEndPreparation = arrayOfInputs(7)
    parScheduleThickness = arrayOfInputs(8)
    parEndStandard = arrayOfInputs(9)
    parPressureRating = arrayOfInputs(10)
    parFlowDirection = arrayOfInputs(11)
    parPortIndex1 = arrayOfInputs(12)
    parNpd1 = arrayOfInputs(13)
    parEndPreparation1 = arrayOfInputs(14)
    parScheduleThickness1 = arrayOfInputs(15)
    parEndStandard1 = arrayOfInputs(16)
    parPressureRating1 = arrayOfInputs(17)
    parFlowDirection1 = arrayOfInputs(18)
    parPortIndex2 = arrayOfInputs(19)
    parNpd2 = arrayOfInputs(20)
    parEndPreparation2 = arrayOfInputs(21)
    parScheduleThickness2 = arrayOfInputs(22)
    parEndStandard2 = arrayOfInputs(23)
    parPressureRating2 = arrayOfInputs(24)
    parFlowDirection2 = arrayOfInputs(25)
    parPortIndex3 = arrayOfInputs(26)
    parNpd3 = arrayOfInputs(27)
    parEndPreparation3 = arrayOfInputs(28)
    parScheduleThickness3 = arrayOfInputs(29)
    parEndStandard3 = arrayOfInputs(30)
    parPressureRating3 = arrayOfInputs(31)
    parFlowDirection3 = arrayOfInputs(32)
    parInsulationThickness = arrayOfInputs(33)
    'Text Inputs
    parId1 = arrayOfInputs(34)
    parId2 = arrayOfInputs(35)
    parId3 = arrayOfInputs(36)
    parNPDUnitType = arrayOfInputs(37)
    parNPDUnitType1 = arrayOfInputs(38)
    parNPDUnitType2 = arrayOfInputs(39)
    parNPDUnitType3 = arrayOfInputs(40)
    iOutput = 0

    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    
    CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, dPipeDiam, dFlangeThick, _
                                            dFlangeDiam, dSptOffset, dDepth

' Insert your code for output 1(Insulation for Body1)
    stPoint.Set -parFace1toCenter, 0, 0
    enPoint.Set 0, 0, 0

    parInsulationDiameter = dPipeDiam(parPortIndex1) + parInsulationThickness * 2
   
    Set ObjInsBody1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody1
    Set ObjInsBody1 = Nothing
    
   'if flangethickness is not zero insert code for insport1
    If CmpDblGreaterthan(dFlangeDiam(parPortIndex1), dPipeDiam(parPortIndex1)) Then
        parInsulationDiameter = dFlangeDiam(parPortIndex1) + parInsulationThickness * 2

        stPoint.Set -parFace1toCenter, 0, 0
        enPoint.Set -parFace1toCenter + dFlangeThick(parPortIndex1) + parInsulationThickness, 0, 0
        
        If CmpDblGreaterthan(dFlangeThick(parPortIndex1) + parInsulationThickness, parFace1toCenter) Then
            enPoint.Set 0, 0, 0
        End If
        
        Set ObjInsPort1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
        
        m_OutputColl.AddOutput "InsPort1", ObjInsPort1
        Set ObjInsPort1 = Nothing
    End If
' Insert your code for output 2( Insulation for Body2 )
    CreateRetrieveDynamicNozzle 2, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd2, _
                                            parNPDUnitType2, parEndPreparation2, parScheduleThickness2, _
                                            parEndStandard2, parPressureRating2, parFlowDirection2, _
                                            parId2, m_OutputColl, m_oCodeListMetadata, _
                                            False, dPipeDiam, dFlangeThick, _
                                            dFlangeDiam, dSptOffset, dDepth

    
    stPoint.Set 0, 0, 0
    enPoint.Set parFace2toCenter * Cos(parWyeAngle / 2), parFace2toCenter * Sin(parWyeAngle / 2), 0
    parInsulationDiameter = dPipeDiam(parPortIndex2) + parInsulationThickness * 2
    Set ObjInsBody2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody2
    Set ObjInsBody2 = Nothing
    
   'if flangethickness is not zero insert code for insport2
    If CmpDblGreaterthan(dFlangeDiam(parPortIndex2), dPipeDiam(parPortIndex2)) Then
        parInsulationDiameter = dFlangeDiam(parPortIndex2) + parInsulationThickness * 2
'        dAbsDistance = parFace2toCenter - parInsulationThickness
        stPoint.Set parFace2toCenter * Cos(parWyeAngle / 2), parFace2toCenter * Sin(parWyeAngle / 2), 0
        enPoint.Set (parFace2toCenter - dFlangeThick(parPortIndex2) - parInsulationThickness) * Cos(parWyeAngle / 2), _
                    (parFace2toCenter - dFlangeThick(parPortIndex2) - parInsulationThickness) * Sin(parWyeAngle / 2), 0
             
        If CmpDblGreaterthan(dFlangeThick(parPortIndex2) + parInsulationThickness, parFace2toCenter) Then
            enPoint.Set 0, 0, 0
        End If
        
        Set ObjInsPort2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
        
        m_OutputColl.AddOutput "InsPort2", ObjInsPort2
        Set ObjInsPort2 = Nothing
    End If
    
    ' Insert your code for output 3( Insulation for Body3 )
    CreateRetrieveDynamicNozzle 3, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd3, _
                                            parNPDUnitType3, parEndPreparation3, parScheduleThickness2, _
                                            parEndStandard3, parPressureRating3, parFlowDirection3, _
                                            parId3, m_OutputColl, m_oCodeListMetadata, _
                                            False, dPipeDiam, dFlangeThick, _
                                            dFlangeDiam, dSptOffset, dDepth

      
    stPoint.Set 0, 0, 0
    enPoint.Set parFace3toCenter * Cos(parWyeAngle / 2), -parFace3toCenter * Sin(parWyeAngle / 2), 0
    parInsulationDiameter = dPipeDiam(parPortIndex3) + parInsulationThickness * 2
    Set ObjInsBody3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsBody3
    Set ObjInsBody3 = Nothing
    
   'if flangethickness is not zero insert code for insport3
    If CmpDblGreaterthan(dFlangeDiam(parPortIndex3), dPipeDiam(parPortIndex3)) Then
        parInsulationDiameter = dFlangeDiam(parPortIndex3) + parInsulationThickness * 2
        stPoint.Set parFace3toCenter * Cos(parWyeAngle / 2), -parFace3toCenter * Sin(parWyeAngle / 2), 0
        enPoint.Set (parFace3toCenter - dFlangeThick(parPortIndex3) - parInsulationThickness) * Cos(parWyeAngle / 2), _
        -(parFace3toCenter - dFlangeThick(parPortIndex3) - parInsulationThickness) * Sin(parWyeAngle / 2), 0
        
        If CmpDblGreaterthan(dFlangeThick(parPortIndex3) + parInsulationThickness, parFace3toCenter) Then
            enPoint.Set 0, 0, 0
        End If
             
        Set ObjInsPort3 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parInsulationDiameter, True)
        
        m_OutputColl.AddOutput "InsPort3", ObjInsPort3
        Set ObjInsPort3 = Nothing
    End If
    
    Set m_oCodeListMetadata = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
     
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
