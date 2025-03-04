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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         dkl
'   Creation Date:  Tuesday, Jan 23 2007
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   This symbol is based on a functionality of a SP3DCSValve with in-built Gear Operator.
'   The in-built operator is based on SP3DOP_291.
'   The 2 ports for the symbol are fully parametric and can be changed on-the-fly.
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------    -----        ------------------
'   23.Jan.2007     dkl    CR-113194. Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private m_oGeomHelper As IJSymbolGeometryHelper

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx
    
    Set m_oGeomHelper = New SymbolServices
  
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oCodeListMetadata = Nothing
    Set m_oGeomHelper = Nothing
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam(1 To 2) As Double
    Dim sptOffset(1 To 2) As Double
    Dim flangeDiam(1 To 2)    As Double
    Dim depth(1 To 2) As Double
    Dim flangeThick(1 To 2) As Double

    Dim iOutput     As Double
    
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parOperatorHeight As Double
    Dim parOperatorDiameter As Double
    Dim parActuatorLength As Double
    Dim parOperatorDepth As Double
    Dim parActuatorDiameter As Double
    Dim parActuatorOffset As Double
    Dim parStemExtension As Double
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
    

' Inputs
   Set oPartFclt = arrayOfInputs(1)
    parFace1toCenter = arrayOfInputs(2)
    parFace2toCenter = arrayOfInputs(3)
    parOperatorHeight = arrayOfInputs(4)
    parOperatorDiameter = arrayOfInputs(5)
    parActuatorLength = arrayOfInputs(6)
    parOperatorDepth = arrayOfInputs(7)
    parActuatorDiameter = arrayOfInputs(8)
    parActuatorOffset = arrayOfInputs(9)
    parStemExtension = arrayOfInputs(10)
    parInsulationThickness = arrayOfInputs(11)
    
    parNPD = arrayOfInputs(12)                       'Input parameters for Dynamic Nozzles
    parNPDUnitType = arrayOfInputs(13)
    parEndPreparation = arrayOfInputs(14)
    parScheduleThickness = arrayOfInputs(15)
    parEndStandard = arrayOfInputs(16)
    parPressureRating = arrayOfInputs(17)
    parFlowDirection = arrayOfInputs(18)
    parId1 = arrayOfInputs(19)
    parPortIndex1 = arrayOfInputs(20)
    parNpd1 = arrayOfInputs(21)
    parNPDUnitType1 = arrayOfInputs(22)
    parEndPreparation1 = arrayOfInputs(23)
    parScheduleThickness1 = arrayOfInputs(24)
    parEndStandard1 = arrayOfInputs(25)
    parPressureRating1 = arrayOfInputs(26)
    parFlowDirection1 = arrayOfInputs(27)
    parId2 = arrayOfInputs(28)
    parPortIndex2 = arrayOfInputs(29)
    parNpd2 = arrayOfInputs(30)
    parNPDUnitType2 = arrayOfInputs(31)
    parEndPreparation2 = arrayOfInputs(32)
    parScheduleThickness2 = arrayOfInputs(33)
    parEndStandard2 = arrayOfInputs(34)
    parPressureRating2 = arrayOfInputs(35)
    parFlowDirection2 = arrayOfInputs(36)
    
    iOutput = 0
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If

'Insert your code for output 2(Insulation Cylinderical Body)
   CreateRetrieveDynamicNozzle 1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            False, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth

   Dim oStPoint As AutoMath.DPosition
   Set oStPoint = New AutoMath.DPosition
   Dim oEndPoint As AutoMath.DPosition
   Set oEndPoint = New AutoMath.DPosition
   Dim dInsulationDia As Double
      
   oStPoint.Set -parFace1toCenter, 0, 0
   oEndPoint.Set parFace2toCenter, 0, 0
   
   dInsulationDia = (2 * parInsulationThickness) + ReturnMax4(pipeDiam(parPortIndex1), flangeDiam(parPortIndex1), _
                                            pipeDiam(parPortIndex2), flangeDiam(parPortIndex2))
   
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEndPoint, dInsulationDia
    Set oStPoint = Nothing
    Set oEndPoint = Nothing
        
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
      
End Sub

Private Function ReturnMax4(A#, B#, C#, D#) As Double
    Dim MaxValue As Double
    
    MaxValue = A
    If CmpDblLessThan(MaxValue, B) Then MaxValue = B
    If CmpDblLessThan(MaxValue, C) Then MaxValue = C
    If CmpDblLessThan(MaxValue, D) Then MaxValue = D
    ReturnMax4 = MaxValue

End Function

