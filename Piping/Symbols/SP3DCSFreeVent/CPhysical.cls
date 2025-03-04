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
'   Author:          MS
'   Creation Date:  Friday, Jan 17 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
' Symbol Model No. is: S11BZ Page No. D-102 PDS Piping Component Data Reference Guide.
' SP3DFreeVent Symbol is created with Two Outputs
' The two are physical aspect outputs created as follows:
'   ObjFreeeVentBody - Using 'PlaceRevolution' function and
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.  
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Private m_oCodeListMetadata     As IJDCodeListMetaData
Private PI           As Double
Private Const MODULE = "Physical:" 'Used for error messages

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
    Dim pipeDiam(1 To 1) As Double
    Dim sptOffset(1 To 1) As Double
    Dim flangeDiam(1 To 1)    As Double
    Dim depth(1 To 1) As Double
    Dim flangeThick(1 To 1) As Double
    
    Dim iOutput     As Double
    Dim ObjFreeeVentBody As Object
    
    Dim parNozztoCenter As Double
    Dim parFacetoCenter As Double
    Dim parVentAngle As Double

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
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parNozztoCenter = arrayOfInputs(2) 'P2
    parFacetoCenter = arrayOfInputs(3) 'P3
    parVentAngle = arrayOfInputs(4) 'P4
    parNPD = arrayOfInputs(5)
    parEndPreparation = arrayOfInputs(6)
    parScheduleThickness = arrayOfInputs(7)
    parEndStandard = arrayOfInputs(8)
    parPressureRating = arrayOfInputs(9)
    parFlowDirection = arrayOfInputs(10)
    parPortIndex1 = arrayOfInputs(11)
    parNpd1 = arrayOfInputs(12)
    parEndPreparation1 = arrayOfInputs(13)
    parScheduleThickness1 = arrayOfInputs(14)
    parEndStandard1 = arrayOfInputs(15)
    parPressureRating1 = arrayOfInputs(16)
    parFlowDirection1 = arrayOfInputs(17)
    parId1 = arrayOfInputs(18)
    parNPDUnitType = arrayOfInputs(19)
    parNPDUnitType1 = arrayOfInputs(20)
    
    iOutput = 0
    If m_oCodeListMetadata Is Nothing Then
        Set m_oCodeListMetadata = m_OutputColl.ResourceManager
    End If
    ' Place Nozzle 1
    Dim oLogicalDistPort    As GSCADNozzleEntities.IJLogicalDistPort
    Dim oDistribPort        As GSCADNozzleEntities.IJDistribPort
    Dim oNozzle             As GSCADNozzleEntities.IJDNozzle
    Dim oDir As New AutoMath.DVector
    Dim oPlacePoint As New AutoMath.DPosition
'CreateRetrieveDynamicNozzle
    Set oNozzle = CreateRetrieveDynamicNozzle(parPortIndex1, parNPD, parNPDUnitType, _
                                            parEndPreparation, parScheduleThickness, parEndStandard, _
                                            parPressureRating, parFlowDirection, parNpd1, _
                                            parNPDUnitType1, parEndPreparation1, parScheduleThickness1, _
                                            parEndStandard1, parPressureRating1, parFlowDirection1, _
                                            parId1, m_OutputColl, m_oCodeListMetadata, _
                                            True, pipeDiam, flangeThick, _
                                            flangeDiam, sptOffset, depth)
    
    Set oLogicalDistPort = oNozzle
    Set oDistribPort = oNozzle
    oNozzle.Length = parNozztoCenter - parFacetoCenter

' Direction of the Nozzle
    oDir.Set 0, -1, 0
    oDistribPort.SetDirectionVector oDir

' Position of the nozzle should be the connect point of the nozzle
    oPlacePoint.Set 0, -sptOffset(parPortIndex1) + depth(parPortIndex1), 0
    oLogicalDistPort.SetCenterLocation oPlacePoint
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), oNozzle
    Set oNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    Set oLogicalDistPort = Nothing
    Set oDistribPort = Nothing
    
 ' Insert your code for output 1(FreeeVentBody)
    Dim objCirc1 As IngrGeom3D.Circle3d
    Dim geomFactory As New IngrGeom3D.GeometryFactory
    
    Dim cirNormalX As Double, cirNormalY As Double, cirNormalZ As Double
    cirNormalX = 0
    cirNormalY = 1
    cirNormalZ = 0
    
    Dim cir1CenterX As Double, cir1CenterY As Double, cir1CenterZ As Double
    cir1CenterX = 0
    cir1CenterY = parNozztoCenter - parFacetoCenter
    cir1CenterZ = 0
    
    Set objCirc1 = geomFactory.Circles3d.CreateByCenterNormalRadius( _
                                                m_OutputColl.ResourceManager, cir1CenterX, _
                                                cir1CenterY, cir1CenterZ, cirNormalX, _
                                                cirNormalY, cirNormalZ, pipeDiam(parPortIndex1) / 2)
 
    Dim centPoint   As New AutoMath.DPosition
    Dim axis        As New AutoMath.DVector
 
    centPoint.Set parFacetoCenter, parNozztoCenter - parFacetoCenter, 0
    axis.Set 0, 0, -1
    
     Set ObjFreeeVentBody = PlaceRevolution(m_OutputColl, objCirc1, _
                                    axis, centPoint, parVentAngle, _
                                    True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjFreeeVentBody
    Set ObjFreeeVentBody = Nothing
    
    Set geomFactory = Nothing
    Set centPoint = Nothing
    Set axis = Nothing

    ' remove curve
    Dim sp3dElem1 As IJDObject
    Set sp3dElem1 = objCirc1
    Set objCirc1 = Nothing
    sp3dElem1.Remove
   
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
