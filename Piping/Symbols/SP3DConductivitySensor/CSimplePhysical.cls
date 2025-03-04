VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2006, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         RRK
'   Creation Date:  Thursday, Jul 13 2006
'   Description:
'    Model CSFT Conductivity Sensor, SENSOR DEVLOPMENT INC at www.sensordev.com/csft.htm
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper

Private PI       As Double
Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo Errx

    Set m_oGeomHelper = New SymbolServices
    
    Exit Sub
    
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    
'   Declaration of temporary variables dX and dY to store distances along
'   X and Y directions respectively
    Dim dX As Double
    Dim dY As Double
    
    Dim dNozzleLength As Double
    
'Inputs
    Dim parFacetoFace As Double
    Dim parInstrumentHeight As Double
    
'Loop index
    Dim iCount As Integer
    
    Dim iOutput     As Double
    iOutput = 0
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoFace = arrayOfInputs(2)
    parInstrumentHeight = arrayOfInputs(3)
    
    m_oGeomHelper.OutputCollection = m_OutputColl
     
' Assumption
    dNozzleLength = parFacetoFace / 80
    
' Place Body of the Sensor (Output 1)
' -------------------------------------------------------------------------------------------------------------------
    Dim stPoint   As AutoMath.DPosition
    Dim enPoint   As AutoMath.DPosition
    
    Set stPoint = New AutoMath.DPosition
    Set enPoint = New AutoMath.DPosition

    stPoint.Set -parFacetoFace / 2, 0, 0
    enPoint.Set parFacetoFace / 2, 0, 0
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
  
    
    'Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, pipeDiam
    
' Place Cylinder 1 (Output 2)
' -------------------------------------------------------------------------------------------------------------------
    Dim dCylLnth1 As Double

' Assumption: The length of the Cylinder 1 is assumed as 1/24 times of Face to Face
    dCylLnth1 = parFacetoFace / 24
    dX = -parFacetoFace / 2 + dNozzleLength
    
    stPoint.Set dX, 0, 0
    enPoint.Set dX + dCylLnth1, 0, 0
    
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 1.5 * pipeDiam
    

' Place Cylinder 2 (Output 3)
' This Cylinder is mirror image of Cylinder 1 about Y-Z Plane
' -------------------------------------------------------------------------------------------------------------------
    stPoint.Set -dX, 0, 0
    enPoint.Set -dX - dCylLnth1, 0, 0
    
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 1.5 * pipeDiam
    

' Place Cylinder 3 (Output 4)
' -------------------------------------------------------------------------------------------------------------------
    Dim dCylLnth3 As Double
    
' Assumption: The height of Cylinder 3 is assumed as 1/48 times of Face to Face
    dCylLnth3 = parFacetoFace / 48
    dX = -parFacetoFace / 2 + dNozzleLength + dCylLnth1
    
    stPoint.Set dX, 0, 0
    enPoint.Set dX + dCylLnth3, 0, 0
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 1.3 * pipeDiam
    
' Place Cylinder 4 (Output 5)
' This Cylinder is mirror image of Cylinder 3 about Y-Z Plane
' -------------------------------------------------------------------------------------------------------------------
    stPoint.Set -dX, 0, 0
    enPoint.Set -dX - dCylLnth3, 0, 0
    
' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 1.3 * pipeDiam
        
' Place T1Header,T2Header and T3Header (Output 6,Output 7 and Output 8)
' -------------------------------------------------------------------------------------------------------------------
    Dim dTHeaderLnth As Double
' Assumption
    dTHeaderLnth = parFacetoFace / 10
    
    For iOutput = 6 To 8
            Select Case iOutput
                Case 6:
                    dX = dX + dCylLnth3
                Case 7:
                    dX = -dX - dTHeaderLnth
                Case 8:
                    dX = -dTHeaderLnth / 2
            End Select
            
        stPoint.Set dX, 0, 0
        enPoint.Set dX + dTHeaderLnth, 0, 0
        
    ' Set the output
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, 1.25 * pipeDiam
    Next
    
' Place all Branches,ConnectorBases and Connectors (Outputs 9 to 17)
' -------------------------------------------------------------------------------------------------------------------
    Dim dTBrHeight As Double
    Dim dTBaseHeight As Double
    Dim dDiam As Double
    Dim ObjBranchCompnt As Object
    
' Assumptions
    dTBrHeight = dTHeaderLnth / 2
    dTBaseHeight = 0.6 * parInstrumentHeight
    
    For iOutput = 9 To 17
        Select Case iOutput
                Case 9:
                'This is the case for T1Branch (Output 9)
                    dX = -parFacetoFace / 2 + dNozzleLength + dCylLnth1 + dCylLnth3 + dTHeaderLnth / 2
                    dY = dTBrHeight
                    dDiam = 1.25 * pipeDiam
                    
                Case 10:
                'This is the case for T1ConnectorBase (Output 10)
                    dY = dTBaseHeight
                    dDiam = 0.9 * pipeDiam
                    
                Case 11:
                'This is the case for T1Connector (Output 11)
                    dY = parInstrumentHeight
                    dDiam = 0.6 * pipeDiam
                    
                Case 12:
                'This is the case for T2Branch (Output 12)
                    dX = -dX
                    dY = dTBrHeight
                    dDiam = 1.25 * pipeDiam
                Case 13:
                'This is the case for T2ConnectorBase (Output 13)
                    dY = dTBaseHeight
                    dDiam = 0.9 * pipeDiam
                    
                Case 14:
                'This is the case for T2Connector (Output 14)
                    dY = parInstrumentHeight
                    dDiam = 0.6 * pipeDiam
                Case 15:
                'This is the case for T3Branch (Output 15)
                    dX = 0
                    dY = dTBrHeight
                    dDiam = 1.25 * pipeDiam
                Case 16:
                'This is the case for T3ConnectorBase (Output 16)
                    dY = dTBaseHeight
                    dDiam = 0.9 * pipeDiam
                    
                Case 17:
                'This is the case for T3Connector (Output 17)
                    dY = parInstrumentHeight
                    dDiam = 0.6 * pipeDiam
            End Select
        
        stPoint.Set dX, 0, 0
        enPoint.Set dX, dY, 0
        
    ' Set the output
        m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), stPoint, enPoint, dDiam
    Next
    iOutput = 17

' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    Set oPlacePoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector
    dX = (parFacetoFace / 2 + sptOffset - depth)
    oPlacePoint.Set -dX, 0, 0
    oDir.Set -1, 0, 0
    
    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

' Place Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth

    dX = (parFacetoFace / 2 + sptOffset - depth)
    oPlacePoint.Set dX, 0, 0
    oDir.Set 1, 0, 0

    If CmpDblLessThan(dNozzleLength, flangeThick) Then dNozzleLength = flangeThick
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                        oPlacePoint, dNozzleLength)
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
        Err.HelpFile, Err.HelpContext
End Sub


