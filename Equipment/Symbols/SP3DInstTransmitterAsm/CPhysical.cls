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
'   Copyright (c) 2006-10, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:    VRK
'   Creation Date:  Wednesday, oct 11 2006
'
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'   21.May.2008     VRK     CR-142762: Provide instrument transmitter and pressure transmitter symbols
'   03.Feb.2010     VKC     TR-CP-157988  TO Do List entries are generated whern placing lighting fixtures
'                            (Modified nozzle code to create with placeholder)
'******************************************************************************

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Public Sub run(ByVal m_outputcoll As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorHandler
    
    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Integer
    
    Dim parInstrumentTubeLength  As Double
    Dim parInstrumentTubeDiameter As Double
    
    Dim ObjInstrumentTube As Object
    Dim ObjTubeCylinder As Object
    Dim oStPoint  As New AutoMath.DPosition
    Dim oEnPoint  As New AutoMath.DPosition

    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parInstrumentTubeLength = arrayOfInputs(2)
    parInstrumentTubeDiameter = arrayOfInputs(3)
    
    iOutput = 0
      
    'Instrument Tube
    oStPoint.Set 0, 0, 0
    oEnPoint.Set parInstrumentTubeLength, 0, 0
    Set ObjInstrumentTube = PlaceCylinder(m_outputcoll, oStPoint, oEnPoint, parInstrumentTubeDiameter, True)
    iOutput = iOutput + 1
    m_outputcoll.AddOutput arrayOfOutputs(iOutput), ObjInstrumentTube
    Set ObjInstrumentTube = Nothing
    
    'TubeCylinderDiameter
    Dim pipedia As Double
    RetrievePipeOD_1 1, oPartFclt, m_outputcoll, pipedia
    Dim dInstrumentTubeRadius As Double
    dInstrumentTubeRadius = parInstrumentTubeDiameter / 2
    
    oStPoint.Set pipedia / 2#, -dInstrumentTubeRadius, -(dInstrumentTubeRadius - pipedia / 2#)
    oEnPoint.Set pipedia / 2#, 0, -(dInstrumentTubeRadius - pipedia / 2#)
    Set ObjTubeCylinder = PlaceCylinder(m_outputcoll, oStPoint, oEnPoint, pipedia, False)
    iOutput = iOutput + 1
    m_outputcoll.AddOutput arrayOfOutputs(iOutput), ObjTubeCylinder
    Set ObjTubeCylinder = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

    ' Insert your code for Conduit Port
    Dim oConduitPortPoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    
    Set oConduitPortPoint = New AutoMath.DPosition
    Set oDir = New AutoMath.DVector

    oDir.Set 0, -1, 0
    oConduitPortPoint.Set pipedia / 2, -dInstrumentTubeRadius, -(dInstrumentTubeRadius - pipedia / 2#)
    
    Dim ObjConduitPort1 As IJConduitPortOcc
    Set ObjConduitPort1 = CreateConduitNozzlePH(oConduitPortPoint, oDir, m_outputcoll, oPartFclt, 1)
    
    ' Set the output
    iOutput = iOutput + 1
    m_outputcoll.AddOutput arrayOfOutputs(iOutput), ObjConduitPort1
    Set ObjConduitPort1 = Nothing

    ' Insert your code for Cable Port
    Dim oCableTrayPortPoint As AutoMath.DPosition
    Set oCableTrayPortPoint = New AutoMath.DPosition

    Dim oRadialOrient As AutoMath.DVector
    Set oRadialOrient = New AutoMath.DVector
    
    oRadialOrient.Set 0, 0, -1
    oCableTrayPortPoint.Set pipedia / 2, -dInstrumentTubeRadius, -(dInstrumentTubeRadius - pipedia / 2#)

    Dim objCableNozzle As CableNozzle
    Dim iDistribPort As IJDistribPort
    Dim iLogicalDistPort As IJLogicalDistPort
    Dim oNozzlePHFactory As NozzlePHFactory
    Set oNozzlePHFactory = New NozzlePHFactory

    Set objCableNozzle = oNozzlePHFactory.CreateCableNozzlePHFromPart(oPartFclt, 2, _
                                                                      m_outputcoll.ResourceManager)

    Set iLogicalDistPort = objCableNozzle
    iLogicalDistPort.SetCenterLocation oCableTrayPortPoint
    Set iDistribPort = objCableNozzle
    iDistribPort.SetDirectionVector oDir

    'Set the output
    iOutput = iOutput + 1
    m_outputcoll.AddOutput arrayOfOutputs(iOutput), objCableNozzle
    Set objCableNozzle = Nothing

    Exit Sub
ErrorHandler:
    ReportUnanticipatedError MODULE, METHOD
End Sub
