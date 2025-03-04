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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:         ACM
'   Creation Date:  Monday, July 27, 2004
'   Description:
'   This is a 2-way diverter valve symbol. This is prepared based on Gemu's catalog.
'   Source: Gemu  Multiport Valves M600-3-2C, Aseptic valve manifold machined from a single block.
'   M600 valve manifold designs, developed and produced according to customer requirements/specifications
'   The symbol has multiple operators. Each operator for each output port
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper
Dim PI As Double
Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim flangeDiam      As Double
    Dim cptOffset       As Double
    Dim depth           As Double
    
    Dim iOutput     As Double
    
    Dim parValveBodyWidth As Double
    Dim parValveBodyDepth As Double
    Dim parValveBodyHeight As Double
    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parValveBodyWidth = arrayOfInputs(2)
    parValveBodyDepth = arrayOfInputs(3)
    parValveBodyHeight = arrayOfInputs(4)
    parFace1toCenter = arrayOfInputs(5)
    parFace2toCenter = arrayOfInputs(6)
    parFace3toCenter = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
    
    iOutput = 0
    
    'Origin is taken at the geometry center of the valve.
    
    m_oGeomHelper.OutputCollection = m_OutputColl
    
'   Insert your code for output 1(Valve Body)
    Dim ObjValveBody As Object
    
    Dim oGeomFactory As New IngrGeom3D.GeometryFactory
    Dim oLineString As IngrGeom3D.LineString3d
    Dim dLinePoints(0 To 20) As Double
    Dim dStepHeight As Double
    Dim oDirVec As New AutoMath.DVector
        
    dStepHeight = 0.1 * parValveBodyHeight ''This is assuemd 10% of the total ValveBody Height
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth
    Dim dValveBodyTopWidth As Double
    ''Width of the Body Top is assumed to be 1.5 times the Pipe diameter of Nozzle 1.
    dValveBodyTopWidth = 1.5 * pipeDiam
    
    dLinePoints(0) = parValveBodyWidth / 2
    dLinePoints(1) = -parValveBodyHeight / 2
    dLinePoints(2) = parValveBodyDepth / 2
    
    dLinePoints(3) = parValveBodyWidth / 2
    dLinePoints(4) = dStepHeight - parValveBodyHeight / 2
    dLinePoints(5) = parValveBodyDepth / 2
            
    dLinePoints(6) = dValveBodyTopWidth / 2
    dLinePoints(7) = parValveBodyHeight / 2
    dLinePoints(8) = parValveBodyDepth / 2
    
    dLinePoints(9) = -dValveBodyTopWidth / 2
    dLinePoints(10) = parValveBodyHeight / 2
    dLinePoints(11) = parValveBodyDepth / 2
    
    dLinePoints(12) = -parValveBodyWidth / 2
    dLinePoints(13) = dStepHeight - parValveBodyHeight / 2
    dLinePoints(14) = parValveBodyDepth / 2
    
    dLinePoints(15) = -parValveBodyWidth / 2
    dLinePoints(16) = -parValveBodyHeight / 2
    dLinePoints(17) = parValveBodyDepth / 2
    
    dLinePoints(18) = parValveBodyWidth / 2
    dLinePoints(19) = -parValveBodyHeight / 2
    dLinePoints(20) = parValveBodyDepth / 2
    
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLinePoints)
            
    oDirVec.Set 0, 0, -1
            
    Set ObjValveBody = PlaceProjection(m_OutputColl, oLineString, oDirVec, parValveBodyDepth, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
    Set ObjValveBody = Nothing


'   Insert your code for output 2(Nozzle 1 with length)
   
    Dim objNozzle As GSCADNozzleEntities.IJDNozzle
    Dim oPlacementPoint As New AutoMath.DPosition
    
    oPlacementPoint.Set 0, parFace1toCenter + cptOffset - depth, 0
    oDirVec.Set 0, 1, 0
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace1toCenter - parValveBodyHeight / 2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

'   Insert your code for output 3(Nozzle 2 with length)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth
                        
    oPlacementPoint.Set parValveBodyWidth / 2 - 0.6 * pipeDiam, _
                        -parFace2toCenter - cptOffset + depth, 0
    oDirVec.Set 0, -1, 0
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace2toCenter - parValveBodyHeight / 2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing

'   Insert your code for output 4(Nozzle 3 with length)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth
                        
    oPlacementPoint.Set -parValveBodyWidth / 2 + 0.6 * pipeDiam, _
                        -parFace2toCenter - cptOffset + depth, 0
    oDirVec.Set 0, -1, 0
    Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, _
                            oDirVec, oPlacementPoint, parFace2toCenter - parValveBodyHeight / 2)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    
' Insert your code for output 15(Valve Operator Body)

    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector

    Dim dOperatorAngle As Double

    dOperatorAngle = Atn((parValveBodyHeight - dStepHeight) / ((parValveBodyWidth - dValveBodyTopWidth) / 2))

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    On Error GoTo ErrorLabel

    Dim oOperatorPart As IJValveOperator
    Dim oMultiValCol As IJMultiPortValveAccCol
    Dim oOperatorOcc   As IJPartOcc

    If Not oPipeComponent Is Nothing Then
      oPipeComponent.GetValveOperatorsForPorts oMultiValCol
    End If

    Dim OpOrigin As IJDPosition
    Set OpOrigin = New DPosition
    OpOrigin.Set 0, 0, 0

''Operator for port 2
      oDirX.Set 0, 0, 1 ''Operator local X-axis matches with Valve Z-axis
      oDirY.Set Sin(dOperatorAngle), Cos(dOperatorAngle), 0
      oDirZ.Set -Cos(dOperatorAngle), Sin(dOperatorAngle), 0

      Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(2)
            If Not oOperatorPart Is Nothing Then
                Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator1", _
                                        oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

            End If
    Set oOperatorOcc = Nothing
    Set oOperatorPart = Nothing

''Operator for port 3
      oDirX.Set 0, 0, 1 ''Operator local X-axis matches with Valve Z-axis
      oDirY.Set -Sin(dOperatorAngle), Cos(dOperatorAngle), 0
      oDirZ.Set -Cos(dOperatorAngle), -Sin(dOperatorAngle), 0

    Set oOperatorPart = oMultiValCol.GetValveOperatorPartPerPort(3)
            If Not oOperatorPart Is Nothing Then
                Set oOperatorOcc = m_oGeomHelper.CreateChildPartOcc("ValveOperator2", _
                                        oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

            End If

    Set oOperatorOcc = Nothing
    Set oOperatorPart = Nothing

    Set oPlacementPoint = Nothing
    Set oDirVec = Nothing
    Set oDirX = Nothing
    Set oDirY = Nothing
    Set oDirZ = Nothing
   
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub
