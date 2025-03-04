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
'   Copyright (c) 2004, Intergraph Corporation. All rights reserved.
'   All Rights Reserved
'
'   CInsulation.cls
'   Author:         ACM
'   Creation Date:  Monday, July 12, 2004
'   Description:
'       This is a multi port diverver valve symbol. This is prepared based on Saunder's catalog.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Dim m_oGeomHelper As IJSymbolGeometryHelper

Private Const MODULE = "SimplePhysical:" 'Used for error messages
Private Sub Class_Terminate()
    Set m_oGeomHelper = Nothing
End Sub

Private Sub Class_Initialize()
    Set m_oGeomHelper = New SymbolServices
    
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
    
'   Insert your code for output 7(Insulated Valve Body)
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
    
    dLinePoints(0) = parInsulationThickness + parValveBodyWidth / 2
    dLinePoints(1) = -parInsulationThickness - parValveBodyHeight / 2
    dLinePoints(2) = parInsulationThickness + parValveBodyDepth / 2
    
    dLinePoints(3) = parInsulationThickness + parValveBodyWidth / 2
    dLinePoints(4) = dStepHeight - parValveBodyHeight / 2
    dLinePoints(5) = dLinePoints(2)
            
    dLinePoints(6) = parInsulationThickness + dValveBodyTopWidth / 2
    dLinePoints(7) = parInsulationThickness + parValveBodyHeight / 2
    dLinePoints(8) = dLinePoints(2)
    
    dLinePoints(9) = -parInsulationThickness - dValveBodyTopWidth / 2
    dLinePoints(10) = parInsulationThickness + parValveBodyHeight / 2
    dLinePoints(11) = dLinePoints(2)
    
    dLinePoints(12) = -parInsulationThickness - parValveBodyWidth / 2
    dLinePoints(13) = dStepHeight - parValveBodyHeight / 2
    dLinePoints(14) = dLinePoints(2)
    
    dLinePoints(15) = -parInsulationThickness - parValveBodyWidth / 2
    dLinePoints(16) = -parInsulationThickness - parValveBodyHeight / 2
    dLinePoints(17) = dLinePoints(2)

    dLinePoints(18) = parInsulationThickness + parValveBodyWidth / 2
    dLinePoints(19) = -parInsulationThickness - parValveBodyHeight / 2
    dLinePoints(20) = parInsulationThickness + parValveBodyDepth / 2
    
    Set oLineString = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 7, dLinePoints)
            
    oDirVec.Set 0, 0, -1
            
    Set ObjValveBody = PlaceProjection(m_OutputColl, oLineString, oDirVec, parValveBodyDepth, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
    Set ObjValveBody = Nothing

'   Insert your code for output 8(Nozzle 1 Insulation)
    Dim dInsulationDia As Double
    Dim oStartPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
   
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth
   
    If pipeDiam < flangeDiam Then
        dInsulationDia = flangeDiam + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set 0, parFace1toCenter - parValveBodyHeight / 2, 0
    oEndPoint.Set 0, parFace1toCenter, 0
   ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia

'   Insert your code for output 9(Nozzle 2 Insulation)
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth
   
    If pipeDiam < flangeDiam Then
        dInsulationDia = flangeDiam + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set parValveBodyWidth / 2 - 0.6 * pipeDiam, -parFace2toCenter + parValveBodyHeight / 2, 0
    oEndPoint.Set parValveBodyWidth / 2 - 0.6 * pipeDiam, -parFace2toCenter, 0
   ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia


'   Insert your code for output 4(Nozzle 3 with length)
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                        flangeThick, flangeDiam, cptOffset, depth
   
    If pipeDiam < flangeDiam Then
        dInsulationDia = flangeDiam + 2 * parInsulationThickness
    Else
        dInsulationDia = pipeDiam + 2 * parInsulationThickness
    End If
    
    oStartPoint.Set -parValveBodyWidth / 2 + 0.6 * pipeDiam, -parFace3toCenter + parValveBodyHeight / 2, 0
    oEndPoint.Set -parValveBodyWidth / 2 + 0.6 * pipeDiam, -parFace3toCenter, 0
   ' Set the output
    iOutput = iOutput + 1
    m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStartPoint, oEndPoint, dInsulationDia

    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
   
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

