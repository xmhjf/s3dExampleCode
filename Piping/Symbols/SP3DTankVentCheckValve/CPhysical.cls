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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author: RRK
'   Creation Date:  Friday, 13, Jul 2007
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    The symbol is modelled for RM-2 Type Tank Vent Check Valve of Winel
'    (source:http://www.winel.nl/Products/tvcv.htm)
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ----   ---------------------------------------
'   13.Jul.2007     RRK     CR-123055: Created the symbol.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'
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

    Dim iOutput     As Double
    
    Dim parHeight As Double
    Dim parDiameter As Double
    Dim parLength As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHeight = arrayOfInputs(2)
    parDiameter = arrayOfInputs(3)
    parLength = arrayOfInputs(4)
    parInsulationThickness = arrayOfInputs(5)
    
    iOutput = 0
'Assumption : Origin is at Port 1.

' Insert your code for Cylindrical portion of the body (output 1)
    'Assumption : The Length of the cylinder is 80% of the Length.
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    oStPoint.Set parHeight - (parDiameter / 2), 0.4 * parLength, 0
    Set oEnPoint = New AutoMath.DPosition
    oEnPoint.Set parHeight - (parDiameter / 2), -0.4 * parLength, 0
    
    Dim objCylinderBody As Object
    Set objCylinderBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parDiameter, True)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objCylinderBody
    Set objCylinderBody = Nothing

' Insert your code for output 2(Side Cover 1)
    'Assumptions : 1) The top of the box is assumed to be at 90% of Height2
    '              2) The bottom of the box is assumed to be at 10% of Diameter from the bottom of the cylinder
    '              3)Box width is assumed to be 0.875 times the chord length at box bottom position
    
    Dim dBoxTopPosX As Double
    Dim dBoxBottomPosX As Double
    Dim dBoxWidth As Double
    Dim dX As Double
    
    dBoxTopPosX = 0.8 * parHeight
    dBoxBottomPosX = parHeight - parDiameter + (0.1 * parDiameter)
    
    dX = parHeight - (parDiameter / 2) - dBoxBottomPosX
    dBoxWidth = 1.75 * Sqr((parDiameter / 2) ^ 2 - dX ^ 2)
    
    oStPoint.Set dBoxBottomPosX, 0.4 * parLength, dBoxWidth / 2
    oEnPoint.Set dBoxTopPosX, 0.5 * parLength, -dBoxWidth / 2
    
    Dim objSideCover1 As Object
    Set objSideCover1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSideCover1
    Set objSideCover1 = Nothing
    
' Insert your code for output 3(Side Cover 2)
    oStPoint.Set dBoxBottomPosX, -0.4 * parLength, dBoxWidth / 2
    oEnPoint.Set dBoxTopPosX, -0.5 * parLength, -dBoxWidth / 2
    
    Dim objSideCover2 As Object
    Set objSideCover2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSideCover2
    Set objSideCover2 = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
' Place Nozzle 1
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir        As AutoMath.DVector
    Dim objNozzle   As GSCADNozzleEntities.IJDNozzle
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, sptOffset, depth
    
    Set oPlacePoint = New AutoMath.DPosition
    oPlacePoint.Set depth - sptOffset, 0, 0
    Set oDir = New AutoMath.DVector
    oDir.Set -1, 0, 0
  
    Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, parHeight - (parDiameter / 2))
    
'   Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
    Set objNozzle = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

