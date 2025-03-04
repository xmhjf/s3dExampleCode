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
'   CSimplePhysical.cls
'   Author:          ACM
'   Creation Date:  Tuesday, May 16, 2003
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003  SymbolTeam(India)   Copyright Information, Header  is added.
'   22.Aug.2003  SymbolTeam(India)   TR 46728 Change Port Index of a nozzle to 15 used to corrupt Equipment.
'                                    Modified port index logic.
'   23.Aug.2003  SymbolTeam(India)       Updated Error Handling Code
'   20.May.2005    MS  CR-76070: Modified the nozzle creation code(as members),added datum points.
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'   This is an Exchanger Shell Body Symbol.
'   This symbol is Fully parametric symbol, consists of 77 inputs and 10 outputs.It has 2 aspects one is Simple Physical
'   another is ReferenceGeometry. Among them 8 are physical and 2 is ReferenceGeometry. It has Five nozzles and all
'   are fully parametric.
'   This is a Dow Equipment E-161 Exchanger Shell symbol.

Option Explicit

Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
     Set m_oSymGeomHelper = New SymbolServices
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
Private Sub Class_Terminate()
    Set m_oSymGeomHelper = Nothing
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    

    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim parVesselDiameter As Double
    Dim parVesselHeight As Double
    Dim parCPtoFace As Double
    Dim parFaceDiameter As Double
    Dim parFaceThickness As Double
    Dim parInsulationThickness As Double
        
    Dim iOutput     As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parVesselHeight = arrayOfInputs(3)
    parCPtoFace = arrayOfInputs(4)
    parFaceDiameter = arrayOfInputs(5)
    parFaceThickness = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
     
    m_oSymGeomHelper.OutputCollection = m_OutputColl
    
    Dim oStPoint As New AutoMath.DPosition
    Dim oEnPoint As New AutoMath.DPosition
    
    Dim dActualVesselDiameter As Double
    Dim dDomeHeight As Double

''Assumptions
    'The input parameter for Vessel Diameter is Internal diameter .We need to take care for Thickness of Body.
    dActualVesselDiameter = parVesselDiameter + 0.0508   ' 2 Inch
    dDomeHeight = dActualVesselDiameter / 4  '2:1 ELiptical Head (TYP)
    
' Insert your code for output 1(Shell Flange)As a Cylinder
    oStPoint.Set 0, 0, -parCPtoFace
    oEnPoint.Set 0, 0, -parCPtoFace - parFaceThickness
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parFaceDiameter
    
' Insert your code for output 2(Cylindrical Exchanger Shell Body)As a Cylinder
    oStPoint.Set 0, 0, -parCPtoFace - parFaceThickness
    oEnPoint.Set 0, 0, -parCPtoFace - parVesselHeight + parFaceThickness
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dActualVesselDiameter

'Insert your code for output 3(Shell Bottom Flange)

    oStPoint.Set 0, 0, -parCPtoFace - parVesselHeight + parFaceThickness
    oEnPoint.Set 0, 0, -parCPtoFace - parVesselHeight
' Create cylinder using m_oGeomHelper method which also sets the output
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parFaceDiameter
    
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Exit Sub
    
ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

