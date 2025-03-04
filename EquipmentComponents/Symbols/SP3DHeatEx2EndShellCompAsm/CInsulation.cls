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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:12:38 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2008, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:GL
'   Creation Date:  Friday, Aug 22 2008
'   Description:Added the insulation aspect as per the TR-110721
'
'
'   Change History:
'   dd.mmm.yyyy     who         change description
'   -----------     ---         ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private m_oSymGeomHelper As IJSymbolGeometryHelper
Private Const MODULE = "Insulation:" 'Used for error messages

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
            
   Dim iOutput As Double
    

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
       
' Code for output 6(Shell Flange Insulation)As a Cylinder
    oStPoint.Set 0, 0, -parCPtoFace
    oEnPoint.Set 0, 0, -parCPtoFace - parFaceThickness - parInsulationThickness
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parFaceDiameter + 2 * parInsulationThickness
    
' Code for output 7(Cylindrical Exchanger Shell Body Insulation)As a Cylinder
    oStPoint.Set 0, 0, -parCPtoFace - parFaceThickness
    oEnPoint.Set 0, 0, -parCPtoFace - parVesselHeight + parFaceThickness
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dActualVesselDiameter + 2 * parInsulationThickness

'Code for output 8(Shell Bottom Flange Insulation)

    oStPoint.Set 0, 0, -parCPtoFace - parVesselHeight + parFaceThickness + parInsulationThickness
    oEnPoint.Set 0, 0, -parCPtoFace - parVesselHeight
    iOutput = iOutput + 1
    m_oSymGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, parFaceDiameter + 2 * parInsulationThickness
        
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Exit Sub
    
ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
