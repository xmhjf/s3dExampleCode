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
'   Author:          RUK
'   Creation Date:  Monday, Aug 27 2007
'   Description:
'    CR-124211 Generic Fire Damper.
'       Source STD-M13.pdf.
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   27-Aug-2007     RUK     Created
'   07.Nov.2007     RUK     DI-CP·56531 Ability for custom weight and cg calculation for parametric standard components.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = Insulation: 'Used for error messages
Private PI       As Double

Private Sub Class_Initialize()
      PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt   As PartFacelets.IJDPart
    Dim iOutput     As Double
    
    Dim parHVACShape As Double
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBladetoPort1 As Double
    Dim parPPtoBlade As Double
    Dim parPPtoPort2 As Double
    Dim parDamperThick As Double
    Dim parBladeAngle As Double
    Dim parInsulationThickness As Double

'   Inputs
    Set oPartFclt = arrayOfInputs(1)
    parHVACShape = arrayOfInputs(2)
    parWidth = arrayOfInputs(3)
    parDepth = arrayOfInputs(4)
    parBladetoPort1 = arrayOfInputs(5)
    parPPtoBlade = arrayOfInputs(6)
    parPPtoPort2 = arrayOfInputs(7)
    parDamperThick = arrayOfInputs(8)   'Not used in the code
    parBladeAngle = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)

    'Origin is taken at the Center of the Damper
    'For Horizontally routed Duct X - Along the flow and Z - towards UP
    iOutput = 0
    
    If parHVACShape = 4 Then 'Round
        parDepth = parWidth
    End If
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    
    Set oGeomFactory = New GeometryFactory
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition

    'Create the Damper (Output 1)
    Dim objInsDamper As Object
    If parHVACShape = Rectangular Then
        oStPoint.Set -(parPPtoBlade + parBladetoPort1), _
                            -(parDepth / 2 + parInsulationThickness), _
                            -(parWidth / 2 + parInsulationThickness)
        oEnPoint.Set parPPtoPort2, (parDepth / 2 + parInsulationThickness), _
                         (parWidth / 2 + parInsulationThickness)
        Set objInsDamper = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
    ElseIf parHVACShape = 4 Then 'Round
        oStPoint.Set -(parPPtoBlade + parBladetoPort1), 0, 0
        oEnPoint.Set parPPtoPort2, 0, 0
        Set objInsDamper = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                parWidth + 2 * parInsulationThickness, True)
    End If

    'Set the Output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsDamper
    Set objInsDamper = Nothing

    'Remove references
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oGeomFactory = Nothing
    
    'New code for WCGOutputs DI-CP·56531
    Dim oGenBOFactory As IJGeneralBusinessObjectsFactory
    Dim oAttributes As IJDAttributes
    Dim oAttribCol As IJDAttributesCol
    Set oGenBOFactory = New GeneralBusinessObjectsFactory
    
    Set oAttributes = oGenBOFactory.CreateWCGPropHolder(m_OutputColl.ResourceManager)

    'In V8 we should use the IJWCGPropHolder interface to set these values
    Set oAttribCol = oAttributes.CollectionOfAttributes(CVar("IJWCGPropHolder"))
    
    'Calculate the Volume
        'Assumed that the thickness of duct is negligible
    Dim dTotVolume As Double
    Dim dHeaderVolume As Double
    Dim dBladeVolume As Double
    Dim dCGX, dCGY, dCGZ As Double
    
'    If parHVACShape = Rectangular Then
'        dHeaderVolume = ((parWidth + 2 * parInsulationThickness) * (parDepth + 2 * parInsulationThickness) - _
'                            (parWidth * parDepth)) * (parBladetoPort1 + parPPtoBlade + parPPtoPort2)
'        dBladeVolume = parWidth * parDepth * 2 * parInsulationThickness
'    ElseIf parHVACShape = 4 Then 'Round
'        dHeaderVolume = (PI * ((parWidth / 2 + parInsulationThickness) ^ 2) - (PI * ((parWidth / 2) ^ 2))) * _
'                                (parBladetoPort1 + parPPtoBlade + parPPtoPort2)
'        dBladeVolume = PI * ((parWidth / 2) ^ 2) * 2 * parInsulationThickness
'    End If
'    dTotVolume = dHeaderVolume + dBladeVolume
'
'    If Not CmpDblEqual(dTotVolume, LINEAR_TOLERANCE) Then
'    dCGX = -((dBladeVolume * (((parBladetoPort1 + parPPtoBlade + parPPtoPort2) / 2) _
'                                            - parBladetoPort1) / dTotVolume) + _
'            (((parBladetoPort1 + parPPtoBlade + parPPtoPort2) / 2) - parPPtoPort2))
'    Else
'        dCGX = ((parBladetoPort1 + parPPtoBlade + parPPtoPort2) / 2) - parPPtoPort2
'    End If
    dCGX = ((parBladetoPort1 + parPPtoBlade + parPPtoPort2) / 2) - parPPtoPort2
    dCGY = 0
    dCGZ = 0
    
    oAttribCol.Item("Volume").Value = dTotVolume '(cubic meters)
    oAttribCol.Item("CGX").Value = dCGX
    oAttribCol.Item("CGY").Value = dCGY
    oAttribCol.Item("CGZ").Value = dCGZ
    
    m_OutputColl.AddOutput "InsulationWCG", oAttributes
    Set oAttribCol = Nothing
    Set oAttributes = Nothing

    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
