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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.

'   ProgID          :  SP3DDiaphragmValveBody.DiaValBody
'   File            :  CInsulation.cls
'   Author          :  PK
'   Creation Date   :  Friday, 26 Oct 2007
'   Description     :  Saunders Standard Forged Diaphragm Valve
'   Change History  :
'   dd.mmm.yyyy     who        change description
'   -----------     -----      ------------------
'   26-OCT-2007      PK        CR-127157:Created the symbol.
'   05-Feb-2008     VRK        CR:135745:Enhance Diaphragm Valve on face-to-face basis for JIS
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
  
    Dim parFacetoFace           As Double
    Dim parValveHeight          As Double
    Dim parValveHeight1         As Double
    Dim parProjectionLength     As Double
    Dim parHandwheelAngle       As Double
    Dim parInsulationThickness  As Double
    
    Dim iOutput As Integer
    
    Dim oCentPos   As AutoMath.DPosition
    Dim oStPoint   As AutoMath.DPosition
    Dim oEnPoint   As AutoMath.DPosition
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    Set oCentPos = New AutoMath.DPosition
    
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoFace = arrayOfInputs(2)
'    parValveHeight = arrayOfInputs(3)
'    parValveHeight1 = arrayOfInputs(4)
'    parProjectionLength = arrayOfInputs(5)
'    parHandwheelAngle = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
    
    oCentPos.Set 0, 0, 0
    iOutput = 0
    
    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                     sptOffset, depth

    Select Case lPartDataBasis

    Case Is <= 1
            parFacetoFace = arrayOfInputs(2)
            parValveHeight = arrayOfInputs(3)
            parValveHeight1 = arrayOfInputs(4)
            parProjectionLength = arrayOfInputs(5)
        
           'Insert your code for Insulation Cylinder
           Dim oInsCyl  As Object
           oStPoint.Set oCentPos.x - parFacetoFace / 2, oCentPos.y, oCentPos.z
           oEnPoint.Set oCentPos.x + parFacetoFace / 2, oCentPos.y, oCentPos.z
           Set oInsCyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + 2 * _
                                                           parInsulationThickness, True)
        
           'Set Output1(Insulation Cylinder)
           iOutput = iOutput + 1
           m_OutputColl.AddOutput arrayOfOutputs(iOutput), oInsCyl
           Set oInsCyl = Nothing
            
           'Insert your code for Insulation body
           Dim oInsBody        As Object
           Dim oTrapezium      As IngrGeom3D.ComplexString3d
           Dim oAxis           As AutoMath.DVector
           Dim lines           As Collection
           Dim oLine           As Object
           Dim oGeomFactory    As IngrGeom3D.GeometryFactory
           Dim dValue1         As Double
           Dim dValue2         As Double
           Dim dValue3         As Double
           
           Set lines = New Collection
           Set oGeomFactory = New IngrGeom3D.GeometryFactory
           Set oAxis = New AutoMath.DVector
           dValue1 = parFacetoFace - 2 * parProjectionLength
           dValue2 = parValveHeight - parValveHeight1
           dValue3 = parValveHeight1 - 0.2 * parValveHeight
           
           Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * (dValue1), _
                           oCentPos.y - (dValue2) - parInsulationThickness, oCentPos.z + 0.3 * _
                           (dValue1) + parInsulationThickness, oCentPos.x - 0.5 * (dValue1), _
                           oCentPos.y - (dValue2) - parInsulationThickness, oCentPos.z - 0.3 * _
                           (dValue1) - parInsulationThickness)
           lines.Add oLine
           Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * (dValue1), _
                           oCentPos.y - (dValue2) - parInsulationThickness, oCentPos.z - 0.3 * _
                           (dValue1) - parInsulationThickness, oCentPos.x - 0.5 * (dValue1), _
                           oCentPos.y + (dValue3) + parInsulationThickness, oCentPos.z - 0.4 * _
                           (dValue1) - parInsulationThickness)
           lines.Add oLine
           Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * (dValue1), _
                           oCentPos.y + (dValue3) + parInsulationThickness, oCentPos.z - 0.4 * _
                           (dValue1) - parInsulationThickness, oCentPos.x - 0.5 * (dValue1), _
                           oCentPos.y + (dValue3) + parInsulationThickness, oCentPos.z + 0.4 * _
                           (dValue1) + parInsulationThickness)
           lines.Add oLine
           Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, oCentPos.x - 0.5 * (dValue1), _
                           oCentPos.y + (dValue3) + parInsulationThickness, oCentPos.z + 0.4 * _
                           (dValue1) + parInsulationThickness, oCentPos.x - 0.5 * (dValue1), _
                           oCentPos.y - (dValue2) - parInsulationThickness, oCentPos.z + 0.3 * _
                           (dValue1) + parInsulationThickness)
           lines.Add oLine
           oStPoint.Set oCentPos.x - 0.5 * (dValue1), oCentPos.y - (dValue2) - _
                           parInsulationThickness, oCentPos.z + 0.3 * (dValue1) + _
                           parInsulationThickness
           
           Set oTrapezium = PlaceTrCString(oStPoint, lines)
               
           oAxis.Set 1, 0, 0
           Set oInsBody = PlaceProjection(m_OutputColl, oTrapezium, oAxis, dValue1, True)
           
           'Set Output2(Insulation body)
           iOutput = iOutput + 1
           m_OutputColl.AddOutput arrayOfOutputs(iOutput), oInsBody
           
           Set oGeomFactory = Nothing
           Set oAxis = Nothing
           Set oTrapezium = Nothing
           Set oLine = Nothing
           Set lines = Nothing
           Set oInsBody = Nothing
                
           'Insert your code for Operator Base
           Dim oInsBase As Object
           oStPoint.Set oCentPos.x - parFacetoFace / 2, oCentPos.y + (dValue3) - _
                           parInsulationThickness, oCentPos.z + 0.5 * (dValue1) + parInsulationThickness
           oEnPoint.Set oCentPos.x + parFacetoFace / 2, oCentPos.y + parValveHeight1 + _
                           parInsulationThickness, oCentPos.z - 0.5 * (dValue1) - parInsulationThickness
            
           Set oInsBase = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
           iOutput = iOutput + 1
           m_OutputColl.AddOutput arrayOfOutputs(iOutput), oInsBase
           Set oInsBase = Nothing
           
      Case 5
    
        parFacetoFace = arrayOfInputs(2)
        parValveHeight = arrayOfInputs(3)
        
        'Cone1(Output 1)
        Dim objInsCone As Object
        oStPoint.Set oCentPos.x - (0.5 * parFacetoFace), oCentPos.y, oCentPos.z
        oEnPoint.Set oCentPos.x, oCentPos.y, oCentPos.z
        Set objInsCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam / 2 + (parInsulationThickness), 0, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCone
        Set objInsCone = Nothing
        
        'Cone2(Output 2)
        oStPoint.Set oCentPos.x + (0.5 * parFacetoFace), oCentPos.y, oCentPos.z
        oEnPoint.Set oCentPos.x, oCentPos.y, oCentPos.z
        Set objInsCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam / 2 + (parInsulationThickness), 0, True)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsCone
        Set objInsCone = Nothing
        
        'Insert your code for output 3(Sphere)
        Dim ObjValveBody As Object
        Set ObjValveBody = PlaceSphere(m_OutputColl, oCentPos, parFacetoFace / 8)
        
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjValveBody
        Set ObjValveBody = Nothing
        
    Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.
        
    End Select
           
           Set oStPoint = Nothing
           Set oEnPoint = Nothing
           Set oCentPos = Nothing
           
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
