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
'   Copyright (c) 2006-2008, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         RRK
'   Creation Date:  12 July, 2006
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'      Source: Warren Controls, http://www.warrencontrols.com/html/pdf/1800ProductSpec.pdf
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   11.FEB.2008     PK      CR-135567       Enhanced the existing 3 way diverter valve to support dimensional basis for
'                                               (i)Face-to-center dimension basis
'                                               (ii)Face-to-Face dimension and Face-to-Center dimension for port 3
'                                               (iii)Asymmetrical Face-to-Center dimension basis
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Insulation:" 'Used for error messages
Const NEGLIGIBLE_THICKNESS = 0.0001

Private Sub Class_Initialize()
    ''
End Sub


Public Sub run(ByVal m_OutputColl As IJDOutputCollection, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim PipeDia1        As Double
    Dim PipeDia2        As Double
    Dim PipeDia3        As Double
    Dim FlangeThick1    As Double
    Dim FlangeThick2    As Double
    Dim FlangeThick3    As Double
    Dim FlangeDia1      As Double
    Dim FlangeDia2      As Double
    Dim FlangeDia3      As Double
    Dim Depth           As Double
    Dim CptOffset       As Double
     
    Dim parFacetoFace       As Double
    Dim parFacetoCenter     As Double
    Dim parInstrumentHeight As Double
    
    Dim parFace1toCenter        As Double
    Dim parFace2toCenter        As Double
    Dim parFace3toCenter        As Double
    Dim parAngle                As Double
    Dim parInsulationThickness  As Double
    
    Dim dFace1toCenter      As Double
    Dim dFace2toCenter      As Double
    Dim dFace3toCenter      As Double
    Dim dAngle              As Double
    Dim dBodyRadius         As Double
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    'parFacetoFace = arrayOfInputs(2)
    'parFacetoCenter = arrayOfInputs(3)
    'parInstrumentHeight = arrayOfInputs(4)
    'parFace1toCenter = arrayOfInputs(5)
    'parFace2toCenter = arrayOfInputs(6)
    'parFace3toCenter = arrayOfInputs(7)
    'parAngle=arrayOfInputs(8)
    'parInsulationThickness=arrayOfInputs(9)

    Dim oPipeComponent      As IJDPipeComponent
    Dim lPartdatabasis      As Long
    
    Set oPipeComponent = oPartFclt
    lPartdatabasis = oPipeComponent.PartDataBasis
    Set oPipeComponent = Nothing
    
    Select Case lPartdatabasis
        
        Case Is <= 1                            'default case
            parFacetoFace = arrayOfInputs(2)
            parFacetoCenter = arrayOfInputs(3)
            parInstrumentHeight = arrayOfInputs(4)
            parInsulationThickness = arrayOfInputs(9)
            dFace1toCenter = parFacetoFace / 2
            dFace2toCenter = parFacetoFace / 2
            dFace3toCenter = parFacetoCenter
                          
        Case 11                                 'Face-to-Face dimension and Face-to-Center dimension for port 3
            parFacetoFace = arrayOfInputs(2)
            parFace3toCenter = arrayOfInputs(7)
            parAngle = arrayOfInputs(8)
            parInsulationThickness = arrayOfInputs(9)
            dFace1toCenter = parFacetoFace / 2
            dFace2toCenter = parFacetoFace / 2
            dFace3toCenter = parFace3toCenter
            dAngle = parAngle
            
        Case 10                                 'Face-to-center dimension basis
            parFacetoCenter = arrayOfInputs(3)
            parAngle = arrayOfInputs(8)
            parInsulationThickness = arrayOfInputs(9)
            dFace1toCenter = parFacetoCenter
            dFace2toCenter = parFacetoCenter
            dFace3toCenter = parFacetoCenter
            dAngle = parAngle
            
        Case 13                                 'Asymmetrical Face-to-Center dimension basis
            parFace1toCenter = arrayOfInputs(5)
            parFace2toCenter = arrayOfInputs(6)
            parFace3toCenter = arrayOfInputs(7)
            parAngle = arrayOfInputs(8)
            parInsulationThickness = arrayOfInputs(9)
            dFace1toCenter = parFace1toCenter
            dFace2toCenter = parFace2toCenter
            dFace3toCenter = parFace3toCenter
            dAngle = parAngle
            If dFace1toCenter = 0 Then
                dFace1toCenter = dFace2toCenter
            ElseIf dFace2toCenter = 0 Then
                dFace2toCenter = dFace1toCenter
            End If
            
        Case Else
            GoTo ErrorLabel:
            
    End Select
 
    'Provide insulation for body of valve (Output 1)
    '-------------------------------------------------------------------------------------
    If lPartdatabasis <= 1 Then
        Dim oInsSphere As Object
        Dim oCentPos   As AutoMath.DPosition
        Dim dInsDiam   As Double
        Set oCentPos = New AutoMath.DPosition
        oCentPos.Set 0, 0, 0
                
        'Assumption: Spherical body diameter is taken to be 80% of parFacetoFace
        
        dBodyRadius = (0.8 * parFacetoFace) / 2
        
        If CmpDblGreaterthan(2 * (dBodyRadius + parInsulationThickness), parFacetoFace) Then
            dInsDiam = parFacetoFace
        Else
            dInsDiam = 2 * (dBodyRadius + parInsulationThickness)
        End If
        
        Set oInsSphere = PlaceSphere(m_OutputColl, oCentPos, dInsDiam / 2)
        
        'Set the output
        m_OutputColl.AddOutput "InsBody", oInsSphere
       
        Set oInsSphere = Nothing
        Set oCentPos = Nothing
        
    ElseIf lPartdatabasis = 10 Or 11 Or 13 Then
        
        RetrieveParameters 1, oPartFclt, m_OutputColl, _
                       PipeDia1, FlangeThick1, FlangeDia1, CptOffset, Depth
        RetrieveParameters 2, oPartFclt, m_OutputColl, _
                       PipeDia2, FlangeThick2, FlangeDia2, CptOffset, Depth
                       
        Dim oInsCylinder    As Object
        Dim dCylLength      As Double
        Dim dInsBodyRadius  As Double
        
        Dim oCylStPoint     As AutoMath.DPosition
        Dim oCylEnPoint     As AutoMath.DPosition
        Set oCylStPoint = New AutoMath.DPosition
        Set oCylEnPoint = New AutoMath.DPosition
        
        'Assumption: Cylinder body radius is taken to be 40% of the minimum of Face 1 to Center and
        'Face 2 to Center and length of cylinder is assumed to be 2 times the maximum of pipe
        'diameter  1 and pipe diameter2
        If CmpDblLessThan(dFace1toCenter, dFace2toCenter) Then
            dBodyRadius = 0.4 * dFace1toCenter
        Else
            dBodyRadius = 0.4 * dFace2toCenter
        End If
        
        'If Insualtion diameter for the body cylinder exceeds the sum of Face 1 to Center and
        'Face 2 to Center then make the insuation radius equal to the maximum of
        'Face 1 to Center and Face 2 to Center
        If CmpDblGreaterthan(2 * dBodyRadius + 2 * parInsulationThickness, dFace1toCenter + dFace2toCenter) Then
            If CmpDblLessThan(dFace1toCenter, dFace2toCenter) Then
                dInsBodyRadius = dFace1toCenter
            Else
                dInsBodyRadius = dFace2toCenter
            End If
        Else
            dInsBodyRadius = dBodyRadius + parInsulationThickness
        End If
        
        If CmpDblGreaterthan(PipeDia1, PipeDia2) Then
            dCylLength = 2 * PipeDia1
        Else
            dCylLength = 2 * PipeDia2
        End If
        
        oCylStPoint.Set 0, 0, dCylLength / 2
        oCylEnPoint.Set 0, 0, -dCylLength / 2
            
        Set oInsCylinder = PlaceCylinder(m_OutputColl, oCylStPoint, oCylEnPoint, 2 * dInsBodyRadius, True)
        
        'Set the output
        m_OutputColl.AddOutput "InsBody", oInsCylinder
    
        Set oInsCylinder = Nothing
        Set oCylStPoint = Nothing
        Set oCylEnPoint = Nothing
    End If
     
    'Place Insulation for Port1
    Dim oStPoint        As AutoMath.DPosition
    Dim oEnPoint        As AutoMath.DPosition
    Dim oInsNozzle1     As Object
    Dim dNozzleLength   As Double
    
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    RetrieveParameters 1, oPartFclt, m_OutputColl, _
                       PipeDia1, FlangeThick1, FlangeDia1, CptOffset, Depth
    
    dNozzleLength = dFace1toCenter - Sqr(Abs(dBodyRadius ^ 2 - (PipeDia1 / 2) ^ 2))
                           
    'If Flange Thickness 1 is equal to zero,then give it some negligible thickness
    If CmpDblEqual(FlangeThick1, 0) Then FlangeThick1 = NEGLIGIBLE_THICKNESS
    
    oStPoint.Set -dFace1toCenter, 0, 0
    
    'Checking whether the insualtion for Port1 flange exceeds nozzle length.If so, make it equal to nozzle length
    If CmpDblGreaterthan(FlangeThick1 + parInsulationThickness, dNozzleLength) Then
        oEnPoint.Set -dFace1toCenter + dNozzleLength, 0, 0
    Else
        oEnPoint.Set -dFace1toCenter + FlangeThick1 + parInsulationThickness, 0, 0
    End If
   
    Set oInsNozzle1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, FlangeDia1 + 2 * parInsulationThickness, True)
    
    'Set the Output
    m_OutputColl.AddOutput "InsNozzle1", oInsNozzle1
    
    Set oInsNozzle1 = Nothing
    
    'Place Insulation for the Nozzle1 Length
    Dim oInsNozzLength1 As Object
  
    oStPoint.Set -dFace1toCenter, 0, 0
    oEnPoint.Set -dFace1toCenter + dNozzleLength, 0, 0
    
    Set oInsNozzLength1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, PipeDia1 + 2 * parInsulationThickness, True)
    
    'Set the Output
    m_OutputColl.AddOutput "InsNozzleLength1", oInsNozzLength1
 
    Set oInsNozzLength1 = Nothing
   
    'Place Insulation for Port2
    RetrieveParameters 2, oPartFclt, m_OutputColl, _
                       PipeDia2, FlangeThick2, FlangeDia2, CptOffset, Depth
    
    Dim oInsNozzle2 As Object
    
    'If Flange Thickness 2 is equal to zero,then give it some negligible thickness
    If CmpDblEqual(FlangeThick2, 0) Then FlangeThick2 = NEGLIGIBLE_THICKNESS
    oStPoint.Set dFace2toCenter, 0, 0
    
    'Checking whether the insualtion for Port2 flange exceeds nozzle length.If so, make it equal to nozzle length
    If CmpDblGreaterthan(FlangeThick2 + parInsulationThickness, dNozzleLength) Then
        oEnPoint.Set dFace2toCenter - dNozzleLength, 0, 0
    Else
        oEnPoint.Set dFace2toCenter - FlangeThick2 - parInsulationThickness, 0, 0
    End If
    
    Set oInsNozzle2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, FlangeDia2 + 2 * parInsulationThickness, True)
    
    'Set the Output
    m_OutputColl.AddOutput "InsNozzle2", oInsNozzle2
    Set oInsNozzle2 = Nothing
    
    'Place Insulation for the Nozzle2 Length
    Dim oInsNozzLength2 As Object
    
    dNozzleLength = dFace2toCenter - Sqr(Abs(dBodyRadius ^ 2 - (PipeDia2 / 2) ^ 2))
    oStPoint.Set dFace2toCenter, 0, 0
    oEnPoint.Set dFace2toCenter - dNozzleLength, 0, 0
    
    Set oInsNozzLength2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, PipeDia2 + 2 * parInsulationThickness, True)
    
    'Set the Output
    m_OutputColl.AddOutput "InsNozzleLength2", oInsNozzLength2
    
    Set oInsNozzLength2 = Nothing
    
    'Place Insulation for Port3
    RetrieveParameters 3, oPartFclt, m_OutputColl, _
                       PipeDia3, FlangeThick3, FlangeDia3, CptOffset, Depth
    
    Dim oInsNozzle3 As Object
   
    'If Flange Thickness 3 is equal to zero,then give it some negligible thickness
    If CmpDblEqual(FlangeThick3, 0) Then FlangeThick3 = NEGLIGIBLE_THICKNESS
    oStPoint.Set 0, -dFace3toCenter, 0
    
    'Checking whether the insualtion for Port3 flange exceeds nozzle length.If so, make it equal to nozzle length
    If CmpDblGreaterthan(FlangeThick3 + parInsulationThickness, dNozzleLength) Then
        oEnPoint.Set 0, -dFace3toCenter + dNozzleLength, 0
    Else
        oEnPoint.Set 0, -dFace3toCenter + FlangeThick3 + parInsulationThickness, 0
    End If
       
    Set oInsNozzle3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, FlangeDia3 + 2 * parInsulationThickness, True)
    
    If lPartdatabasis <= 1 Then
        dAngle = 0
    Else
        dAngle = -parAngle
    End If
           
    Call Rotate(oInsNozzle3, dAngle)
    
    'Set the Output
    m_OutputColl.AddOutput "InsNozzle3", oInsNozzle3
    
    Set oInsNozzle3 = Nothing
    
    'Place Insulation for the Nozzle3 Length
    Dim oInsNozzLength3 As Object

    dNozzleLength = dFace3toCenter - Sqr(Abs(dBodyRadius ^ 2 - (PipeDia3 / 2) ^ 2))
    oStPoint.Set 0, -dFace3toCenter, 0
    oEnPoint.Set 0, -dFace3toCenter + dNozzleLength, 0
    
    Set oInsNozzLength3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, PipeDia3 + 2 * parInsulationThickness, True)

    If lPartdatabasis <= 1 Then
        dAngle = 0
    Else
        dAngle = -parAngle
    End If
    
    Call Rotate(oInsNozzLength3, dAngle)
    
    'Set the Output
    m_OutputColl.AddOutput "InsNozzleLength3", oInsNozzLength3
    
    Set oInsNozzLength3 = Nothing
    
    If lPartdatabasis <= 1 Then
        'Provide Insulation for Cylinder below the Operator
        Dim dCylinderDia    As Double
        Dim dCyl1Height     As Double
        Dim dMeetDistance   As Double

        dCylinderDia = 1.2 * PipeDia3
        dMeetDistance = Sqr(Abs(dBodyRadius ^ 2 - (dCylinderDia / 2) ^ 2))
        dCyl1Height = 0.6 * parInstrumentHeight
        oStPoint.Set 0, dMeetDistance, 0
        oEnPoint.Set 0, dCyl1Height, 0
                
        Dim oCylinder1 As Object
        Set oCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dCylinderDia + 2 * parInsulationThickness, True)
        
        'Set the output
        m_OutputColl.AddOutput "Cylinder1", oCylinder1
        Set oCylinder1 = Nothing
        
        'Provide Insulation for Flange 1
        Dim dFlange1Thk As Double
        Dim oInsFlange  As Object
        
        dFlange1Thk = 0.2 * parInstrumentHeight
        oStPoint.Set 0, dCyl1Height, 0
        oEnPoint.Set 0, dCyl1Height + dFlange1Thk, 0
          
        Set oInsFlange = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.5 * PipeDia3 + 2 * parInsulationThickness, True)
    
        'Set the output
        m_OutputColl.AddOutput "InsFlange1", oInsFlange
        Set oInsFlange = Nothing
        
        'Provide insulation for Flange 2
        Dim dCyl2Height As Double

        If CmpDblGreaterthan(0.5 * parFacetoCenter, dCyl1Height) Then
           dCyl2Height = dCyl1Height
        Else
           dCyl2Height = 0.5 * parFacetoCenter
        End If
        
        oStPoint.Set 0, -dCyl2Height, 0
        oEnPoint.Set 0, -dCyl2Height - dFlange1Thk, 0
        Set oInsFlange = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 1.5 * PipeDia3 + 2 * parInsulationThickness, True)
    
        Call Rotate(oInsFlange, dAngle)
        
        'Set the output
        m_OutputColl.AddOutput "InsFlange2", oInsFlange
        Set oInsFlange = Nothing
    End If
        
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
 Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub

Private Sub Rotate(Port3Obj As Object, RotAngle As Double)
    
    Const METHOD = "Rotate"
    Dim oTransmat      As AutoMath.DT4x4
    Dim oAxisVec       As AutoMath.DVector
    
    Set oAxisVec = New AutoMath.DVector
    Set oTransmat = New AutoMath.DT4x4
    oTransmat.LoadIdentity
    oAxisVec.Set 0, 0, -1
    oTransmat.Rotate RotAngle, oAxisVec
    Port3Obj.Transform oTransmat
    
    Set oTransmat = Nothing
    Set oAxisVec = Nothing
 Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
