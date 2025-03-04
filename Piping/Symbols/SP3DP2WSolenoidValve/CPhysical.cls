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
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Sep 23 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is: I7AZ Page No. 355 of PDS Piping Component Data Reference Guide pcd.pdf.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  06.Dec.2004     SymbolTeam(India) Removed the small value (0.0000001) that was added to
'                                    the end point for cones in symbol code (Which was a temporary fix).
'  08.SEP.2006     KKC  DI-95670     Replace names with initials in all revision history sheets and symbols.
'  18.Apr.2007     dkl  CR-118302    Enabled associating the symbol with different actuators.
'                                    Implemented part data basis to create geometry for Solenoid valves
'                                    Type D and Type P as per JIS B8471.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI       As Double
Private Const E_NOOPERATORDEFINED = &H80040400 'Corresponds to E_PIPECOMP_NOOPERATORDEFINED in RefDataPipingFacelets.idl
Private Const E_NOOPERATORPARTFOUND = &H80040401 'Corresponds to E_PIPECOMP_NOVALVEPARTFOUND in RefDataPipingFacelets.idl
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
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
    
    Dim ObjLeftCone As Object
    Dim ObjRightCone As Object
    Dim ObjNozzle1 As Object
    Dim ObjNozzle2 As Object
    Dim parFacetoCenter As Double
    Dim parFace1toCenter As Double
    Dim parInsulationThickness As Double
    Dim parHandwheelAngle As Double
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFacetoCenter = arrayOfInputs(2)
'    parFace1toCenter = arrayOfInputs(3)
    parInsulationThickness = arrayOfInputs(4)
'    parFacetoFace = arrayOfInputs(5)
'    parHandwheelAngle = arrayOfInputs(6)

    iOutput = 0
    
    'The operator can be associated with Part data basis values 1, 255 and 260. The existing symbols
    'not associated with Hand-wheel angle continue to function.
    If UBound(arrayOfInputs) >= 6 Then
        parHandwheelAngle = arrayOfInputs(6)
    End If
    
'The center is assumed at the center of the center-line joining the two ports.

    Dim oPipeComponent As IJDPipeComponent
    Set oPipeComponent = oPartFclt
    Dim lPartDataBasis As Integer
    lPartDataBasis = oPipeComponent.PartDataBasis

' This symbol is based on the following three Part data basis values that govern its geometry,
' 'Default' - 1, 'Solenoid Valve type 1 - 255 , and 'Solenoid Valve type 2 - 260.

    Select Case lPartDataBasis
        
        Case Is <= 1 'The default state, output as two cones and input parameters - Face to Center
                     ' and Face 1 to Center.
            parFacetoCenter = arrayOfInputs(2)
            parFace1toCenter = arrayOfInputs(3)
            
            'Insert your code for output 1(Left Cone)
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                        sptOffset, depth
             
            oStPoint.Set -parFacetoCenter + flangeThick, 0#, 0#
            oEnPoint.Set 0#, 0#, 0#
            
            Set ObjLeftCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam / 2, _
                                                                             0, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjLeftCone
            Set ObjLeftCone = Nothing
            
            'Insert your code for output 2(Right Cone)
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                        sptOffset, depth
            oStPoint.Set parFace1toCenter - flangeThick, 0#, 0#
            oEnPoint.Set 0#, 0#, 0#
            Set ObjRightCone = PlaceCone(m_OutputColl, oStPoint, oEnPoint, pipeDiam / 2, _
                                                                            0, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjRightCone
            Set ObjRightCone = Nothing
    
        Case 255 'Solenoid Valve type 1 (Type D, Ref JIS B8471).
            Dim parFacetoFace As Double
            parFacetoFace = arrayOfInputs(5)
            parFacetoCenter = parFacetoFace / 2
            parFace1toCenter = parFacetoFace / 2

            'Constructing Smaller diameter horizontal cylinderical body in between the sockets / flanges
            'Assumptions: The diameter Smaller diameter horizontal cylinder is equal to Pipe OD.
            
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                        sptOffset, depth
                       
            oStPoint.Set -(parFacetoFace / 2 - flangeThick), 0, 0
            oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
            
            Dim objHorSmallerCylinder As Object
            Set objHorSmallerCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                    pipeDiam, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHorSmallerCylinder
            Set objHorSmallerCylinder = Nothing
        
            'Constructing Larger diameter horizontal cylinderical body
            'Assumption: 1. The diameter of the larger cylinder is equal to 1.2 x flange dia.
            ' 2. The Projection of Smaller cylinder outside the Larger horizontal cylinder
            '    is equal to Hub/Flange Thickness.
            
            Dim dProjectionofSmallerCyl As Double
            If CmpDblEqual(flangeThick, 0) Then
                dProjectionofSmallerCyl = parFacetoFace / 50
            Else
                dProjectionofSmallerCyl = flangeThick
            End If
            
            Dim dLengthofLargerCyl As Double
            dLengthofLargerCyl = 2 * (parFacetoFace / 2 - dProjectionofSmallerCyl - flangeThick)
                    
            If CmpDblEqual(flangeDiam, 0) Then flangeDiam = 2 * pipeDiam
                    
            Dim dRadiusofLargerCyl As Double
            dRadiusofLargerCyl = 1.2 * flangeDiam
            
            If CmpDblGreaterthan(dRadiusofLargerCyl, dLengthofLargerCyl / 2) Then
            dRadiusofLargerCyl = dLengthofLargerCyl / 3
            End If
            
            oStPoint.Set -dLengthofLargerCyl / 2, 0, 0
            oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
                    
            Dim objHorLargerCylinder As Object
            Set objHorLargerCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                2 * dRadiusofLargerCyl, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHorLargerCylinder
            Set objHorLargerCylinder = Nothing

            'Constructing Lower vertical cylinderical body
            'Assumptions:
            '1. The diameter of the Lower vertical cylinderical body is equal to the
            '   diameter of the larger cylinder.
            '2. The height of the Lower vertical cylinderical body is equal to half of the radius
            '   of Larger diameter horizontal cylinderical body.
                    
            oStPoint.Set 0, 0, 0
            oEnPoint.Set oStPoint.x, dRadiusofLargerCyl * 3 / 2, oStPoint.z
                    
            Dim objLowerVerticalCylinder As Object
            Set objLowerVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                2 * dRadiusofLargerCyl, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLowerVerticalCylinder
            Set objLowerVerticalCylinder = Nothing
   
            'Constructing Upper vertical cylinderical body
            'Assumptions:
            '1. The diameter of the Upper vertical cylinderical body is equal to the
            '   1.2 times the diameter of the lower cylinderical body.
            '2. The height of the Upper vertical cylinderical body is equal to half of the height
            '   of the lower cylinderical body.
                    
            oStPoint.Set 0, dRadiusofLargerCyl * 3 / 2, 0
            oEnPoint.Set oStPoint.x, oStPoint.y + dRadiusofLargerCyl / 4, oStPoint.z
                    
            Dim objUpperVerticalCylinder As Object
            Set objUpperVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                            2.4 * dRadiusofLargerCyl, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objUpperVerticalCylinder
            Set objUpperVerticalCylinder = Nothing
            
        Case 260 'Solenoid Valve type 2(Type P, Ref JIS B8471).
            parFacetoFace = arrayOfInputs(5)
            parFacetoCenter = parFacetoFace / 2
            parFace1toCenter = parFacetoFace / 2
            
            'Constructing smaller diameter horizontal cylinderical body in between the sockets / flanges
            'Assumptions: The diameter Smaller diameter horizontal cylinder is equal to Pipe OD.
            
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                        sptOffset, depth
                                                                                
            Set oStPoint = New AutoMath.DPosition
            Set oEnPoint = New AutoMath.DPosition
            
            oStPoint.Set -(parFacetoFace / 2 - flangeThick), 0, 0
            oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
            
            Set objHorSmallerCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                                    pipeDiam, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHorSmallerCylinder
            Set objHorSmallerCylinder = Nothing
   
            ' Constructing elliptical body of the valve
            ' Assumptions:
            ' Height of elliptical body of the valve (from the center) is 1.2 x Hub outside radius.
            
            If CmpDblEqual(flangeThick, 0) Then
                dProjectionofSmallerCyl = parFacetoFace / 50
            Else
                dProjectionofSmallerCyl = flangeThick
            End If
            
            If CmpDblEqual(flangeDiam, 0) Then flangeDiam = 2 * pipeDiam
            
            Dim dHeightofEllipticalArc As Double
            dHeightofEllipticalArc = 1.2 * flangeDiam
            
            Dim dHalfBodyWidth As Double
            dHalfBodyWidth = parFacetoFace / 2 - dProjectionofSmallerCyl - flangeThick
            
            Dim oGeomFactory As IngrGeom3D.GeometryFactory
            Set oGeomFactory = New IngrGeom3D.GeometryFactory

            Dim objEllipticalArc As IngrGeom3D.EllipticalArc3d
            Set objEllipticalArc = _
                    oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                    0, pipeDiam / 2, 0, _
                    0, 0, 1, _
                    dHalfBodyWidth, 0, 0, _
                    ((dHeightofEllipticalArc - pipeDiam / 2) / dHalfBodyWidth), 0, PI)

            Dim oAxisVec As AutoMath.DVector
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Dim oCenPoint As AutoMath.DPosition
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0.0001, 0
            
            Dim objValveBody As Object
            Set objValveBody = PlaceRevolution(m_OutputColl, objEllipticalArc, oAxisVec, _
                                            oCenPoint, 2 * PI, True)
            'Set the output 3
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveBody
            Set objValveBody = Nothing
            Set oAxisVec = Nothing
            Set oCenPoint = Nothing
            Set objEllipticalArc = Nothing
            Set oGeomFactory = Nothing
      
            'Constructing Lower vertical cylinderical body
            'Assumptions:
            '1. The diameter of the Lower vertical cylinderical body is equal to the
            '   Half width of Elliptical Arc (of the ellptical body created above) / 1.4
            '2. The height of the Lower vertical cylinderical body is equal to
            '   Height of Elliptical Arc (of the ellptical body created above) / 3
            
            Dim dDiaLowerVerticalCylinder As Double
            dDiaLowerVerticalCylinder = dHalfBodyWidth / 1.4
            
            oStPoint.Set 0, 0, 0
            oEnPoint.Set oStPoint.x, oStPoint.y + dHeightofEllipticalArc * 4 / 3, oStPoint.z
                    
            Set objLowerVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                       dDiaLowerVerticalCylinder, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLowerVerticalCylinder
            Set objLowerVerticalCylinder = Nothing
         
            'Constructing Upper vertical cylinderical body
            'Assumptions:
            '1. The diameter of the Upper vertical cylinderical body is equal to the
            '   1.2 times the diameter of the lower cylinderical body.
            '2. The height of the Upper vertical cylinderical body is equal to half of the height
            '   of the lower cylinderical body.
            '3. This cylinder is offset to the Lower vertical cylinder by 20% of the diameter of
            '   the Lower vertical cylinder.
                        
            Dim dOffset As Double
            dOffset = dDiaLowerVerticalCylinder * 0.2
            
            oStPoint.Set -dOffset, oEnPoint.y, 0
            oEnPoint.Set -dOffset, oStPoint.y + dHeightofEllipticalArc / 8, oStPoint.z
                    
            Set objUpperVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                            2 * dOffset + dDiaLowerVerticalCylinder, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objUpperVerticalCylinder
            Set objUpperVerticalCylinder = Nothing
        
        Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.
        
    End Select
    
    Set oStPoint = Nothing
    Set oEnPoint = Nothing

' Insert your code for Nozzle 1
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                        sptOffset, depth
   
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector
    Set oDir = New AutoMath.DVector
    Set oPlacePoint = New AutoMath.DPosition
    
    oPlacePoint.Set -parFacetoCenter - sptOffset + depth, 0#, 0#
    oDir.Set -1, 0#, 0#
    
    Set ObjNozzle1 = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle1
    Set ObjNozzle1 = Nothing
    
 ' Insert your code for Nozzle 2
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                        sptOffset, depth
    oPlacePoint.Set parFace1toCenter + sptOffset - depth, 0#, 0#
    oDir.Set 1, 0#, 0#
    
    Set ObjNozzle2 = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjNozzle2
    Set ObjNozzle2 = Nothing
    Set oDir = Nothing
    Set oPlacePoint = Nothing

  ' Insert your code for Valve Actuator
'-----------------------------------------------------------------------------------------------------
    Dim oSymbolHelper As IJSymbolGeometryHelper
    Set oSymbolHelper = New SP3DSymbolHelper.SymbolServices
    oSymbolHelper.OutputCollection = m_OutputColl

    On Error Resume Next
    Dim oDirX As IJDVector
    Dim oDirY As IJDVector
    Dim oDirZ As IJDVector
    Set oDirX = New DVector
    Set oDirY = New DVector
    Set oDirZ = New DVector

    oDirX.Set Cos(parHandwheelAngle), 0, Sin(parHandwheelAngle)
    oDirY.Set 0, 1, 0
    oDirZ.Set -Sin(parHandwheelAngle), 0, Cos(parHandwheelAngle)
    
    Dim oOperatorPart As IJDPart
    Dim oOperatorOcc   As IJPartOcc
    
    If Not oPipeComponent Is Nothing Then
'       Note: oPipeComponent.GetValveOperatorPart method call may fail in following two cases:
'        1. If valve operator catalog part number is not defined (to check for E_NOOPERATORDEFINED) or
'        2. If valve operator part is not found (to check for E_NOOPERATORPARTFOUND)
'        symbol should raise error for case 2.
        On Error GoTo HandleOperatorCodeLabel

        Set oOperatorPart = oPipeComponent.GetValveOperatorPart

        If Not oOperatorPart Is Nothing Then
            Dim OpOrigin As IJDPosition
            Set OpOrigin = New DPosition
            OpOrigin.Set 0, 0, 0
            Set oOperatorOcc = oSymbolHelper.CreateChildPartOcc("ValveOperator", oOperatorPart, OpOrigin, oDirX, oDirY, oDirZ)

        End If
    End If
HandleOperatorCodeLabel:
'   The symbol code execution should not raise error for case 1 i.e. If Err.Number = E_NOOPERATORDEFINED
    If Err.Number = E_NOOPERATORPARTFOUND Then          'Case 2
        Err.Raise Err.Number, MODULE & METHOD, "No operator part found."
    End If
  
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
End Sub
