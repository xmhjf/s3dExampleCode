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
'   Copyright (c) 2004-07, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Sep 23 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    Symbol Model No. is: I7AZ Page No. 355 of PDS Piping Component Data Reference Guide pcd.pdf.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols.
'  18.Apr.2007     dkl  CR-118302    Implemented part data basis to create geometry for Solenoid valves
'                                    Type D and Type P as per JIS B8471.
'                                    Used appropriate CmpdblXXX functions for comparing doubles.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Insulation:" 'Used for error messages
Private PI       As Double


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
            'Insulation cylinder common to all Part Data Bases defined at the end.
                
        Case 255 'Solenoid Valve type 1 (Type D, Ref JIS B8471).
            Dim parFacetoFace As Double
            parFacetoFace = arrayOfInputs(5)
            parFacetoCenter = parFacetoFace / 2
            parFace1toCenter = parFacetoFace / 2
            
            'The insualtion for Smaller diameter horizontal cylinderical body in between the sockets / flanges
            ' is Insulation cylinder common to all Part Data Bases at the end.
            
            'Constructing insulation for Larger diameter horizontal cylinderical body
            'Assumptions:
            ' 1. The diameter of the larger cylinder is equal to 1.2 x flange dia.
            ' 2. The length of Smaller diameter horizontal cylinder is equal to Hub/Flange Thickness.
            
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                        sptOffset, depth
            
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
            
            'The ends of the cylinder need to be insulated. However if insualtion thickness exceeds
            'the length of the small cylinder, restrict the insulation along length to length of small cylinder.
            
            Dim dInsThkAlongLength As Double
            If CmpDblLessThan(dProjectionofSmallerCyl, parInsulationThickness) Then
                dInsThkAlongLength = dProjectionofSmallerCyl
            Else
                dInsThkAlongLength = parInsulationThickness
            End If
            
            oStPoint.Set -dLengthofLargerCyl / 2 - dInsThkAlongLength, 0, 0
            oEnPoint.Set -oStPoint.x, oStPoint.y, oStPoint.z
                    
            Dim objHorLargerCylinder As Object
            Set objHorLargerCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                    2 * (dRadiusofLargerCyl + parInsulationThickness), True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objHorLargerCylinder
            Set objHorLargerCylinder = Nothing

            'Constructing insulation for Lower vertical cylinderical body
            'Assumptions:
            '1. The diameter of the Lower vertical cylinderical body is equal to the
            '   diameter of the larger cylinder.
            '2. The height of the Lower vertical cylinderical body is equal to half of the radius
            '   of Larger diameter horizontal cylinderical body.
                    
            oStPoint.Set 0, 0, 0
            oEnPoint.Set oStPoint.x, dRadiusofLargerCyl * 3 / 2, oStPoint.z
                    
            Dim objLowerVerticalCylinder As Object
            Set objLowerVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                2 * dRadiusofLargerCyl + 2 * parInsulationThickness, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLowerVerticalCylinder
            Set objLowerVerticalCylinder = Nothing
   
            'Constructing insulation for Upper vertical cylinderical body
            'Assumptions:
            '1. The diameter of the Upper vertical cylinderical body is equal to the
            '   1.2 times the diameter of the lower cylinderical body.
            '2. The height of the Upper vertical cylinderical body is equal to half of the height
            '   of the lower cylinderical body.
                    
            oStPoint.Set 0, oEnPoint.y, 0
            oEnPoint.Set oStPoint.x, oStPoint.y + dRadiusofLargerCyl / 4, oStPoint.z
                    
            Dim objUpperVerticalCylinder As Object
            Set objUpperVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                2.4 * dRadiusofLargerCyl + 2 * parInsulationThickness, True)
        
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objUpperVerticalCylinder
            Set objUpperVerticalCylinder = Nothing
            
        Case 260 'Solenoid Valve type 2(Type P, Ref JIS B8471).
            parFacetoFace = arrayOfInputs(5)
            parHandwheelAngle = arrayOfInputs(6)
            parFacetoCenter = parFacetoFace / 2
            parFace1toCenter = parFacetoFace / 2
            
            'The insualtion for Smaller diameter horizontal cylinderical body in between the sockets / flanges
            ' is Insulation cylinder common to all Part Data Bases at the end.
   
            ' Constructing insulation for elliptical body of the valve
            ' Assumptions:
            ' 1. Height of elliptical body of the valve (from the center) is 1.2 x Hub outside radius.
            ' 2. The length of Smaller diameter horizontal cylinder is equal to Hub/Flange Thickness.
            
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                        sptOffset, depth
            
            If CmpDblEqual(flangeDiam, 0) Then flangeDiam = 2 * pipeDiam
            
            Dim dHeightofEllipticalArc As Double
            dHeightofEllipticalArc = 1.2 * flangeDiam

            If CmpDblEqual(flangeThick, 0) Then
                dProjectionofSmallerCyl = parFacetoFace / 50
            Else
                dProjectionofSmallerCyl = flangeThick
            End If
            
            Dim dHalfBodyWidth As Double
            dHalfBodyWidth = parFacetoFace / 2 - dProjectionofSmallerCyl - flangeThick
            
            Dim oGeomFactory As IngrGeom3D.GeometryFactory
            Set oGeomFactory = New IngrGeom3D.GeometryFactory

            'The ends of the cylinder need to be insulated. However if insualtion thickness exceeds
            'the length of the small cylinder, restrict the insulation along length to length of small cylinder.
            
            If CmpDblLessThan(dProjectionofSmallerCyl, parInsulationThickness) Then
                dInsThkAlongLength = dProjectionofSmallerCyl
            Else
                dInsThkAlongLength = parInsulationThickness
            End If
  
            Dim objEllipticalArc As IngrGeom3D.EllipticalArc3d
            Set objEllipticalArc = _
                    oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                    0, pipeDiam / 2, 0, _
                    0, 0, 1, _
                    dHalfBodyWidth + dInsThkAlongLength, 0, 0, _
                    ((dHeightofEllipticalArc - pipeDiam / 2 + dInsThkAlongLength) / (dHalfBodyWidth + dInsThkAlongLength)), _
                     0, PI)

            Dim oAxisVec As AutoMath.DVector
            Set oAxisVec = New AutoMath.DVector
            oAxisVec.Set 1, 0, 0
            Dim oCenPoint As AutoMath.DPosition
            Set oCenPoint = New AutoMath.DPosition
            oCenPoint.Set 0, 0.0001, 0
            
            Dim objValveBody As Object
            Set objValveBody = PlaceRevolution(m_OutputColl, objEllipticalArc, oAxisVec, _
                                            oCenPoint, 2 * PI, True)
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objValveBody
            Set objValveBody = Nothing
            Set oAxisVec = Nothing
            Set oCenPoint = Nothing
            Set objEllipticalArc = Nothing
            Set oGeomFactory = Nothing
      
            'Constructing insulation for Lower vertical cylinderical body
            'Assumptions:
            '1. The diameter of the Lower vertical cylinderical body is equal to the
            '   Height of Elliptical Arc (of the ellptical body created above) / 1.4
            '2. The height of the Lower vertical cylinderical body is equal to
            '   Height of Elliptical Arc (of the ellptical body created above) / 3
            
            Dim dDiaLowerVerticalCylinder As Double
            dDiaLowerVerticalCylinder = dHalfBodyWidth / 1.4
            
            oStPoint.Set 0, 0, 0
            oEnPoint.Set oStPoint.x, oStPoint.y + dHeightofEllipticalArc * 4 / 3, oStPoint.z
                    
            Set objLowerVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                            dDiaLowerVerticalCylinder + 2 * parInsulationThickness, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objLowerVerticalCylinder
            Set objLowerVerticalCylinder = Nothing
         
            'Constructing insulation for Upper vertical cylinderical body
            'Assumptions:
            '1. The diameter of the Upper vertical cylinderical body is equal to the
            '   1.2 times the diameter of the lower cylinderical body.
            '2. The height of the Upper vertical cylinderical body is equal to half of the height
            '   of the lower cylinderical body.
            '3. This cylinder is offset to the Lower vertical cylinder by 20% of the diameter of
            '   the Lower vertical cylinder.
                        
            Dim dOffset As Double
            dOffset = dDiaLowerVerticalCylinder * 0.2
                        
            oStPoint.Set -dDiaLowerVerticalCylinder * 0.2, oEnPoint.y, 0
            oEnPoint.Set -dDiaLowerVerticalCylinder * 0.2, _
                                        oStPoint.y + dHeightofEllipticalArc / 8, oStPoint.z
                    
            Set objUpperVerticalCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                    2 * dOffset + dDiaLowerVerticalCylinder + 2 * parInsulationThickness, True)
            
            'Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objUpperVerticalCylinder
            Set objUpperVerticalCylinder = Nothing
            
        Case Else
            GoTo ErrorLabel:      ' Invalid Part data basis.
        
    End Select
    
' Insert your code for output Insulation Cylinder from port 1 to port 2
    oStPoint.Set -parFacetoCenter, 0, 0
    oEnPoint.Set parFace1toCenter, 0, 0
    
    Dim InsulationDiameter As Double
    Dim InsulationDiameter1 As Double
    Dim InsulationDiameter2 As Double

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                                    flangeDiam, sptOffset, depth
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        InsulationDiameter1 = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDiameter1 = pipeDiam + parInsulationThickness * 2
    End If
    
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, _
                                                    flangeDiam, sptOffset, depth
    If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
        InsulationDiameter2 = flangeDiam + parInsulationThickness * 2
    Else
        InsulationDiameter2 = pipeDiam + parInsulationThickness * 2
    End If
    
    If CmpDblGreaterthan(InsulationDiameter1, InsulationDiameter2) Then
        InsulationDiameter = InsulationDiameter1
    Else
        InsulationDiameter = InsulationDiameter2
    End If
    
    Dim ObjInsulationCylinder As Object
    Set ObjInsulationCylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                        InsulationDiameter, True)
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulationCylinder
    Set ObjInsulationCylinder = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    
End Sub
