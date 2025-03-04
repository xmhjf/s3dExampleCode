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
'   Copyright (c) 2007, Intergraph Corporation. All rights reserved.

'   CInsulation.cls
'   Author:         RRK
'   Creation Date:  Monday, May 14 2007
'   Description:
'   This symbol is created for Straight And Conical laterals of McGill Air Flow Corporation
'   The symbol supports the following two part data basis values
'   1.)Straight Lateral
'   2.)Conical Lateral
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
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
    Dim CP As New AutoMath.DPosition
    
    Dim iOutput     As Double
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parAngle As Double
    Dim parHVACShape As Double
    Dim parInsulationThickness As Double
    
    Dim Inch As Double
    Inch = 0.0254
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parBDepth = arrayOfInputs(5)
    parAngle = arrayOfInputs(6)
    parHVACShape = arrayOfInputs(7)
    parInsulationThickness = arrayOfInputs(8)
        
    iOutput = 0
        
    Dim dHlength As Double 'Header Length
    Dim dCentFrmPort2 As Double 'Distance from Center to Port 2
    Dim dBLength As Double 'Branch Length form center along the axis
    Dim dBranchInt2FromPort3 As Double 'Distance of intersection point 2 of branch(with header) and branch port along branch axis
    
    Dim dBottomBrWidth As Double 'Width of Branch at the Center(bottom)
    Dim dBottomBrDepth As Double 'Depth of Branch at the Center(bottom)

    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim oBranchCurve1 As Object
    Dim oBranchCurve2 As Object
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim objInsHeader As Object
    Dim objInsBranch As Object
    
    Dim Dir As AutoMath.DVector
    Set Dir = New AutoMath.DVector
    
    'Check for part data basis
    Dim lPartdatabasis As Double
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartdatabasis = oHvacPart.PartDataBasis

'Checking for maximum values of Branch Width and Branch Depth based on part data basis
'--------------------------------------------
'    parWidth               Maximum BWidth
'--------------------------------------------
'    <=10inch               parWidth-3inch
'    >10inch and <=42inch   parWidth-4inch
'    >42inch                parWidth-5inch
    
    If (lPartdatabasis <= 1 Or lPartdatabasis = 45) Then
        If CmpDblGreaterthan(parBWidth, parWidth) Then
            parBWidth = parWidth
        End If
        
        If (Not parHVACShape = 4) And CmpDblGreaterthan(parBDepth, parDepth) Then
            parBDepth = parDepth
        End If
        
    ElseIf (lPartdatabasis = 50) Then
    
        If CmpDblLessThanOrEqualTo(parWidth, 10 * Inch) And CmpDblGreaterthan(parBWidth, parWidth - 3 * Inch) Then
            parBWidth = parWidth - (3 * Inch)
        ElseIf CmpDblGreaterthan(parWidth, 10 * Inch) And CmpDblLessThanOrEqualTo(parWidth, 42 * Inch) And CmpDblGreaterthan(parBWidth, parWidth - 4 * Inch) Then
            parBWidth = parWidth - (4 * Inch)
        ElseIf CmpDblGreaterthan(parWidth, 42 * Inch) And CmpDblGreaterthan(parBWidth, parWidth - 5 * Inch) Then
            parBWidth = parWidth - (5 * Inch)
        End If
        
        If (Not parHVACShape = 4) And CmpDblGreaterthan(parBDepth, parDepth - 2 * Inch) Then
            parBDepth = parDepth - 2 * Inch
        End If
        
    End If
        
    'In case of round shape (where Depth is optional) making Depth equal to Width
    If CmpDblEqual(parDepth, 0) Then
        parDepth = parWidth
    End If
    
    'In case of round shape (where BDepth is optional) making BDepth equal to BWidth
    If CmpDblEqual(parBDepth, 0) Then
        parBDepth = parBWidth
    End If
    
    Dim dAlpha As Double 'Half of the Difference between the branch width and width at intersection point of branch (with header)
    Dim dTemp As Double
        
'Calculations of header length and other parameters as given in McGill catalog
    If (lPartdatabasis <= 1 Or lPartdatabasis = 45) Then 'Straight Lateral
        dHlength = (parBWidth / Sin(parAngle)) + (2 * Inch)
        dAlpha = 0
        dBranchInt2FromPort3 = 2 * Inch
    ElseIf (lPartdatabasis = 50) Then 'Conical Lateral
        dHlength = ((parBWidth + (2 * Inch)) / Sin(parAngle)) + (2 * Inch)
        dTemp = ((parBWidth / Inch) + 2) / (4 * Tan(parAngle))
        dAlpha = 2 * Inch / (dTemp + 2)
        dBranchInt2FromPort3 = 4 * Inch
    End If

'Calculations of Branch Length as given in McGill Catalog
    dBLength = (parWidth / (2 * Sin(parAngle))) + _
              ((parBWidth + (2 * dAlpha)) / (2 * Tan(parAngle))) + dBranchInt2FromPort3

'Calculation of distance of port 2 from center as given in McGill catalog
    dCentFrmPort2 = (parWidth / (2 * Tan(parAngle))) + _
              ((parBWidth + (2 * dAlpha)) / (2 * Sin(parAngle))) + Inch
    
'Assigning branch width and branch depth at the center(bottom) based on part data basis
    If (lPartdatabasis <= 1 Or lPartdatabasis = 45) Then 'Straight Lateral
         dBottomBrWidth = parBWidth
         dBottomBrDepth = parBDepth
    ElseIf (lPartdatabasis = 50) Then 'Conical Lateral
        dBottomBrWidth = parBWidth + ((dBLength * dAlpha) / (2 * Inch))
        dBottomBrDepth = parBDepth + (2 * Inch)
    End If
           
    
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    
'Creating geometry based on shape of the lateral
    If parHVACShape = FlatOval Then
        Dim oHeaderCurve As Object
        
        'Creating flat oval header curve
        CP.Set -dHlength + dCentFrmPort2 - (2 * Inch), 0, 0
        Set oHeaderCurve = CreFlatOval(CP, parWidth + (2 * parInsulationThickness), _
                                           parDepth + (2 * parInsulationThickness), PI / 2)
        
        'Creating header
        oAxisVec.Set 1, 0, 0
        Set objInsHeader = PlaceProjection(m_OutputColl, oHeaderCurve, oAxisVec, dHlength + (4 * Inch), True)
        Set oHeaderCurve = Nothing
        
        'Creating branch curve at the center
        CP.Set 0, 0, 0
        Set oBranchCurve1 = CreFltOvlBranchNormaltoZ(CP, dBottomBrWidth + (2 * parInsulationThickness), _
                                                        dBottomBrDepth + (2 * parInsulationThickness), PI / 2 - (parAngle))
        
        'Creating branch curve at branch port
        CP.Set 0, 0, dBLength + (2 * Inch)
        Set oBranchCurve2 = CreFltOvlBranchNormaltoZ(CP, parBWidth + (2 * parInsulationThickness), _
                                                         parBDepth + (2 * parInsulationThickness), PI / 2 - (parAngle))
        
        'Creating Branch by joining two branch curves
        Set objInsBranch = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                             oBranchCurve1, oBranchCurve2, False)
        Set oBranchCurve1 = Nothing
        Set oBranchCurve2 = Nothing
    
    ElseIf parHVACShape = 4 Then 'Round=4
        
        'Creating Header
        oStPoint.Set -dHlength + dCentFrmPort2 - (2 * Inch), 0, 0
        oEnPoint.Set dCentFrmPort2 + (2 * Inch), 0, 0
        Set objInsHeader = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth + (2 * parInsulationThickness), True)
        
        'Check to see that branch width at center doesn't exceed width of the round duct
        If CmpDblGreaterthan(dBottomBrWidth, parWidth) Then
            dBottomBrWidth = parWidth
        End If
        
        'Creating Branch curve at the center
        CP.Set 0, 0, 0
        Dir.Set Cos(parAngle), 0, Sin(parAngle)
        Set oBranchCurve1 = PlaceTrCircleByCenter(CP, Dir, (dBottomBrWidth / 2) + parInsulationThickness)
            
        'Creating Branch curve at the branch port
        CP.Set (dBLength + 2 * Inch) * Cos(parAngle), 0, (dBLength + 2 * Inch) * Sin(parAngle)
        Dir.Set Cos(parAngle), 0, Sin(parAngle)
        Set oBranchCurve2 = PlaceTrCircleByCenter(CP, Dir, (parBWidth / 2) + parInsulationThickness)
        
        'Creating Branch by joining two branch curves
        Set objInsBranch = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                             oBranchCurve1, oBranchCurve2, True)
        
        Set oBranchCurve1 = Nothing
        Set oBranchCurve2 = Nothing
        
    ElseIf parHVACShape = Rectangular Then
        
        'Creating Header
        oStPoint.Set -dHlength + dCentFrmPort2 - 2 * Inch, (parDepth / 2) + parInsulationThickness, _
                                                            (parWidth / 2) + parInsulationThickness
        oEnPoint.Set dCentFrmPort2 + 2 * Inch, -(parDepth / 2 + parInsulationThickness), _
                                               -(parWidth / 2 + parInsulationThickness)
        Set objInsHeader = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
            
        'Creating branch curve at the center
        CP.Set dCentFrmPort2 - (dHlength / 2), 0, parWidth / 2
        Set oBranchCurve1 = CreRectBranchNormaltoZ(CP, dHlength - (2 * Inch) + (2 * parInsulationThickness / Sin(parAngle)), _
                                                     dBottomBrDepth + (2 * parInsulationThickness), 0)
        
        'Creating branch curve at the branch port
        CP.Set 0, 0, dBLength + (2 * Inch)
        Set oBranchCurve2 = CreRectBranchNormaltoZ(CP, parBWidth + (2 * parInsulationThickness), _
                                                        parBDepth + (2 * parInsulationThickness), PI / 2 - (parAngle))
        
        'Creating Branch by joining two branch curves
        Set objInsBranch = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                             oBranchCurve1, oBranchCurve2, True)
                             
        Set oBranchCurve1 = Nothing
        Set oBranchCurve2 = Nothing
        
    End If

    ' Set the Output 12(Insulation for Header)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsHeader
    Set objInsHeader = Nothing
    
    ' Set the output 13(Insulation for Branch)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBranch
    Set objInsBranch = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError2 MODULE, METHOD
End Sub

