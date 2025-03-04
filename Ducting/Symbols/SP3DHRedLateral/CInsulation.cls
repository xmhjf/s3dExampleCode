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
'   This symbol is created for Reducing Straight And Conical laterals of McGill Air Flow Corporation
'   The symbol supports the following two part data basis values
'   1.)Reducing Lateral,Straight
'   2.)Reducing Lateral,Conical
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
    Dim parWidth2 As Double
    Dim parDepth2 As Double
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
    parWidth2 = arrayOfInputs(4)
    parDepth2 = arrayOfInputs(5)
    parBWidth = arrayOfInputs(6)
    parBDepth = arrayOfInputs(7)
    parAngle = arrayOfInputs(8)
    parHVACShape = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)

    Dim dRedHeaderLength As Double
    dRedHeaderLength = parWidth - parWidth2
    
    'Check to see that length of the reducing portion is not less than 1 inch (as per McGill Catalog)
    If CmpDblLessThan(dRedHeaderLength, 1 * Inch) Then
        dRedHeaderLength = 1 * Inch
    End If
    
    'Check to see that length of the reducing portion is not greater than 12 inch (as per McGill Catalog)
    If CmpDblGreaterthan(dRedHeaderLength, 12 * Inch) Then
        dRedHeaderLength = 12 * Inch
    End If
    
    iOutput = 0
       
    
        
    Dim dHlength As Double 'Header Length
    Dim dCentFrmStHeaderEnd As Double 'Distance from Center to End position of straight header
    Dim dBLength As Double 'Branch Length form center along the axis
    Dim dBranchInt2FromPort3 As Double 'Distance of intersection point 2 of branch(with header) and branch port along branch axis
    
    Dim dBottomBrWidth As Double 'Width of Branch at the Center(bottom)
    Dim dBottomBrDepth As Double 'Depth of Branch at the Center(bottom)

    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim objHeaderTakeoff1 As Object
    Dim objHeaderTakeoff2 As Object
    Dim objBranchTakeoff As Object
    Dim objHeaderStrip1 As Object
    Dim objHeaderStrip2 As Object
    Dim objBranchStrip As Object
        
    Dim oBranchStripCurve As Object
    Dim oBranchTakeoffCurve As Object
    
    Dim oBranchCurve1 As Object
    Dim oBranchCurve2 As Object
    
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Set oStPoint = New AutoMath.DPosition
    Set oEnPoint = New AutoMath.DPosition
    
    Dim objInsStHeader As Object
    Dim objInsBranch As Object
    Dim objInsRedHeader As Object
    
    Dim Dir As AutoMath.DVector
    Set Dir = New AutoMath.DVector

'Check for part data basis
    Dim lPartdatabasis As Double
    Dim oHvacPart As IJDHvacPart
    Set oHvacPart = oPartFclt
    lPartdatabasis = oHvacPart.PartDataBasis
    
'Checking for maximum values of Branch Width based on part data basis
'--------------------------------------------
'    parWidth               Maximum BWidth
'--------------------------------------------
'    <=10inch               parWidth-3inch
'    >10inch and <=42inch   parWidth-4inch
'    >42inch                parWidth-5inch
    
    If (lPartdatabasis <= 1 Or lPartdatabasis = 55) Then
        
        If CmpDblGreaterthan(parBWidth, parWidth) Then
            parBWidth = parWidth
        End If
        
        If (Not parHVACShape = 4) And CmpDblGreaterthan(parBDepth, parDepth) Then
            parBDepth = parDepth
        End If
        
    ElseIf (lPartdatabasis = 60) Then
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
    
    'In case of round shape (where Depth2 is optional) making Depth2 equal to Width2
    If CmpDblEqual(parDepth2, 0) Then
        parDepth2 = parWidth2
    End If
    
    'In case of round shape (where BDepth is optional) making BDepth equal to BWidth
    If CmpDblEqual(parBDepth, 0) Then
        parBDepth = parBWidth
    End If
    
    Dim dAlpha As Double 'Half of the Difference between the branch width and width at intersection point of branch (with header)
    Dim dTemp As Double
    
    
'Calculations of header length, branch length and distance of port 2 from center based on McGill catalog
    If (lPartdatabasis <= 1 Or lPartdatabasis = 55) Then 'Straight Lateral
        dHlength = (parBWidth / Sin(parAngle)) + (2 * Inch)
        dAlpha = 0
        dBranchInt2FromPort3 = 2 * Inch
    ElseIf (lPartdatabasis = 60) Then 'Conical Lateral
        dHlength = ((parBWidth + (2 * Inch)) / Sin(parAngle)) + (2 * Inch)
        dTemp = ((parBWidth / Inch) + 2) / (4 * Tan(parAngle))
        dAlpha = 2 * Inch / (dTemp + 2)
        dBranchInt2FromPort3 = 4 * Inch
    End If
           
    dBLength = (parWidth / (2 * Sin(parAngle))) + _
              ((parBWidth + (2 * dAlpha)) / (2 * Tan(parAngle))) + dBranchInt2FromPort3
    dCentFrmStHeaderEnd = (parWidth / (2 * Tan(parAngle))) + _
              ((parBWidth + (2 * dAlpha)) / (2 * Sin(parAngle))) + Inch
    
'Assigning branch width and branch depth at the center(bottom) based on part data basis
    If (lPartdatabasis <= 1 Or lPartdatabasis = 55) Then 'Straight Lateral
         dBottomBrWidth = parBWidth
         dBottomBrDepth = parBDepth
    ElseIf (lPartdatabasis = 60) Then 'Conical Lateral
        dBottomBrWidth = parBWidth + ((dBLength * dAlpha) / (2 * Inch))
        dBottomBrDepth = parBDepth + (2 * Inch)
    End If
           
    
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    
'Creating geometry based on shape of the lateral
        
    If parHVACShape = 4 Then 'Round=4
        
        'Creating Header
        oStPoint.Set -dHlength + dCentFrmStHeaderEnd - 2 * Inch, 0, 0
        oEnPoint.Set dCentFrmStHeaderEnd, 0, 0
        Set objInsStHeader = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parWidth + (2 * parInsulationThickness), False)
        
        'Creating Reducing Header
        oStPoint.Set dCentFrmStHeaderEnd, 0, 0
        oEnPoint.Set dCentFrmStHeaderEnd + dRedHeaderLength + 2 * Inch, 0, 0
        Set objInsRedHeader = PlaceCone(m_OutputColl, oStPoint, oEnPoint, _
                                        (parWidth / 2) + parInsulationThickness, _
                                        (parWidth2 / 2) + parInsulationThickness, False)

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
        oStPoint.Set -dHlength + dCentFrmStHeaderEnd - 2 * Inch, (parDepth / 2) + parInsulationThickness, _
                                                    (parWidth / 2) + parInsulationThickness
        oEnPoint.Set dCentFrmStHeaderEnd, (-parDepth / 2) - parInsulationThickness, _
                                          (-parWidth / 2) - parInsulationThickness
        Set objInsStHeader = PlaceBox(m_OutputColl, oStPoint, oEnPoint)
        
        'Creating Reducing Header
        oStPoint.Set dCentFrmStHeaderEnd, 0, 0
        Set objInsRedHeader = PlaceTrapezoid(m_OutputColl, oStPoint, parWidth + (2 * parInsulationThickness), parDepth + (2 * parInsulationThickness), _
                                                parWidth2 + (2 * parInsulationThickness), parDepth2 + (2 * parInsulationThickness), _
                                                dRedHeaderLength + 2 * Inch, False, 0, PI / 2, 0)
            
        'Creating branch curve at the center
        CP.Set dCentFrmStHeaderEnd - (dHlength / 2), 0, parWidth / 2
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

    ' Set the Output 13(Insulation for Header)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsStHeader
    Set objInsStHeader = Nothing
    
    
    ' Set the Output 14(Insulation Reducing Header)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsRedHeader
    Set objInsRedHeader = Nothing
    
    ' Set the output 15(Insulation Branch)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBranch
    Set objInsBranch = Nothing
    
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError2 MODULE, METHOD
End Sub

