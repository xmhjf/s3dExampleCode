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
'   ProgID:         SP3DHTaperBodyLateral.HTapLateral
'   Author:         VRK
'   Creation Date:  Thursday,June 28 2007
'   Description:
'   This symbol is prepared for Tapered Body Lateral of McGill Air flow corporation as per CR-120452
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     ---     ------------------
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "CSimplePhysical" 'Used for error messages
Private m_oGeomHelper As IJSymbolGeometryHelper
Private m_GeomFactory As IngrGeom3D.GeometryFactory

Private Const E_FAIL = &H80004005
Const NEGLIGIBLE_THICKNESS = 0.0001
Private PI  As Double

Private Sub Class_Initialize()
    Const METHOD = "Class_Initialize"
    On Error GoTo ErrorHandler
    
    PI = 4 * Atn(1)
    Set m_oGeomHelper = New SymbolServices
    Set m_GeomFactory = New IngrGeom3D.GeometryFactory
    
    Exit Sub
ErrorHandler:
    ReportUnanticipatedError2 MODULE, METHOD
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt As PartFacelets.IJDPart
    Dim iOutput  As Double
    
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    Dim parWidth As Double
    Dim parDepth As Double
    Dim parBWidth As Double
    Dim parBDepth As Double
    Dim parHVACShape As Long
    Dim parWidth2 As Double
    Dim parDepth2 As Double
    Dim parAngle As Double
    Dim parInsulationThickness As Double
    
    Dim Inch As Double
    Inch = 0.0254
    
    'Inputs
    Set oPartFclt = arrayOfInputs(1)
    parWidth = arrayOfInputs(2)
    parDepth = arrayOfInputs(3)
    parBWidth = arrayOfInputs(4)
    parBDepth = arrayOfInputs(5)
    parHVACShape = arrayOfInputs(6)
    parWidth2 = arrayOfInputs(7)
    parDepth2 = arrayOfInputs(8)
    parAngle = arrayOfInputs(9)
    parInsulationThickness = arrayOfInputs(10)
            
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
    
    If CmpDblGreaterthan(parBDepth, parDepth2) Then
        parBDepth = parDepth2
    End If

    Dim dHlength As Double 'Header Length
    Dim dCentFrmPort2 As Double 'Distance from Center to Port 2
    Dim dBLength As Double 'Branch Length form center along the axis
    Dim dTemp As Double 'Variable used for temporary calculations
    
    'The following calculations are made as per McGill catalog
    dTemp = (parWidth - parWidth2) / ((2 * Tan(parAngle)))
    dHlength = 2 * Inch + (parBWidth / (Sin(parAngle))) - dTemp
    dCentFrmPort2 = 1 * Inch + (parBWidth / ((2 * Sin(parAngle)))) + (parWidth2 / (2 * Tan(parAngle))) + (Inch * dTemp / dHlength)
    dBLength = 2 * Inch + (Inch * dTemp / dHlength) + (parWidth2 / (2 * Sin(parAngle))) + (parBWidth / (2 * Tan(parAngle)))


    Dim stPoint  As New AutoMath.DPosition
    Dim enPoint  As New AutoMath.DPosition
    Dim objInsBranch As Object
    Dim objInsTaperedBody As Object
    Dim objInsHeaderTakeoff1 As Object
    Dim objInsHeaderTakeoff2 As Object
    
    Dim oHeaderCurve1 As Object
    Dim oHeaderCurve2 As Object
    Dim oBranchCurve1 As Object
    Dim oBranchCurve2 As Object
        
    Dim CP As AutoMath.DPosition
    Set CP = New AutoMath.DPosition
    
    Dim oAxisVec As AutoMath.DVector
    Set oAxisVec = New AutoMath.DVector
    
    If parHVACShape = 4 Then 'Round
    
        'Creating Insulation for Tapered Body
        stPoint.Set (dCentFrmPort2 - dHlength), 0, 0
        enPoint.Set dCentFrmPort2, 0, 0
        Set objInsTaperedBody = PlaceCone(m_OutputColl, stPoint, enPoint, parWidth / 2 + (parInsulationThickness), parWidth2 / 2 + (parInsulationThickness), True)
        
        'Creating Insulation for Branch
        stPoint.Set 0, 0, 0
        enPoint.Set (dBLength + 2 * Inch) * Cos(parAngle), 0, (dBLength + 2 * Inch) * Sin(parAngle)
        Set objInsBranch = PlaceCylinder(m_OutputColl, stPoint, enPoint, parBWidth + (2 * parInsulationThickness), True)
        
        'Creating Insulation for Header Take-off 1
        stPoint.Set (dCentFrmPort2 - dHlength), 0, 0
        enPoint.Set (dCentFrmPort2 - dHlength - 2 * Inch), 0, 0
        Set objInsHeaderTakeoff1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWidth + (2 * parInsulationThickness), True)
        
        'Creating Insulation for Header Take-off 2
        stPoint.Set dCentFrmPort2, 0, 0
        enPoint.Set dCentFrmPort2 + (2 * Inch), 0, 0
        Set objInsHeaderTakeoff2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parWidth2 + (2 * parInsulationThickness), True)
        
    
    ElseIf parHVACShape = 1 Then
        
        'Creating Insulation for Tapered Body
        CP.Set dCentFrmPort2 - dHlength, 0, 0
        Set oHeaderCurve1 = CreRectangle(CP, parWidth + (2 * parInsulationThickness), parDepth + (2 * parInsulationThickness), PI / 2)
        
        CP.Set dCentFrmPort2, 0, 0
        Set oHeaderCurve2 = CreRectangle(CP, parWidth2 + (2 * parInsulationThickness), parDepth2 + (2 * parInsulationThickness), PI / 2)
        
        Set objInsTaperedBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            oHeaderCurve1, oHeaderCurve2, True)
        
        'Creating Insulation for Branch
        CP.Set 0, 0, 0
        Set oBranchCurve1 = CreRectBranchNormaltoZ(CP, parBWidth + (2 * parInsulationThickness), parBDepth + (2 * parInsulationThickness), PI / 2 - parAngle)
        
        CP.Set 0, 0, dBLength + (2 * Inch)
        Set oBranchCurve2 = CreRectBranchNormaltoZ(CP, parBWidth + (2 * parInsulationThickness), parBDepth + (2 * parInsulationThickness), PI / 2 - parAngle)
        
        Set objInsBranch = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            oBranchCurve1, oBranchCurve2, True)
        
        'Creating Insulation for Header Take-off 1
        stPoint.Set (dCentFrmPort2 - dHlength), (parDepth / 2) + parInsulationThickness, _
                                                (parWidth / 2) + parInsulationThickness
        enPoint.Set (dCentFrmPort2 - dHlength - 2 * Inch), (-parDepth / 2) - parInsulationThickness, _
                                                            (-parWidth / 2) - parInsulationThickness
        Set objInsHeaderTakeoff1 = PlaceBox(m_OutputColl, stPoint, enPoint)
        
        'Creating Insulation for Header Take-off 2
        stPoint.Set dCentFrmPort2, (parDepth2 / 2) + parInsulationThickness, _
                                   (parWidth2 / 2) + parInsulationThickness
        enPoint.Set dCentFrmPort2 + (2 * Inch), -(parDepth2 / 2) - parInsulationThickness, _
                                                -(parWidth2 / 2) - parInsulationThickness
        Set objInsHeaderTakeoff2 = PlaceBox(m_OutputColl, stPoint, enPoint)
        

    ElseIf parHVACShape = FlatOval Then
        
        'Creating Insulation for Tapered Body
        CP.Set dCentFrmPort2 - dHlength, 0, 0
        Set oHeaderCurve1 = CreFlatOval(CP, parWidth + (2 * parInsulationThickness), parDepth + (2 * parInsulationThickness), PI / 2)
        
        CP.Set dCentFrmPort2, 0, 0
        Set oHeaderCurve2 = CreFlatOval(CP, parWidth2 + (2 * parInsulationThickness), parDepth2 + (2 * parInsulationThickness), PI / 2)

        Set objInsTaperedBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            oHeaderCurve1, oHeaderCurve2, True)
        
        'Creating Insulation for Branch
        CP.Set 0, 0, 0
        Set oBranchCurve1 = CreFltOvlBranchNormaltoZ(CP, parBWidth + (2 * parInsulationThickness), parBDepth + (2 * parInsulationThickness), PI / 2 - parAngle)
        
        CP.Set 0, 0, dBLength + (2 * Inch)
        Set oBranchCurve2 = CreFltOvlBranchNormaltoZ(CP, parBWidth + (2 * parInsulationThickness), parBDepth + (2 * parInsulationThickness), PI / 2 - parAngle)
        
        Set objInsBranch = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
            oBranchCurve1, oBranchCurve2, True)
        
        Set oBranchCurve1 = Nothing
        Set oBranchCurve2 = Nothing
            
        'Creating Insulation for Header Take-off 1
        oAxisVec.Set -1, 0, 0
        Set objInsHeaderTakeoff1 = PlaceProjection(m_OutputColl, oHeaderCurve1, oAxisVec, 2 * Inch, True)
        Set oHeaderCurve1 = Nothing
        
        'Creating Insulation for Header Take-off 2
        oAxisVec.Set 1, 0, 0
        Set objInsHeaderTakeoff2 = PlaceProjection(m_OutputColl, oHeaderCurve2, oAxisVec, 2 * Inch, True)
        Set oHeaderCurve2 = Nothing
        
    End If
    
    'Set output 12(Insulation for Tapered Body)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsTaperedBody
    Set objInsTaperedBody = Nothing
    
    'Set the Output 13(Insulation for Branch)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsBranch
    Set objInsBranch = Nothing
    
    'Set the Output 14(Insulation for Header Take off 1)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsHeaderTakeoff1
    Set objInsHeaderTakeoff1 = Nothing
    
    'Set the Output 15(Insulation for Header Take off 2)
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsHeaderTakeoff2
    Set objInsHeaderTakeoff2 = Nothing
    
    Exit Sub
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub





