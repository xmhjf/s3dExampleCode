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
'   Copyright (c) 2007-08, Intergraph Corporation. All rights reserved.
'
'   CInsulation.cls
'   ProgID:         SP3DDoubleY.DoubleY
'   Author:         VRG
'   Creation Date:  Wednesday, December 5,2007
'   Description:
'   The Double Wye symbol in which the following Part data basis cases are addressed
'               i>  Double Wye
'                   (a)Double Wye, PDB-525:Source:(ABI Foundry Catalog.pdf-Page.No:6)
'                   (b)Double Wye, PDB-526:Source:(Tyler Soil Pipe Catalog.pdf-Page.No:7)
'                   (c)Double Wye, PDB-527:Source:(ABI Foundry Catalog.pdf-Page.No:6)
'               ii>  Double Wye with 1/8 Bend
'                   (a)1/8 th Bend Double Wye, PDB-528:Source:(ABI Foundry Catalog.pdf-Page.No:6)
'                   (b)1/8 th Bend Double Wye, PDB-529:Source:(A861.26224-1.pdf-Page.No:7,fIG:11 and _
'                       Tyler Soil Pipe Catalog.pdf-Page.No:7)
'                   (c)1/8 th Bend Double Wye, PDB-530:Source:(Tyler Soil Pipe Catalog.pdf-Page.No:34)
'                   (d)1/8 th Bend Double Wye, PDB-531:Source:(A74.17725-1.pdf-Page.No:16,table:14)
'                   (e)1/8 th Bend Double Wye, PDB-532:Source:(A74.17725-1.pdf-Page.No:16,table:14)
'   Change History:
'   dd.mmm.yyyy     who          change description
'   -----------     ---          ------------------
'   18.Feb.2008     VRK          CR-127564  Created Double Wye symbol  with options 45 degree and 1/8 bend.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private m_oGeomHelper As IJSymbolGeometryHelper
Private PI As Double
Private Const MODULE = "Insulation:" 'Used for error messages

Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
    On Error GoTo Errx
    
    Set m_oGeomHelper = New SymbolServices
    PI = Atn(1) * 4
     
    Exit Sub
Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim cptoffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim flangeDiam2      As Double
    Dim cptoffset2      As Double
    Dim depth2          As Double
    Dim flangeThick2    As Double
    Dim pipeDiam3        As Double
    Dim flangeDiam3      As Double
    Dim cptoffset3      As Double
    Dim depth3          As Double
    Dim flangeThick3    As Double
    Dim flangeDiam4      As Double
    Dim cptoffset4      As Double
    Dim depth4          As Double
    Dim flangeThick4    As Double
    Dim iOutput         As Integer
    Dim Centerpos       As AutoMath.DPosition
    Set Centerpos = New DPosition
    
    Centerpos.Set 0, 0, 0

    Dim parFace1toCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFacetoFace As Double
    Dim parPortttoPortOffset As Double
    Dim parTangentLength1 As Double
    Dim parTangentLength3 As Double
    Dim parFacetoBendCenter As Double
    Dim parCentertoPortOffset As Double
    Dim parBendRadius As Double
    Dim parSeattoSeat As Double
    Dim parSeat1toCenter As Double
    Dim parSeat3toCenter As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
'    parFace1toCenter = arrayOfInputs(2)
'    parFace2toCenter = arrayOfInputs(3)
'    parFace3toCenter = arrayOfInputs(4)
'    parFacetoFace = arrayOfInputs(5)
'    parPortttoPortOffset = arrayOfInputs(6)
'    parTangentLength1 = arrayOfInputs(7)
'    parTangentLength3 = arrayOfInputs(8)
'    parFacetoBendCenter = arrayOfInputs(9)
'    parCentertoPortOffset = arrayOfInputs(10)
'    parSeattoSeat = arrayOfInputs(11)
'    parSeat1toCenter = arrayOfInputs(12)
'    parSeat3toCenter = arrayOfInputs(13)
    parInsulationThickness = arrayOfInputs(14)
    parAngle = arrayOfInputs(15)

    iOutput = 0
    m_oGeomHelper.OutputCollection = m_OutputColl
    
    'Declaration of Variables for Creation of bends
    Dim Surfset As IngrGeom3D.IJElements
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oBranchPoint As AutoMath.DPosition
    Dim oNormal As AutoMath.DVector
    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Dim oTraceStr As IngrGeom3D.ComplexString3d
    Dim oCircle As IngrGeom3D.Circle3d
    Dim iCount As Integer
    Dim oCollection As Collection
    Dim oLine As IngrGeom3D.Line3d
    Dim oArc As IngrGeom3D.Arc3d
    Dim stnorm() As Double
    Dim ednorm() As Double
    Dim oCenter As AutoMath.DPosition
    Dim ObjSurface As Object
    Dim dArcY, dLength As Double
    Dim dArcRadius As Double
    Dim dYlength As Double
    Dim dFacetoface As Double
    Dim dInsulationDiameter As Double
    Dim dInsulationDiameter3 As Double
    Dim oTransMatrix As IJDT4x4
    Dim dVarAngle As Double

    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oBranchPoint = New DPosition
    
  'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis
        
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, _
                                                                            cptoffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick2, flangeDiam2, _
                                                                            cptoffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, _
                                                                            cptoffset3, depth3
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, flangeThick4, flangeDiam4, _
                                                                            cptoffset4, depth4
   
'    dInsulationDiameter = pipeDiam + parInsulationThickness * 2
'    dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
    
    'To make trimmable
    If CmpDblEqual(parAngle, 0) Then
        If (lPartDataBasis = 525 Or lPartDataBasis = 526 Or lPartDataBasis = 527) Then
            parAngle = PI / 4
        ElseIf (lPartDataBasis = 528 Or lPartDataBasis = 529 Or lPartDataBasis = 530 Or _
             lPartDataBasis = 531 Or lPartDataBasis = 532) Then
            parAngle = PI / 2
        End If
    End If
    
    Select Case lPartDataBasis
        
        'Double Wye
        Case 525
            
            parFace3toCenter = arrayOfInputs(4)     'D
            parFacetoFace = arrayOfInputs(5)        'F
            parTangentLength1 = arrayOfInputs(7)    'G
            dVarAngle = parAngle - PI / 4
            
            'Distance from Center to Tangent Length start position is taken as 1 / 10 * (parFacetoFace - parTangentLength1)
            'Insert your code for output 1(Insulated Body)
            oStPoint.Set -(1 / 10 * (parFacetoFace - parTangentLength1) + _
                        parTangentLength1 - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                    -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick, 0) Then
                oStPoint.Set -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, Centerpos, dInsulationDiameter
            
            'Insert your code for output 2(Insulated Port1)
            oStPoint.Set -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1) * Cos(dVarAngle), _
                 -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1) * Sin(dVarAngle), 0
            oEnPoint.Set -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                 -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0

            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2
            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                 -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 3(Insulated Branch Port2)
            oStPoint.Set (9 / 10 * (parFacetoFace - parTangentLength1) - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                   (9 / 10 * (parFacetoFace - parTangentLength1) - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick2, 0) Then
                oStPoint.Set (9 / 10 * (parFacetoFace - parTangentLength1) - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                    (9 / 10 * (parFacetoFace - parTangentLength1) - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), Centerpos, oStPoint, dInsulationDiameter
            
            'Insert your code for output 4(Insulated Port2)
            oStPoint.Set 9 / 10 * (parFacetoFace - parTangentLength1) * Cos(dVarAngle), _
              9 / 10 * (parFacetoFace - parTangentLength1) * Sin(dVarAngle), 0
            oEnPoint.Set (9 / 10 * (parFacetoFace - parTangentLength1) - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                    (9 / 10 * (parFacetoFace - parTangentLength1) - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam) Then dInsulationDiameter = flangeDiam2 + parInsulationThickness * 2
            If CmpDblEqual(flangeThick2, 0) Then
                oEnPoint.Set (9 / 10 * (parFacetoFace - parTangentLength1) - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                    (9 / 10 * (parFacetoFace - parTangentLength1) - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 7(Insulated Branch Port3)
            oBranchPoint.Set ((parFace3toCenter - flangeThick3 - parInsulationThickness) * Cos(PI / 4)), _
                        (parFace3toCenter - flangeThick3 - parInsulationThickness) * Sin(PI / 4), 0
            If CmpDblEqual(flangeThick3, 0) Then
                oBranchPoint.Set ((parFace3toCenter - NEGLIGIBLE_THICKNESS) * Cos(PI / 4)), _
                        (parFace3toCenter - NEGLIGIBLE_THICKNESS) * Sin(PI / 4), 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            'Set the output
            m_oGeomHelper.CreateCylinder "InsulatedBranchPort3", Centerpos, oBranchPoint, dInsulationDiameter3
            
            'Insert your code for output 5 (Insulated Port3)
            oStPoint.Set (parFace3toCenter * Cos(PI / 4)), parFace3toCenter * Sin(PI / 4), 0
            oEnPoint.Set ((parFace3toCenter - flangeThick3 - parInsulationThickness) * Cos(PI / 4)), _
                        (parFace3toCenter - flangeThick3 - parInsulationThickness) * Sin(PI / 4), 0
            If CmpDblGreaterthan(flangeDiam3, pipeDiam3) Then dInsulationDiameter3 = flangeDiam3 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick3, 0) Then
                oEnPoint.Set ((parFace3toCenter - NEGLIGIBLE_THICKNESS) * Cos(PI / 4)), _
                        (parFace3toCenter - NEGLIGIBLE_THICKNESS) * Sin(PI / 4), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
        
            'Insert your code for output 8(Insulated Branch Port4)
            oBranchPoint.Set ((parFace3toCenter - flangeThick4 - parInsulationThickness) * _
                         Cos(PI / 4)), -(parFace3toCenter - flangeThick4 - parInsulationThickness) * Sin(PI / 4), 0
            If CmpDblEqual(flangeThick4, 0) Then
                oBranchPoint.Set ((parFace3toCenter - NEGLIGIBLE_THICKNESS) * _
                         Cos(PI / 4)), -(parFace3toCenter - NEGLIGIBLE_THICKNESS) * Sin(PI / 4), 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            'Set the output
            m_oGeomHelper.CreateCylinder "InsulatedBranchPort4", Centerpos, oBranchPoint, dInsulationDiameter3
        
            'Insert your code for output 6 (Insulated Port4)
            oStPoint.Set ((parFace3toCenter) * Cos(PI / 4)), -(parFace3toCenter) * Sin(PI / 4), 0
            oEnPoint.Set ((parFace3toCenter - flangeThick4 - parInsulationThickness) * _
                         Cos(PI / 4)), -(parFace3toCenter - flangeThick4 - parInsulationThickness) * Sin(PI / 4), 0
            If CmpDblGreaterthan(flangeDiam4, pipeDiam3) Then dInsulationDiameter3 = flangeDiam4 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick4, 0) Then
                oEnPoint.Set ((parFace3toCenter - NEGLIGIBLE_THICKNESS) * _
                         Cos(PI / 4)), -(parFace3toCenter - NEGLIGIBLE_THICKNESS) * Sin(PI / 4), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
         
         
       'Double Wye
        Case 526
            
            parFace1toCenter = arrayOfInputs(2)     'G
            parFace3toCenter = arrayOfInputs(4)     'D
            parFacetoFace = arrayOfInputs(5)        'F
            dVarAngle = parAngle - PI / 4
            
            'Insert your code for output 1(Insulated Body)
            oStPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                 -(parFace1toCenter - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick, 0) Then
            oStPoint.Set -(parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                 -(parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, Centerpos, dInsulationDiameter
            
            'Insert your code for output 2(Insulated Port1)
            oStPoint.Set -parFace1toCenter * Cos(dVarAngle), -parFace1toCenter * Sin(dVarAngle), 0
            oEnPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                 -(parFace1toCenter - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2
            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set -(parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                 -(parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 3(Insulated Branch Port2)
            oStPoint.Set (parFacetoFace - parFace1toCenter - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                (parFacetoFace - parFace1toCenter - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick2, 0) Then
                oStPoint.Set (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), Centerpos, oStPoint, dInsulationDiameter
            
            'Insert your code for output 4(Insulated Port2)
            oStPoint.Set (parFacetoFace - parFace1toCenter) * Cos(dVarAngle), _
                            (parFacetoFace - parFace1toCenter) * Sin(dVarAngle), 0
            oEnPoint.Set (parFacetoFace - parFace1toCenter - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                    (parFacetoFace - parFace1toCenter - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam) Then dInsulationDiameter = flangeDiam2 + parInsulationThickness * 2
            If CmpDblEqual(flangeThick2, 0) Then
                oEnPoint.Set (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 7(Insulated Branch Port3)
            oBranchPoint.Set (parFace3toCenter - flangeThick3 - parInsulationThickness) * Cos(PI / 4), _
                            (parFace3toCenter - flangeThick3 - parInsulationThickness) * Sin(PI / 4), 0
            If CmpDblEqual(flangeThick3, 0) Then
                oBranchPoint.Set (parFace3toCenter - NEGLIGIBLE_THICKNESS) * Cos(PI / 4), _
                                 (parFace3toCenter - NEGLIGIBLE_THICKNESS) * Sin(PI / 4), 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            'Set the output
            m_oGeomHelper.CreateCylinder "InsulatedBranchPort3", Centerpos, oBranchPoint, dInsulationDiameter3
            
            'Insert your code for output 5 (Insulated Port3)
            oStPoint.Set (parFace3toCenter) * Cos(PI / 4), (parFace3toCenter) * Sin(PI / 4), 0
            oEnPoint.Set (parFace3toCenter - flangeThick3 - parInsulationThickness) * Cos(PI / 4), _
                         (parFace3toCenter - flangeThick3 - parInsulationThickness) * Sin(PI / 4), 0
            If CmpDblGreaterthan(flangeDiam3, pipeDiam3) Then dInsulationDiameter3 = flangeDiam3 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick3, 0) Then
                oEnPoint.Set (parFace3toCenter - NEGLIGIBLE_THICKNESS) * Cos(PI / 4), _
                             (parFace3toCenter - NEGLIGIBLE_THICKNESS) * Sin(PI / 4), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
        
            'Insert your code for output 8(Insulated Branch Port4)
            oBranchPoint.Set (parFace3toCenter - flangeThick4 - parInsulationThickness) * Cos(PI / 4), _
                         -(parFace3toCenter - flangeThick4 - parInsulationThickness) * Sin(PI / 4), 0
            If CmpDblEqual(flangeThick4, 0) Then
                oBranchPoint.Set (parFace3toCenter - NEGLIGIBLE_THICKNESS) * Cos(PI / 4), _
                             -(parFace3toCenter - NEGLIGIBLE_THICKNESS) * Sin(PI / 4), 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            'Set the output
            m_oGeomHelper.CreateCylinder "InsulatedBranchPort4", Centerpos, oBranchPoint, dInsulationDiameter3
            
            'Insert your code for output 6 (Insulated Port4)
            oStPoint.Set (parFace3toCenter) * Cos(PI / 4), -(parFace3toCenter) * Sin(PI / 4), 0
            oEnPoint.Set (parFace3toCenter - flangeThick4 - parInsulationThickness) * Cos(PI / 4), _
                         -(parFace3toCenter - flangeThick4 - parInsulationThickness) * Sin(PI / 4), 0
            If CmpDblGreaterthan(flangeDiam4, pipeDiam3) Then dInsulationDiameter3 = flangeDiam4 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick4, 0) Then
                oEnPoint.Set (parFace3toCenter - NEGLIGIBLE_THICKNESS) * Cos(PI / 4), _
                             -(parFace3toCenter - NEGLIGIBLE_THICKNESS) * Sin(PI / 4), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
         
     'DoubleWye
        Case 527
            
            parSeat1toCenter = arrayOfInputs(12)     'G
            parSeattoSeat = arrayOfInputs(11)       'X
            parSeat3toCenter = arrayOfInputs(13)    'X'
            
            parFace1toCenter = parSeat1toCenter + depth - cptoffset
            dVarAngle = parAngle - PI / 4

            'Insert your code for output 1(Insulated Body)
            oStPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                     -(parFace1toCenter - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick, 0) Then
                oStPoint.Set -(parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                -(parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, Centerpos, dInsulationDiameter
            
            'Insert your code for output 2(Insulated Port1)
            oStPoint.Set -parFace1toCenter * Cos(dVarAngle), _
                        -parFace1toCenter * Sin(dVarAngle), 0
            oEnPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
            -(parFace1toCenter - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2
            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set -(parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                -(parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 3(Insulated Branch Port2)
            dFacetoface = parSeattoSeat + (depth - cptoffset) + (depth2 - cptoffset2)
            oStPoint.Set (dFacetoface - parFace1toCenter - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                    (dFacetoface - parFace1toCenter - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick2, 0) Then
                oStPoint.Set (dFacetoface - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                (dFacetoface - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), Centerpos, oStPoint, dInsulationDiameter
            
            'Insert your code for output 4(Insulated Port2)
            oStPoint.Set (dFacetoface - parFace1toCenter) * Cos(dVarAngle), _
                    (dFacetoface - parFace1toCenter) * Sin(dVarAngle), 0
            oEnPoint.Set (dFacetoface - parFace1toCenter - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                (dFacetoface - parFace1toCenter - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam) Then dInsulationDiameter = flangeDiam2 + parInsulationThickness * 2
            If CmpDblEqual(flangeThick2, 0) Then
                oEnPoint.Set (dFacetoface - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                (dFacetoface - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 7(Insulated Branch Port3)
            oBranchPoint.Set (parSeat3toCenter + depth3 - cptoffset3 - flangeThick3 - parInsulationThickness) * Cos(PI / 4), _
                         (parSeat3toCenter + depth3 - cptoffset3 - flangeThick3 - parInsulationThickness) * Sin(PI / 4), 0
            If CmpDblEqual(flangeThick3, 0) Then
                oBranchPoint.Set (parSeat3toCenter + depth3 - cptoffset3 - NEGLIGIBLE_THICKNESS) * Cos(PI / 4), _
                                 (parSeat3toCenter + depth3 - cptoffset3 - NEGLIGIBLE_THICKNESS) * Sin(PI / 4), 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            'Set the output
            m_oGeomHelper.CreateCylinder "InsulatedBranchPort3", Centerpos, oBranchPoint, dInsulationDiameter3
            
            'Insert your code for output 5 (Insulated Port3)
            oStPoint.Set (parSeat3toCenter + depth3 - cptoffset3) * Cos(PI / 4), (parSeat3toCenter + depth3 - cptoffset3) * Sin(PI / 4), 0
            oEnPoint.Set (parSeat3toCenter + depth3 - cptoffset3 - flangeThick3 - parInsulationThickness) * Cos(PI / 4), _
                         (parSeat3toCenter + depth3 - cptoffset3 - flangeThick3 - parInsulationThickness) * Sin(PI / 4), 0
            If CmpDblGreaterthan(flangeDiam3, pipeDiam3) Then dInsulationDiameter3 = flangeDiam3 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick3, 0) Then
                oEnPoint.Set (parSeat3toCenter + depth3 - cptoffset3 - NEGLIGIBLE_THICKNESS) * Cos(PI / 4), _
                             (parSeat3toCenter + depth3 - cptoffset3 - NEGLIGIBLE_THICKNESS) * Sin(PI / 4), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
        
            'Insert your code for output 8(Insulated Branch Port4)
            oBranchPoint.Set (parSeat3toCenter + depth4 - cptoffset4 - flangeThick4 - parInsulationThickness) * Cos(PI / 4), _
                            -(parSeat3toCenter + depth4 - cptoffset4 - flangeThick4 - parInsulationThickness) * Sin(PI / 4), 0
            If CmpDblEqual(flangeThick4, 0) Then
                oBranchPoint.Set (parSeat3toCenter + depth4 - cptoffset4 - NEGLIGIBLE_THICKNESS) * Cos(PI / 4), _
                             -(parSeat3toCenter + depth4 - cptoffset4 - NEGLIGIBLE_THICKNESS) * Sin(PI / 4), 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            'Set the output
            m_oGeomHelper.CreateCylinder "InsulatedBranchPort4", Centerpos, oBranchPoint, dInsulationDiameter3
            
            'Insert your code for output 6 (Insulated Port4)
            oStPoint.Set (parSeat3toCenter + depth4 - cptoffset4) * Cos(PI / 4), -(parSeat3toCenter + depth4 - cptoffset4) * Sin(PI / 4), 0
            oEnPoint.Set (parSeat3toCenter + depth4 - cptoffset4 - flangeThick4 - parInsulationThickness) * Cos(PI / 4), _
                         -(parSeat3toCenter + depth4 - cptoffset4 - flangeThick4 - parInsulationThickness) * Sin(PI / 4), 0
            If CmpDblGreaterthan(flangeDiam4, pipeDiam3) Then dInsulationDiameter3 = flangeDiam4 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick4, 0) Then
                oEnPoint.Set (parSeat3toCenter + depth4 - cptoffset4 - NEGLIGIBLE_THICKNESS) * Cos(PI / 4), _
                             -(parSeat3toCenter + depth4 - cptoffset4 - NEGLIGIBLE_THICKNESS) * Sin(PI / 4), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
            
           
        '1/8 Bend DoubleWye
        Case 528
        
            parFace3toCenter = arrayOfInputs(4)     'H
            parFacetoFace = arrayOfInputs(5)        'F
            parPortttoPortOffset = arrayOfInputs(6) 'D
            parTangentLength1 = arrayOfInputs(7)    'G
            dVarAngle = parAngle - PI / 2
            
            'Distance from Center to Tangent Length start position is taken as 1 / 10 * (parFacetoFace - parTangentLength1)
            'Insert your code for output 1(Insulated Body)
            oStPoint.Set -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                     -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick, 0) Then
                oStPoint.Set -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            oEnPoint.Set 0, 0, 0
        
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 2(Insulated Port1)
            oStPoint.Set -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1) * Cos(dVarAngle), _
                 -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1) * Sin(dVarAngle), 0
            oEnPoint.Set -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
             -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2
            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                -(1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1 - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 3(Insulated Branch Port2)
            oBranchPoint.Set (9 / 10 * (parFacetoFace - parTangentLength1) - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                   (9 / 10 * (parFacetoFace - parTangentLength1) - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick2, 0) Then
                oBranchPoint.Set (9 / 10 * (parFacetoFace - parTangentLength1) - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                    (9 / 10 * (parFacetoFace - parTangentLength1) - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), Centerpos, oBranchPoint, dInsulationDiameter
            
            'Insert your code for output 4(Insulated Port2)
            oStPoint.Set 9 / 10 * (parFacetoFace - parTangentLength1) * Cos(dVarAngle), _
                      9 / 10 * (parFacetoFace - parTangentLength1) * Sin(dVarAngle), 0
            oEnPoint.Set (9 / 10 * (parFacetoFace - parTangentLength1) - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                    (9 / 10 * (parFacetoFace - parTangentLength1) - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam) Then dInsulationDiameter = flangeDiam2 + parInsulationThickness * 2
            If CmpDblEqual(flangeThick2, 0) Then
                oEnPoint.Set (9 / 10 * (parFacetoFace - parTangentLength1) - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                    (9 / 10 * (parFacetoFace - parTangentLength1) - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 5 (Insulated Port3)
            oStPoint.Set (parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1)), parFace3toCenter, 0
            oEnPoint.Set ((parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1))), _
                        (parFace3toCenter - flangeThick3 - parInsulationThickness), 0
            If CmpDblGreaterthan((flangeThick3 + parInsulationThickness), 0.25 * parFace3toCenter) Then
                 oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), 0.75 * parFace3toCenter, 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam3, pipeDiam3) Then dInsulationDiameter3 = flangeDiam3 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick3, 0) Then
                oEnPoint.Set ((parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1))), _
                        (parFace3toCenter - NEGLIGIBLE_THICKNESS), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
            
            'Insert your code for output 6 (Insulated Port4)
            oStPoint.Set (parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1)), -parFace3toCenter, 0
            oEnPoint.Set ((parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1))), _
                        -(parFace3toCenter - flangeThick4 - parInsulationThickness), 0
            If CmpDblGreaterthan((flangeThick4 + parInsulationThickness), 0.25 * parFace3toCenter) Then
                 oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), -0.75 * parFace3toCenter, 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam4, pipeDiam3) Then dInsulationDiameter3 = flangeDiam4 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick4, 0) Then
                oEnPoint.Set ((parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1))), _
                        -(parFace3toCenter - NEGLIGIBLE_THICKNESS), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
        
            
            'Create Bend wye
            Set oCenter = New DPosition
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oStPoint = New DPosition
            Set oEnPoint = New DPosition
            Set oTraceStr = New ComplexString3d
            Set oCollection = New Collection
            Set oLine = New Line3d
            Set oArc = New Arc3d
            Set oNormal = New DVector
            
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            oCenter.Set (parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1)), parFace3toCenter, 0
            oNormal.Set 0, 1, 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                  oCenter.x, oCenter.y, oCenter.z, _
                                                 oNormal.x, oNormal.y, oNormal.z, dInsulationDiameter3 / 2)


            oStPoint.Set (parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1)), parFace3toCenter, 0
            oEnPoint.Set (parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1)), 0.75 * parFace3toCenter, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine

            dYlength = 0.25 * parFace3toCenter * ((1 / Cos(PI / 4)) - 1)
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                         (parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1) - (0.25 * parFace3toCenter / Cos(PI / 4))), 0.75 * parFace3toCenter, 0, _
                         (parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1)), 0.75 * parFace3toCenter, 0, _
                         (parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1) - dYlength), 0.5 * parFace3toCenter, 0)
            oCollection.Add oArc

            oStPoint.Set (parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1) - dYlength), 0.5 * parFace3toCenter, 0
            oEnPoint.Set 0, 0, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine

            oStPoint.Set (parPortttoPortOffset - (1 / 10 * (parFacetoFace - parTangentLength1) + parTangentLength1)), parFace3toCenter, 0
            Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)
                                 
             For Each ObjSurface In Surfset
                 m_OutputColl.AddOutput "InsElbow_", ObjSurface
             Next ObjSurface
             
             Set ObjSurface = Nothing
             Set Surfset = Nothing
             Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)
             For Each ObjSurface In Surfset
                Set oTransMatrix = New AutoMath.DT4x4
                oTransMatrix.LoadIdentity
                oTransMatrix.IndexValue(5) = -1
                ObjSurface.Transform oTransMatrix
                m_OutputColl.AddOutput "InsElbow_", ObjSurface
             Next ObjSurface
            
            'Remove References
             For iCount = 1 To oCollection.Count
                 oCollection.Remove 1
             Next iCount
             For iCount = 1 To Surfset.Count
                 Surfset.Remove 1
             Next iCount
             
        '1/8 Bend DoubleWye
        Case 529
            
            parFace1toCenter = arrayOfInputs(2)     'G
            parFace3toCenter = arrayOfInputs(4)     'E
            parFacetoFace = arrayOfInputs(5)        'F
            parPortttoPortOffset = arrayOfInputs(6) 'D
            dVarAngle = parAngle - PI / 2
            
            'Insert your code for output 1(Insulated Body)
            oStPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                 -(parFace1toCenter - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick, 0) Then
                oStPoint.Set (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                  (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            oEnPoint.Set 0, 0, 0
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 2(Insulated Port1)
            oStPoint.Set -parFace1toCenter * Cos(dVarAngle), -parFace1toCenter * Sin(dVarAngle), 0
            oEnPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                    -(parFace1toCenter - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2
            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 3(Insulated Branch Port2)
            oBranchPoint.Set (parFacetoFace - parFace1toCenter - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                   (parFacetoFace - parFace1toCenter - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick2, 0) Then
                oBranchPoint.Set (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), Centerpos, oBranchPoint, dInsulationDiameter
            
            'Insert your code for output 4(Insulated Port2)
            oStPoint.Set (parFacetoFace - parFace1toCenter) * Cos(dVarAngle), _
                        (parFacetoFace - parFace1toCenter) * Sin(dVarAngle), 0
            oEnPoint.Set ((parFacetoFace - parFace1toCenter) - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                    ((parFacetoFace - parFace1toCenter) - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam) Then dInsulationDiameter = flangeDiam2 + parInsulationThickness * 2
            If CmpDblEqual(flangeThick2, 0) Then
                oEnPoint.Set (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
          
            'Insert your code for output 5 (Insulated Port3)
            oStPoint.Set (parPortttoPortOffset - (parFace1toCenter)), parFace3toCenter, 0
            oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), _
                        (parFace3toCenter - flangeThick3 - parInsulationThickness), 0
            If CmpDblGreaterthan((flangeThick3 + parInsulationThickness), 0.25 * parFace3toCenter) Then
                 oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), 0.75 * parFace3toCenter, 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam3, pipeDiam3) Then dInsulationDiameter3 = flangeDiam3 + parInsulationThickness * 2
            If CmpDblEqual(flangeThick3, 0) Then
                oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), _
                        (parFace3toCenter - NEGLIGIBLE_THICKNESS), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
            
            'Insert your code for output 6 (Insulated Port4)
            oStPoint.Set (parPortttoPortOffset - (parFace1toCenter)), -parFace3toCenter, 0
            oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), _
                        -(parFace3toCenter - flangeThick4 - parInsulationThickness), 0
            If CmpDblGreaterthan((flangeThick4 + parInsulationThickness), 0.25 * parFace3toCenter) Then
                 oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), -0.75 * parFace3toCenter, 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam4, pipeDiam3) Then dInsulationDiameter3 = flangeDiam4 + parInsulationThickness * 2
            If CmpDblEqual(flangeThick4, 0) Then
                oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), _
                        -(parFace3toCenter - NEGLIGIBLE_THICKNESS), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
            
            'Create Bend wye
            Set oCenter = New DPosition
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oStPoint = New DPosition
            Set oEnPoint = New DPosition
            Set oTraceStr = New ComplexString3d
            Set oCollection = New Collection
            Set oLine = New Line3d
            Set oArc = New Arc3d
            Set oNormal = New DVector
            
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            oCenter.Set (parPortttoPortOffset - parFace1toCenter), parFace3toCenter, 0
            oNormal.Set 0, 1, 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                  oCenter.x, oCenter.y, oCenter.z, _
                                                 oNormal.x, oNormal.y, oNormal.z, dInsulationDiameter3 / 2)

            oStPoint.Set (parPortttoPortOffset - parFace1toCenter), parFace3toCenter, 0
            oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), 0.75 * parFace3toCenter, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
            
            dYlength = 0.25 * parFace3toCenter * ((1 / Cos(PI / 4)) - 1)
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                         (parPortttoPortOffset - parFace1toCenter - (0.25 * parFace3toCenter / Cos(PI / 4))), 0.75 * parFace3toCenter, 0, _
                         (parPortttoPortOffset - parFace1toCenter), 0.75 * parFace3toCenter, 0, _
                         (parPortttoPortOffset - parFace1toCenter - dYlength), 0.5 * parFace3toCenter, 0)
            oCollection.Add oArc

            oStPoint.Set (parPortttoPortOffset - parFace1toCenter - dYlength), 0.5 * parFace3toCenter, 0
            oEnPoint.Set 0, 0, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine

            oStPoint.Set (parPortttoPortOffset - parFace1toCenter), parFace3toCenter, 0
            Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)

             For Each ObjSurface In Surfset
                 m_OutputColl.AddOutput "InsElbow_", ObjSurface
             Next ObjSurface
             
             Set ObjSurface = Nothing
             Set Surfset = Nothing
             Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)
             For Each ObjSurface In Surfset
                Set oTransMatrix = New AutoMath.DT4x4
                oTransMatrix.LoadIdentity
                oTransMatrix.IndexValue(5) = -1
                ObjSurface.Transform oTransMatrix
                m_OutputColl.AddOutput "InsElbow_", ObjSurface
             Next ObjSurface
            
            'Remove References
             For iCount = 1 To oCollection.Count
                 oCollection.Remove 1
             Next iCount
             For iCount = 1 To Surfset.Count
                 Surfset.Remove 1
             Next iCount
             
         '1/8 Bend DoubleWye
        Case 530
            
            parSeat1toCenter = arrayOfInputs(12)     'B
            parPortttoPortOffset = arrayOfInputs(6) 'G
            parSeattoSeat = arrayOfInputs(11)       'X
            parSeat3toCenter = arrayOfInputs(13)    'E
            
            parFace1toCenter = parSeat1toCenter + depth - cptoffset
            parFace3toCenter = parSeat3toCenter + depth3 - cptoffset3
            dVarAngle = parAngle - PI / 2
             
            'Insert your code for output 1(Insulated Body)
            oStPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                     -(parFace1toCenter - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick, 0) Then
                oStPoint.Set (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                     (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            oEnPoint.Set 0, 0, 0
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 2(Insulated Port1)
            oStPoint.Set -parFace1toCenter * Cos(dVarAngle), -parFace1toCenter * Sin(dVarAngle), 0
            oEnPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                       -(parFace1toCenter - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2
            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                       (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 3(Insulated Branch Port2)
            dFacetoface = parSeattoSeat + (depth - cptoffset) + (depth2 - cptoffset2)
            oBranchPoint.Set (dFacetoface - parFace1toCenter - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                   (dFacetoface - parFace1toCenter - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick2, 0) Then
                oBranchPoint.Set (dFacetoface - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                  (dFacetoface - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), Centerpos, oBranchPoint, dInsulationDiameter
            
            'Insert your code for output 4(Insulated Port2)
            oStPoint.Set (dFacetoface - parFace1toCenter) * Cos(dVarAngle), (dFacetoface - parFace1toCenter) * Sin(dVarAngle), 0
            oEnPoint.Set ((dFacetoface - parFace1toCenter) - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                    ((dFacetoface - parFace1toCenter) - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam) Then dInsulationDiameter = flangeDiam2 + parInsulationThickness * 2
            If CmpDblEqual(flangeThick2, 0) Then
                oEnPoint.Set (dFacetoface - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                 (dFacetoface - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 5 (Insulated Port3)
            oStPoint.Set (parPortttoPortOffset - (parFace1toCenter)), parSeat3toCenter + (depth3 - cptoffset3), 0
            oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), _
                        (parFace3toCenter - flangeThick3 - parInsulationThickness), 0
            If CmpDblGreaterthan((flangeThick3 + parInsulationThickness), 0.25 * parSeat3toCenter) Then
                 oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), 0.75 * parSeat3toCenter, 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam3, pipeDiam3) Then dInsulationDiameter3 = flangeDiam3 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick3, 0) Then
                oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), _
                        (parFace3toCenter - NEGLIGIBLE_THICKNESS), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
            
            'Insert your code for output 6 (Insulated Port4)
            oStPoint.Set (parPortttoPortOffset - (parFace1toCenter)), -(parSeat3toCenter + (depth4 - cptoffset4)), 0
            oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), _
                        -(parFace3toCenter - flangeThick4 - parInsulationThickness), 0
            If CmpDblGreaterthan((flangeThick4 + parInsulationThickness), 0.25 * parSeat3toCenter) Then
                 oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), -0.75 * parSeat3toCenter, 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam4, pipeDiam3) Then dInsulationDiameter3 = flangeDiam4 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick4, 0) Then
                oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), _
                        -(parFace3toCenter - NEGLIGIBLE_THICKNESS), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
        
            
            'Create Bend wye
            Set oCenter = New DPosition
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oStPoint = New DPosition
            Set oEnPoint = New DPosition
            Set oTraceStr = New ComplexString3d
            Set oCollection = New Collection
            Set oLine = New Line3d
            Set oArc = New Arc3d
            Set oNormal = New DVector
            
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            oCenter.Set (parPortttoPortOffset - parFace1toCenter), parSeat3toCenter, 0
            oNormal.Set 0, 1, 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                  oCenter.x, oCenter.y, oCenter.z, _
                                                 oNormal.x, oNormal.y, oNormal.z, dInsulationDiameter3 / 2)

            oStPoint.Set (parPortttoPortOffset - parFace1toCenter), parSeat3toCenter, 0
            oEnPoint.Set (parPortttoPortOffset - parFace1toCenter), 0.75 * parSeat3toCenter, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
            
            dYlength = 0.25 * parSeat3toCenter * ((1 / Cos(PI / 4)) - 1)
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                         (parPortttoPortOffset - parFace1toCenter - (0.25 * parSeat3toCenter / Cos(PI / 4))), 0.75 * parSeat3toCenter, 0, _
                         (parPortttoPortOffset - parFace1toCenter), 0.75 * parSeat3toCenter, 0, _
                         (parPortttoPortOffset - parFace1toCenter - dYlength), 0.5 * parSeat3toCenter, 0)
            oCollection.Add oArc

            oStPoint.Set (parPortttoPortOffset - parFace1toCenter - dYlength), 0.5 * parSeat3toCenter, 0
            oEnPoint.Set 0, 0, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine

            oStPoint.Set (parPortttoPortOffset - parFace1toCenter), parSeat3toCenter, 0
            Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)
             For Each ObjSurface In Surfset
                 m_OutputColl.AddOutput "InsElbow_", ObjSurface
             Next ObjSurface
             
             Set ObjSurface = Nothing
             Set Surfset = Nothing
             Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)
             For Each ObjSurface In Surfset
                Set oTransMatrix = New AutoMath.DT4x4
                oTransMatrix.LoadIdentity
                oTransMatrix.IndexValue(5) = -1
                ObjSurface.Transform oTransMatrix
                m_OutputColl.AddOutput "InsElbow_", ObjSurface
             Next ObjSurface
            
            'Remove References
             For iCount = 1 To oCollection.Count
                 oCollection.Remove 1
             Next iCount
             For iCount = 1 To Surfset.Count
                 Surfset.Remove 1
             Next iCount
             
        '1/8 Bend DoubleWye
        Case 531
            
            parTangentLength1 = arrayOfInputs(7)        'B
            parTangentLength3 = arrayOfInputs(8)        'A'
            parFacetoBendCenter = arrayOfInputs(9)      'C'
            parCentertoPortOffset = arrayOfInputs(10)   'H
            parSeattoSeat = arrayOfInputs(11)            'X
            parSeat1toCenter = arrayOfInputs(12)         'G
            parSeat3toCenter = arrayOfInputs(13)         'X'
            parBendRadius = oPipeComponent.BendRadius   'R'
                        
            parFace1toCenter = parSeat1toCenter - cptoffset + depth
            parFacetoFace = parSeattoSeat + (depth - cptoffset) + (depth2 - cptoffset2)
            parFace2toCenter = parFacetoFace - parFace1toCenter
            parFace3toCenter = parSeat3toCenter - cptoffset3 + depth3
            dVarAngle = parAngle - PI / 2
            
            'Insert your code for output 1(Insulated Body)
            oStPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                           -(parFace1toCenter - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick, 0) Then
                oStPoint.Set (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                        (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            oEnPoint.Set 0, 0, 0
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 2(Insulated Port1)
            oStPoint.Set -parFace1toCenter * Cos(dVarAngle), -parFace1toCenter * Sin(dVarAngle), 0
            oEnPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                                 -(parFace1toCenter - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2
            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 3(Insulated Branch Port2)
            oBranchPoint.Set (parFacetoFace - parFace1toCenter - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                   (parFacetoFace - parFace1toCenter - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick2, 0) Then
                oBranchPoint.Set (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                    (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), Centerpos, oBranchPoint, dInsulationDiameter
            
            'Insert your code for output 4(Insulated Port2)
            oStPoint.Set (parFacetoFace - parFace1toCenter) * Cos(dVarAngle), _
                            (parFacetoFace - parFace1toCenter) * Sin(dVarAngle), 0
            oEnPoint.Set ((parFacetoFace - parFace1toCenter) - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                    ((parFacetoFace - parFace1toCenter) - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam) Then dInsulationDiameter = flangeDiam2 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick2, 0) Then
                oEnPoint.Set (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 5 (Insulated Port3)
            oStPoint.Set parCentertoPortOffset, parFace3toCenter, 0
            oEnPoint.Set parCentertoPortOffset, _
                        (parFace3toCenter - flangeThick3 - parInsulationThickness), 0
            If CmpDblGreaterthan((flangeThick3 + parInsulationThickness), parTangentLength3) Then
                 oEnPoint.Set parCentertoPortOffset, (parFace3toCenter - parTangentLength3), 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam3, pipeDiam3) Then dInsulationDiameter3 = flangeDiam3 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick3, 0) Then
                oEnPoint.Set parCentertoPortOffset, _
                        (parFace3toCenter - NEGLIGIBLE_THICKNESS), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
            
            'Insert your code for output 6 (Insulated Port4)
            oStPoint.Set parCentertoPortOffset, -parFace3toCenter, 0
            oEnPoint.Set parCentertoPortOffset, _
                        -(parFace3toCenter - flangeThick4 - parInsulationThickness), 0
            If CmpDblGreaterthan((flangeThick4 + parInsulationThickness), parTangentLength3) Then
                 oEnPoint.Set parCentertoPortOffset, -(parFace3toCenter - parTangentLength3), 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam4, pipeDiam3) Then dInsulationDiameter3 = flangeDiam4 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick4, 0) Then
                oEnPoint.Set parCentertoPortOffset, _
                        -(parFace3toCenter - NEGLIGIBLE_THICKNESS), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
             
            'Create Bend wye
            Set oCenter = New DPosition
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oStPoint = New DPosition
            Set oEnPoint = New DPosition
            Set oTraceStr = New ComplexString3d
            Set oCollection = New Collection
            Set oLine = New Line3d
            Set oArc = New Arc3d
            Set oNormal = New DVector
            
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            oCenter.Set parCentertoPortOffset, parSeat3toCenter, 0
            oNormal.Set 0, 1, 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                  oCenter.x, oCenter.y, oCenter.z, _
                                                 oNormal.x, oNormal.y, oNormal.z, dInsulationDiameter3 / 2)
                                                 
            oStPoint.Set parCentertoPortOffset, parSeat3toCenter, 0
            oEnPoint.Set parCentertoPortOffset, parSeat3toCenter + depth3 - cptoffset3 - parTangentLength3, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
                        
            oStPoint.Set parCentertoPortOffset, parSeat3toCenter + depth3 - cptoffset3 - parTangentLength3, 0
            oCenter.Set parCentertoPortOffset - parBendRadius, parSeat3toCenter + depth3 - cptoffset3 - parTangentLength3, 0
            oEnPoint.Set parCentertoPortOffset - parBendRadius * (1 - Sin(PI / 4)), parSeat3toCenter + depth3 - cptoffset3 - parTangentLength3 - parBendRadius * Cos(PI / 4), 0
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                    oCenter.x, oCenter.y, oCenter.z, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            oCollection.Add oArc
            
            oStPoint.Set parCentertoPortOffset - parBendRadius * (1 - Sin(PI / 4)), parSeat3toCenter + depth3 - cptoffset3 - parTangentLength3 - parBendRadius * Cos(PI / 4), 0
            oEnPoint.Set 0, 0, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
            
            oStPoint.Set parCentertoPortOffset, parSeat3toCenter, 0
            Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)
                                 
             For Each ObjSurface In Surfset
                 m_OutputColl.AddOutput "InsElbow_", ObjSurface
             Next ObjSurface
            
            Set ObjSurface = Nothing
            Set Surfset = Nothing
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)
             For Each ObjSurface In Surfset
                Set oTransMatrix = New AutoMath.DT4x4
                oTransMatrix.LoadIdentity
                oTransMatrix.IndexValue(5) = -1
                ObjSurface.Transform oTransMatrix
                m_OutputColl.AddOutput "InsElbow_", ObjSurface
             Next ObjSurface
            
            'Remove References
             For iCount = 1 To oCollection.Count
                 oCollection.Remove 1
             Next iCount
             For iCount = 1 To Surfset.Count
                 Surfset.Remove 1
             Next iCount
             
        '1/8 Bend DoubleWye
        Case 532
            
            parFace1toCenter = arrayOfInputs(2)         'G
            parFace2toCenter = arrayOfInputs(3)         'E
            parFace3toCenter = arrayOfInputs(4)         'E'
            parFacetoFace = arrayOfInputs(5)            'F
            parTangentLength1 = arrayOfInputs(7)        'B
            parTangentLength3 = arrayOfInputs(8)        'A'
            parFacetoBendCenter = arrayOfInputs(9)      'C'
            parCentertoPortOffset = arrayOfInputs(10)   'H
            parBendRadius = oPipeComponent.BendRadius   'R'
            
            parFace2toCenter = parFacetoFace - parFace1toCenter
            dVarAngle = parAngle - PI / 2
            
            'Insert your code for output 1(Insulated Body)
            oStPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                      -(parFace1toCenter - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick, 0) Then
                oStPoint.Set (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                            (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            oEnPoint.Set 0, 0, 0
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 2(Insulated Port1)
            oStPoint.Set -parFace1toCenter * Cos(dVarAngle), -parFace1toCenter * Sin(dVarAngle), 0
            oEnPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness) * Cos(dVarAngle), _
                          -(parFace1toCenter - flangeThick - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam, pipeDiam) Then dInsulationDiameter = flangeDiam + parInsulationThickness * 2
            If CmpDblEqual(flangeThick, 0) Then
                oEnPoint.Set (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                        (-parFace1toCenter + NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
            
            'Insert your code for output 3(Insulated Branch Port2)
            oBranchPoint.Set (parFacetoFace - parFace1toCenter - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                   (parFacetoFace - parFace1toCenter - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblEqual(flangeThick2, 0) Then
                oBranchPoint.Set (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                     (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            dInsulationDiameter = pipeDiam + parInsulationThickness * 2
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), Centerpos, oBranchPoint, dInsulationDiameter
        
            'Insert your code for output 4(Insulated Port2)
            oStPoint.Set (parFacetoFace - parFace1toCenter) * Cos(dVarAngle), _
                            (parFacetoFace - parFace1toCenter) * Sin(dVarAngle), 0
            oEnPoint.Set ((parFacetoFace - parFace1toCenter) - flangeThick2 - parInsulationThickness) * Cos(dVarAngle), _
                   ((parFacetoFace - parFace1toCenter) - flangeThick2 - parInsulationThickness) * Sin(dVarAngle), 0
            If CmpDblGreaterthan(flangeDiam2, pipeDiam) Then dInsulationDiameter = flangeDiam2 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick2, 0) Then
                oEnPoint.Set (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Cos(dVarAngle), _
                                (parFacetoFace - parFace1toCenter - NEGLIGIBLE_THICKNESS) * Sin(dVarAngle), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter
                      
            'Insert your code for output 5 (Insulated Port3)
            oStPoint.Set parCentertoPortOffset, parFace3toCenter, 0
            oEnPoint.Set parCentertoPortOffset, _
                        (parFace3toCenter - flangeThick3 - parInsulationThickness), 0
            If CmpDblGreaterthan((flangeThick3 + parInsulationThickness), parTangentLength3) Then
                 oEnPoint.Set parCentertoPortOffset, (parFace3toCenter - parTangentLength3), 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam3, pipeDiam3) Then dInsulationDiameter3 = flangeDiam3 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick3, 0) Then
                oEnPoint.Set parCentertoPortOffset, _
                        (parFace3toCenter - NEGLIGIBLE_THICKNESS), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
            
            'Insert your code for output 6 (Insulated Port4)
            oStPoint.Set parCentertoPortOffset, -parFace3toCenter, 0
            oEnPoint.Set parCentertoPortOffset, _
                        -(parFace3toCenter - flangeThick4 - parInsulationThickness), 0
            If CmpDblGreaterthan((flangeThick4 + parInsulationThickness), parTangentLength3) Then
                 oEnPoint.Set parCentertoPortOffset, -(parFace3toCenter - parTangentLength3), 0
            End If
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            If CmpDblGreaterthan(flangeDiam4, pipeDiam3) Then dInsulationDiameter3 = flangeDiam4 + parInsulationThickness * 2
             If CmpDblEqual(flangeThick4, 0) Then
                oEnPoint.Set parCentertoPortOffset, _
                        -(parFace3toCenter - NEGLIGIBLE_THICKNESS), 0
            End If
            'Set the output
            iOutput = iOutput + 1
            m_oGeomHelper.CreateCylinder arrayOfOutputs(iOutput), oStPoint, oEnPoint, dInsulationDiameter3
            
            
            'Create Bend wye
            Set oCenter = New DPosition
            Set oGeomFactory = New IngrGeom3D.GeometryFactory
            Set oStPoint = New DPosition
            Set oEnPoint = New DPosition
            Set oTraceStr = New ComplexString3d
            Set oCollection = New Collection
            Set oLine = New Line3d
            Set oArc = New Arc3d
            Set oNormal = New DVector
            
            dInsulationDiameter3 = pipeDiam3 + parInsulationThickness * 2
            oCenter.Set parCentertoPortOffset, parFace3toCenter - (depth3 - cptoffset3), 0
            oNormal.Set 0, 1, 0
            Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                  oCenter.x, oCenter.y, oCenter.z, _
                                                 oNormal.x, oNormal.y, oNormal.z, dInsulationDiameter3 / 2)
                                                 
            oStPoint.Set parCentertoPortOffset, parFace3toCenter - (depth3 - cptoffset3), 0
            oEnPoint.Set parCentertoPortOffset, parFace3toCenter - parTangentLength3, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
                        
            oStPoint.Set parCentertoPortOffset, parFace3toCenter - parTangentLength3, 0
            oCenter.Set parCentertoPortOffset - parBendRadius, parFace3toCenter - parTangentLength3, 0
            oEnPoint.Set parCentertoPortOffset - parBendRadius * (1 - Sin(PI / 4)), parFace3toCenter - parTangentLength3 - parBendRadius * Cos(PI / 4), 0
            Set oArc = oGeomFactory.Arcs3d.CreateByCenterStartEnd(Nothing, _
                                    oCenter.x, oCenter.y, oCenter.z, _
                                    oStPoint.x, oStPoint.y, oStPoint.z, _
                                    oEnPoint.x, oEnPoint.y, oEnPoint.z)
            oCollection.Add oArc
            
            oStPoint.Set parCentertoPortOffset - parBendRadius * (1 - Sin(PI / 4)), parFace3toCenter - parTangentLength3 - parBendRadius * Cos(PI / 4), 0
            oEnPoint.Set 0, 0, 0
            Set oLine = PlaceTrLine(oStPoint, oEnPoint)
            oCollection.Add oLine
            
            oStPoint.Set parCentertoPortOffset, parFace3toCenter - depth3 + cptoffset3, 0
            Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
            Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)
                                 
             For Each ObjSurface In Surfset
                 m_OutputColl.AddOutput "InsElbow_", ObjSurface
             Next ObjSurface
             
             Set ObjSurface = Nothing
             Set Surfset = Nothing
             Set Surfset = oGeomFactory.GeometryServices.CreateBySingleSweep( _
                                   m_OutputColl.ResourceManager, oTraceStr, oCircle, _
                                 CircularCorner, 0, stnorm, ednorm, True)
             For Each ObjSurface In Surfset
                Set oTransMatrix = New AutoMath.DT4x4
                oTransMatrix.LoadIdentity
                oTransMatrix.IndexValue(5) = -1
                ObjSurface.Transform oTransMatrix
                m_OutputColl.AddOutput "InsElbow_", ObjSurface
             Next ObjSurface
            
            'Remove References
             For iCount = 1 To oCollection.Count
                 oCollection.Remove 1
             Next iCount
             For iCount = 1 To Surfset.Count
                 Surfset.Remove 1
             Next iCount
            
        Case Else
            GoTo ErrorLabel:
    End Select
            Set m_oGeomHelper = Nothing
            Set oArc = Nothing
            Set oBranchPoint = Nothing
            Set oCenter = Nothing
            Set oCircle = Nothing
            Set oCollection = Nothing
            Set oEnPoint = Nothing
            Set oGeomFactory = Nothing
            Set oLine = Nothing
            Set oNormal = Nothing
            Set oPipeComponent = Nothing
            Set oStPoint = Nothing
            Set oTraceStr = Nothing
            Set Surfset = Nothing
            Set ObjSurface = Nothing
            Set oTransMatrix = Nothing

    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub

