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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'   CInsulation.cls
'   Author:          NN
'   Creation Date:  Sunday, Nov 5 2000
'   Description:
'   This class module is the place for user to implement graphical part of VBSymbol for this aspect
'   Change History:
'   dd.mmm.yyyy       who                   change description
'   -----------       -----                  ------------------
'   09.Jul.2003      SymbolTeam(India)     Copyright Information, Header  is added.
'   16.Feb.2005      svsmylav              TR-59827 Added check to ensure non-zero diameter and non-zero length for
'                                           cylinder to avoid PlaceCylinder function failure.
'   11.Sep.2006      KKC                   CR 104167 Request to update tee symbol to optionally support 2 dimensions
'   14.Sep.2006      KKC                   DI-95670  Replace names with initials in all revision history sheets and symbols")
'   21-Dec-06        svsmylav              TR-112718: Face2toCenter and Face3toCenter are optional inputs added in V7 -
'                                          these are moved to the end of first two V6.1 existing inputs (otherwise mirror operation failed).
'   19-Feb-07        svsmylav              CR-113435: Optional input 'Angle' is added to support Trimmable Tee.
'   06-Mar-07        svsmylav              TR-116016: Added logic to consider optional inputs based on part data basis values
'   06.Aug.2007      VRG                   TR-124961 Determination of flange diameter based on flange thickness is avoided
'   18-Mar-08        KKC                   CR-127565 and CR-127567. Created the Sanitary Tee symbol with various gemotric options
'   12-May-08        RRK                   CR-141196:Enhanced the symbol to support long sweep tees with part data basis values of 862, 863 and 864
'   16-Jul-08        RRK                   CR-CP-145694 Enhanced the symbol to support outlet tee of JIS, by making face 3 to center as optional
'                                          for the part data basis- 'Sanitary tee, 90o side outlet on center, specified by face-to-center 1 and face-to-center 3'(829)
'   05.July.08       RUK                   CR-CP-146405  Enhance Tee symbol for vortex tee geometry per JIS G 5527
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Const NEGLIGIBLE_THICKNESS = 0.0001
Private Const MODULE = "Insulation:"    'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim flangeThick As Double
    Dim cptoffset As Double
    Dim flangeDiam As Double
    Dim depth As Double
    Dim CenterPos As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0

    Dim iOutput As Integer

    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedBranch As Object
    Dim ObjInsulatedPort1 As Object
    Dim ObjInsulatedPort2 As Object
    Dim ObjInsulatedPort3 As Object

    Dim parFacetoCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
    Dim parFacetoFace As Double
    Dim parFace1toCenter As Double
    Dim parInsulationThickness As Double
    Dim parAngle As Double
    Dim parSeat3toCenter As Double
    Dim parPlugOuttoCenter As Double
    Dim parBranchDirection As Double
    Dim parSeat1toSeat2 As Double
    Dim parSeat1toCenter As Double
    Dim parBranchRadius As Double
    Dim parTangentLength1 As Double
    Dim parTangentLength3 As Double
    Dim parSeat2toCenter As Double
    Dim parBasetoCenter As Double
    Dim parBaseDiam As Double
    Dim parBaseWidth As Double
    Dim parBaseThickness As Double
    Dim parRibThickness As Double
    Dim parOutFacetoCenter As Double
    Dim parOffBranch1toBranch2 As Double
    Dim parOutOffsetPort1 As Double
    Dim parOutSeattoCenter As Double
    Dim parDiameter As Double
    Dim parLength As Double
    Dim parSeattoSeat As Double

    Dim pipeDiam1 As Double
    Dim flangeThick1 As Double
    Dim flangeDiam1 As Double
    Dim depth1 As Double
    Dim CptOffset1 As Double

    Dim pipeDiam2 As Double
    Dim flangeThick2 As Double
    Dim CptOffset2 As Double
    Dim flangeDiam2 As Double
    Dim depth2 As Double

    Dim pipeDiam3 As Double
    Dim flangeThick3 As Double
    Dim CptOffset3 As Double
    Dim flangeDiam3 As Double
    Dim depth3 As Double

    Dim pipeDiam4 As Double
    Dim flangeThick4 As Double
    Dim CptOffset4 As Double
    Dim flangeDiam4 As Double
    Dim depth4 As Double

    Dim pipeDiam5 As Double
    Dim flangeThick5 As Double
    Dim CptOffset5 As Double
    Dim flangeDiam5 As Double
    Dim depth5 As Double
    
    Dim dSeattoSeat As Double
    Dim dSeat1toCenter As Double
    Dim dSeat2toCenter As Double
    Dim dSeat3toCenter As Double
    Dim dNozzleLength As Double
    
    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)

    iOutput = 0

    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    'Retrieve the Nozzles properties
    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, flangeThick1, flangeDiam1, _
                       CptOffset1, depth1
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, flangeThick2, flangeDiam2, _
                       CptOffset2, depth2
    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, flangeThick3, flangeDiam3, _
                       CptOffset3, depth3
    RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, flangeThick4, flangeDiam4, _
                       CptOffset4, depth4

    Select Case lPartDataBasis
    Case Is <= 1 Or 190     ''Default' or 'Tee, Symmetrical, specified by face-to-center 1 190'
        parFace2toCenter = parFacetoCenter
        parFace3toCenter = parFacetoCenter

    Case 192    ''Tee, Asymmetrical, specified by face-to-center 1, and face-to-center 2'
        parFace2toCenter = arrayOfInputs(4)
        parFace3toCenter = parFacetoCenter

    Case 194    ''Tee, Asymmetrical, specified by face-to-center 1, and face-to-center 3'
        parFace2toCenter = parFacetoCenter
        parFace3toCenter = arrayOfInputs(5)

    Case 196    'Tee, Asymmetrical, specified by face-to-center 1, face-to-center 2, and face-to-center 3
        parFace2toCenter = arrayOfInputs(4)
        parFace3toCenter = arrayOfInputs(5)

    Case 825, 826, 827   ' Sanitary tee with 45 side opening above center
        'Sanitary tee, 45o right side outlet above center, specified by header face-to-face dimension, face-to-center 1, face-to-center 3, outlet face-to-center, and outlet offset from port 1
        'Sanitary tee, 45o left side outlet above center, specified by header face-to-face dimension, face-to-center 1, face-to-center 3, outlet face-to-center, and outlet offset from port 1
        'Sanitary tee, two 45o side outlets above center, specified by header face-to-face dimension, face-to-center 1, face-to-center 3, outlet face-to-center, and outlet offset from port 1
        parFacetoFace = arrayOfInputs(13)
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parOutFacetoCenter = arrayOfInputs(26)
        parOutOffsetPort1 = arrayOfInputs(16)
        parInsulationThickness = arrayOfInputs(3)

    Case 828    ' Sanitary Tee - Fig. 37 A861
        ' Sanitary tee, specified by header seat-to-seat dimension, seat-to-center 1, and seat-to-center 3
        parSeat1toCenter = arrayOfInputs(18)
        parSeat1toSeat2 = arrayOfInputs(17)
        parSeat3toCenter = arrayOfInputs(9)
        parInsulationThickness = arrayOfInputs(3)

    Case 829    ' Sanitary tee, 90o side outlet on center, specified by face-to-center 1 and face-to-center 3
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parInsulationThickness = arrayOfInputs(3)

        'Setting the default value of face to center dimension for side outlet(face 3 to center) equal to face 1 to center, when it is not specified
        If CmpDblEqual(parFace3toCenter, 0) Then parFace3toCenter = parFace1toCenter

    Case 830  ' Sanitary tee with cleanout plug on main
        ' Sanitary T branches with cleanout plug, specified by header face-to-face dimension, face-to-center 1, and face-to-center 3
        parFacetoFace = arrayOfInputs(13)
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parInsulationThickness = arrayOfInputs(3)

    Case 831
        ' Sanitary T branches with cleanout plug, specified by face-to-center 1, face-to-center 2, and face-to-center 3
        parFace1toCenter = arrayOfInputs(14)
        parFace2toCenter = arrayOfInputs(4)
        parFace3toCenter = arrayOfInputs(5)
        parInsulationThickness = arrayOfInputs(3)

    Case 832    ' Sanitary Tapped Tee Page 7 tyler Soil pipe
        'Sanitary tee, specified by header face-to-face dimension, face-to-center 1, and face-to-center 3
        parFacetoFace = arrayOfInputs(13)
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parInsulationThickness = arrayOfInputs(3)

    Case 833, 834, 835   ' Sanitary tee with side opening on center Page 7 tyler Soil pipe
        ' Sanitary tee, 90o right side outlet above center, specified by header face-to-face dimension, face-to-center 1, face-to-center 3, outlet face-to-center, and outlet offset from port 1'
        parFacetoFace = arrayOfInputs(13)
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parOutFacetoCenter = arrayOfInputs(26)
        parOutOffsetPort1 = arrayOfInputs(16)
        parInsulationThickness = arrayOfInputs(3)

    Case 836    ' Sanitary tapped tee, double horizontal Page 8 tyler pipe
        parFacetoFace = arrayOfInputs(13)
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parInsulationThickness = arrayOfInputs(3)

    Case 837    ' Sanitary tapped tee, double horizontal specified by Face 1 to center, Face 2 to center and Face 3 to Center
        parFace1toCenter = arrayOfInputs(14)
        parFace2toCenter = arrayOfInputs(4)
        parFace3toCenter = arrayOfInputs(5)
        parInsulationThickness = arrayOfInputs(3)

    Case 838    ' Sanitary tapped tee, double Vertical Page 8 tyler pipe
        parFacetoFace = arrayOfInputs(13)
        parFace3toCenter = arrayOfInputs(5)
        parOffBranch1toBranch2 = arrayOfInputs(21)
        parInsulationThickness = arrayOfInputs(3)

    Case 839    ' Sanitary tee, with square base, specified by face-to-center 1, face-to-center 3, center-to-base, base thickness, base width, and rib thickness
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(9)
        parBaseWidth = arrayOfInputs(23)
        parBasetoCenter = arrayOfInputs(19)
        parBaseThickness = arrayOfInputs(24)
        parRibThickness = arrayOfInputs(25)
        parInsulationThickness = arrayOfInputs(3)

    Case 841    ' T branches, Single, specified by header face-to-face dimension, face-to-center 1, and face-to-center 3.
        parFacetoFace = arrayOfInputs(13)
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parInsulationThickness = arrayOfInputs(3)

    Case 842    'T branches specified by face-to-center 1, face-to-center 2, and face-to-center 3.
        parFace1toCenter = arrayOfInputs(14)
        parFace2toCenter = arrayOfInputs(4)
        parFace3toCenter = arrayOfInputs(5)
        parInsulationThickness = arrayOfInputs(3)

    Case 843    'T branches specified by header seat-to-seat dimension, seat-to-center 1, and seat-to-center 3
        parSeat1toSeat2 = arrayOfInputs(17)
        parSeat1toCenter = arrayOfInputs(18)
        parSeat3toCenter = arrayOfInputs(9)
        parInsulationThickness = arrayOfInputs(3)

    Case 845    ' Sanitary T branches, tapped , Single, specified by header seat-to-seat dimension, seat-to-center 1, seat-to-center 3, branch radius, and tangent length 3
        parSeat1toSeat2 = arrayOfInputs(17)
        parSeat1toCenter = arrayOfInputs(18)
        parSeat3toCenter = arrayOfInputs(9)
        parBranchRadius = arrayOfInputs(22)
        parTangentLength3 = arrayOfInputs(12)
        parInsulationThickness = arrayOfInputs(3)

    Case 846    ' Sanitary T branches,tapped ,Single,  specified by header face-to-face dimension, face-to-center 1, face-to-center 3, branch radius, and tangent length 3
        parFacetoFace = arrayOfInputs(13)
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parBranchRadius = arrayOfInputs(22)
        parTangentLength3 = arrayOfInputs(12)
        parInsulationThickness = arrayOfInputs(3)

    Case 848    ' Sanitary T branches,tapped ,Single, specified by face-to-center 1, face-to-center 2, face-to-center 3, branch radius, and tangent length 3
        parFace1toCenter = arrayOfInputs(14)
        parFace2toCenter = arrayOfInputs(4)
        parFace3toCenter = arrayOfInputs(5)
        parBranchRadius = arrayOfInputs(22)
        parTangentLength3 = arrayOfInputs(12)
        parInsulationThickness = arrayOfInputs(3)

    Case 847    ' Sanitary T branches,tapped ,Single, specified by header face-to-face dimension, tangent length 1, branch radius, face-to-center 3, and tangent length 3
        parFacetoFace = arrayOfInputs(13)
        parTangentLength1 = arrayOfInputs(11)
        parBranchRadius = arrayOfInputs(22)
        parFace3toCenter = arrayOfInputs(5)
        parTangentLength3 = arrayOfInputs(12)
        parInsulationThickness = arrayOfInputs(3)

    Case 849    ' Sanitary T branches,tapped ,Single, specified by tangent length 1, branch radius, face-to-center 2, face-to-center 3, and tangent length 3
        parTangentLength1 = arrayOfInputs(11)
        parBranchRadius = arrayOfInputs(22)
        parFace2toCenter = arrayOfInputs(4)
        parFace3toCenter = arrayOfInputs(5)
        parTangentLength3 = arrayOfInputs(12)
        parInsulationThickness = arrayOfInputs(3)

    Case 850    ' Sanitary T branches,tapped ,Single, specified by header seat-to-seat dimension, seat-to-center 1, face-to-center 3, and branch radius
        parSeat1toSeat2 = arrayOfInputs(17)
        parSeat1toCenter = arrayOfInputs(18)
        parFace3toCenter = arrayOfInputs(5)
        parBranchRadius = arrayOfInputs(22)
        parInsulationThickness = arrayOfInputs(3)

    Case 844    ' T branches, single, specified by header seat-to-seat dimension, seat-to-center 1, and seat-to-center 3
        parSeat1toCenter = arrayOfInputs(18)
        parSeat2toCenter = arrayOfInputs(15)
        parSeat3toCenter = arrayOfInputs(9)
        parInsulationThickness = arrayOfInputs(3)

    Case 851    ' Sanitary tee, with round base, specified by face-to-center 1, face-to-center 3, center-to-base, base thickness, base width, and rib thickness
        parSeat1toCenter = arrayOfInputs(18)
        parSeat2toCenter = arrayOfInputs(15)
        parSeat3toCenter = arrayOfInputs(9)
        parBasetoCenter = arrayOfInputs(19)
        parBaseWidth = arrayOfInputs(23)
        parBaseThickness = arrayOfInputs(24)
        parRibThickness = arrayOfInputs(25)
        parInsulationThickness = arrayOfInputs(3)

    Case 852    ' Test tee, specified by header face-to-face dimension, face-to-center 1, and face-to-center 3
        parFacetoFace = arrayOfInputs(13)
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parInsulationThickness = arrayOfInputs(3)

    Case 853    ' Test tee, specified by header seat-to-seat dimension, seat-to-center 1, and seat-to-center 3
        parSeat1toCenter = arrayOfInputs(18)
        parSeat1toSeat2 = arrayOfInputs(17)
        parSeat3toCenter = arrayOfInputs(9)
        parInsulationThickness = arrayOfInputs(3)

    Case 840  ' Sanitary tee, with square base, specified by seat-to-center 1, seat-to-center 2, seat-to-center 3, center-to-base, base thickness, base width, and rib thickness
        parSeat1toCenter = arrayOfInputs(18)
        parSeat2toCenter = arrayOfInputs(15)
        parSeat3toCenter = arrayOfInputs(9)
        parBasetoCenter = arrayOfInputs(19)
        parBaseWidth = arrayOfInputs(23)
        parBaseThickness = arrayOfInputs(24)
        parRibThickness = arrayOfInputs(25)
        parInsulationThickness = arrayOfInputs(3)

    Case 854, 855, 856   ' 854 Sanitary tee, 90o right side outlet above center, specified by header seat-to-seat dimension, seat-to-center 1, seat-to-center 3, outlet seat-to-center, and outlet offset from port 1
        '               855 Sanitary tee, 90o left side outlet above center, specified by header seat-to-seat dimension, seat-to-center 1, seat-to-center 3, outlet seat-to-center, and outlet offset from port 1'
        '               856 Sanitary tee, two 90o side outlets above center, specified by header seat-to-seat dimension, seat-to-center 1, seat-to-center 3, outlet seat-to-center, and outlet offset from port 1
        parSeat1toCenter = arrayOfInputs(18)
        parSeat1toSeat2 = arrayOfInputs(17)
        parSeat3toCenter = arrayOfInputs(9)
        parOutOffsetPort1 = arrayOfInputs(16)
        parOutSeattoCenter = arrayOfInputs(27)
        parInsulationThickness = arrayOfInputs(3)

    Case 857, 858, 859
        ' (857) Sanitary tee, 90o right side outlet on center, specified by header face-to-face dimension, face-to-center 1, face-to-center 3, and outlet face-to-center
        ' (858) Sanitary tee, 90o left side outlet on center, specified by header face-to-face dimension, face-to-center 1, face-to-center 3, and outlet face-to-center
        ' (859) Sanitary tee, two 90o side outlets on center, specified by header face-to-face dimension, face-to-center 1, face-to-center 3, and outlet face-to-center
        parFacetoFace = arrayOfInputs(13)
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parOutFacetoCenter = arrayOfInputs(26)
        parInsulationThickness = arrayOfInputs(3)

    Case 860    ' Sanitary Tee specified by Face 1 to Center and Face 3 to Center
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parInsulationThickness = arrayOfInputs(3)

    Case 862    '  Tee, long sweep full flow

        Dim dFace1toCenter As Double
        Dim dBodyHeaderLength As Double

        parFacetoFace = arrayOfInputs(13)
        parFace3toCenter = arrayOfInputs(5)
        parBranchRadius = arrayOfInputs(22)
        parDiameter = arrayOfInputs(28)
        parInsulationThickness = arrayOfInputs(3)

        dFace1toCenter = parFacetoFace / 2

        'Setting the values of nozzle length
        If CmpDblEqual(flangeThick1, 0) Then
            dBodyHeaderLength = 0.95 * parFacetoFace
            dNozzleLength = (parFacetoFace - dBodyHeaderLength) / 2
        Else
            dBodyHeaderLength = parFacetoFace - 2 * flangeThick1
            dNozzleLength = flangeThick1
        End If

    Case 863    '  Tee, long sweep
        parFacetoFace = arrayOfInputs(13)
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parBranchRadius = arrayOfInputs(22)
        parDiameter = arrayOfInputs(28)
        parInsulationThickness = arrayOfInputs(3)

        dFace1toCenter = parFace1toCenter

        'Setting the default value of face 1 to center equal to face 3 to center
        If CmpDblEqual(dFace1toCenter, 0) Then dFace1toCenter = parFace3toCenter

        'Setting the values of nozzle length
        If CmpDblEqual(flangeThick1, 0) Then
            dBodyHeaderLength = 0.95 * parFacetoFace
            dNozzleLength = (parFacetoFace - dBodyHeaderLength) / 2
        Else
            dBodyHeaderLength = parFacetoFace - 2 * flangeThick1
            dNozzleLength = flangeThick1
        End If

    Case 864    '  Tee, long sweep with integral bull plug
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parBranchRadius = arrayOfInputs(22)
        parDiameter = arrayOfInputs(28)
        parLength = arrayOfInputs(29)
        parInsulationThickness = arrayOfInputs(3)

        'Setting the default value of face 1 to center equal to face 3 to center
        dFace1toCenter = parFace1toCenter
        If CmpDblEqual(dFace1toCenter, 0) Then dFace1toCenter = parFace3toCenter

        'Setting the values of nozzle length
        If CmpDblEqual(flangeThick1, 0) Then
            dBodyHeaderLength = 0.975 * parLength
            dNozzleLength = (parLength - dBodyHeaderLength)
        Else
            dBodyHeaderLength = parLength - flangeThick1
            dNozzleLength = flangeThick1
        End If

    Case 1043   'Vortex tee, specified by seat-to-seat, seat-to-center 1, and seat-to-center 3
        parSeattoSeat = arrayOfInputs(30)
        parSeat1toCenter = arrayOfInputs(18)
        parSeat3toCenter = arrayOfInputs(9)

        dSeat1toCenter = parSeat1toCenter
        dSeat2toCenter = parSeattoSeat - parSeat1toCenter
        dSeat3toCenter = parSeat3toCenter

    Case 1044   'Vortex tee, specified by seat-to-seat, seat-to-center 2, and seat-to-center 3
        parSeattoSeat = arrayOfInputs(30)
        parSeat2toCenter = arrayOfInputs(15)
        parSeat3toCenter = arrayOfInputs(9)

        dSeat1toCenter = parSeattoSeat - parSeat2toCenter
        dSeat2toCenter = parSeat2toCenter
        dSeat3toCenter = parSeat3toCenter

    Case 1045   'Vortex tee, specified by seat-to-center 1, seat-to-center 2, and seat-to-center 3
        parSeat1toCenter = arrayOfInputs(18)
        parSeat2toCenter = arrayOfInputs(15)
        parSeat3toCenter = arrayOfInputs(9)

        dSeat1toCenter = parSeat1toCenter
        dSeat2toCenter = parSeat2toCenter
        dSeat3toCenter = parSeat3toCenter

    End Select

    If UBound(arrayOfInputs) >= 6 Then
        parAngle = arrayOfInputs(6)
    Else
        parAngle = PI / 2    'To suppport the prior to V7 instances
    End If

    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oCenter As AutoMath.DPosition
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oVector As AutoMath.DVector
    Dim oAxisVect As AutoMath.DVector
    Dim oCircle As IngrGeom3D.Circle3d
    Dim oCP As AutoMath.DPosition
    Dim ocv As AutoMath.DVector

    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oCenter = New AutoMath.DPosition
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oVector = New DVector
    Set oAxisVect = New AutoMath.DVector
    Set oCircle = New IngrGeom3D.Circle3d
    Set oCP = New AutoMath.DPosition
    Set ocv = New AutoMath.DVector

    Dim ObjCurvedSurface As Object

    Select Case lPartDataBasis
    Case Is <= 1, 190, 192, 194, 196
        ' Insert your code for output 1(Insulated Body)
        Dim parInsulationDiameter As Double
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        oStPoint.Set -(parFacetoCenter - flangeThick), 0, 0
        oEnPoint.Set parFace2toCenter - flangeThick, 0, 0
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing

        ' Insert your code for output 2(Insulated Port1)
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set -parFacetoCenter, 0, 0
        oEnPoint.Set -(parFacetoCenter - flangeThick - parInsulationThickness), 0, 0

        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
        Set ObjInsulatedPort1 = Nothing

        ' Insert your code for output 3(Insulated Port2)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set parFace2toCenter, 0, 0
        oEnPoint.Set (parFace2toCenter - flangeThick - parInsulationThickness), 0, 0

        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
        Set ObjInsulatedPort2 = Nothing

        ' Insert your code for output 4(Insulated Branch)
        Dim BranchPoint As New AutoMath.DPosition
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        BranchPoint.Set (parFace3toCenter - flangeThick) * Cos(parAngle), _
                        (parFace3toCenter - flangeThick) * Sin(parAngle), CenterPos.z
        Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, CenterPos, BranchPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
        Set ObjInsulatedBranch = Nothing

        ' Insert your code for output 5 (Insulated Port3)
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), 0
        oEnPoint.Set (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(parAngle), _
                     (parFace3toCenter - flangeThick - parInsulationThickness) * Sin(parAngle), 0

        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
        Set ObjInsulatedPort3 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing

    Case 825, 826, 827     '45 Deg Side Opening above center

        ' Insert your code for output 1(45 Deg Side Opening above center)
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        If lPartDataBasis = 825 Then
            oStPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, 0
            oEnPoint.Set parOutOffsetPort1 - parFace1toCenter, parOutFacetoCenter * _
                                                               Cos(PI / 4), parOutFacetoCenter * Sin(PI / 4)
            Dim objInclinedCylinder1 As Object
            Set objInclinedCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, _
                                                     oEnPoint, pipeDiam + 2 * parInsulationThickness, True)

            '   Set the output
            m_OutputColl.AddOutput "IncCyl_", objInclinedCylinder1
            Set objInclinedCylinder1 = Nothing

        ElseIf lPartDataBasis = 826 Then
            ' Insert your code for output 2(45 Deg Side Opening above center 2)
            RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, cptoffset, depth

            oStPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, 0
            oEnPoint.Set parOutOffsetPort1 - parFace1toCenter, parOutFacetoCenter * _
                                                               Cos(PI / 4), -parOutFacetoCenter * Sin(PI / 4)
            Dim objInclinedCylinder2 As Object
            Set objInclinedCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, _
                                                     oEnPoint, pipeDiam + 2 * parInsulationThickness, True)

            '   Set the output
            m_OutputColl.AddOutput "IncCyl2", objInclinedCylinder2
            Set objInclinedCylinder2 = Nothing

        ElseIf lPartDataBasis = 827 Then
            oStPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, 0
            oEnPoint.Set parOutOffsetPort1 - parFace1toCenter, parOutFacetoCenter * _
                                                               Cos(PI / 4), parOutFacetoCenter * Sin(PI / 4)

            Set objInclinedCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, _
                                                     oEnPoint, pipeDiam + 2 * parInsulationThickness, True)

            '   Set the output
            m_OutputColl.AddOutput "IncCyl_", objInclinedCylinder1
            Set objInclinedCylinder1 = Nothing

            ' Insert your code for output 2(45 Deg Side Opening above center 2)
            RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, cptoffset, depth

            oStPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, 0
            oEnPoint.Set parOutOffsetPort1 - parFace1toCenter, parOutFacetoCenter * _
                                                               Cos(PI / 4), -parOutFacetoCenter * Sin(PI / 4)

            Set objInclinedCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, _
                                                     oEnPoint, pipeDiam + 2 * parInsulationThickness, True)

            '   Set the output
            m_OutputColl.AddOutput "IncCyl2", objInclinedCylinder2
            Set objInclinedCylinder2 = Nothing

        End If
        Dim dInsDia As Double
        'Create the Insulation of Port 1
        Dim dPortLen As Double
        Dim dfacetoface As Double
        dfacetoface = parFacetoFace
        dFace1toCenter = parFace1toCenter
        If Not CmpDblEqual(flangeThick1, 0) Then

            If CmpDblGreaterthan(dfacetoface, flangeThick1 + parInsulationThickness) Then
                dPortLen = flangeThick1 + parInsulationThickness
            Else
                dPortLen = dfacetoface
            End If

            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Dim objInsPort As Object
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        'Create the Insulation for Body between Nozzle 1 and Nozzle 2

        dInsDia = pipeDiam1 + 2 * parInsulationThickness
        oStPoint.Set -dFace1toCenter + flangeThick1, 0, 0
        oEnPoint.Set dfacetoface - dFace1toCenter, 0, 0
        Dim ObjMainRunInsul As Object
        Set ObjMainRunInsul = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        m_OutputColl.AddOutput "ObjIns_", ObjMainRunInsul
        Set ObjMainRunInsul = Nothing
        ' Insulation for Port 2

        If Not CmpDblEqual(flangeThick2, 0) Then
            oStPoint.Set dfacetoface - dFace1toCenter, 0, 0
            oEnPoint.Set dfacetoface - dFace1toCenter - flangeThick2, 0, 0
            dPortLen = flangeThick2 + parInsulationThickness
            dInsDia = flangeDiam2 + parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        Dim dface3tocenter As Double
        dface3tocenter = parFace3toCenter
        If Not CmpDblEqual(flangeThick3, 0) Then
            dInsDia = flangeDiam3 + 2 * parInsulationThickness
            oStPoint.Set dface3tocenter * Cos(parAngle), dface3tocenter * Sin(parAngle), 0
            oEnPoint.Set (dface3tocenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (dface3tocenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), 0
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        ' Part body neat port 3
        oStPoint.Set (parFace3toCenter + cptoffset - depth) * Cos(parAngle), (parFace3toCenter + cptoffset - depth) * Sin(parAngle), 0
        oEnPoint.Set 0, 0, 0

        dInsDia = pipeDiam3 + 2 * parInsulationThickness
        Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort
        Set objInsPort = Nothing
        ' Insulation for Port 4
        dInsDia = flangeDiam4
        If lPartDataBasis = 825 Then
            If Not flangeThick4 = 0 Then
                oStPoint.Set parOutOffsetPort1 - parFace3toCenter, _
                             parOutFacetoCenter * Cos(PI / 4), _
                             parOutFacetoCenter * Sin(PI / 4)
                oEnPoint.Set parOutOffsetPort1 - parFace3toCenter, _
                             (parOutFacetoCenter - flangeThick4) * Cos(PI / 4), _
                             (parOutFacetoCenter - flangeThick4) * Sin(PI / 4)


                Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
                '   Set the output
                m_OutputColl.AddOutput "ObjIns_", objInsPort
                Set objInsPort = Nothing
            End If
        ElseIf lPartDataBasis = 826 Then
            If Not flangeThick4 = 0 Then
                oStPoint.Set parOutOffsetPort1 - parFace3toCenter, _
                             parOutFacetoCenter * Cos(PI / 4), _
                             -parOutFacetoCenter * Sin(PI / 4)
                oEnPoint.Set parOutOffsetPort1 - parFace3toCenter, _
                             (parOutFacetoCenter - flangeThick4) * Cos(PI / 4), _
                             -(parOutFacetoCenter - flangeThick4) * Sin(PI / 4)


                Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
                '   Set the output
                m_OutputColl.AddOutput "ObjIns_", objInsPort
                Set objInsPort = Nothing
            End If
        ElseIf lPartDataBasis = 827 Then
            If Not flangeThick4 = 0 Then
                oStPoint.Set parOutOffsetPort1 - parFace3toCenter, _
                             parOutFacetoCenter * Cos(PI / 4), _
                             parOutFacetoCenter * Sin(PI / 4)
                oEnPoint.Set parOutOffsetPort1 - parFace3toCenter, _
                             (parOutFacetoCenter - flangeThick4) * Cos(PI / 4), _
                             (parOutFacetoCenter - flangeThick4) * Sin(PI / 4)


                Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
                '   Set the output
                m_OutputColl.AddOutput "ObjIns_", objInsPort
                Set objInsPort = Nothing
            End If
            ' Place
            If Not flangeThick5 = 0 Then
                oStPoint.Set parOutOffsetPort1 - parFace3toCenter, _
                             parOutFacetoCenter * Cos(PI / 4), _
                             -parOutFacetoCenter * Sin(PI / 4)
                oEnPoint.Set parOutOffsetPort1 - parFace3toCenter, _
                             (parOutFacetoCenter - flangeThick4) * Cos(PI / 4), _
                             -(parOutFacetoCenter - flangeThick4) * Sin(PI / 4)

                dInsDia = flangeDiam5
                Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
                '   Set the output
                m_OutputColl.AddOutput "ObjIns_", objInsPort
                Set objInsPort = Nothing
            End If
        End If


    Case 828     ' Sanitary Tee - Fig. 36 A861

        If CmpDblGreaterthan(pipeDiam1, pipeDiam2) Then pipeDiam = pipeDiam2
        Dim dInsulationRadius As Double
        Dim dElbowRadius As Double
        Dim dStartingPoint As Double
        dInsulationRadius = pipeDiam2 / 2 + parInsulationThickness
        If CmpDblGreaterthan(parAngle, PI / 2) Then
            dElbowRadius = parSeat3toCenter
            dStartingPoint = parSeat3toCenter
        Else
            dElbowRadius = parSeat3toCenter + depth
            dStartingPoint = parSeat3toCenter + depth
        End If
        If CmpDblLessThan(dElbowRadius, dInsulationRadius) Then
            dInsulationRadius = dElbowRadius - 0.0001
        End If
        Set oCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, dInsulationRadius)

        '   Revolve the circle about the Z axis passing by the origin

        oCP.Set -dElbowRadius, dElbowRadius, 0
        ocv.Set 0, 0, 1

        Set ObjCurvedSurface = PlaceRevolution(m_OutputColl, oCircle, ocv, _
                                               oCP, parAngle, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur_", ObjCurvedSurface
        Set ObjCurvedSurface = Nothing
        Set oCP = Nothing
        Set ocv = Nothing
        dfacetoface = parSeat1toSeat2 + depth1 + depth2 - CptOffset1 - CptOffset2
        dFace1toCenter = parSeat1toCenter + depth - CptOffset1
        If Not CmpDblEqual(flangeThick1, 0) Then

            If CmpDblGreaterthan(dfacetoface, flangeThick1 + parInsulationThickness) Then
                dPortLen = flangeThick1 + parInsulationThickness
            Else
                dPortLen = dfacetoface
            End If

            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        'Create the Insulation for Body between Nozzle 1 and Nozzle 2

        dInsDia = pipeDiam1 + 2 * parInsulationThickness
        oStPoint.Set -dFace1toCenter + flangeThick1, 0, 0
        oEnPoint.Set dfacetoface - dFace1toCenter, 0, 0
        Set ObjMainRunInsul = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        m_OutputColl.AddOutput "ObjIns_", ObjMainRunInsul
        Set ObjMainRunInsul = Nothing
        ' Insulation for Port 2

        If Not CmpDblEqual(flangeThick2, 0) Then
            oStPoint.Set dfacetoface - dFace1toCenter, 0, 0
            oEnPoint.Set dfacetoface - dFace1toCenter - flangeThick2, 0, 0
            dPortLen = flangeThick2 + parInsulationThickness
            dInsDia = flangeDiam2 + parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

        dface3tocenter = parSeat3toCenter + depth3
        If Not CmpDblEqual(flangeThick3, 0) Then
            dInsDia = flangeDiam3 + 2 * parInsulationThickness
            oStPoint.Set dface3tocenter * Cos(parAngle), dface3tocenter * Sin(parAngle), 0
            oEnPoint.Set (dface3tocenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (dface3tocenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), 0
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

    Case 829    ' Sanitary tee, 90o side outlet on center, specified by face-to-center 1 and face-to-center 3

        ' Insulation for Nozzle 1
        dFace1toCenter = parFace1toCenter
        dfacetoface = 2 * parFace1toCenter
        If Not CmpDblEqual(flangeThick1, 0) Then

            If CmpDblGreaterthan(dfacetoface, flangeThick1 + parInsulationThickness) Then
                dPortLen = flangeThick1 + parInsulationThickness
            Else
                dPortLen = dfacetoface
            End If

            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        'Create the Insulation for Body between Nozzle 1 and Nozzle 2

        dInsDia = pipeDiam1 + 2 * parInsulationThickness
        oStPoint.Set -dFace1toCenter + flangeThick1, 0, 0
        oEnPoint.Set dfacetoface - dFace1toCenter, 0, 0
        Set ObjMainRunInsul = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        m_OutputColl.AddOutput "ObjIns_", ObjMainRunInsul
        Set ObjMainRunInsul = Nothing
        ' Insulation for Port 2
        If Not flangeThick2 = 0 Then

            dPortLen = flangeThick2 + parInsulationThickness
            If CmpDblGreaterthan(dPortLen, dfacetoface) Then dPortLen = dfacetoface

            oStPoint.Set dfacetoface - dFace1toCenter, 0, 0
            oEnPoint.Set dfacetoface - dFace1toCenter - dPortLen, 0, 0

            dInsDia = flangeDiam2 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

        dface3tocenter = parFace3toCenter
        If Not CmpDblEqual(flangeThick3, 0) Then
            dInsDia = flangeDiam3 + 2 * parInsulationThickness
            oStPoint.Set dface3tocenter * Cos(parAngle), dface3tocenter * Sin(parAngle), 0
            oEnPoint.Set (dface3tocenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (dface3tocenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), 0
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

        If Not CmpDblEqual(flangeThick4, 0) Then
            dInsDia = flangeDiam4 + 2 * parInsulationThickness
            oStPoint.Set 0, 0, -parFace3toCenter
            oEnPoint.Set 0, 0, -(parFace3toCenter - flangeThick4 - parInsulationThickness)
            dInsDia = flangeDiam4 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort
            Set objInsPort = Nothing
        End If

        ' Insulation for the Body (Port 3)
        oStPoint.Set (parFace3toCenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (parFace3toCenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), 0
        oEnPoint.Set 0, 0, 0
        dInsDia = pipeDiam3 + 2 * parInsulationThickness
        Dim InsCylinder1 As Object
        Set InsCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), InsCylinder1
        Set InsCylinder1 = Nothing

        ' Insulation for the Body (Port 4)

        oStPoint.Set 0, 0, -(parFace3toCenter - flangeThick4 - parInsulationThickness)
        oEnPoint.Set 0, 0, 0
        dInsDia = pipeDiam4 + 2 * parInsulationThickness
        Dim InsCylinder2 As Object
        Set InsCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "ObjCyl2_", InsCylinder2
        Set InsCylinder2 = Nothing

    Case 830, 831      'T Branches, Cleanout Plug on Main
        If lPartDataBasis = 830 Then
            parFace2toCenter = parFacetoFace - parFace1toCenter
        End If

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        ' Insert your code for the Cylinder 2(Header)
        oStPoint.Set 0.7 * parFace2toCenter, 0, 0
        oEnPoint.Set parFace2toCenter, 0, 0
        Dim ObjCylinder1 As Object
        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                         1.4 * pipeDiam + 2 * parInsulationThickness, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder1
        Set ObjCylinder1 = Nothing
        parSeat3toCenter = parFace3toCenter - depth
        dInsulationRadius = pipeDiam2 / 2 + parInsulationThickness
        If CmpDblGreaterthan(parAngle, PI / 2) Then
            dElbowRadius = parSeat3toCenter
            dStartingPoint = parSeat3toCenter
        Else
            dElbowRadius = parSeat3toCenter + depth
            dStartingPoint = parSeat3toCenter + depth
        End If
        If CmpDblLessThan(dElbowRadius, dInsulationRadius) Then
            dInsulationRadius = dElbowRadius - 0.0001
        End If
        Set oCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, dInsulationRadius)

        '   Revolve the circle about the Z axis passing by the origin

        oCP.Set -dElbowRadius, dElbowRadius, 0
        ocv.Set 0, 0, 1

        Set ObjCurvedSurface = PlaceRevolution(m_OutputColl, oCircle, ocv, _
                                               oCP, parAngle, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur_", ObjCurvedSurface
        Set ObjCurvedSurface = Nothing
        Set oCP = Nothing
        Set ocv = Nothing
        If lPartDataBasis = 830 Then
            dfacetoface = parFacetoFace
        ElseIf lPartDataBasis = 831 Then
            dfacetoface = parFace1toCenter + parFace2toCenter
        End If
        dFace1toCenter = parFace1toCenter
        If Not CmpDblEqual(flangeThick1, 0) Then
            If CmpDblGreaterthan(dfacetoface, flangeThick1 + parInsulationThickness) Then
                dPortLen = flangeThick1 + parInsulationThickness
            Else
                dPortLen = dfacetoface
            End If

            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        'Create the Insulation for Body between Nozzle 1 and Nozzle 2

        dInsDia = pipeDiam1 + 2 * parInsulationThickness
        oStPoint.Set -dFace1toCenter + flangeThick1, 0, 0
        oEnPoint.Set dfacetoface - dFace1toCenter, 0, 0
        Set ObjMainRunInsul = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        m_OutputColl.AddOutput "ObjIns_", ObjMainRunInsul
        Set ObjMainRunInsul = Nothing
        ' Insulation for Port 2

        dface3tocenter = parFace3toCenter
        If Not CmpDblEqual(flangeThick2, 0) Then
            dInsDia = flangeDiam2 + 2 * parInsulationThickness
            oStPoint.Set dface3tocenter * Cos(parAngle), dface3tocenter * Sin(parAngle), 0
            oEnPoint.Set (dface3tocenter - flangeThick2 - parInsulationThickness) * Cos(parAngle), (dface3tocenter - flangeThick2 - parInsulationThickness) * Sin(parAngle), 0
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

    Case 832    ' Sanitary Tapped Tee
        If lPartDataBasis = 826 Then
            parFacetoFace = parFace1toCenter + parFace2toCenter
        End If
        If CmpDblGreaterthanOrEqualTo(pipeDiam3, pipeDiam1) Then pipeDiam3 = pipeDiam1
        parSeat3toCenter = parFace3toCenter - depth
        dInsulationRadius = pipeDiam3 / 2 + parInsulationThickness
        If CmpDblGreaterthan(parAngle, PI / 2) Then
            dElbowRadius = parSeat3toCenter
            dStartingPoint = parSeat3toCenter
        Else
            dElbowRadius = parSeat3toCenter + depth
            dStartingPoint = parSeat3toCenter + depth
        End If
        dInsulationRadius = pipeDiam3 + 2 * parInsulationThickness
        If CmpDblLessThan(dElbowRadius, dInsulationRadius) Then
            dInsulationRadius = dElbowRadius - 0.0001
        End If

        Set oCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, dInsulationRadius)

        '   Revolve the circle about the Z axis passing by the origin

        oCP.Set -dElbowRadius, dElbowRadius, 0
        ocv.Set 0, 0, 1

        Set ObjCurvedSurface = PlaceRevolution(m_OutputColl, oCircle, ocv, _
                                               oCP, parAngle, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur_", ObjCurvedSurface
        Set ObjCurvedSurface = Nothing
        Set oCP = Nothing
        Set ocv = Nothing
        dFace1toCenter = parFace1toCenter
        dfacetoface = parFacetoFace
        If Not CmpDblEqual(flangeThick1, 0) Then    '832

            If CmpDblGreaterthan(dfacetoface, flangeThick1 + parInsulationThickness) Then
                dPortLen = flangeThick1 + parInsulationThickness
            Else
                dPortLen = dfacetoface
            End If

            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        'Create the Insulation for Body between Nozzle 1 and Nozzle 2

        dInsDia = pipeDiam1 + 2 * parInsulationThickness
        oStPoint.Set -dFace1toCenter + flangeThick1, 0, 0
        oEnPoint.Set dfacetoface - dFace1toCenter, 0, 0
        Set ObjMainRunInsul = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        m_OutputColl.AddOutput "ObjIns_", ObjMainRunInsul
        Set ObjMainRunInsul = Nothing
        ' Insulation for Port 2

        If Not CmpDblEqual(flangeThick2, 0) Then
            oStPoint.Set dfacetoface - dFace1toCenter, 0, 0
            oEnPoint.Set dfacetoface - dFace1toCenter - flangeThick2, 0, 0
            dPortLen = flangeThick2 + parInsulationThickness
            dInsDia = flangeDiam2 + parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

        dface3tocenter = parFace3toCenter
        If Not CmpDblEqual(flangeThick3, 0) Then
            dInsDia = flangeDiam3 + 2 * parInsulationThickness
            oStPoint.Set dface3tocenter * Cos(parAngle), dface3tocenter * Sin(parAngle), 0
            oEnPoint.Set (dface3tocenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (dface3tocenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), 0
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

    Case 833, 834, 835, 854, 855, 856   ' Sanitary tapped tee with tap on side
        If lPartDataBasis = 854 Or lPartDataBasis = 855 Or lPartDataBasis = 856 Then
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, cptoffset, depth
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, CptOffset2, depth2
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, CptOffset3, depth3
            parFace1toCenter = parSeat1toCenter + depth - cptoffset
            parFacetoFace = parSeat1toSeat2 + depth - cptoffset + depth2 - CptOffset2
            parFace3toCenter = parSeat3toCenter + depth3 - CptOffset3
        End If
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        '   Revolve the circle about the Z axis passing by the origin for the branch
        dInsDia = pipeDiam1 + 2 * parInsulationThickness
        If CmpDblGreaterthanOrEqualTo(pipeDiam2, pipeDiam) Then pipeDiam2 = pipeDiam
        parSeat3toCenter = parFace3toCenter - depth3
        dInsulationRadius = pipeDiam2 / 2 + parInsulationThickness
        If CmpDblGreaterthan(parAngle, PI / 2) Then
            dElbowRadius = parSeat3toCenter
            dStartingPoint = parSeat3toCenter
        Else
            dElbowRadius = parSeat3toCenter + depth
            dStartingPoint = parSeat3toCenter + depth
        End If
        If CmpDblLessThan(dElbowRadius, dInsulationRadius) Then
            dInsulationRadius = dElbowRadius - 0.0001
        End If
        Set oCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, dInsulationRadius)

        '   Revolve the circle about the Z axis passing by the origin

        oCP.Set -dElbowRadius, dElbowRadius, 0
        ocv.Set 0, 0, 1

        Set ObjCurvedSurface = PlaceRevolution(m_OutputColl, oCircle, ocv, _
                                               oCP, parAngle, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur_", ObjCurvedSurface
        Set ObjCurvedSurface = Nothing
        Set oCP = Nothing
        Set ocv = Nothing

        ' Insert Your Code for Output 3 ( Second Branch)

        Dim oEllipse As IngrGeom3D.Ellipse3d
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, _
                           flangeThick, flangeDiam, cptoffset, depth
        Dim dMMRatio As Double
        dMMRatio = 0.5 + parInsulationThickness / 0.7 + parInsulationThickness
        If lPartDataBasis = 833 Or lPartDataBasis = 854 Then
            Set oEllipse = _
            oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                   parOutOffsetPort1 - parFace1toCenter - 0.2 * pipeDiam4, 0, 0, 0, 0, -1, parOutOffsetPort1 - parFace1toCenter + 0.7 * pipeDiam4 + parInsulationThickness, 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, parOutOffsetPort1 - parFace1toCenter, 0, _
                                                              0.7 * parOutFacetoCenter, 0, 0, 1, pipeDiam4 / 2 + parInsulationThickness)


            Dim oBranch2 As Object
            Set oBranch2 = _
            oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                        oEllipse, oCircle, True)

            '   Set the output
            m_OutputColl.AddOutput "ObjCur_", oBranch2
            Set oBranch2 = Nothing
            Set oCircle = Nothing
        ElseIf lPartDataBasis = 834 Or lPartDataBasis = 855 Then
            Set oEllipse = _
            oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                   parOutOffsetPort1 - parFace1toCenter - 0.2 * pipeDiam4, 0, 0, 0, 0, 1, -(parOutOffsetPort1 - parFace1toCenter + 0.7 * pipeDiam4) + parInsulationThickness, 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, parOutOffsetPort1 - parFace1toCenter, 0, _
                                                              -0.7 * parOutFacetoCenter, 0, 0, -1, pipeDiam4 / 2 + parInsulationThickness)

            Set oBranch2 = _
            oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                        oEllipse, oCircle, True)

            '   Set the output
            m_OutputColl.AddOutput "ObjCur_", oBranch2
            Set oBranch2 = Nothing
            Set oCircle = Nothing
            Set oEllipse = Nothing
        ElseIf lPartDataBasis = 835 Or lPartDataBasis = 856 Then
            Set oEllipse = _
            oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                   parOutOffsetPort1 - parFace1toCenter - 0.2 * pipeDiam4, 0, 0, 0, 0, -1, parOutOffsetPort1 - parFace1toCenter + 0.7 * pipeDiam4 + parInsulationThickness, 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, parOutOffsetPort1 - parFace1toCenter, 0, _
                                                              0.7 * parOutFacetoCenter, 0, 0, 1, pipeDiam4 / 2 + parInsulationThickness)

            Set oBranch2 = _
            oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                        oEllipse, oCircle, True)

            '   Set the output
            m_OutputColl.AddOutput "ObjCur_", oBranch2
            Set oBranch2 = Nothing
            Set oCircle = Nothing
            Set oEllipse = Nothing

            Set oEllipse = _
            oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                   parOutOffsetPort1 - parFace1toCenter - 0.2 * pipeDiam4, 0, 0, 0, 0, 1, -(parOutOffsetPort1 - parFace1toCenter + 0.7 * pipeDiam4 + parInsulationThickness), 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, parOutOffsetPort1 - parFace1toCenter, 0, _
                                                              -0.7 * parOutFacetoCenter, 0, 0, -1, pipeDiam4 / 2 + parInsulationThickness)

            Dim oBranch3 As Object
            Set oBranch3 = _
            oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                        oEllipse, oCircle, True)

            '   Set the output
            m_OutputColl.AddOutput "ObjCur_", oBranch3
            Set oBranch2 = Nothing
            Set oCircle = Nothing
            Set oEllipse = Nothing

        End If
        dfacetoface = parFacetoFace
        dFace1toCenter = parFace1toCenter
        If Not CmpDblEqual(flangeThick1, 0) Then    '832

            If CmpDblGreaterthan(dfacetoface, flangeThick1 + parInsulationThickness) Then
                dPortLen = flangeThick1 + parInsulationThickness
            Else
                dPortLen = dfacetoface
            End If

            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        'Create the Insulation for Body between Nozzle 1 and Nozzle 2

        dInsDia = pipeDiam1 + 2 * parInsulationThickness
        oStPoint.Set -dFace1toCenter + flangeThick1, 0, 0
        oEnPoint.Set dfacetoface - dFace1toCenter, 0, 0
        Set ObjMainRunInsul = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        m_OutputColl.AddOutput "ObjIns_", ObjMainRunInsul
        Set ObjMainRunInsul = Nothing
        ' Insulation for Port 2

        If Not CmpDblEqual(flangeThick2, 0) Then
            oStPoint.Set dfacetoface - dFace1toCenter, 0, 0
            oEnPoint.Set dfacetoface - dFace1toCenter - flangeThick2, 0, 0
            dPortLen = flangeThick2 + parInsulationThickness
            dInsDia = flangeDiam2 + parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

        dface3tocenter = parFace3toCenter
        If Not CmpDblEqual(flangeThick3, 0) Then
            dInsDia = flangeDiam3 + 2 * parInsulationThickness
            oStPoint.Set dface3tocenter * Cos(parAngle), dface3tocenter * Sin(parAngle), 0
            oEnPoint.Set (dface3tocenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (dface3tocenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), 0
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        If Not CmpDblEqual(flangeThick4, 0) Then
            dInsDia = flangeDiam4 + 2 * parInsulationThickness
            oStPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, parOutFacetoCenter + flangeThick4
            oEnPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, parOutFacetoCenter
            dInsDia = flangeThick4 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort
            Set objInsPort = Nothing
        End If

        If Not CmpDblEqual(flangeThick5, 0) Then
            dInsDia = flangeDiam5 + 2 * parInsulationThickness
            oStPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, -(parOutFacetoCenter + flangeThick4)
            oEnPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, -parOutFacetoCenter
            dInsDia = flangeThick5 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort
            Set objInsPort = Nothing
        End If
        ' Insert code for body (Port 4)
        If lPartDataBasis = 833 Or lPartDataBasis = 854 Then
            oStPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, parOutFacetoCenter + flangeThick4
            oEnPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, 0.7 * parOutFacetoCenter
        ElseIf lPartDataBasis = 834 Or lPartDataBasis = 835 Or lPartDataBasis = 855 Or lPartDataBasis = 856 Then
            oStPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, -(parOutFacetoCenter + flangeThick4)
            oEnPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, -0.7 * parOutFacetoCenter
        End If
        dInsDia = pipeDiam4 + 2 * parInsulationThickness
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing

        If lPartDataBasis = 835 Or lPartDataBasis = 856 Then
            oStPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, parOutFacetoCenter + flangeThick4
            oEnPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, 0.7 * parOutFacetoCenter

            Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
            m_OutputColl.AddOutput "ObjInsBody_", ObjInsulatedBody
            Set ObjInsulatedBody = Nothing
        End If

    Case 836, 837    ' Horizontal tapped Tee

        If lPartDataBasis = 836 Then
            parFace2toCenter = parFacetoFace - parFace1toCenter
        End If
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, _
                           flangeThick, flangeDiam, cptoffset, depth

        ' Insert your code for the Horizontal Branch 1
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        If CmpDblEqual(flangeDiam, 0) Then flangeDiam = pipeDiam
        oStPoint.Set 0, -pipeDiam1 / 2, -flangeDiam / 2
        oEnPoint.Set parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), -flangeDiam / 2

        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                         pipeDiam + 2 * parInsulationThickness, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder1
        Set ObjCylinder1 = Nothing

        ' Insert your code for the Horizontal Branch 2
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        If CmpDblEqual(flangeDiam, 0) Then flangeDiam = pipeDiam
        oStPoint.Set 0, -pipeDiam1 / 2, flangeDiam / 2
        oEnPoint.Set parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), flangeDiam / 2
        Dim ObjCylinder2 As Object
        Set ObjCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                         pipeDiam + 2 * parInsulationThickness, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder2
        Set ObjCylinder2 = Nothing
        ' Insulation for the ports

        dFace1toCenter = parFace1toCenter
        dfacetoface = parFacetoFace
        If lPartDataBasis = 837 Then dfacetoface = parFace1toCenter + parFace2toCenter
        If Not CmpDblEqual(flangeThick1, 0) Then

            If CmpDblGreaterthan(dfacetoface, flangeThick1 + parInsulationThickness) Then
                dPortLen = flangeThick1 + parInsulationThickness
            Else
                dPortLen = dfacetoface
            End If

            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        'Create the Insulation for Body between Nozzle 1 and Nozzle 2

        dInsDia = pipeDiam1 + 2 * parInsulationThickness
        oStPoint.Set -dFace1toCenter + flangeThick1, 0, 0
        oEnPoint.Set dfacetoface - dFace1toCenter, 0, 0
        Set ObjMainRunInsul = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        m_OutputColl.AddOutput "ObjIns_", ObjMainRunInsul
        Set ObjMainRunInsul = Nothing
        ' Insulation for Port 2

        If Not CmpDblEqual(flangeThick2, 0) Then
            oStPoint.Set dfacetoface - dFace1toCenter, 0, 0
            oEnPoint.Set dfacetoface - dFace1toCenter - flangeThick2, 0, 0
            dPortLen = flangeThick2 + parInsulationThickness
            dInsDia = flangeDiam2 + parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        If Not CmpDblEqual(flangeThick3, 0) Then
            dInsDia = flangeDiam3 + 2 * parInsulationThickness
            oStPoint.Set parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), flangeDiam3 / 2
            oEnPoint.Set (parFace3toCenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (parFace3toCenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), flangeDiam3 / 2
            dInsDia = flangeThick3 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort
            Set objInsPort = Nothing
        End If
        If Not CmpDblEqual(flangeThick4, 0) Then
            dInsDia = flangeDiam4 + 2 * parInsulationThickness
            oStPoint.Set parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), flangeDiam4 / 2
            oEnPoint.Set (parFace3toCenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (parFace3toCenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), flangeDiam4 / 2
            dInsDia = flangeThick3 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort
            Set objInsPort = Nothing
        End If

    Case 838  ' Vertical tapped Tee

        ' Insert your code for the Horizontal Branch 1
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oStPoint.Set -parOffBranch1toBranch2 / 2, 0, 0
        oEnPoint.Set -parOffBranch1toBranch2 / 2 + parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), 0

        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                         pipeDiam + 2 * parInsulationThickness, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder1
        Set ObjCylinder1 = Nothing

        ' Insert your code for the Horizontal Branch 2
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        If CmpDblEqual(flangeDiam, 0) Then flangeDiam = pipeDiam
        oStPoint.Set parOffBranch1toBranch2 / 2, 0, 0
        oEnPoint.Set parOffBranch1toBranch2 / 2 + parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), 0

        Set ObjCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                         pipeDiam + 2 * parInsulationThickness, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder2
        Set ObjCylinder2 = Nothing

        dFace1toCenter = parFacetoFace / 2
        dfacetoface = parFacetoFace
        If Not CmpDblEqual(flangeThick1, 0) Then

            If CmpDblGreaterthan(dfacetoface, flangeThick1 + parInsulationThickness) Then
                dPortLen = flangeThick1 + parInsulationThickness
            Else
                dPortLen = dfacetoface
            End If

            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        'Create the Insulation for Body between Nozzle 1 and Nozzle 2

        dInsDia = pipeDiam1 + 2 * parInsulationThickness
        oStPoint.Set -dFace1toCenter + flangeThick1, 0, 0
        oEnPoint.Set dfacetoface - dFace1toCenter, 0, 0
        Set ObjMainRunInsul = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        m_OutputColl.AddOutput "ObjIns_", ObjMainRunInsul
        Set ObjMainRunInsul = Nothing
        ' Insulation for Port 2

        If Not CmpDblEqual(flangeThick2, 0) Then
            oStPoint.Set dfacetoface - dFace1toCenter, 0, 0
            oEnPoint.Set dfacetoface - dFace1toCenter - flangeThick2, 0, 0
            dPortLen = flangeThick2 + parInsulationThickness
            dInsDia = flangeDiam2 + parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        If Not CmpDblEqual(flangeThick3, 0) Then
            dInsDia = flangeDiam3 + 2 * parInsulationThickness
            oStPoint.Set -parOffBranch1toBranch2 / 2 + parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), 0
            oEnPoint.Set -parOffBranch1toBranch2 / 2 + (parFace3toCenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (parFace3toCenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), 0
            dInsDia = flangeThick3 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort
            Set objInsPort = Nothing
        End If
        If Not CmpDblEqual(flangeThick4, 0) Then
            dInsDia = flangeDiam4 + 2 * parInsulationThickness
            oStPoint.Set parOffBranch1toBranch2 / 2 + parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), 0
            oEnPoint.Set parOffBranch1toBranch2 / 2 + (parFace3toCenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (parFace3toCenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), 0
            dInsDia = flangeThick4 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objInsPort
            Set objInsPort = Nothing
        End If

    Case 841    ' T branches, Single, specified by header face-to-face dimension, face-to-center 1, and face-to-center 3

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        oStPoint.Set -(parFace1toCenter - flangeThick), 0, 0
        oEnPoint.Set parFacetoFace - parFace1toCenter - flangeThick, 0, 0
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing

        ' Insert your code for output 2(Insulated Port1)
        If CmpDblEqual(flangeThick1, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set -parFace1toCenter, 0, 0
        oEnPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness), 0, 0

        If CmpDblGreaterthan(flangeDiam1, pipeDiam1) Then
            parInsulationDiameter = flangeDiam1 + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam1 + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
        Set ObjInsulatedPort1 = Nothing

        ' Insert your code for output 3(Insulated Port2)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set parFacetoFace - parFace1toCenter, 0, 0
        oEnPoint.Set (parFacetoFace - parFace1toCenter - flangeThick - parInsulationThickness), 0, 0

        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
        Set ObjInsulatedPort2 = Nothing

        ' Insert your code for output 4(Insulated Branch)

        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        BranchPoint.Set (parFace3toCenter - flangeThick) * Cos(parAngle), _
                        (parFace3toCenter - flangeThick) * Sin(parAngle), CenterPos.z
        Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, CenterPos, BranchPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
        Set ObjInsulatedBranch = Nothing

        ' Insert your code for output 5 (Insulated Port3)
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), 0
        oEnPoint.Set (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(parAngle), _
                     (parFace3toCenter - flangeThick - parInsulationThickness) * Sin(parAngle), 0

        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
        Set ObjInsulatedPort3 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing

    Case 842, 860  'T branches, single, specified by face-to-center 1, face-to-center 2, and face-to-center 3.

        If lPartDataBasis = 860 Then parFace2toCenter = parFace1toCenter
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        oStPoint.Set -(parFace1toCenter - flangeThick), 0, 0
        oEnPoint.Set parFace2toCenter - flangeThick, 0, 0
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing

        ' Insert your code for output 2(Insulated Port1)
        If CmpDblEqual(flangeThick1, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set -parFace1toCenter, 0, 0
        oEnPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness), 0, 0

        If CmpDblGreaterthan(flangeDiam1, pipeDiam1) Then
            parInsulationDiameter = flangeDiam1 + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam1 + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
        Set ObjInsulatedPort1 = Nothing

        ' Insert your code for output 3(Insulated Port2)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set parFace2toCenter, 0, 0
        oEnPoint.Set (parFace2toCenter - flangeThick - parInsulationThickness), 0, 0

        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
        Set ObjInsulatedPort2 = Nothing

        ' Insert your code for output 4(Insulated Branch)

        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        BranchPoint.Set (parFace3toCenter - flangeThick) * Cos(parAngle), _
                        (parFace3toCenter - flangeThick) * Sin(parAngle), CenterPos.z
        Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, CenterPos, BranchPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
        Set ObjInsulatedBranch = Nothing

        ' Insert your code for output 5 (Insulated Port3)
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), 0
        oEnPoint.Set (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(parAngle), _
                     (parFace3toCenter - flangeThick - parInsulationThickness) * Sin(parAngle), 0

        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
        Set ObjInsulatedPort3 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing

    Case 843    'T branches specified by header seat-to-seat dimension, seat-to-center 1, and seat-to-center 3

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        oStPoint.Set -parSeat1toCenter, 0, 0
        oEnPoint.Set parSeat1toSeat2 - parSeat1toCenter, 0, 0
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing

        ' Insert your code for output 2(Insulated Port1)
        If CmpDblEqual(flangeThick1, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        parFace1toCenter = parSeat1toCenter + depth - cptoffset
        parFace2toCenter = parSeat1toSeat2 - parSeat1toCenter + depth2 - CptOffset2
        parFace3toCenter = parSeat3toCenter + depth3 - CptOffset3
        oStPoint.Set -parFace1toCenter, 0, 0
        oEnPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness), 0, 0

        If CmpDblGreaterthan(flangeDiam1, pipeDiam1) Then
            parInsulationDiameter = flangeDiam1 + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam1 + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
        Set ObjInsulatedPort1 = Nothing

        ' Insert your code for output 3(Insulated Port2)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set parFace2toCenter, 0, 0
        oEnPoint.Set (parFace2toCenter - flangeThick - parInsulationThickness), 0, 0

        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
        Set ObjInsulatedPort2 = Nothing

        ' Insert your code for output 4(Insulated Branch)

        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        BranchPoint.Set (parFace3toCenter - flangeThick) * Cos(parAngle), _
                        (parFace3toCenter - flangeThick) * Sin(parAngle), CenterPos.z
        Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, CenterPos, BranchPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
        Set ObjInsulatedBranch = Nothing

        ' Insert your code for output 5 (Insulated Port3)
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), 0
        oEnPoint.Set (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(parAngle), _
                     (parFace3toCenter - flangeThick - parInsulationThickness) * Sin(parAngle), 0

        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
        Set ObjInsulatedPort3 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing

    Case 845, 846, 848, 847, 849, 850

        Select Case lPartDataBasis

        Case 845
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, cptoffset, depth
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, CptOffset2, depth2
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, CptOffset3, depth3

            parFace1toCenter = parSeat1toCenter - cptoffset + depth
            parFace2toCenter = parSeat1toSeat2 + depth + depth2 - _
                               cptoffset - CptOffset2 - parFace1toCenter
            parFace3toCenter = parSeat3toCenter + depth3 - CptOffset3
            parTangentLength1 = parFace1toCenter - parBranchRadius

        Case 846
            parFace2toCenter = parFacetoFace - parFace1toCenter
            parTangentLength1 = parFace1toCenter - parBranchRadius

        Case 848
            parTangentLength1 = parFace1toCenter - parBranchRadius

        Case 847
            parFace1toCenter = parBranchRadius + parTangentLength1
            parFace2toCenter = parFacetoFace - parFace1toCenter

        Case 849
            parFace1toCenter = parBranchRadius + parTangentLength1

        Case 850
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, cptoffset, depth
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, CptOffset2, depth2

            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, CptOffset3, depth3

            parFace1toCenter = parSeat1toCenter - cptoffset + depth
            parFace2toCenter = parSeat1toSeat2 + depth + depth2 - _
                               cptoffset - CptOffset2 - parFace1toCenter
            parTangentLength1 = parFace1toCenter - parBranchRadius
            parTangentLength3 = depth3

        End Select
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        ' Insert your code for the Curved Branch Part
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, _
                           flangeThick, flangeDiam, CptOffset3, depth3
        dInsulationRadius = pipeDiam3 / 2 + parInsulationThickness
        If CmpDblGreaterthan(parAngle, PI / 2) Then
            dElbowRadius = parFace3toCenter - parTangentLength3
            dStartingPoint = parFace3toCenter - parTangentLength3
        Else
            dElbowRadius = parFace3toCenter
            dStartingPoint = parFace3toCenter
        End If
        If CmpDblLessThan(parBranchRadius, dInsulationRadius) Then
            dInsulationRadius = dElbowRadius - 0.0001
        End If
        Set oCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, dInsulationRadius)

        '   Revolve the circle about the Z axis passing by the origin

        oCP.Set -dElbowRadius, dElbowRadius, 0
        ocv.Set 0, 0, 1

        Set ObjCurvedSurface = PlaceRevolution(m_OutputColl, oCircle, ocv, _
                                               oCP, parAngle, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur_", ObjCurvedSurface
        Set ObjCurvedSurface = Nothing
        Set oCP = Nothing
        Set ocv = Nothing

        dFace1toCenter = parFace1toCenter
        dfacetoface = parFace1toCenter + parFace2toCenter
        If Not CmpDblEqual(flangeThick1, 0) Then

            If CmpDblGreaterthan(dfacetoface, flangeThick1 + parInsulationThickness) Then
                dPortLen = flangeThick1 + parInsulationThickness
            Else
                dPortLen = dfacetoface
            End If

            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        'Create the Insulation for Body between Nozzle 1 and Nozzle 2

        dInsDia = pipeDiam1 + 2 * parInsulationThickness
        oStPoint.Set -dFace1toCenter + flangeThick1, 0, 0
        oEnPoint.Set dfacetoface - dFace1toCenter, 0, 0
        Set ObjMainRunInsul = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        m_OutputColl.AddOutput "ObjIns_", ObjMainRunInsul
        Set ObjMainRunInsul = Nothing
        ' Insulation for Port 2

        If Not CmpDblEqual(flangeThick2, 0) Then
            oStPoint.Set dfacetoface - dFace1toCenter, 0, 0
            oEnPoint.Set dfacetoface - dFace1toCenter - flangeThick2, 0, 0
            dPortLen = flangeThick2 + parInsulationThickness
            dInsDia = flangeDiam2 + parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        dface3tocenter = parFace3toCenter
        If Not CmpDblEqual(flangeThick3, 0) Then
            dInsDia = flangeDiam3 + 2 * parInsulationThickness
            oStPoint.Set dface3tocenter * Cos(parAngle), dface3tocenter * Sin(parAngle), 0
            oEnPoint.Set (dface3tocenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (dface3tocenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), 0
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

    Case 851

        ' Insert your code for the base part
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oStPoint.Set -pipeDiam / 2 - parInsulationThickness, 0, -parRibThickness / 2 - parInsulationThickness
        oEnPoint.Set pipeDiam / 2 + parInsulationThickness, -parBasetoCenter - parInsulationThickness, parRibThickness / 2 + parInsulationThickness

        Dim ObjBasePart1 As Object
        Set ObjBasePart1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        m_OutputColl.AddOutput "Basepart_", ObjBasePart1
        Set ObjBasePart1 = Nothing

        ' Insert code for the base part 2
        oStPoint.Set -parRibThickness / 2 - parInsulationThickness, 0, -pipeDiam / 2 - parInsulationThickness
        oEnPoint.Set parRibThickness / 2 + parInsulationThickness, -parBasetoCenter - parInsulationThickness, pipeDiam / 2 + parInsulationThickness

        Dim ObjBasePart2 As Object
        Set ObjBasePart2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        m_OutputColl.AddOutput "Basepart_", ObjBasePart2
        Set ObjBasePart2 = Nothing

        ' Insert code for the base part 3
        oStPoint.Set 0, -parBasetoCenter + parBaseThickness + parInsulationThickness, 0
        oEnPoint.Set 0, -parBasetoCenter - parInsulationThickness, 0

        Dim ObjBasePart3 As Object

        Set ObjBasePart3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                         parBaseWidth + 2 * parInsulationThickness, True)

        'Set the output
        m_OutputColl.AddOutput "Basepart_", ObjBasePart3
        Set ObjBasePart3 = Nothing
        dFace1toCenter = parSeat1toCenter + depth1 - CptOffset1
        parFace2toCenter = parSeat2toCenter + depth2 - CptOffset2
        parFace3toCenter = parSeat3toCenter + depth3 - CptOffset3
        dfacetoface = dFace1toCenter + parFace2toCenter
        If Not CmpDblEqual(flangeThick1, 0) Then

            If CmpDblGreaterthan(dfacetoface, flangeThick1 + parInsulationThickness) Then
                dPortLen = flangeThick1 + parInsulationThickness
            Else
                dPortLen = dfacetoface
            End If

            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        'Create the Insulation for Body between Nozzle 1 and Nozzle 2

        dInsDia = pipeDiam1 + 2 * parInsulationThickness
        oStPoint.Set -dFace1toCenter + flangeThick1, 0, 0
        oEnPoint.Set dfacetoface - dFace1toCenter, 0, 0
        Set ObjMainRunInsul = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        m_OutputColl.AddOutput "ObjIns_", ObjMainRunInsul
        Set ObjMainRunInsul = Nothing
        ' Insulation for Port 2

        If Not CmpDblEqual(flangeThick2, 0) Then
            oStPoint.Set dfacetoface - dFace1toCenter, 0, 0
            oEnPoint.Set dfacetoface - dFace1toCenter - flangeThick2, 0, 0
            dPortLen = flangeThick2 + parInsulationThickness
            dInsDia = flangeDiam2 + parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

        dface3tocenter = parFace3toCenter
        If Not CmpDblEqual(flangeThick3, 0) Then
            dInsDia = flangeDiam3 + 2 * parInsulationThickness
            oStPoint.Set dface3tocenter * Cos(parAngle), dface3tocenter * Sin(parAngle), 0
            oEnPoint.Set (dface3tocenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (dface3tocenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), 0
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

    Case 852, 853
        If CmpDblEqual(parFace3toCenter, 0) Then parFace3toCenter = 0.8 * pipeDiam
        If lPartDataBasis = 853 Then
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, cptoffset, depth
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, CptOffset2, depth2
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, CptOffset3, depth3
            parFace1toCenter = parSeat1toCenter + depth - cptoffset
            parFacetoFace = parSeat1toSeat2 + depth - cptoffset + depth2 - CptOffset2
            parFace3toCenter = parSeat3toCenter + depth3 - CptOffset3
        End If
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        oStPoint.Set -(parFace1toCenter - flangeThick), 0, 0
        oEnPoint.Set parFacetoFace - parFace1toCenter - flangeThick, 0, 0
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing

        ' Insert your code for output 2(Insulated Port1)
        If CmpDblEqual(flangeThick1, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set -parFace1toCenter, 0, 0
        oEnPoint.Set -(parFace1toCenter - flangeThick - parInsulationThickness), 0, 0

        If CmpDblGreaterthan(flangeDiam1, pipeDiam1) Then
            parInsulationDiameter = flangeDiam1 + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam1 + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
        Set ObjInsulatedPort1 = Nothing

        ' Insert your code for output 3(Insulated Port2)
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set parFacetoFace - parFace1toCenter, 0, 0
        oEnPoint.Set (parFacetoFace - parFace1toCenter - flangeThick - parInsulationThickness), 0, 0

        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
        Set ObjInsulatedPort2 = Nothing

        ' Insert your code for output 4(Insulated Branch)

        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        BranchPoint.Set (parFace3toCenter - flangeThick) * Cos(parAngle), _
                        (parFace3toCenter - flangeThick) * Sin(parAngle), CenterPos.z
        Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, CenterPos, BranchPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
        Set ObjInsulatedBranch = Nothing

        ' Insert your code for output 5 (Insulated Port3)
        If CmpDblEqual(flangeThick, 0) Then
            flangeThick = NEGLIGIBLE_THICKNESS
        End If
        oStPoint.Set parFace3toCenter * Cos(parAngle), parFace3toCenter * Sin(parAngle), 0
        oEnPoint.Set (parFace3toCenter - flangeThick - parInsulationThickness) * Cos(parAngle), _
                     (parFace3toCenter - flangeThick - parInsulationThickness) * Sin(parAngle), 0

        If CmpDblGreaterthan(flangeDiam, pipeDiam) Then
            parInsulationDiameter = flangeDiam + parInsulationThickness * 2
        Else
            parInsulationDiameter = pipeDiam + parInsulationThickness * 2
        End If
        Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, parInsulationDiameter, True)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
        Set ObjInsulatedPort3 = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing


    Case 840, 839

        If lPartDataBasis = 840 Then
            RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, cptoffset, depth
            RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, CptOffset2, depth2
            RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, CptOffset3, depth3
            parFace1toCenter = parSeat1toCenter + depth - cptoffset
            parFace2toCenter = parSeat2toCenter + depth2 - CptOffset2
            parFace3toCenter = parSeat3toCenter + depth3 - CptOffset3
        End If
        If lPartDataBasis = 839 Then parFace2toCenter = parFace1toCenter

        ' Insert your code for the base part
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oStPoint.Set -parBaseWidth / 2 - parInsulationThickness, 0, -parRibThickness / 2 - parInsulationThickness
        oEnPoint.Set parBaseWidth / 2 + parInsulationThickness, -parBasetoCenter - parInsulationThickness, parRibThickness / 2 + parInsulationThickness

        Set ObjBasePart1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBasePart1
        Set ObjBasePart1 = Nothing

        ' Insert code for the base part 2

        oStPoint.Set -parRibThickness / 2 - parInsulationThickness, 0, -parBaseWidth / 2 - parInsulationThickness
        oEnPoint.Set parRibThickness / 2 + parInsulationThickness, -parBasetoCenter - parInsulationThickness, parBaseWidth / 2 + parInsulationThickness

        Set ObjBasePart2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBasePart2
        Set ObjBasePart2 = Nothing

        ' Insert code for the base part 3

        oStPoint.Set -parBaseWidth / 2 - parInsulationThickness, -parBasetoCenter + parBaseThickness + parInsulationThickness, -parBaseWidth / 2 - parInsulationThickness
        oEnPoint.Set parBaseWidth / 2 + parInsulationThickness, -parBasetoCenter - parInsulationThickness, parBaseWidth / 2 + parInsulationThickness

        Set ObjBasePart3 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBasePart3
        Set ObjBasePart3 = Nothing

        dfacetoface = parFace1toCenter + parFace2toCenter
        dFace1toCenter = parFace1toCenter
        If Not CmpDblEqual(flangeThick1, 0) Then

            If CmpDblGreaterthan(dfacetoface, flangeThick1 + parInsulationThickness) Then
                dPortLen = flangeThick1 + parInsulationThickness
            Else
                dPortLen = dfacetoface
            End If

            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        'Create the Insulation for Body between Nozzle 1 and Nozzle 2

        dInsDia = pipeDiam1 + 2 * parInsulationThickness
        oStPoint.Set -dFace1toCenter + flangeThick1, 0, 0
        oEnPoint.Set dfacetoface - dFace1toCenter, 0, 0
        Set ObjMainRunInsul = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

        ' Set the output
        m_OutputColl.AddOutput "ObjIns_", ObjMainRunInsul
        Set ObjMainRunInsul = Nothing
        ' Insulation for Port 2

        If Not CmpDblEqual(flangeThick2, 0) Then
            oStPoint.Set dfacetoface - dFace1toCenter, 0, 0
            oEnPoint.Set dfacetoface - dFace1toCenter - flangeThick2, 0, 0
            dPortLen = flangeThick2 + parInsulationThickness
            dInsDia = flangeDiam2 + parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

        dface3tocenter = parFace3toCenter
        If Not CmpDblEqual(flangeThick3, 0) Then
            dInsDia = flangeDiam3 + 2 * parInsulationThickness
            oStPoint.Set dface3tocenter * Cos(parAngle), dface3tocenter * Sin(parAngle), 0
            oEnPoint.Set (dface3tocenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (dface3tocenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), 0
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

    Case 857, 858, 859

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        ' Insert your code for the Cylinder 1(Header)
        oStPoint.Set -parFace1toCenter, 0, 0
        oEnPoint.Set parFacetoFace - parFace1toCenter, 0, 0
        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                             pipeDiam + 2 * parInsulationThickness, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjInsulatedBody
        Set ObjInsulatedBody = Nothing

        '   Revolve the circle about the Z axis passing by the origin for the branch
        If CmpDblGreaterthan(parAngle, PI / 2) Then
            dElbowRadius = parFace3toCenter - depth
            dStartingPoint = parFace3toCenter - depth
        Else
            dElbowRadius = parFace3toCenter
            dStartingPoint = parFace3toCenter
        End If
        dInsulationRadius = pipeDiam3 + parInsulationThickness
        If CmpDblLessThan(dElbowRadius, dInsulationRadius) Then
            dInsulationRadius = dElbowRadius - 0.0001
        End If
        Set oCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, dInsulationRadius)

        '   Revolve the circle about the Z axis passing by the origin

        oCP.Set -dElbowRadius, dElbowRadius, 0
        ocv.Set 0, 0, 1

        Set ObjCurvedSurface = PlaceRevolution(m_OutputColl, oCircle, ocv, _
                                               oCP, parAngle, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur_", ObjCurvedSurface
        Set ObjCurvedSurface = Nothing
        Set oCP = Nothing
        Set ocv = Nothing

        ' Insert Your Code for Output 3 ( Second Branch)

        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam4, _
                           flangeThick, flangeDiam, cptoffset, depth
        dMMRatio = (0.5 + parInsulationThickness) / (0.7 + parInsulationThickness)
        If lPartDataBasis = 857 Then
            Set oEllipse = _
            oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                   -0.2 * pipeDiam4, 0, 0, 0, 0, -1, (0.7 * pipeDiam4 + parInsulationThickness), 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, _
                                                              0.7 * parOutFacetoCenter, 0, 0, 1, pipeDiam4 / 2 + parInsulationThickness)

            Set oBranch2 = _
            oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                        oEllipse, oCircle, True)

            '   Set the output
            m_OutputColl.AddOutput "ObjCur_", oBranch2
            Set oBranch2 = Nothing
            Set oCircle = Nothing

        ElseIf lPartDataBasis = 858 Then
            Set oEllipse = _
            oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                   -0.2 * pipeDiam4, 0, 0, 0, 0, 1, (0.7 * pipeDiam4 + parInsulationThickness), 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, _
                                                              -0.7 * parOutFacetoCenter, 0, 0, -1, pipeDiam4 / 2 + parInsulationThickness)

            Set oBranch2 = _
            oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                        oEllipse, oCircle, True)

            '   Set the output
            m_OutputColl.AddOutput "ObjCur_", oBranch2
            Set oBranch2 = Nothing
            Set oCircle = Nothing
            Set oEllipse = Nothing

        ElseIf lPartDataBasis = 859 Then
            Set oEllipse = _
            oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                   -0.2 * pipeDiam4, 0, 0, 0, 0, -1, 0.7 * pipeDiam4 + parInsulationThickness, 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, _
                                                              0.7 * parOutFacetoCenter, 0, 0, 1, pipeDiam4 / 2 + parInsulationThickness)
            Set oBranch2 = _
            oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                        oEllipse, oCircle, True)

            '   Set the output
            m_OutputColl.AddOutput "ObjCur_", oBranch2
            Set oBranch2 = Nothing
            Set oCircle = Nothing
            Set oEllipse = Nothing

            Set oEllipse = _
            oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                   -0.2 * pipeDiam4, 0, 0, 0, 0, 1, 0.7 * pipeDiam4 + parInsulationThickness, 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, _
                                                              -0.7 * parOutFacetoCenter, 0, 0, -1, pipeDiam4 / 2 + parInsulationThickness)


            Set oBranch3 = _
            oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                        oEllipse, oCircle, True)

            '   Set the output
            m_OutputColl.AddOutput "ObjCur_", oBranch3
            Set oBranch2 = Nothing
            Set oCircle = Nothing
            Set oEllipse = Nothing

        End If
        dfacetoface = parFacetoFace
        dFace1toCenter = parFace1toCenter
        If Not CmpDblEqual(flangeThick1, 0) Then

            If CmpDblGreaterthan(dfacetoface, flangeThick1 + parInsulationThickness) Then
                dPortLen = flangeThick1 + parInsulationThickness
            Else
                dPortLen = dfacetoface
            End If

            oStPoint.Set -dFace1toCenter, 0, 0
            oEnPoint.Set -dFace1toCenter + dPortLen, 0, 0
            dInsDia = flangeDiam1 + 2 * parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        ' Insulation for Port 2

        If Not CmpDblEqual(flangeThick2, 0) Then
            oStPoint.Set dfacetoface - dFace1toCenter, 0, 0
            oEnPoint.Set dfacetoface - dFace1toCenter - flangeThick2, 0, 0
            dPortLen = flangeThick2 + parInsulationThickness
            dInsDia = flangeDiam2 + parInsulationThickness
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)

            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If

        dface3tocenter = parFace3toCenter
        If Not CmpDblEqual(flangeThick3, 0) Then
            dInsDia = flangeDiam3 + 2 * parInsulationThickness
            oStPoint.Set dface3tocenter * Cos(parAngle), dface3tocenter * Sin(parAngle), 0
            oEnPoint.Set (dface3tocenter - flangeThick3 - parInsulationThickness) * Cos(parAngle), (dface3tocenter - flangeThick3 - parInsulationThickness) * Sin(parAngle), 0
            Set objInsPort = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
            ' Set the output
            m_OutputColl.AddOutput "ObjIns_", objInsPort
            Set objInsPort = Nothing
        End If
        ' Place Nozzle 4
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        If lPartDataBasis = 857 Then
            oStPoint.Set 0, 0, parOutFacetoCenter + cptoffset - depth
            oEnPoint.Set 0, 0, 0
        ElseIf lPartDataBasis = 858 Or lPartDataBasis = 859 Then
            oStPoint.Set 0, 0, -(parOutFacetoCenter + cptoffset - depth)
            oEnPoint.Set 0, 0, 0
        End If
        Dim objinscyl As Object
        Set objinscyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + 2 * parInsulationThickness, True)
        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objinscyl
        Set objinscyl = Nothing

        If lPartDataBasis = 859 Then
            oStPoint.Set 0, 0, parOutFacetoCenter + cptoffset - depth
            oEnPoint.Set 0, 0, 0

            Set objinscyl = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, pipeDiam + 2 * parInsulationThickness, True)
            m_OutputColl.AddOutput "InsCyl_", objinscyl
            Set objinscyl = Nothing
        End If

    Case 862, 863, 864
        Dim dBranchRadius As Double
        Dim dDiameter As Double

        Dim dBodyBranchLength As Double

        dBranchRadius = parBranchRadius
        dDiameter = parDiameter

        'Setting the default values for branch radius and body diameter for the following cases
        'Case where branch radius is not specified and diameter specified
        If CmpDblEqual(dBranchRadius, 0) And Not CmpDblEqual(dDiameter, 0) Then
            dBranchRadius = 0.9 * SmallerDim(dFace1toCenter, parFace3toCenter)
            If CmpDblGreaterthan(dDiameter / 2, dBranchRadius) Then _
               dBranchRadius = dDiameter / 2
        End If

        'Case where diameter is not specified and branch radius specified
        If CmpDblEqual(dDiameter, 0) And Not CmpDblEqual(dBranchRadius, 0) Then
            dDiameter = 1.1 * pipeDiam1
            If CmpDblGreaterthan(dDiameter / 2, dBranchRadius) Then _
               dDiameter = 2 * dBranchRadius
        End If

        'Case where both diameter and branch radius are not specified
        If CmpDblEqual(dDiameter, 0) And CmpDblEqual(dBranchRadius, 0) Then
            dDiameter = 1.1 * pipeDiam1
            dBranchRadius = 0.9 * SmallerDim(dFace1toCenter, parFace3toCenter)
            If CmpDblGreaterthan(dDiameter / 2, dBranchRadius) Then _
               dBranchRadius = dDiameter / 2
        End If

        dBodyBranchLength = parFace3toCenter - dNozzleLength

        'Creating the output for the insulation of the curved surface, only
        'when branch radius is greater than the radius of the insulated header cylinder
        If CmpDblGreaterthan(dBranchRadius, dDiameter / 2 + parInsulationThickness) Then
            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                              -dBranchRadius, 0, 0, -1, 0, 0, dDiameter / 2 + parInsulationThickness)

            '   Revolve the circle about the Z axis passing by the origin
            oCP.Set -dBranchRadius, dBranchRadius, 0
            ocv.Set 0, 0, 1

            Set ObjCurvedSurface = PlaceRevolution(m_OutputColl, oCircle, ocv, _
                                                   oCP, PI / 2, True)

            '   Set the output
            m_OutputColl.AddOutput "InsCurSur1", ObjCurvedSurface
            Set ObjCurvedSurface = Nothing
        End If

        'Creating the out put for the insulation of header cylinder
        oStPoint.Set -dFace1toCenter + dNozzleLength, 0, 0

        'Setting the end point based on part data basis value
        If lPartDataBasis = 864 Then
            oEnPoint.Set parLength - dFace1toCenter - dDiameter / 2, 0, 0
        Else
            oEnPoint.Set parFacetoFace - dFace1toCenter - dNozzleLength, 0, 0
        End If

        Set ObjInsulatedBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                             dDiameter + 2 * parInsulationThickness, True)


        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBody
        Set ObjInsulatedBody = Nothing

        'Creating the output for the branch cylinder
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, dBodyBranchLength, 0

        Set ObjInsulatedBranch = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                               dDiameter + 2 * parInsulationThickness, True)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedBranch
        Set ObjInsulatedBranch = Nothing

        'Creating the output for insulation of Port 1
        oStPoint.Set -dFace1toCenter, 0, 0

        'Setting the end points based on flange thickness
        If CmpDblEqual(flangeThick1, 0) Then
            oEnPoint.Set -dFace1toCenter + dNozzleLength, 0, 0
        Else
            oEnPoint.Set -dFace1toCenter + dNozzleLength + parInsulationThickness, 0, 0
        End If

        'Check to see that insulation doesn't exceed face to face in the cases where part data basis values are 862 and 863
        If lPartDataBasis = 862 Or lPartDataBasis = 863 Then
            If CmpDblGreaterthan(flangeThick1 + parInsulationThickness, parFacetoFace) Then
                oEnPoint.Set parFacetoFace - dFace1toCenter, 0, 0
            End If
        End If

        Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                              GreaterDim(flangeDiam1, pipeDiam1) + 2 * parInsulationThickness, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort1
        Set ObjInsulatedPort1 = Nothing

        'Creating outputs based on part data basis values
        If lPartDataBasis = 864 Then
            'Insulation for Port 2
            oStPoint.Set 0, parFace3toCenter, 0

            'Setting the end points based on flange thickness
            If CmpDblEqual(flangeThick2, 0) Then
                oEnPoint.Set 0, parFace3toCenter - dNozzleLength, 0
            Else
                oEnPoint.Set 0, parFace3toCenter - flangeThick2 - parInsulationThickness, 0
            End If

            'Check to see that branch insulation doesn't exceed face to center 3 dimension
            If CmpDblGreaterthan(flangeThick2 + parInsulationThickness, parFace3toCenter) Then
                oEnPoint.Set 0, 0, 0
            End If

            Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                  GreaterDim(flangeDiam2, pipeDiam2) + 2 * parInsulationThickness, True)

            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
            Set ObjInsulatedPort2 = Nothing

            'Creating the output for the insulation of sphere (bull plug)
            Dim objSphere As Object
            oCP.Set parLength - dFace1toCenter - dDiameter / 2, 0, 0
            Set objSphere = PlaceSphere(m_OutputColl, oCP, dDiameter / 2 + parInsulationThickness)

            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSphere
            Set objSphere = Nothing
        Else

            'Insulation for Port 2
            oStPoint.Set parFacetoFace - dFace1toCenter, 0, 0

            'Setting the end points based on flange thickness
            If CmpDblEqual(flangeThick2, 0) Then
                oEnPoint.Set parFacetoFace - dFace1toCenter - dNozzleLength, 0, 0
            Else
                oEnPoint.Set parFacetoFace - dFace1toCenter - flangeThick2 - parInsulationThickness, 0, 0
            End If

            'Check to see that insulation doesn't exceed face to face
            If CmpDblGreaterthan(flangeThick2 + parInsulationThickness, parFacetoFace) Then
                oEnPoint.Set parFacetoFace - dFace1toCenter, 0, 0
            End If

            Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                  GreaterDim(flangeDiam2, pipeDiam2) + 2 * parInsulationThickness, True)

            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort2
            Set ObjInsulatedPort2 = Nothing

            'Insulation for Port 3
            oStPoint.Set 0, parFace3toCenter, 0

            'Setting the end points based on flange thickness
            If CmpDblEqual(flangeThick3, 0) Then
                oEnPoint.Set 0, parFace3toCenter - dNozzleLength, 0
            Else
                oEnPoint.Set 0, parFace3toCenter - flangeThick3 - parInsulationThickness, 0
            End If

            'Check to see that branch insulation doesn't exceed face to center 3 dimension
            If CmpDblGreaterthan(flangeThick3 + parInsulationThickness, parFace3toCenter) Then
                oEnPoint.Set 0, 0, 0
            End If

            Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                                  GreaterDim(flangeDiam3, pipeDiam3) + 2 * parInsulationThickness, True)

            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjInsulatedPort3
            Set ObjInsulatedPort3 = Nothing
        End If

        'Creating the outputs for the insulation of curved surfaces based on part data basis values
        If lPartDataBasis = 862 Then
            'Creating the outputs only when branch radius is greater than the radius of the insulated header cylinder
            If CmpDblGreaterthan(dBranchRadius, dDiameter / 2 + parInsulationThickness) Then
                'Creating output for insulation of curved surface of the branch, which is a mirror of the first curved surface of the branch in Y-Z plane
                Dim ObjCurvedSurface2 As Object
                Set oCircle = _
                oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                                  dBranchRadius, 0, 0, 1, 0, 0, dDiameter / 2 + parInsulationThickness)

                Set oGeomFactory = Nothing

                oCP.Set dBranchRadius, dBranchRadius, 0
                ocv.Set 0, 0, -1

                Set ObjCurvedSurface2 = PlaceRevolution(m_OutputColl, oCircle, ocv, _
                                                        oCP, PI / 2, True)

                '   Set the output
                m_OutputColl.AddOutput "InsCurSur2", ObjCurvedSurface2
                Set ObjCurvedSurface2 = Nothing

                'Setting the points for the line string
                ReDim dPoints(0 To 8)

                'Creating points for the complex string
                'Point 1
                dPoints(0) = -dBranchRadius
                dPoints(1) = 0
                dPoints(2) = dDiameter / 2 + parInsulationThickness

                'Point 2
                dPoints(3) = 0
                dPoints(4) = dBranchRadius
                dPoints(5) = dPoints(2)

                'Point 3
                dPoints(6) = -dPoints(0)
                dPoints(7) = 0
                dPoints(8) = dPoints(2)

                Dim objCollection As Collection
                Set objCollection = New Collection

                Set oGeomFactory = New IngrGeom3D.GeometryFactory

                'Creating circular arc
                Dim oArc As IngrGeom3D.Arc3d
                Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                                       -dBranchRadius, dBranchRadius, dDiameter / 2 + parInsulationThickness, _
                                                                       0, 0, 1, _
                                                                       dPoints(0), dPoints(1), dPoints(2), _
                                                                       dPoints(3), dPoints(4), dPoints(5))
                objCollection.Add oArc

                'Creating circular arc
                Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                                       dBranchRadius, dBranchRadius, dDiameter / 2 + parInsulationThickness, _
                                                                       0, 0, 1, _
                                                                       dPoints(3), dPoints(4), dPoints(5), _
                                                                       dPoints(6), dPoints(7), dPoints(8))

                objCollection.Add oArc

                'Creating line 1
                Dim oLine As IngrGeom3D.Line3d
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                                 dPoints(6), dPoints(7), dPoints(8), _
                                                                 dPoints(0), dPoints(1), dPoints(2))
                objCollection.Add oLine

                oStPoint.Set dPoints(0), dPoints(1), dPoints(2)

                Dim objBodyOutline As IngrGeom3D.ComplexString3d
                Set objBodyOutline = PlaceTrCString(oStPoint, objCollection)

                ocv.Set 0, 0, -1
                Dim ObjCurvedSurface3 As Object
                Set ObjCurvedSurface3 = PlaceProjection(m_OutputColl, objBodyOutline, ocv, _
                                                        dDiameter + 2 * parInsulationThickness, True)

                '   Set the output
                m_OutputColl.AddOutput "ObjInsCur3", ObjCurvedSurface3
                Set ObjCurvedSurface3 = Nothing
                Set objBodyOutline = Nothing
                Set objCollection = Nothing
                Set oLine = Nothing
                Set oArc = Nothing
            End If

        ElseIf lPartDataBasis = 863 Or lPartDataBasis = 864 Then
            'Creating the outputs only when branch radius is greater than the radius of the insulated header cylinder
            If CmpDblGreaterthan(dBranchRadius, dDiameter / 2 + parInsulationThickness) Then
                'Setting the points for the line string
                ReDim dPoints(0 To 8)

                'Creating points for the complex string
                'Point 1
                dPoints(0) = -dBranchRadius
                dPoints(1) = 0
                dPoints(2) = dDiameter / 2 + parInsulationThickness

                'Point 2
                dPoints(3) = 0
                dPoints(4) = dBranchRadius
                dPoints(5) = dPoints(2)

                dPoints(6) = 0
                dPoints(7) = 0
                dPoints(8) = dPoints(2)

                Set objCollection = New Collection

                Set oGeomFactory = New IngrGeom3D.GeometryFactory

                'Creating circular arc
                Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                                       -dBranchRadius, dBranchRadius, dDiameter / 2 + parInsulationThickness, _
                                                                       0, 0, 1, _
                                                                       dPoints(0), dPoints(1), dPoints(2), _
                                                                       dPoints(3), dPoints(4), dPoints(5))
                objCollection.Add oArc

                'Creating line 1
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                                 dPoints(3), dPoints(4), dPoints(5), _
                                                                 dPoints(6), dPoints(7), dPoints(8))
                objCollection.Add oLine

                'Creating line 1
                Set oLine = oGeomFactory.Lines3d.CreateBy2Points(Nothing, _
                                                                 dPoints(6), dPoints(7), dPoints(8), _
                                                                 dPoints(0), dPoints(1), dPoints(2))
                objCollection.Add oLine


                oStPoint.Set dPoints(0), dPoints(1), dPoints(2)

                Set objBodyOutline = PlaceTrCString(oStPoint, objCollection)

                ocv.Set 0, 0, -1

                Set ObjCurvedSurface2 = PlaceProjection(m_OutputColl, objBodyOutline, ocv, _
                                                        dDiameter + 2 * parInsulationThickness, True)

                '   Set the output
                m_OutputColl.AddOutput "ObjInsCur2", ObjCurvedSurface2
                Set ObjCurvedSurface2 = Nothing
                Set objBodyOutline = Nothing
                Set objCollection = Nothing
                Set oLine = Nothing
                Set oArc = Nothing
                Set oGeomFactory = Nothing
            End If
        End If
        
    Case 1043, 1044, 1045
        
        'Create Insulated Port1 Body
        dInsDia = pipeDiam1 + 2 * parInsulationThickness
        oStPoint.Set -dSeat1toCenter, 0, 0
        oEnPoint.Set 0, 0, 0
        Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "InsulatedBody", ObjInsulatedPort1
        Set ObjInsulatedPort1 = Nothing
        
        'Create Insulated Port 2 Body
        dInsDia = pipeDiam2 + 2 * parInsulationThickness
        oStPoint.Set dSeat2toCenter, 0, 0
        oEnPoint.Set 0, 0, 0
        Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "InsulatedBody2", ObjInsulatedPort2
        Set ObjInsulatedPort2 = Nothing
        
        'Create Insulated Port 3 Body
        dInsDia = pipeDiam3 + 2 * parInsulationThickness
        oStPoint.Set 0, dSeat3toCenter, 0
        oEnPoint.Set 0, 0, 0
        Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput "InsulatedBranch", ObjInsulatedPort3
        Set ObjInsulatedPort3 = Nothing
        
        'Create the Insulated Port 1
        dInsDia = IIf(CmpDblGreaterthan(flangeDiam1, pipeDiam1), flangeDiam1, pipeDiam1) _
                                + 2 * parInsulationThickness
        dNozzleLength = IIf(CmpDblGreaterthan(flangeThick1, NEGLIGIBLE_THICKNESS), flangeThick1, _
                            NEGLIGIBLE_THICKNESS)
                                    
        oStPoint.Set -(dSeat1toCenter - CptOffset1 + depth1), 0, 0
        oEnPoint.Set -(dSeat1toCenter - CptOffset1 + depth1 - dNozzleLength), 0, 0
        Set ObjInsulatedPort1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
        
        'Set the output
        m_OutputColl.AddOutput "InsulatedPort1", ObjInsulatedPort1
        Set ObjInsulatedPort1 = Nothing
        
        'Create the Insulated Port 2
        dInsDia = IIf(CmpDblGreaterthan(flangeDiam2, pipeDiam2), flangeDiam2, pipeDiam2) _
                                + 2 * parInsulationThickness
        dNozzleLength = IIf(CmpDblGreaterthan(flangeThick2, NEGLIGIBLE_THICKNESS), flangeThick2, _
                            NEGLIGIBLE_THICKNESS)
                                    
        oStPoint.Set (dSeat2toCenter - CptOffset2 + depth2), 0, 0
        oEnPoint.Set (dSeat2toCenter - CptOffset2 + depth2 - dNozzleLength), 0, 0
        Set ObjInsulatedPort2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
        
        'Set the output
        m_OutputColl.AddOutput "InsulatedPort2", ObjInsulatedPort2
        Set ObjInsulatedPort2 = Nothing
        
        'Create the Insulated Port 3
        dInsDia = IIf(CmpDblGreaterthan(flangeDiam3, pipeDiam3), flangeDiam3, pipeDiam3) _
                                + 2 * parInsulationThickness
        dNozzleLength = IIf(CmpDblGreaterthan(flangeThick3, NEGLIGIBLE_THICKNESS), flangeThick3, _
                            NEGLIGIBLE_THICKNESS)
                                    
        oStPoint.Set 0, (dSeat3toCenter - CptOffset3 + depth3), 0
        oEnPoint.Set 0, (dSeat3toCenter - CptOffset3 + depth3 - dNozzleLength), 0
        Set ObjInsulatedPort3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, dInsDia, True)
        
        'Set the output
        m_OutputColl.AddOutput "InsulatedPort3", ObjInsulatedPort3
        Set ObjInsulatedPort3 = Nothing
        
        'Create the Insulated vortex shape
        Dim objVortexBody As Object
        Dim dArcRadius As Double
        
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                                    flangeThick, flangeDiam, cptoffset, depth3
        
        dArcRadius = pipeDiam / 2 + pipeDiam3 / 2
        
        oStPoint.Set (pipeDiam3 / 2) + 2 * parInsulationThickness, 0, -pipeDiam3 / 2
        oEnPoint.Set -((pipeDiam3 / 2) + 2 * parInsulationThickness), 0, -pipeDiam3 / 2
        Set objVortexBody = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, 2 * dArcRadius, True)
        
        'Set the Output
        m_OutputColl.AddOutput "InsVortexBody", objVortexBody
        Set objVortexBody = Nothing
        
    End Select
    
    Set oGeomFactory = Nothing
    Set oCenter = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oVector = Nothing
    Set oAxisVect = Nothing
    Set oCircle = Nothing
    Set oCP = Nothing
    Set ocv = Nothing
    Set oCenter = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
End Sub
Private Function GreaterDim(ByVal Dim1 As Double, ByVal Dim2 As Double) As Double
    If CmpDblGreaterthanOrEqualTo(Dim1, Dim2) Then
        GreaterDim = Dim1
    Else
        GreaterDim = Dim2
    End If

End Function

