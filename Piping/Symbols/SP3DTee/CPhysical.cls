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
'   Copyright (c) 2003-08, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          NN
'   Creation Date:  Sunday, Nov 5 2000
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who                 change description
'   -----------     -----               ------------------
'   09.Jul.2003    SymbolTeam(India)      Copyright Information, Header  is added.
'   24.Oct.2003    SymbolTeam(India)      TR 51102 Removed duplicated Header/Copyright Information.
'   11.Sep.2006    KKC                    CR 104167 Request to update tee symbol to optionally support 2 dimensions
'    14-SEP-06     KKC                    DI-95670  Replace names with initials in all revision history sheets and symbols")
'   21-Dec-06      svsmylav               TR-112718: Face2toCenter and Face3toCenter are optional inputs added in V7 -
'                                          these are moved to the end of first two V6.1 existing inputs (otherwise mirror operation failed).
'   19-Feb-07      svsmylav               CR-113435: Optional input 'Angle' is added to support Trimmable Tee.
'   06-Mar-07      svsmylav               TR-116016: Added logic to consider optional inputs based on part data basis values
'   18-Mar-08      KKC                    CR-127565 and CR-127567. Created the Sanitary Tee symbol with various gemotric options
'   28-Apr-08      KKC                    TR-139699  Unable to route from second branch of Sanitary Tee
'   12-May-08      RRK                    CR-141196:Enhanced the symbol to support long sweep tees with part data basis values of 862, 863 and 864
'   16-Jul-08      RRK                    CR-CP-145694 Enhanced the symbol to support outlet tee of JIS, by making face 3 to center as optional
'                                         for the part data basis- 'Sanitary tee, 90o side outlet on center, specified by face-to-center 1 and face-to-center 3'(829)
'   05.July.08     RUK                    CR-CP-146405  Enhance Tee symbol for vortex tee geometry per JIS G 5527
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Const MODULE = "Physical:"    'Used for error messages
Private PI As Double

Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), _
               arrayOfOutputs() As String)

    Const METHOD = "run"
    On Error GoTo ErrorLabel

    Dim oPartFclt As PartFacelets.IJDPart
    Dim pipeDiam As Double
    Dim pipeDiam2 As Double
    Dim pipeDiam3 As Double
    Dim pipeDiam4 As Double
    Dim flangeThick As Double
    Dim flangeThick2 As Double
    Dim flangeThick3 As Double
    Dim cptoffset As Double
    Dim flangeDiam As Double
    Dim depth As Double
    Dim CptOffset2 As Double
    Dim depth2 As Double
    Dim CptOffset3 As Double
    Dim depth3 As Double

    Dim iOutput As Double

    Dim ObjInsulatedBody As Object
    Dim ObjInsulatedBranch As Object
    Dim parFacetoFace As Double
    Dim parFace1toCenter As Double
    Dim parFacetoCenter As Double
    Dim parFace2toCenter As Double
    Dim parFace3toCenter As Double
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

    ' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parFacetoCenter = arrayOfInputs(2)
    parInsulationThickness = arrayOfInputs(3)
    
    Dim oCP As AutoMath.DPosition
    Set oCP = New AutoMath.DPosition
    Dim ocv As AutoMath.DVector
    Set ocv = New AutoMath.DVector
    Dim ObjCurvedSurface As Object
    iOutput = 0

    'Checking for the PartDataBasis
    Dim oPipeComponent As IJDPipeComponent
    Dim lPartDataBasis As Long
    Set oPipeComponent = oPartFclt
    lPartDataBasis = oPipeComponent.PartDataBasis

    RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                       flangeThick, flangeDiam, cptoffset, depth
    RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, _
                       flangeThick, flangeDiam, cptoffset, depth2

    RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, _
                       flangeThick, flangeDiam, cptoffset, depth3

    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Dim oCenter As AutoMath.DPosition
    Dim oStPoint As AutoMath.DPosition
    Dim oEnPoint As AutoMath.DPosition
    Dim oVector As AutoMath.DVector
    Dim oCircle As Object
    Dim oArc As IngrGeom3D.Arc3d
    Dim oLine As IngrGeom3D.Line3d
    
    Set oGeomFactory = New GeometryFactory
    Set oCenter = New DPosition
    Set oStPoint = New DPosition
    Set oEnPoint = New DPosition
    Set oVector = New DVector

    Dim dSeattoSeat As Double
    Dim dSeat1toCenter As Double
    Dim dSeat2toCenter As Double
    Dim dSeat3toCenter As Double

    Select Case lPartDataBasis
    Case Is <= 1, 190      ''Default' or 'Tee, Symmetrical, specified by face-to-center 1 190'
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

    Case 828    ' Sanitary Tee - Fig. 36 A861
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

    Case 845    ' Sanitary T branches, Single, specified by header seat-to-seat dimension, seat-to-center 1, seat-to-center 3, branch radius, and tangent length 3
        parSeat1toSeat2 = arrayOfInputs(17)
        parSeat1toCenter = arrayOfInputs(18)
        parSeat3toCenter = arrayOfInputs(9)
        parBranchRadius = arrayOfInputs(22)
        parTangentLength3 = arrayOfInputs(12)
        parInsulationThickness = arrayOfInputs(3)

    Case 846    ' Sanitary T branches,Single,  specified by header face-to-face dimension, face-to-center 1, face-to-center 3, branch radius, and tangent length 3
        parFacetoFace = arrayOfInputs(13)
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parBranchRadius = arrayOfInputs(22)
        parTangentLength3 = arrayOfInputs(12)
        parInsulationThickness = arrayOfInputs(3)

    Case 848    ' Sanitary T branches, Single, specified by face-to-center 1, face-to-center 2, face-to-center 3, branch radius, and tangent length 3
        parFace1toCenter = arrayOfInputs(14)
        parFace2toCenter = arrayOfInputs(4)
        parFace3toCenter = arrayOfInputs(5)
        parBranchRadius = arrayOfInputs(22)
        parTangentLength3 = arrayOfInputs(12)
        parInsulationThickness = arrayOfInputs(3)

    Case 847    ' Sanitary T branches,Single, specified by header face-to-face dimension, tangent length 1, branch radius, face-to-center 3, and tangent length 3
        parFacetoFace = arrayOfInputs(13)
        parTangentLength1 = arrayOfInputs(11)
        parBranchRadius = arrayOfInputs(22)
        parFace3toCenter = arrayOfInputs(5)
        parTangentLength3 = arrayOfInputs(12)
        parInsulationThickness = arrayOfInputs(3)

    Case 849    ' Sanitary T branches ,Single, specified by tangent length 1, branch radius, face-to-center 2, face-to-center 3, and tangent length 3
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
        Dim dNozzleLength As Double

        parFacetoFace = arrayOfInputs(13)
        parFace3toCenter = arrayOfInputs(5)
        parBranchRadius = arrayOfInputs(22)
        parDiameter = arrayOfInputs(28)
        parInsulationThickness = arrayOfInputs(3)

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        dFace1toCenter = parFacetoFace / 2

        'Setting the values of nozzle length
        If CmpDblEqual(flangeThick, 0) Then
            dBodyHeaderLength = 0.95 * parFacetoFace
            dNozzleLength = (parFacetoFace - dBodyHeaderLength) / 2
        Else
            dBodyHeaderLength = parFacetoFace - 2 * flangeThick
            dNozzleLength = flangeThick
        End If

    Case 863    '  Tee, long sweep
        parFacetoFace = arrayOfInputs(13)
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parBranchRadius = arrayOfInputs(22)
        parDiameter = arrayOfInputs(28)
        parInsulationThickness = arrayOfInputs(3)

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        dFace1toCenter = parFace1toCenter

        'Setting the default value of face 1 to center equal to face 3 to center
        If CmpDblEqual(dFace1toCenter, 0) Then dFace1toCenter = parFace3toCenter

        'Setting the values of nozzle length
        If CmpDblEqual(flangeThick, 0) Then
            dBodyHeaderLength = 0.95 * parFacetoFace
            dNozzleLength = (parFacetoFace - dBodyHeaderLength) / 2
        Else
            dBodyHeaderLength = parFacetoFace - 2 * flangeThick
            dNozzleLength = flangeThick
        End If

    Case 864    '  Tee, long sweep with integral bull plug
        parFace1toCenter = arrayOfInputs(14)
        parFace3toCenter = arrayOfInputs(5)
        parBranchRadius = arrayOfInputs(22)
        parDiameter = arrayOfInputs(28)
        parLength = arrayOfInputs(29)
        parInsulationThickness = arrayOfInputs(3)

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        'Setting the default value of face 1 to center equal to face 3 to center
        dFace1toCenter = parFace1toCenter
        If CmpDblEqual(dFace1toCenter, 0) Then dFace1toCenter = parFace3toCenter

        'Setting the values of nozzle length
        If CmpDblEqual(flangeThick, 0) Then
            dBodyHeaderLength = 0.975 * parLength
            dNozzleLength = (parLength - dBodyHeaderLength)
        Else
            dBodyHeaderLength = parLength - flangeThick
            dNozzleLength = flangeThick
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
    
    Dim oPlacePoint As AutoMath.DPosition
    Dim oDir As AutoMath.DVector

    Set oPlacePoint = New DPosition
    Set oDir = New DVector

    Select Case lPartDataBasis
    Case Is <= 1, 190, 192, 194, 196
        'The Body of the Tee is part of the Port1

        ' Place Nozzle 1

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Dim objNozzle As GSCADNozzleEntities.IJDNozzle
        Dim FaceToCenter As Double

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set -parFacetoCenter - cptoffset + depth, 0, 0
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        Dim Port1Length As Double
        Port1Length = parFacetoCenter + parFace2toCenter - flangeThick
        oDir.Set -1, 0, 0

        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                                               oDir, oPlacePoint, Port1Length)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2

        oPlacePoint.Set parFace2toCenter + cptoffset - depth, 0, 0
        oDir.Set 1, 0, 0

        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set (parFace3toCenter + cptoffset - depth) * _
                        Cos(parAngle), (parFace3toCenter + cptoffset - depth) * _
                                       Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0
        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, _
                                               oDir, oPlacePoint, parFace3toCenter)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

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
                                                     oEnPoint, pipeDiam, True)

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
                                                     oEnPoint, pipeDiam, True)

            '   Set the output
            m_OutputColl.AddOutput "IncCyl2", objInclinedCylinder2
            Set objInclinedCylinder2 = Nothing

        ElseIf lPartDataBasis = 827 Then
            oStPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, 0
            oEnPoint.Set parOutOffsetPort1 - parFace1toCenter, parOutFacetoCenter * _
                                                               Cos(PI / 4), parOutFacetoCenter * Sin(PI / 4)

            Set objInclinedCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, _
                                                     oEnPoint, pipeDiam, True)

            '   Set the output
            m_OutputColl.AddOutput "IncCsyl_", objInclinedCylinder1
            Set objInclinedCylinder1 = Nothing

            ' Insert your code for output 2(45 Deg Side Opening above center 2)
            RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, cptoffset, depth

            oStPoint.Set parOutOffsetPort1 - parFace1toCenter, 0, 0
            oEnPoint.Set parOutOffsetPort1 - parFace1toCenter, parOutFacetoCenter * _
                                                               Cos(PI / 4), -parOutFacetoCenter * Sin(PI / 4)

            Set objInclinedCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, _
                                                     oEnPoint, pipeDiam, True)

            '   Set the output
            m_OutputColl.AddOutput "IncCyl2", objInclinedCylinder2
            Set objInclinedCylinder2 = Nothing

        End If

        ' Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set -parFace1toCenter - cptoffset + depth, 0, 0
        oDir.Set -1, 0, 0

        dNozzleLength = parFace1toCenter
        If dNozzleLength < flangeThick Then dNozzleLength = flangeThick

        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                                               oDir, oPlacePoint, dNozzleLength)

        '   Set the output
        m_OutputColl.AddOutput "Nozz_", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set parFacetoFace - parFace1toCenter + cptoffset - _
                        depth, 0, 0
        oDir.Set 1, 0, 0
        dNozzleLength = parFacetoFace - parFace1toCenter
        Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, _
                                               oDir, oPlacePoint, dNozzleLength)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set (parFace3toCenter + cptoffset - depth) * Cos(parAngle), (parFace3toCenter + cptoffset - depth) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        dNozzleLength = parFace3toCenter
        If dNozzleLength < flangeThick Then dNozzleLength = flangeThick

        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, _
                                               oDir, oPlacePoint, dNozzleLength)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 4
        If lPartDataBasis = 825 Then
            RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, cptoffset, depth

            oPlacePoint.Set parOutOffsetPort1 - parFace1toCenter, _
                            (parOutFacetoCenter - cptoffset + depth) * Cos(PI / 4), _
                            (parOutFacetoCenter - cptoffset + depth) * Sin(PI / 4)
            oDir.Set 0, Cos(PI / 4), Sin(PI / 4)

            dNozzleLength = parFace3toCenter
            If dNozzleLength < flangeThick Then dNozzleLength = flangeThick

            Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDir, _
                                         oPlacePoint)

            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing

        ElseIf lPartDataBasis = 826 Then
            ' Place Nozzle 5
            RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, cptoffset, depth

            oPlacePoint.Set parOutOffsetPort1 - parFace1toCenter, _
                            (parOutFacetoCenter - cptoffset + depth) * Cos(PI / 4), _
                            -(parOutFacetoCenter - cptoffset + depth) * Sin(PI / 4)
            oDir.Set 0, Cos(PI / 4), -Sin(PI / 4)

            dNozzleLength = parFace3toCenter
            If dNozzleLength < flangeThick Then dNozzleLength = flangeThick

            Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDir, _
                                         oPlacePoint)

            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
        ElseIf lPartDataBasis = 827 Then

            RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, cptoffset, depth

            oPlacePoint.Set parOutOffsetPort1 - parFace1toCenter, _
                            (parOutFacetoCenter - cptoffset + depth) * Cos(PI / 4), _
                            (parOutFacetoCenter - cptoffset + depth) * Sin(PI / 4)
            oDir.Set 0, Cos(PI / 4), Sin(PI / 4)

            dNozzleLength = parFace3toCenter
            If dNozzleLength < flangeThick Then dNozzleLength = flangeThick

            Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDir, _
                                         oPlacePoint)

            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            ' Place Nozzle 5
            RetrieveParameters 5, oPartFclt, m_OutputColl, pipeDiam, _
                               flangeThick, flangeDiam, cptoffset, depth

            oPlacePoint.Set parOutOffsetPort1 - parFace1toCenter, _
                            (parOutFacetoCenter - cptoffset + depth) * Cos(PI / 4), _
                            -(parOutFacetoCenter - cptoffset + depth) * Sin(PI / 4)
            oDir.Set 0, Cos(PI / 4), -Sin(PI / 4)

            dNozzleLength = parFace3toCenter
            If dNozzleLength < flangeThick Then dNozzleLength = flangeThick

            Set objNozzle = CreateNozzle(5, oPartFclt, m_OutputColl, oDir, _
                                         oPlacePoint)

            '   Set the output
            m_OutputColl.AddOutput "ObjNoz_", objNozzle
            Set objNozzle = Nothing

        End If
        Set oPlacePoint = Nothing
        Set oDir = Nothing
        Set oStPoint = Nothing
        Set oEnPoint = Nothing

    Case 828     ' Sanitary Tee - Fig. 36 A861

        ' Insert your code for the Curved Branch Part
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam2, _
                           flangeThick, flangeDiam, cptoffset, depth

        If CmpDblGreaterthan(pipeDiam2, pipeDiam) Then pipeDiam2 = pipeDiam
        Dim dElbowRadius As Double
        Dim dStartingPoint As Double
        If CmpDblGreaterthan(parAngle, PI / 2) Then
            dElbowRadius = parSeat3toCenter
            dStartingPoint = parSeat3toCenter
        Else
            dElbowRadius = parSeat3toCenter + depth
            dStartingPoint = parSeat3toCenter + depth
        End If
        Set oCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, pipeDiam2 / 2)

        Set oGeomFactory = Nothing
        '   Revolve the circle about the Z axis passing by the origin

        If CmpDblLessThan(parSeat3toCenter, pipeDiam2 / 2) Then _
           dElbowRadius = pipeDiam2 / 2
        Set ocv = New AutoMath.DVector

        oCP.Set -dElbowRadius, dElbowRadius, 0
        ocv.Set 0, 0, 1
        Set ObjCurvedSurface = PlaceRevolution(m_OutputColl, oCircle, ocv, _
                                               oCP, parAngle, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur_", ObjCurvedSurface
        Set ObjCurvedSurface = Nothing
        Set oCP = Nothing
        Set ocv = Nothing
        ' Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set -parSeat1toCenter, 0, 0
        oDir.Set -1, 0, 0

        dNozzleLength = parSeat1toSeat2
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set parSeat1toSeat2 - parSeat1toCenter, 0, 0

        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set parSeat3toCenter * Cos(parAngle), parSeat3toCenter * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

    Case 829    ' Sanitary tee, 90o side outlet on center, specified by face-to-center 1 and face-to-center 3

        ' Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set -parFace1toCenter - cptoffset + depth, 0, 0
        oDir.Set -1, 0, 0
        dNozzleLength = 2 * parFace1toCenter - flangeThick
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)

        '   Set the output
        m_OutputColl.AddOutput "Noz_", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set parFace1toCenter + cptoffset - depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set (parFace3toCenter + cptoffset - depth) * Cos(parAngle), (parFace3toCenter + cptoffset - depth) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, parFace3toCenter)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 4
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set 0, 0, -(parFace3toCenter + cptoffset - depth)
        oDir.Set 0, 0, -1

        Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, parFace3toCenter)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

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
                                         1.4 * pipeDiam, True)
        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder1
        Set ObjCylinder1 = Nothing

        If CmpDblGreaterthan(parAngle, PI / 2) Then
            dElbowRadius = parFace3toCenter - depth3
            dStartingPoint = parFace3toCenter - depth3
        Else
            dElbowRadius = parFace3toCenter
            dStartingPoint = parFace3toCenter
        End If

        Set oCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, pipeDiam2 / 2)

        Set oGeomFactory = Nothing
        '   Revolve the circle about the Z axis passing by the origin

        oCP.Set -dElbowRadius, dElbowRadius, 0
        ocv.Set 0, 0, 1
        Set ObjCurvedSurface = PlaceRevolution(m_OutputColl, oCircle, ocv, _
                                               oCP, parAngle, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCurvedSurface
        Set ObjCurvedSurface = Nothing
        Set oCP = Nothing
        Set ocv = Nothing

        ' Place Nozzle 1

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set (-parFace1toCenter - cptoffset + depth), 0, 0
        oDir.Set -1, 0, 0
        dNozzleLength = parFace1toCenter + parFace2toCenter + 0.0254    ' Added 1 " for Plug
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set (parFace3toCenter - depth + cptoffset) * Cos(parAngle), (parFace3toCenter - depth + cptoffset) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

    Case 832     ' Sanitary Tapped Tee

        If CmpDblGreaterthan(parAngle, PI / 2) Then
            dElbowRadius = parFace3toCenter - depth3
            dStartingPoint = parFace3toCenter - depth3
        Else
            dElbowRadius = parFace3toCenter
            dStartingPoint = parFace3toCenter
        End If
        If CmpDblGreaterthan(pipeDiam3, pipeDiam) Then pipeDiam3 = pipeDiam

        Set oCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, pipeDiam3 / 2)

        Set oGeomFactory = Nothing
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
        ' Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set (-parFace1toCenter - cptoffset + depth), 0, 0
        oDir.Set -1, 0, 0
        dNozzleLength = parFacetoFace
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set (parFacetoFace - parFace1toCenter + cptoffset - depth), 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set (parFace3toCenter + cptoffset - depth) * Cos(parAngle), (parFace3toCenter + cptoffset - depth) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

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

        If CmpDblGreaterthan(parAngle, PI / 2) Then
            dElbowRadius = parFace3toCenter - depth3
            dStartingPoint = parFace3toCenter - depth3
        Else
            dElbowRadius = parFace3toCenter
            dStartingPoint = parFace3toCenter
        End If


        Set oCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, pipeDiam3 / 2)

        Set oGeomFactory = Nothing
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
        dMMRatio = 0.5 / 0.7
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        If lPartDataBasis = 833 Or lPartDataBasis = 854 Then
            Set oEllipse = _
            oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                   parOutOffsetPort1 - parFace1toCenter - 0.2 * pipeDiam4, 0, 0, 0, 0, -1, parOutOffsetPort1 - parFace1toCenter + 0.7 * pipeDiam4, 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, parOutOffsetPort1 - parFace1toCenter, 0, _
                                                              0.7 * parOutFacetoCenter, 0, 0, 1, pipeDiam4 / 2)


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
                                                                   parOutOffsetPort1 - parFace1toCenter - 0.2 * pipeDiam4, 0, 0, 0, 0, 1, -(parOutOffsetPort1 - parFace1toCenter + 0.7 * pipeDiam4), 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, parOutOffsetPort1 - parFace1toCenter, 0, _
                                                              -0.7 * parOutFacetoCenter, 0, 0, -1, pipeDiam4 / 2)

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
                                                                   parOutOffsetPort1 - parFace1toCenter - 0.2 * pipeDiam4, 0, 0, 0, 0, -1, parOutOffsetPort1 - parFace1toCenter + 0.7 * pipeDiam4, 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, parOutOffsetPort1 - parFace1toCenter, 0, _
                                                              0.7 * parOutFacetoCenter, 0, 0, 1, pipeDiam4 / 2)

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
                                                                   parOutOffsetPort1 - parFace1toCenter - 0.2 * pipeDiam4, 0, 0, 0, 0, 1, -(parOutOffsetPort1 - parFace1toCenter + 0.7 * pipeDiam4), 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, parOutOffsetPort1 - parFace1toCenter, 0, _
                                                              -0.7 * parOutFacetoCenter, 0, 0, -1, pipeDiam4 / 2)

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
        ' Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set (-parFace1toCenter - cptoffset + depth), 0, 0
        dNozzleLength = parFacetoFace
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)

        '   Set the output
        m_OutputColl.AddOutput "ObjNozz_", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set (parFacetoFace - parFace1toCenter + cptoffset - depth), 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set (parFace3toCenter + cptoffset - depth) * Cos(parAngle), (parFace3toCenter + cptoffset - depth) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        ' Place Nozzle 4
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        If lPartDataBasis = 833 Or lPartDataBasis = 854 Then
            oPlacePoint.Set parOutOffsetPort1 - parFace1toCenter, 0, parOutFacetoCenter + cptoffset - depth
            oDir.Set 0, 0, 1
        ElseIf lPartDataBasis = 834 Or lPartDataBasis = 835 Or lPartDataBasis = 855 Or lPartDataBasis = 856 Then
            oPlacePoint.Set parOutOffsetPort1 - parFace1toCenter, 0, -(parOutFacetoCenter + cptoffset - depth)
            oDir.Set 0, 0, -1
        End If
        Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, 0.3 * parOutFacetoCenter)
        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        If lPartDataBasis = 835 Or lPartDataBasis = 856 Then
            oPlacePoint.Set parOutOffsetPort1 - parFace1toCenter, 0, parOutFacetoCenter + cptoffset - depth
            oDir.Set 0, 0, 1

            Set objNozzle = CreateNozzleWithLength(5, oPartFclt, m_OutputColl, oDir, _
                                                   oPlacePoint, 0.3 * parOutFacetoCenter)
            m_OutputColl.AddOutput "ObjNoz_", objNozzle
            Set objNozzle = Nothing
        End If

    Case 836, 837    ' Horizontal tapped Tee

        If lPartDataBasis = 836 Then
            parFace2toCenter = parFacetoFace - parFace1toCenter
        End If
        Dim pipeDiam1 As Double
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam1, _
                           flangeThick, flangeDiam, cptoffset, depth

        ' Insert your code for the Horizontal Branch 1
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        If CmpDblEqual(flangeDiam, 0) Then flangeDiam = pipeDiam
        oStPoint.Set 0, -pipeDiam1 / 2, -flangeDiam / 2
        oEnPoint.Set (parFace3toCenter - depth + cptoffset) * Cos(parAngle), (parFace3toCenter - depth + cptoffset) * Sin(parAngle), -flangeDiam / 2

        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                         pipeDiam, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder1
        Set ObjCylinder1 = Nothing

        ' Insert your code for the Horizontal Branch 2
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        If CmpDblEqual(flangeDiam, 0) Then flangeDiam = pipeDiam
        oStPoint.Set 0, -pipeDiam1 / 2, flangeDiam / 2
        oEnPoint.Set (parFace3toCenter - depth + cptoffset) * Cos(parAngle), (parFace3toCenter - depth + cptoffset) * Sin(parAngle), flangeDiam / 2
        Dim ObjCylinder2 As Object
        Set ObjCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                         pipeDiam, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", ObjCylinder2
        Set ObjCylinder2 = Nothing
        Dim ObjCylinder3 As Object
        ' Place Nozzle 1

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set (-parFace1toCenter - cptoffset + depth), 0, 0
        oDir.Set -1, 0, 0
        If lPartDataBasis = 837 Then parFacetoFace = parFace1toCenter + parFace2toCenter
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, parFacetoFace)

        '   Set the output
        m_OutputColl.AddOutput "ObjNozz_", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set (parFace2toCenter + cptoffset - depth), 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        If CmpDblEqual(flangeDiam, 0) Then flangeDiam = pipeDiam
        oPlacePoint.Set (parFace3toCenter + cptoffset - depth) * Cos(parAngle), (parFace3toCenter + cptoffset - depth) * Sin(parAngle), _
                        -flangeDiam / 2
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        ' Place Nozzle 4
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        If CmpDblEqual(flangeDiam, 0) Then flangeDiam = pipeDiam
        oPlacePoint.Set (parFace3toCenter + cptoffset - depth) * Cos(parAngle), (parFace3toCenter + cptoffset - depth) * Sin(parAngle), _
                        flangeDiam / 2
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

    Case 838  ' Vertical tapped Tee
        ' Insert your code for the Horizontal Branch 1
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oStPoint.Set -parOffBranch1toBranch2 / 2, 0, 0
        oEnPoint.Set -parOffBranch1toBranch2 / 2 + (parFace3toCenter - depth) * Cos(parAngle), (parFace3toCenter - depth) * Sin(parAngle), 0

        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                         pipeDiam, True)

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
                                         pipeDiam, True)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjCylinder2
        Set ObjCylinder2 = Nothing

        ' Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set (-parFacetoFace / 2 - cptoffset + depth) * Sin(parAngle), 0, 0
        oDir.Set -1, 0, 0

        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, parFacetoFace)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set (parFacetoFace / 2 + cptoffset - depth) * Sin(parAngle), 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        m_OutputColl.AddOutput "ObjNozz_", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set -parOffBranch1toBranch2 / 2 + (parFace3toCenter + cptoffset - depth) * Cos(parAngle), (parFace3toCenter + cptoffset - depth) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        '   Set the output
        m_OutputColl.AddOutput "ObjNozz_", objNozzle
        Set objNozzle = Nothing
        ' Place Nozzle 4
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set parOffBranch1toBranch2 / 2 + (parFace3toCenter + cptoffset - depth) * Cos(parAngle), (parFace3toCenter + cptoffset - depth) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        Set objNozzle = CreateNozzle(4, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        '   Set the output
        m_OutputColl.AddOutput "ObjNoz_", objNozzle
        Set objNozzle = Nothing

    Case 841    ' T branches, Single, specified by header face-to-face dimension, face-to-center 1, and face-to-center 3.

        ' Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set -parFace1toCenter - cptoffset + depth, 0, 0
        oDir.Set -1, 0, 0

        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, parFacetoFace)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set parFacetoFace - parFace1toCenter + cptoffset - _
                        depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set (parFace3toCenter + cptoffset - depth) * _
                        Cos(parAngle), (parFace3toCenter + cptoffset - depth) * _
                                       Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, parFace3toCenter)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

    Case 842, 860   'T branches, single, specified by face-to-center 1, face-to-center 2, and face-to-center 3.

        If lPartDataBasis = 860 Then parFace2toCenter = parFace1toCenter
        ' Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set -parFace1toCenter - cptoffset + depth, 0, 0
        oDir.Set -1, 0, 0
        dNozzleLength = parFace1toCenter + parFace2toCenter
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set parFace2toCenter + cptoffset - depth, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set (parFace3toCenter + cptoffset - depth) * _
                        Cos(parAngle), (parFace3toCenter + cptoffset - depth) * _
                                       Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, parFace3toCenter)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

    Case 843    'T branches specified by header seat-to-seat dimension, seat-to-center 1, and seat-to-center 3

        ' Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set -parSeat1toCenter - cptoffset + depth, 0, 0
        oDir.Set -1, 0, 0
        dNozzleLength = parSeat1toSeat2
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set parSeat1toSeat2 - parSeat1toCenter, 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set (parSeat3toCenter) * _
                        Cos(parAngle), (parSeat3toCenter) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, (parSeat3toCenter + depth - cptoffset))

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

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
        If CmpDblGreaterthanOrEqualTo(parAngle, PI / 2) Then
            dElbowRadius = parFace3toCenter - parTangentLength3
            dStartingPoint = parFace3toCenter - parTangentLength3
        Else
            dElbowRadius = parFace3toCenter
            dStartingPoint = parFace3toCenter
        End If
        Set oCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, pipeDiam3 / 2)

        Set oGeomFactory = Nothing
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

        ' Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set (-parFace1toCenter - cptoffset + depth), 0, 0
        oDir.Set -1, 0, 0
        dNozzleLength = parFace1toCenter + parFace2toCenter - flangeThick
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set (parFace2toCenter + cptoffset - depth), 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set (parFace3toCenter + cptoffset - depth) * Cos(parAngle), (parFace3toCenter + cptoffset - depth) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, parTangentLength3)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

    Case 844
        ' Place Nozzle 1

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, CptOffset2, depth2

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set -parSeat1toCenter, 0, 0
        Port1Length = parSeat1toCenter + parSeat2toCenter + depth + depth2 - cptoffset - CptOffset2
        oDir.Set -1, 0, 0

        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                                               oDir, oPlacePoint, Port1Length)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set parSeat2toCenter, 0, 0
        oDir.Set 1, 0, 0

        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set (parSeat3toCenter) * Cos(parAngle), _
                        (parSeat3toCenter) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0
        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, _
                                               oDir, oPlacePoint, parSeat3toCenter + depth)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

    Case 851

        ' Insert your code for the base part
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oStPoint.Set -pipeDiam / 2, 0, -parRibThickness / 2
        oEnPoint.Set pipeDiam / 2, -parBasetoCenter, parRibThickness / 2

        Dim ObjBasePart1 As Object
        Set ObjBasePart1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        m_OutputColl.AddOutput "Basepart_", ObjBasePart1
        Set ObjBasePart1 = Nothing

        ' Insert code for the base part 2
        oStPoint.Set -parRibThickness / 2, -0, -pipeDiam / 2
        oEnPoint.Set parRibThickness / 2, -parBasetoCenter, pipeDiam / 2

        Dim ObjBasePart2 As Object
        Set ObjBasePart2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        m_OutputColl.AddOutput "Basepart_", ObjBasePart2
        Set ObjBasePart2 = Nothing

        ' Insert code for the base part 3
        oStPoint.Set 0, -parBasetoCenter + parBaseThickness, 0
        oEnPoint.Set 0, -parBasetoCenter, 0

        Dim ObjBasePart3 As Object

        Set ObjBasePart3 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                         parBaseWidth, True)


        'Set the output
        m_OutputColl.AddOutput "Basepart_", ObjBasePart3
        Set ObjBasePart3 = Nothing

        ' Place Nozzle 1

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth


        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set (-parSeat1toCenter) * Sin(parAngle), 0, 0
        oDir.Set -1, 0, 0
        Port1Length = parSeat1toCenter + parSeat2toCenter + depth
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, _
                                               oDir, oPlacePoint, Port1Length)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set parSeat2toCenter, 0, 0
        oDir.Set 1, 0, 0

        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set parSeat3toCenter * Cos(parAngle), parSeat3toCenter * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0
        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, _
                                               oDir, oPlacePoint, parSeat3toCenter + depth)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

    Case 852, 853

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

        If CmpDblEqual(parFace3toCenter, 0) Then parFace3toCenter = 0.8 * pipeDiam
        ' Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set (-parFace1toCenter - cptoffset + depth), _
                        0, 0
        oDir.Set -1, 0, 0
        dNozzleLength = parFacetoFace
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set (parFacetoFace - parFace1toCenter + cptoffset - depth) _
                      , 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set (parFace3toCenter - depth) * Cos(parAngle), (parFace3toCenter - depth) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0
        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, parFace3toCenter)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

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

        oStPoint.Set -parBaseWidth / 2, -pipeDiam / 2, -parRibThickness / 2
        oEnPoint.Set parBaseWidth / 2, -parBasetoCenter, parRibThickness / 2


        Set ObjBasePart1 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBasePart1
        Set ObjBasePart1 = Nothing

        ' Insert code for the base part 2

        oStPoint.Set -parRibThickness / 2, -pipeDiam / 2, -parBaseWidth / 2
        oEnPoint.Set parRibThickness / 2, -parBasetoCenter, parBaseWidth / 2

        Set ObjBasePart2 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBasePart2
        Set ObjBasePart2 = Nothing

        ' Insert code for the base part 3

        oStPoint.Set -parBaseWidth / 2, -parBasetoCenter + parBaseThickness, -parBaseWidth / 2
        oEnPoint.Set parBaseWidth / 2, -parBasetoCenter, parBaseWidth / 2

        Set ObjBasePart3 = PlaceBox(m_OutputColl, oStPoint, oEnPoint)

        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBasePart3
        Set ObjBasePart3 = Nothing

        ' Place Nozzle 1

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth


        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set -parFace1toCenter - cptoffset + depth, 0, 0


        Port1Length = parFace1toCenter + parFace2toCenter
        oDir.Set -1, 0, 0

        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, oPlacePoint, Port1Length)
        ' Set the output
        m_OutputColl.AddOutput "ObjNoz_", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set parFace2toCenter + cptoffset - depth, 0, 0
        oDir.Set 1, 0, 0

        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, oPlacePoint)

        ' Set the output
        m_OutputColl.AddOutput "ObjNoz_", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set (parFace3toCenter + cptoffset - depth) * Cos(parAngle), (parFace3toCenter + cptoffset - depth) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0
        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, oPlacePoint, parFace3toCenter)
        ' Set the output
        m_OutputColl.AddOutput "ObjNoz_", objNozzle
        Set objNozzle = Nothing
    Case 857, 858, 859

        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        ' Insert your code for the Cylinder 1(Header)
        oStPoint.Set -parFace1toCenter, 0, 0
        oEnPoint.Set parFacetoFace - parFace1toCenter, 0, 0
        Dim objcylinder As Object
        Set objcylinder = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                        pipeDiam, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl_", objcylinder
        Set objcylinder = Nothing

        '   Revolve the circle about the Z axis passing by the origin for the branch

        If CmpDblGreaterthan(parAngle, PI / 2) Then
            dElbowRadius = parFace3toCenter - depth3
            dStartingPoint = parFace3toCenter - depth3
        Else
            dElbowRadius = parFace3toCenter
            dStartingPoint = parFace3toCenter
        End If


        Set oCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dStartingPoint, 0, 0, -1, 0, 0, pipeDiam3 / 2)

        Set oGeomFactory = Nothing
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
        dMMRatio = 0.6 / 0.7
        Set oGeomFactory = New IngrGeom3D.GeometryFactory
        If lPartDataBasis = 857 Then
            Set oEllipse = _
            oGeomFactory.Ellipses3d.CreateByCenterNormMajAxisRatio(Nothing, _
                                                                   -0.2 * pipeDiam4, 0, 0, 0, 0, -1, 0.7 * pipeDiam4, 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, _
                                                              0.7 * parOutFacetoCenter, 0, 0, 1, pipeDiam4 / 2)

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
                                                                   -0.2 * pipeDiam4, 0, 0, 0, 0, 1, 0.7 * pipeDiam4, 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, _
                                                              -0.7 * parOutFacetoCenter, 0, 0, -1, pipeDiam4 / 2)

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
                                                                   -0.2 * pipeDiam4, 0, 0, 0, 0, -1, 0.7 * pipeDiam4, 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, _
                                                              0.7 * parOutFacetoCenter, 0, 0, 1, pipeDiam4 / 2)


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
                                                                   -0.2 * pipeDiam4, 0, 0, 0, 0, 1, 0.7 * pipeDiam4, 0, 0, _
                                                                   dMMRatio)

            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, 0, 0, _
                                                              -0.7 * parOutFacetoCenter, 0, 0, -1, pipeDiam4 / 2)

            Set oBranch3 = _
            oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                                        oEllipse, oCircle, True)

            '   Set the output
            m_OutputColl.AddOutput "ObjCur_", oBranch3
            Set oBranch2 = Nothing
            Set oCircle = Nothing
            Set oEllipse = Nothing

        End If
        ' Place Nozzle 1
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector

        oPlacePoint.Set (-parFace1toCenter - cptoffset + depth), 0, 0
        oDir.Set -1, 0, 0

        Set objNozzle = CreateNozzle(1, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        '   Set the output
        m_OutputColl.AddOutput "ObjNozz_", objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 2
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        oPlacePoint.Set (parFacetoFace - parFace1toCenter + cptoffset - depth) * Sin(parAngle), 0, 0
        oDir.Set 1, 0, 0
        Set objNozzle = CreateNozzle(2, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        ' Place Nozzle 3
        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth

        oPlacePoint.Set (parFace3toCenter + cptoffset - depth) * Cos(parAngle), (parFace3toCenter + cptoffset - depth) * Sin(parAngle), 0
        oDir.Set Cos(parAngle), Sin(parAngle), 0

        Set objNozzle = CreateNozzle(3, oPartFclt, m_OutputColl, oDir, _
                                     oPlacePoint)

        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        ' Place Nozzle 4
        RetrieveParameters 4, oPartFclt, m_OutputColl, pipeDiam, _
                           flangeThick, flangeDiam, cptoffset, depth
        If lPartDataBasis = 857 Then
            oPlacePoint.Set 0, 0, parOutFacetoCenter + cptoffset - depth
            oDir.Set 0, 0, 1
        ElseIf lPartDataBasis = 858 Or lPartDataBasis = 859 Then
            oPlacePoint.Set 0, 0, -(parOutFacetoCenter + cptoffset - depth)
            oDir.Set 0, 0, -1
        End If
        Set objNozzle = CreateNozzleWithLength(4, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, 0.3 * parOutFacetoCenter)
        '   Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        If lPartDataBasis = 859 Then
            oPlacePoint.Set 0, 0, parOutFacetoCenter + cptoffset - depth
            oDir.Set 0, 0, 1

            Set objNozzle = CreateNozzleWithLength(5, oPartFclt, m_OutputColl, oDir, _
                                                   oPlacePoint, 0.3 * parOutFacetoCenter)
            m_OutputColl.AddOutput "ObjNozz_", objNozzle
            Set objNozzle = Nothing
        End If

    Case 862, 863, 864
        RetrieveParameters 2, oPartFclt, m_OutputColl, pipeDiam2, _
                           flangeThick2, flangeDiam, CptOffset2, depth2

        RetrieveParameters 3, oPartFclt, m_OutputColl, pipeDiam3, _
                           flangeThick3, flangeDiam, CptOffset3, depth3

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
            dDiameter = 1.1 * pipeDiam
            If CmpDblGreaterthan(dDiameter / 2, dBranchRadius) Then _
               dDiameter = 2 * dBranchRadius
        End If

        'Case where both diameter and branch radius are not specified
        If CmpDblEqual(dDiameter, 0) And CmpDblEqual(dBranchRadius, 0) Then
            dDiameter = 1.1 * pipeDiam
            dBranchRadius = 0.9 * SmallerDim(dFace1toCenter, parFace3toCenter)
            If CmpDblGreaterthan(dDiameter / 2, dBranchRadius) Then _
               dBranchRadius = dDiameter / 2
        End If

        dBodyBranchLength = parFace3toCenter - dNozzleLength

        'Creating the output for the curved surface of the branch
        'Create a circle
        Set oCircle = _
        oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                          -dBranchRadius, 0, 0, -1, 0, 0, dDiameter / 2)

        ' Revolve the circle about the Z axis passing by the origin
        oCP.Set -dBranchRadius, dBranchRadius, 0
        ocv.Set 0, 0, 1

        Set ObjCurvedSurface = PlaceRevolution(m_OutputColl, oCircle, ocv, _
                                               oCP, PI / 2, True)

        '   Set the output
        m_OutputColl.AddOutput "CurSur1", ObjCurvedSurface
        Set ObjCurvedSurface = Nothing

        'Creating the output for the header cylinder
        oStPoint.Set -dFace1toCenter + dNozzleLength, 0, 0

        'Setting the end point based on part data basis value
        If lPartDataBasis = 864 Then
            oEnPoint.Set parLength - dFace1toCenter - dDiameter / 2, 0, 0
        Else
            oEnPoint.Set parFacetoFace - dFace1toCenter - dNozzleLength, 0, 0
        End If

        Set ObjCylinder1 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                         dDiameter, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl1", ObjCylinder1
        Set ObjCylinder1 = Nothing

        'Creating the output for the branch cylinder
        oStPoint.Set 0, 0, 0
        oEnPoint.Set 0, dBodyBranchLength, 0

        Set ObjCylinder2 = PlaceCylinder(m_OutputColl, oStPoint, oEnPoint, _
                                         dDiameter, True)

        '   Set the output
        m_OutputColl.AddOutput "ObjCyl2", ObjCylinder2
        Set ObjCylinder2 = Nothing

        'Place Nozzle 1
        Set oPlacePoint = New AutoMath.DPosition
        Set oDir = New AutoMath.DVector
        oPlacePoint.Set -dFace1toCenter - cptoffset + depth, 0, 0
        oDir.Set -1, 0, 0
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)
        'Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing

        'Creating other outputs based on part data basis values. For part data basis value of 864,nozzle 2 and sphere(bull plug) are required.
        'For part data basis values of 862 and 863 two nozzle 2 and nozzle 3 are required
        If lPartDataBasis = 864 Then
            'Place Nozzle 2
            oPlacePoint.Set 0, parFace3toCenter + CptOffset2 - depth2, 0
            oDir.Set 0, 1, 0

            If CmpDblLessThan(dNozzleLength, flangeThick2) Then dNozzleLength = flangeThick2
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                                                   oPlacePoint, dNozzleLength)
            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing

            'Creating the output for sphere
            Dim objSphere As Object
            oCP.Set parLength - dFace1toCenter - dDiameter / 2, 0, 0
            Set objSphere = PlaceSphere(m_OutputColl, oCP, dDiameter / 2)

            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objSphere
            Set objSphere = Nothing
        Else
            'Place Nozzle 2
            oPlacePoint.Set parFacetoFace - dFace1toCenter + CptOffset2 - depth2, 0, 0
            oDir.Set 1, 0, 0

            If CmpDblLessThan(dNozzleLength, flangeThick2) Then dNozzleLength = flangeThick2
            Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                                                   oPlacePoint, dNozzleLength)
            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing

            'Place Nozzle 3
            oPlacePoint.Set 0, parFace3toCenter + CptOffset3 - depth3, 0
            oDir.Set 0, 1, 0
            If CmpDblLessThan(dNozzleLength, flangeThick3) Then dNozzleLength = flangeThick3
            Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, _
                                                   oPlacePoint, dNozzleLength)
            '   Set the output
            iOutput = iOutput + 1
            m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
            Set objNozzle = Nothing
            Set oPlacePoint = Nothing
            Set oDir = Nothing
        End If

        'Creating outputs for curved surfaces based on part data basis values. For part data basis value of 862,two curved surfaces are required.
        'For part data basis values of 863 and 864 only one curved surface is required
        If lPartDataBasis = 862 Then
            'Creating output for curved surface of the branch, which is a mirror of the first curved surface of the branch in Y-Z plane
            Dim ObjCurvedSurface2 As Object
            Set oCircle = _
            oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                                                              dBranchRadius, 0, 0, 1, 0, 0, dDiameter / 2)

            Set oGeomFactory = Nothing

            oCP.Set dBranchRadius, dBranchRadius, 0
            ocv.Set 0, 0, -1

            Set ObjCurvedSurface2 = PlaceRevolution(m_OutputColl, oCircle, ocv, _
                                                    oCP, PI / 2, True)

            '   Set the output
            m_OutputColl.AddOutput "CurSur2", ObjCurvedSurface2
            Set ObjCurvedSurface2 = Nothing

            'Creating the output for the curved surface which is the projection along z axis
            'Setting the points for the line string
            ReDim dPoints(0 To 8)

            'Creating points for the complex string
            'Point 1
            dPoints(0) = -dBranchRadius
            dPoints(1) = 0
            dPoints(2) = dDiameter / 2

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
            Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                                   -dBranchRadius, dBranchRadius, dDiameter / 2, _
                                                                   0, 0, 1, _
                                                                   dPoints(0), dPoints(1), dPoints(2), _
                                                                   dPoints(3), dPoints(4), dPoints(5))
            objCollection.Add oArc

            'Creating circular arc
            Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                                   dBranchRadius, dBranchRadius, dDiameter / 2, _
                                                                   0, 0, 1, _
                                                                   dPoints(3), dPoints(4), dPoints(5), _
                                                                   dPoints(6), dPoints(7), dPoints(8))

            objCollection.Add oArc

            'Creating line 1
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
                                                    dDiameter, True)

            '   Set the output
            m_OutputColl.AddOutput "ObjCur3", ObjCurvedSurface3
            Set ObjCurvedSurface3 = Nothing
            Set objBodyOutline = Nothing
            Set oLine = Nothing
            Set oArc = Nothing
            Set oGeomFactory = Nothing
            Set objCollection = Nothing
        ElseIf lPartDataBasis = 863 Or lPartDataBasis = 864 Then
            'Creating the output for the curved surface which is the projection along z axis
            'Setting the points for the line string
            ReDim dPoints(0 To 8)
            '
            'Creating points for the complex string
            'Point 1
            dPoints(0) = -dBranchRadius
            dPoints(1) = 0
            dPoints(2) = dDiameter / 2

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
                                                                   -dBranchRadius, dBranchRadius, dDiameter / 2, _
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
                                                    dDiameter, True)

            '   Set the output
            m_OutputColl.AddOutput "ObjCur2", ObjCurvedSurface2
            Set ObjCurvedSurface2 = Nothing
            Set objCollection = Nothing
            Set objBodyOutline = Nothing
            Set oArc = Nothing
            Set oLine = Nothing
            Set oGeomFactory = Nothing
        End If

    Case 1043, 1044, 1045
        'Create Port1
        oPlacePoint.Set -dSeat1toCenter, 0, 0
        oDir.Set -1, 0, 0
        dNozzleLength = dSeat1toCenter + depth
        Set objNozzle = CreateNozzleWithLength(1, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Create Port 2
        oPlacePoint.Set dSeat2toCenter, 0, 0
        oDir.Set 1, 0, 0
        dNozzleLength = dSeat2toCenter
        Set objNozzle = CreateNozzleWithLength(2, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Create Port 3
        oPlacePoint.Set 0, dSeat3toCenter, 0
        oDir.Set 0, 1, 0
        dNozzleLength = dSeat3toCenter
        Set objNozzle = CreateNozzleWithLength(3, oPartFclt, m_OutputColl, oDir, _
                                               oPlacePoint, dNozzleLength)
        
        'Set the Output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), objNozzle
        Set objNozzle = Nothing
        
        'Create the vortex shape
        Dim objVortexBody As Object
        Dim oSurfset   As IngrGeom3D.IJElements
        Dim iCount As Integer
        
        'Create the Sweep Curve
        RetrieveParameters 1, oPartFclt, m_OutputColl, pipeDiam, _
                                    flangeThick, flangeDiam, cptoffset, depth3
        Dim oCollection As Collection
        Dim oTraceStr As IngrGeom3D.ComplexString3d
        Dim dArcRadius As Double
        Dim stnorm() As Double
        Dim ednorm() As Double
        
        Set oCollection = New Collection
        Set oTraceStr = New ComplexString3d
        
        dArcRadius = pipeDiam / 2 + pipeDiam3 / 2
        
        'Create the Crosssection objects
        Dim oTrObj As IJElements
        Dim oCsObj As IJElements
        
        Set oTrObj = New JObjectCollection
        Set oCsObj = New JObjectCollection
        
        oCenter.Set 0, dSeat3toCenter, 0
        oVector.Set 0, 1, 0
        Set oCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, _
                    oCenter.x, oCenter.y, oCenter.z, oVector.x, oVector.y, oVector.z, pipeDiam3 / 2)
        oCsObj.Add oCircle
        
        'Create the Trace String
        'Line
        oStPoint.Set 0, dSeat3toCenter, -pipeDiam3 / 2
        oEnPoint.Set 0, dSeat3toCenter - dArcRadius, -pipeDiam3 / 2
        Set oLine = PlaceTrLine(oStPoint, oEnPoint)
        oCollection.Add oLine
        Set oLine = Nothing
        
        'Arc
        oCenter.Set 0, 0, -pipeDiam3 / 2
        oVector.Set -1, 0, 0
        oStPoint.Set 0, dSeat3toCenter - dArcRadius, -pipeDiam3 / 2
        oEnPoint.Set 0, 0, dArcRadius - pipeDiam3 / 2
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, oCenter.x, oCenter.y, oCenter.z, _
                            oVector.x, oVector.y, oVector.z, oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
        oCollection.Add oArc
        Set oArc = Nothing
        
        oStPoint.Set 0, dSeat3toCenter, -pipeDiam3 / 2
        Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
        oTrObj.Add oTraceStr
        
        'Create the Trace String
        For iCount = 1 To oCollection.Count
            oCollection.Remove 1
        Next iCount
        
        'Line
        oStPoint.Set 0, dSeat3toCenter, pipeDiam3 / 2
        oEnPoint.Set 0, Sqr((pipeDiam / 2) ^ 2 - (pipeDiam3 / 2) ^ 2), pipeDiam3 / 2
        Set oLine = PlaceTrLine(oStPoint, oEnPoint)
        oCollection.Add oLine
        Set oLine = Nothing
        
        'Arc
        oCenter.Set 0, 0, 0
        oVector.Set -1, 0, 0
        oStPoint.Set 0, oEnPoint.y, pipeDiam3 / 2
        oEnPoint.Set 0, 0, pipeDiam / 2
        
        Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, oCenter.x, oCenter.y, oCenter.z, _
                            oVector.x, oVector.y, oVector.z, oStPoint.x, oStPoint.y, oStPoint.z, _
                            oEnPoint.x, oEnPoint.y, oEnPoint.z)
        oCollection.Add oArc
        Set oArc = Nothing
        
        oStPoint.Set 0, dSeat3toCenter, pipeDiam3 / 2
        Set oTraceStr = PlaceTrCString(oStPoint, oCollection)
        oTrObj.Add oTraceStr
        
        Set oSurfset = oGeomFactory.GeometryServices.CreateBySkinning(m_OutputColl.ResourceManager, _
                        oTrObj, oCsObj, False)
        
        'Set the Output
        For Each objVortexBody In oSurfset
            If Not objVortexBody Is Nothing Then
                m_OutputColl.AddOutput "VortexBody", objVortexBody
            End If
        Next objVortexBody

        Set objVortexBody = Nothing
        Set oCircle = Nothing
        oTraceStr.RemoveCurve True
        Set oTraceStr = Nothing
        For iCount = 1 To oSurfset.Count
            oSurfset.Remove 1
        Next iCount
        Set oSurfset = Nothing
        For iCount = 1 To oCollection.Count
            oCollection.Remove 1
        Next iCount
        Set oCollection = Nothing
        
    End Select
    
    Set oGeomFactory = Nothing
    Set oCenter = Nothing
    Set oStPoint = Nothing
    Set oEnPoint = Nothing
    Set oArc = Nothing
    Set oLine = Nothing
    Set oCircle = Nothing
    Set ocv = Nothing
    Set oCP = Nothing
    Set oPlacePoint = Nothing
    Set oDir = Nothing
    
    Exit Sub

ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD

End Sub
