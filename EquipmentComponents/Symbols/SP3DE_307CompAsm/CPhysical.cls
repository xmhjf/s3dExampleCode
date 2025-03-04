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
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:13:37 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004 Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         svsmylav
'   Creation Date:  Thursday, Apr 1 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is Kettle Exchanger Component symbol.
'    Symbol details are taken from PDS Equipment Modeling User's Guide,
'    E307 Symbol in Page no 304. Exchanger End E319 -type A/C/D/N in Page no 310 is taken.
'    Symbol uses variable outputs for supports.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------     -----   ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages
Private PI As Double
Private Sub Class_Initialize()
    PI = 4 * Atn(1)
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjExchangerRHSEnd As Object
    Dim ObjExchangerBody As Object
    Dim ObjExTaperBody As Object
    Dim ObjExneckBody As Object
    Dim ObjExchangerFlange As Object
    Dim ObjBundleFlange As Object
    Dim ObjExchFrontEndRHSFlange As Object
    Dim ObjExFrontEndBody As Object
    Dim ObjExFrontEndBodyIFlange As Object
    Dim ObjExchFrontEndLHSFlange As Object
    Dim ObjExchangerSup1 As Object
    Dim ObjExchangerSup2 As Object
    Dim ObjDefSurface2 As Object

    Dim parExchangerLength As Double
    Dim parExchangerNeckLength As Double
    Dim parExchangerNeckDiameter As Double
    Dim parExchangerTaperLength As Double
    Dim parExchangerDiameter As Double
    Dim parBundleFlangeTk As Double
    Dim parBundleFlangeDia As Double
    Dim parExchangerFlangeTk1 As Double
    Dim parBundlePullingLength As Double
    Dim parBotSupportCenFromPP As Double
    Dim parBottomSupportCentoCen As Double
    Dim parSupport1Thickness As Double
    Dim parSupport2Thickness As Double
    Dim parBottomSupportHeight As Double
    Dim parSupportLength As Double
    Dim parFrontEndFlangeDia As Double
    Dim parFrontEndFlangeTk1 As Double
    Dim parFrontEndLength1 As Double
    Dim parFrontEndLength2 As Double
    Dim parFrontEndFlangeTk2 As Double
    Dim parFrontEndFlangeTk3 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parExchangerLength = arrayOfInputs(2)           'P1
    parExchangerNeckLength = arrayOfInputs(3)       'P2
    parExchangerNeckDiameter = arrayOfInputs(4)     'P3
    parExchangerTaperLength = arrayOfInputs(5)      'P4
    parExchangerDiameter = arrayOfInputs(6)         'P5
    parBundleFlangeTk = arrayOfInputs(7)            'P6
    parBundleFlangeDia = arrayOfInputs(8)           'P7
    parExchangerFlangeTk1 = arrayOfInputs(9)        'P8
    parBundlePullingLength = arrayOfInputs(10)      'P9
    parBotSupportCenFromPP = arrayOfInputs(11)      'P10
    parBottomSupportCentoCen = arrayOfInputs(12)    'P11
    parSupport1Thickness = arrayOfInputs(13)        'P12
    parSupport2Thickness = arrayOfInputs(14)        'P13
    parBottomSupportHeight = arrayOfInputs(15)      'P14
    parSupportLength = arrayOfInputs(16)            'P15
    parFrontEndFlangeDia = arrayOfInputs(17)        'P30
    parFrontEndFlangeTk1 = arrayOfInputs(18)        'P31
    parFrontEndLength1 = arrayOfInputs(19)          'P32
    parFrontEndLength2 = arrayOfInputs(20)          'P33
    parFrontEndFlangeTk2 = arrayOfInputs(21)        'P34
    parFrontEndFlangeTk3 = arrayOfInputs(22)        'P35
    parInsulationThickness = arrayOfInputs(23)
    
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory

    iOutput = 0
    
'   Origin is taken on axis of Exchanger End E319, at the junction of P32 and P33 dimensions.

 ' Insert your code for output 1(ExchangerRHSEnd)
    'Place Elliptical Arc on Right Side
    Dim cenX As Double
    Dim cenY As Double
    Dim cenZ As Double
    Dim MajorX As Double
    Dim MajorY As Double
    Dim MajorZ As Double
    Dim mMRatio As Double
    Dim StartAngle As Double
    Dim SweepAngle As Double
    Dim norX As Double
    Dim norY As Double
    Dim norZ As Double

    Dim oEqpArc As IngrGeom3D.EllipticalArc3d
    Dim axisVect As New AutoMath.DVector
    Dim revCenPt As New AutoMath.DPosition
    
    cenX = parFrontEndLength1 + parExchangerLength
    cenY = 0
    cenZ = parExchangerDiameter / 2 - parExchangerNeckDiameter / 2
    
    MajorX = 0
    MajorY = 0
    MajorZ = parExchangerDiameter / 2
    
    mMRatio = 0.5
    StartAngle = 1.5 * PI
    SweepAngle = PI / 2
    
    norX = 0
    norY = -1
    norZ = 0
    
    Set oEqpArc = oGeomFactory.EllipticalArcs3d.CreateByCenterNormalMajAxisRatioAngle(Nothing, _
                                                        cenX, cenY, cenZ, norX, norY, norZ, MajorX, MajorY, MajorZ, mMRatio, _
                                                        StartAngle, SweepAngle)
    axisVect.Set 1, 0, 0
    revCenPt.Set cenX, cenY, cenZ

    'Revolve it about X-Axis
    Set ObjExchangerRHSEnd = PlaceRevolution(m_OutputColl, oEqpArc, axisVect, revCenPt, PI * 2, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerRHSEnd
    Set ObjExchangerRHSEnd = Nothing

 ' Insert your code for output 2(ExchangerBody)
    Dim stPoint As New AutoMath.DPosition
    Dim enPoint As New AutoMath.DPosition
    
    stPoint.Set parFrontEndLength1 + parExchangerLength, 0, _
                    (parExchangerDiameter / 2 - parExchangerNeckDiameter / 2)
    enPoint.Set parFrontEndLength1 + parExchangerNeckLength + parExchangerTaperLength, 0, _
                    (parExchangerDiameter / 2 - parExchangerNeckDiameter / 2)

    Set ObjExchangerBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                            parExchangerDiameter, False)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerBody
    Set ObjExchangerBody = Nothing
 
 ' Insert your code for output 3(ExchangerTaperBody)
    Dim topCenter As New AutoMath.DPosition
    Dim baseCenter As New AutoMath.DPosition
    Dim obaseCircle As IngrGeom3D.Circle3d
    Dim otopCircle As IngrGeom3D.Circle3d

    'X-coordinate: P2+P32
    baseCenter.Set (parExchangerNeckLength + parFrontEndLength1), 0, 0

    'X-coordinate: P32+P2+P4
    topCenter.Set (parExchangerNeckLength + parExchangerTaperLength + parFrontEndLength1), _
                            0, _
                            parExchangerDiameter / 2 - parExchangerNeckDiameter / 2

    Set obaseCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, baseCenter.x, _
                                            baseCenter.y, baseCenter.z, 1, 0, 0, parExchangerNeckDiameter / 2)

    Set otopCircle = oGeomFactory.Circles3d.CreateByCenterNormalRadius(Nothing, topCenter.x, _
                                            topCenter.y, topCenter.z, 1, 0, 0, parExchangerDiameter / 2)

    Set ObjExTaperBody = oGeomFactory.RuledSurfaces3d.CreateByCurves(m_OutputColl.ResourceManager, _
                                            obaseCircle, otopCircle, False)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExTaperBody
    Set ObjExTaperBody = Nothing
    Set obaseCircle = Nothing
    Set otopCircle = Nothing
    
 ' Insert your code for output 4(Exchanger Neck Portion)
    'X-coordinate: P32 + P2
    stPoint.Set (parExchangerNeckLength + parFrontEndLength1), 0, 0

    'X-coordinate: P32+P6+P8
    enPoint.Set (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1), _
                0, 0

    Set ObjExneckBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                            parExchangerNeckDiameter, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExneckBody
    Set ObjExneckBody = Nothing


 ' Insert your code for output 5(ExchangerFlange)
    'X-coordinate: P32+P6+P8
    stPoint.Set (parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1), 0, 0
   'X-coordinate: P32+P6
    enPoint.Set (parFrontEndLength1 + parBundleFlangeTk), 0, 0
                
    Set ObjExchangerFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                            parFrontEndFlangeDia, False)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerFlange
    Set ObjExchangerFlange = Nothing

 ' Insert your code for output 6(BundleFlange)
   'X-coordinate: P32+P6
    stPoint.Set (parFrontEndLength1 + parBundleFlangeTk), 0, 0
    
   'X-coordinate: P32
    enPoint.Set parFrontEndLength1, 0, 0
                
    Set ObjBundleFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                            parBundleFlangeDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBundleFlange
    Set ObjBundleFlange = Nothing

' Insert your code for output 7(ExchFrontEndRHSFlange)
   'X-coordinate: P32
    stPoint.Set parFrontEndLength1, 0, 0
    
   'X-coordinate: P32 - P31
    enPoint.Set parFrontEndLength1 - parFrontEndFlangeTk1, 0, 0
                
    Set ObjExchFrontEndRHSFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                             parFrontEndFlangeDia, True)
                            
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchFrontEndRHSFlange
    Set ObjExchFrontEndRHSFlange = Nothing

' Insert your code for output 8(ExFrontEndBody)
   'X-coordinate: P32 - P31
    stPoint.Set parFrontEndLength1 - parFrontEndFlangeTk1, 0, 0
    
   'X-coordinate: -P33 + P35
    enPoint.Set -parFrontEndLength2 + parFrontEndFlangeTk3, 0, 0
                
    Set ObjExFrontEndBody = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                             parExchangerNeckDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExFrontEndBody
    Set ObjExFrontEndBody = Nothing
    
' Insert your code for output 9(ExFrontEndBodyIFlange)
   'X-coordinate: -P33 + P35
    stPoint.Set -parFrontEndLength2 + parFrontEndFlangeTk3, 0, 0
    
   'X-coordinate: -P33
    enPoint.Set -parFrontEndLength2, 0, 0
    
    Set ObjExFrontEndBodyIFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                             parFrontEndFlangeDia, True)

' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExFrontEndBodyIFlange
    Set ObjExFrontEndBodyIFlange = Nothing
    
' Insert your code for output 10(ExchFrontEndLHSFlange)
   'X-coordinate: -P33
    stPoint.Set -parFrontEndLength2, 0, 0
    
   'X-coordinate: -P33 -P34
    enPoint.Set -parFrontEndLength2 - parFrontEndFlangeTk2, 0, 0
    
    Set ObjExchFrontEndLHSFlange = PlaceCylinder(m_OutputColl, stPoint, enPoint, _
                             parFrontEndFlangeDia, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchFrontEndLHSFlange
    Set ObjExchFrontEndLHSFlange = Nothing

 ' Insert your code for output (Exchanger Support)
    Dim oSupLine As IngrGeom3D.LineString3d
    Dim dblSupPts(0 To 14) As Double
    Dim projVect As New AutoMath.DVector
    Dim ActualProj As Double
    Dim Sup1L As IJDPosition
    Dim Sup1H As IJDPosition
    Set Sup1L = New DPosition
    Set Sup1H = New DPosition
    Dim SupLZ As Double
    Dim blnIsDefSurfaceCreated As Boolean
    blnIsDefSurfaceCreated = False
    If (parBottomSupportHeight > parExchangerNeckDiameter / 2 And parSupportLength > 0) Then
        If parSupport1Thickness > 0 Then
            SupLZ = parBottomSupportHeight - cenZ
            Sup1L.Set (parBotSupportCenFromPP - parSupport1Thickness / 2), _
                        -(parSupportLength / 2), -SupLZ

            Sup1H.Set (parBotSupportCenFromPP + parSupport1Thickness / 2), _
                      (parSupportLength / 2), 0
            Set ObjExchangerSup1 = PlaceBox(m_OutputColl, Sup1L, Sup1H)
            
            'Set the output
            m_OutputColl.AddOutput "Support1", ObjExchangerSup1
            Set ObjExchangerSup1 = Nothing
        End If
        
        ' Insert your code for output (Exchanger Support2)
        
        If (parBottomSupportCentoCen > parSupport1Thickness And parSupport2Thickness > 0) Then
            'Point1
            dblSupPts(0) = (parBotSupportCenFromPP + parBottomSupportCentoCen - parSupport2Thickness / 2)
            dblSupPts(1) = -(parSupportLength / 2)
            dblSupPts(2) = -SupLZ
            'Point2
            dblSupPts(3) = dblSupPts(0)
            dblSupPts(4) = parSupportLength / 2
            dblSupPts(5) = -SupLZ
             'Point3
            dblSupPts(6) = (parBotSupportCenFromPP + parBottomSupportCentoCen + parSupport2Thickness / 2)
            dblSupPts(7) = (parSupportLength / 2)
            dblSupPts(8) = -SupLZ
            'Point4
            dblSupPts(9) = dblSupPts(6)
            dblSupPts(10) = dblSupPts(1)
            dblSupPts(11) = -SupLZ
            'Point5
            dblSupPts(12) = dblSupPts(0)
            dblSupPts(13) = dblSupPts(1)
            dblSupPts(14) = dblSupPts(2)

            projVect.Set 0, 0, 1
            ActualProj = SupLZ
            Set oSupLine = oGeomFactory.LineStrings3d.CreateByPoints(Nothing, 5, dblSupPts)
            Set ObjExchangerSup2 = PlaceProjection(m_OutputColl, oSupLine, projVect, ActualProj, True)
             
            ' Set the output
            m_OutputColl.AddOutput "Support2", ObjExchangerSup2
            Set ObjExchangerSup2 = Nothing
        End If
    End If
    
    Set oGeomFactory = Nothing
    Set axisVect = Nothing
    Set revCenPt = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Set topCenter = Nothing
    Set baseCenter = Nothing
    Set projVect = Nothing
    Set Sup1L = Nothing
    Set Sup1H = Nothing
    Set oEqpArc = Nothing
    Set oSupLine = Nothing
    Set ObjDefSurface2 = Nothing

    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.description, _
       Err.HelpFile, Err.HelpContext
    
End Sub

