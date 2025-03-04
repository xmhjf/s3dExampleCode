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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:12:37 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:         ACM
'   Creation Date:  Wednesday, Feb 25 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This Symbol detail is taken from PDS Equipment Modeling User's Guide,
'    E305 Symbol in Page no 299.  Physical and Insulation aspect outputs are variable outputs.
'
'   Change History:
'   dd.mmm.yyyy     who               change description
'   -----------     ---               ------------------
'   29.Nov.2004     V6UpgradeSO       Made compatible with Smart Occurrence based Equipments
'   14.DEC.2004     MS/BBA            TR 68409 Corrected the top first support endpoint location to have P13 and not p14
'   26.Apr.2005     svsmylav          TR-75775 End point coordinates are modified so that 'PlaceBox' function
'                                     places box with the surface normals outward.
'   11.Jul.2006     kkc               DI 95670-Replaced names with initials in the revision history.
'   26.Sep.2006     dkl               TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
'   17.Sep.2007     PK                TR-117405 Modified code to ensure that symbol does not fail when
'                                     P34,P35,P43 and P44 (flange thicknesses) are made zero.
'   25.Aug.2008     VKC               TR-110721 At the rear end RearEndFlangeTk2 is changed to RearEndFlangeTk3
'   30.Sep.2008     VKC               TR-151340 Message Box Removed.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    
    Dim parExchangerLength As Double
    Dim parExchangerDiameter As Double
    Dim parBundleFlangeTk As Double
    Dim parBundleFlangeDia As Double
    Dim parExchangerFlangeTk1 As Double
    Dim parExchangerFlangeTk2 As Double
    Dim parExpansionJointPosition As Double
    Dim parExpansionJointThickness As Double
    Dim parExpansionJointDiameter As Double
    Dim parBundlePullingLength As Double
    Dim parBotSupportCenFromPP As Double
    Dim parBottomSupportCentoCen As Double
    Dim parSupport1Thickness As Double
    Dim parSupport2Thickness As Double
    Dim parBottomSupportHeight As Double
    Dim parSupportLength As Double
    Dim parTopSupportCenFromPP As Double
    Dim parTopSupportCentoCen As Double
    Dim parTopSupportHeight As Double
    Dim parFrontEndFlangeDia As Double
    Dim parFrontEndFlangeTk1 As Double
    Dim parFrontEndLength1 As Double
    Dim parFrontEndLength2 As Double
    Dim parFrontEndFlangeTk2 As Double
    Dim parFrontEndFlangeTk3 As Double
    Dim parRearEndFlangeDia As Double
    Dim parRearEndFlangeTk1 As Double
    Dim parRearEndLength As Double
    Dim parRearEndFlangeTk2 As Double
    Dim parRearEndFlangeTk3 As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parExchangerLength = arrayOfInputs(2)                       'P1
    parExchangerDiameter = arrayOfInputs(3)                     'P2
    parBundleFlangeTk = arrayOfInputs(4)                            'P3
    parBundleFlangeDia = arrayOfInputs(5)                           'P4
    parExchangerFlangeTk1 = arrayOfInputs(6)                    'P5
    parExchangerFlangeTk2 = arrayOfInputs(7)                    'P6
    parExpansionJointPosition = arrayOfInputs(8)                'P7
    parExpansionJointThickness = arrayOfInputs(9)               'P8
    parExpansionJointDiameter = arrayOfInputs(10)               'P9
    parBundlePullingLength = arrayOfInputs(11)                  'P10
    parBotSupportCenFromPP = arrayOfInputs(12)              'P11
    parBottomSupportCentoCen = arrayOfInputs(13)            'P12
    parSupport1Thickness = arrayOfInputs(14)                        'P13
    parSupport2Thickness = arrayOfInputs(15)                        'P14
    parBottomSupportHeight = arrayOfInputs(16)                  'P15
    parSupportLength = arrayOfInputs(17)                                'P16
    parTopSupportCenFromPP = arrayOfInputs(18)                  'P17
    parTopSupportCentoCen = arrayOfInputs(19)                   'P18
    parTopSupportHeight = arrayOfInputs(20)                         'P19
    parFrontEndFlangeDia = arrayOfInputs(21)                    'P30
    parFrontEndFlangeTk1 = arrayOfInputs(22)                    'P31
    parFrontEndLength1 = arrayOfInputs(23)          'P32
    parFrontEndLength2 = arrayOfInputs(24)          'P33
    parFrontEndFlangeTk2 = arrayOfInputs(25)      'P34
    parFrontEndFlangeTk3 = arrayOfInputs(26)      'P35
    parRearEndFlangeDia = arrayOfInputs(27)        'P40
    parRearEndFlangeTk1 = arrayOfInputs(28)         'P41
    parRearEndLength = arrayOfInputs(29)                'P42
    parRearEndFlangeTk2 = arrayOfInputs(30)             'P43
    parRearEndFlangeTk3 = arrayOfInputs(31)            'P44
    parInsulationThickness = arrayOfInputs(32)
   
    iOutput = 0
       
    
''Origin is at DP/PP1( Channel Inlet )
    
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    If CmpDblEqual(parBundlePullingLength, LINEAR_TOLERANCE) Then parBundlePullingLength = parExchangerLength

 ' Insert your code for output 1(Exchanger Body Cylinder)
    Dim oStartPoint As New AutoMath.DPosition
    Dim oEndPoint As New AutoMath.DPosition
    Dim objExchangerCyl As Object
    
    oStartPoint.Set parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1, 0, 0
    oEndPoint.Set parFrontEndLength1 + parExchangerLength - parExchangerFlangeTk2, 0, 0
    
    Set objExchangerCyl = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parExchangerDiameter, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), objExchangerCyl
    Set objExchangerCyl = Nothing
                            
 ' Insert your code for output 2(Exchanger Rear Flange)
    Dim ObjExchangerRearFlange As Object
    
    oStartPoint.Set parFrontEndLength1 + parExchangerLength - parExchangerFlangeTk2, 0, 0
    oEndPoint.Set parFrontEndLength1 + parExchangerLength, 0, 0
    
    Set ObjExchangerRearFlange = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parRearEndFlangeDia, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerRearFlange
    Set ObjExchangerRearFlange = Nothing
 
 ' Insert your code for output 3(Exchanger Front Flange1)
    Dim ObjExchangerFrontFlange1 As Object
    
    oStartPoint.Set parFrontEndLength1 + parBundleFlangeTk + parExchangerFlangeTk1, 0, 0
    oEndPoint.Set parFrontEndLength1 + parBundleFlangeTk, 0, 0
    
    Set ObjExchangerFrontFlange1 = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parRearEndFlangeDia, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerFrontFlange1
    Set ObjExchangerFrontFlange1 = Nothing
 
 ' Insert your code for output 4(Exchanger Front Flange2)
    Dim ObjExchangerFrontFlange2 As Object
    
    oStartPoint.Set parFrontEndLength1 + parBundleFlangeTk, 0, 0
    oEndPoint.Set parFrontEndLength1, 0, 0
    
    Set ObjExchangerFrontFlange2 = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parBundleFlangeDia, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchangerFrontFlange2
    Set ObjExchangerFrontFlange2 = Nothing
 
  ' Insert your code for output 5(ExchFrontEndFlange1)
    Dim ObjExchFrontEndFlange1 As Object
    
    oStartPoint.Set parFrontEndLength1, 0, 0
    oEndPoint.Set parFrontEndLength1 - parFrontEndFlangeTk1, 0, 0
    
    Set ObjExchFrontEndFlange1 = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parFrontEndFlangeDia, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchFrontEndFlange1
    Set ObjExchFrontEndFlange1 = Nothing
  
  ' Insert your code for output 6(Exchanger Front End Cylinder)
    Dim ObjExchFrontEndCyl As Object
    
    oStartPoint.Set parFrontEndLength1 - parFrontEndFlangeTk1, 0, 0
    oEndPoint.Set -parFrontEndLength2 + parFrontEndFlangeTk3, 0, 0
    
    Set ObjExchFrontEndCyl = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parExchangerDiameter, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchFrontEndCyl
    Set ObjExchFrontEndCyl = Nothing
    
  'The If condition is required here to ensure that the symbol occurence does not fail when
  'parFrontEndFlangeTk3(P35) or parFrontEndFlangeTk2(P34) become zero
  
    ' Insert your code for output 7(ExchFrontEndFlange2)
    If CmpDblLessThanOrEqualTo(parFrontEndFlangeTk3, 0) Then
        iOutput = iOutput + 1
    Else
        Dim ObjExchFrontEndFlange2 As Object
        
        oStartPoint.Set -parFrontEndLength2 + parFrontEndFlangeTk3, 0, 0
        oEndPoint.Set -parFrontEndLength2, 0, 0
        
        Set ObjExchFrontEndFlange2 = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parFrontEndFlangeDia, True)
        
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchFrontEndFlange2
        Set ObjExchFrontEndFlange2 = Nothing
    End If
 
    ' Insert your code for output 8(ExchFrontEndFlange3)
    If CmpDblLessThanOrEqualTo(parFrontEndFlangeTk2, 0) Then
        iOutput = iOutput + 1
    Else
        Dim ObjExchFrontEndFlange3 As Object
        
        oStartPoint.Set -parFrontEndLength2, 0, 0
        oEndPoint.Set -parFrontEndLength2 - parFrontEndFlangeTk2, 0, 0
        
        Set ObjExchFrontEndFlange3 = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parFrontEndFlangeDia, True)
        
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchFrontEndFlange3
        Set ObjExchFrontEndFlange3 = Nothing
    End If
 
  ' Insert your code for output 9(ExchRearEndFlange1)
    Dim ObjExchRearEndFlange1 As Object
    
    oStartPoint.Set parFrontEndLength1 + parExchangerLength, 0, 0
    oEndPoint.Set parFrontEndLength1 + parExchangerLength + parRearEndFlangeTk1, 0, 0
    
    Set ObjExchRearEndFlange1 = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parRearEndFlangeDia, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchRearEndFlange1
    Set ObjExchRearEndFlange1 = Nothing
 
  ' Insert your code for output 10(ExchRearEndCyl)
    Dim ObjExchRearEndCyl As Object
    
    oStartPoint.Set parFrontEndLength1 + parExchangerLength + parRearEndFlangeTk1, 0, 0
    oEndPoint.Set parFrontEndLength1 + parExchangerLength + parRearEndLength - parRearEndFlangeTk3, 0, 0
    
    Set ObjExchRearEndCyl = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parExchangerDiameter, True)
    
    ' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchRearEndCyl
    Set ObjExchRearEndCyl = Nothing
    
  'The If condition is required here to ensure that the symbol occurence does not fail when
  'parRearEndFlangeTk3(P44) or parRearEndFlangeTk2(P43) become zero
  
    ' Insert your code for output 11(ExchRearEndFlange2)
    If CmpDblLessThanOrEqualTo(parRearEndFlangeTk3, 0) Then
        iOutput = iOutput + 1
    Else
        Dim ObjExchRearEndFlange2 As Object
        
        oStartPoint.Set parFrontEndLength1 + parExchangerLength + parRearEndLength - parRearEndFlangeTk3, 0, 0
        oEndPoint.Set parFrontEndLength1 + parExchangerLength + parRearEndLength, 0, 0
        
        Set ObjExchRearEndFlange2 = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parRearEndFlangeDia, True)
        
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchRearEndFlange2
        Set ObjExchRearEndFlange2 = Nothing
    End If
    
    ' Insert your code for output 12(ExchRearEndFlange3)
    If CmpDblLessThanOrEqualTo(parRearEndFlangeTk2, 0) Then
        iOutput = iOutput + 1
    Else
        Dim ObjExchRearEndFlange3 As Object
        
        oStartPoint.Set parFrontEndLength1 + parExchangerLength + parRearEndLength, 0, 0
        oEndPoint.Set parFrontEndLength1 + parExchangerLength + parRearEndLength + parRearEndFlangeTk3, 0, 0
        
        Set ObjExchRearEndFlange3 = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parRearEndFlangeDia, True)
        
        ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjExchRearEndFlange3
        Set ObjExchRearEndFlange3 = Nothing
    End If
    
    'Expansion Joint
    'P7 and P8 are defined provide expansion Joint
    If CmpDblGreaterthan(parExpansionJointPosition, LINEAR_TOLERANCE) And CmpDblGreaterthan(parExpansionJointThickness, LINEAR_TOLERANCE) Then
         Dim ObjExpanJt  As Object
        
         oStartPoint.Set parExpansionJointPosition, CenterPos.y, CenterPos.z
         oEndPoint.Set parExpansionJointPosition + parExpansionJointThickness, CenterPos.y, CenterPos.z
        
         Set ObjExpanJt = PlaceCylinder(m_OutputColl, oStartPoint, oEndPoint, parExpansionJointDiameter, True)
        
        ' Set the output
         m_OutputColl.AddOutput "Expansion_", ObjExpanJt
        
         Set ObjExpanJt = Nothing
    End If
    
    'P16  is defined then place Supports
    If CmpDblGreaterthan(parSupportLength, LINEAR_TOLERANCE) Then

''        'P15 is defined Place Bottom Support
        If CmpDblGreaterthan(parBottomSupportHeight, LINEAR_TOLERANCE) Then
''            'P11 is defined Place First Bottom Support
            If CmpDblGreaterthan(parBotSupportCenFromPP, LINEAR_TOLERANCE) Then
                Dim ObjSupport1 As Object
                
                oStartPoint.Set parBotSupportCenFromPP - parSupport1Thickness / 2, _
                                        -(parSupportLength / 2), _
                                        -parBottomSupportHeight
                oEndPoint.Set oStartPoint.x + parSupport1Thickness, _
                                        oStartPoint.y + parSupportLength, _
                                        oStartPoint.z + parBottomSupportHeight
                                        
                Set ObjSupport1 = PlaceBox(m_OutputColl, oStartPoint, oEndPoint)
                
                ' Set the output
                m_OutputColl.AddOutput "Support_", ObjSupport1
                Set ObjSupport1 = Nothing
                       
            End If              'P11 is Defined
                        
            'P12 is defined Place 2nd Bottom Support
            If CmpDblGreaterthan(parBottomSupportCentoCen, LINEAR_TOLERANCE) Then
                ' Insert your code for output 3(Support2) Bottom Support
                
                Dim ObjSupport2 As Object
                
                oStartPoint.Set (parBotSupportCenFromPP + parBottomSupportCentoCen - parSupport2Thickness / 2), _
                                       -(parSupportLength / 2), _
                                       -parBottomSupportHeight
                oEndPoint.Set oStartPoint.x + parSupport2Thickness, _
                                        oStartPoint.y + parSupportLength, _
                                        oStartPoint.z + parBottomSupportHeight
                                        
                Set ObjSupport2 = PlaceBox(m_OutputColl, oStartPoint, oEndPoint)
                
                ' Set the output
                m_OutputColl.AddOutput "Support_", ObjSupport2
                Set ObjSupport2 = Nothing
                
            End If          '' P12 is defined
            
        End If          'P15 is Defined
        ' Insert your code for output 4(Support3) Top Support
        'P19 is Defined Provide Top Support
        If CmpDblGreaterthan(parTopSupportHeight, LINEAR_TOLERANCE) Then
            'if P17 is defined then provide first Top Support
            If CmpDblGreaterthan(parTopSupportCenFromPP, LINEAR_TOLERANCE) Then
                Dim ObjSupport3 As Object
                
                oStartPoint.Set (parTopSupportCenFromPP - parSupport1Thickness / 2), _
                                       -(parSupportLength / 2), _
                                       0
                oEndPoint.Set oStartPoint.x + parSupport1Thickness, _
                                        oStartPoint.y + parSupportLength, _
                                        oStartPoint.z + parTopSupportHeight
                                        
                Set ObjSupport3 = PlaceBox(m_OutputColl, oStartPoint, oEndPoint)
                
                ' Set the output
                m_OutputColl.AddOutput "Support_", ObjSupport3
                Set ObjSupport3 = Nothing
                
            End If
            'P18 is defined Provide second Top Support
            If CmpDblGreaterthan(parTopSupportCentoCen, LINEAR_TOLERANCE) Then
                ' Insert your code for output 5(Support4)
                
                Dim ObjSupport4 As Object
                
                oStartPoint.Set (parTopSupportCenFromPP + parTopSupportCentoCen - parSupport2Thickness / 2), _
                                       -(parSupportLength / 2), _
                                       0
                oEndPoint.Set oStartPoint.x + parSupport2Thickness, _
                                        oStartPoint.y + parSupportLength, _
                                        oStartPoint.z + parTopSupportHeight
                                        
                Set ObjSupport4 = PlaceBox(m_OutputColl, oStartPoint, oEndPoint)
                
                ' Set the output
                m_OutputColl.AddOutput "Support_", ObjSupport4
                Set ObjSupport4 = Nothing
                                
            End If                              'P18 is defined
        End If                                      'P19 is defined
    End If                                           'P16 is Defined
    
    Set oGeomFactory = Nothing
    Set oStartPoint = Nothing
    Set oEndPoint = Nothing
    Set CenterPos = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext

End Sub

