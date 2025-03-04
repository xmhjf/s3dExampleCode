VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:04:46 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2004-06, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:         MS
'   Creation Date:  Thursday, May 13 2004
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This Symbol detail is based on the PDS Equipment Modeling User's Guide,
'    E305 Symbol in Page no 299.  Physical and Insulation aspect outputs are variable outputs.
'
'   Change History:
'   dd.mmm.yyyy     who                change description
'   -----------     ---                ------------------
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006     kkc       DI 95670-Replaced names with initials in the revision history.
'   26.Sep.2006     dkl       TR-84724 If variables of type double are compared, appropriate CmpDblXXX functions are used.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Maintenance:" 'Used for error messages

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
    Dim ObjBundlePulling As Object
    
    Dim parExchangerLength As Double
    Dim parChannelDiameter As Double
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
    Dim parFrontEndLength1 As Double
    Dim parRearEndFlangeDia As Double
        
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parExchangerLength = arrayOfInputs(2)                       'P1
    parChannelDiameter = arrayOfInputs(3)                     'P2
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
    parFrontEndLength1 = arrayOfInputs(22)          'P32
    parRearEndFlangeDia = arrayOfInputs(23)        'P40
    parInsulationThickness = arrayOfInputs(24)
    
    iOutput = 0
    
   If CmpDblEqual(parBundlePullingLength, LINEAR_TOLERANCE) Then parBundlePullingLength = parExchangerLength
' Insert your code for output 8(BundlePulling)
    Dim startPoint As New AutoMath.DPosition    'P32-P3
    Dim endPoint As New AutoMath.DPosition      'P10-(P32-P3)
    
    startPoint.Set (parFrontEndLength1), _
                            CenterPos.y, _
                            CenterPos.z
    endPoint.Set -(parBundlePullingLength - parFrontEndLength1), _
                            CenterPos.y, _
                            CenterPos.z
                                
    Set ObjBundlePulling = PlaceCylinder(m_OutputColl, startPoint, endPoint, parChannelDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBundlePulling
    Set ObjBundlePulling = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Set CenterPos = Nothing
    
    Exit Sub
    
ErrorLabel:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
