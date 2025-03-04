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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:14:08 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-05, Intergraph Corporation. All rights reserved.
'
'   CMaintenance.cls
'   Author:          BG
'   Creation Date:  Thursday, Apr 18 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   11.May.2005     kkk         CR-76070: Removed the Nozzle related parameters.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
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
    Dim pipeDiam        As Double
    Dim flangeThick     As Double
    Dim sptOffset       As Double
    Dim flangeDiam      As Double
    Dim depth           As Double
    Dim CenterPos       As New AutoMath.DPosition

    CenterPos.Set 0, 0, 0
    
    Dim iOutput     As Double
    Dim ObjExchangerBody As Object
    Dim ObjExTaperBody As Object
    Dim ObjExneckBody As Object
    Dim ObjExFrontEnd As Object
    Dim ObjExchangerSup1 As Object
    Dim ObjExchangerSup2 As Object
    Dim ObjDefSurface1 As Object
    Dim ObjDefSurface2 As Object
    Dim ObjExBodyInsul As Object
    Dim ObjExTapBodyInsul As Object
    Dim ObjExneckBodyInsul As Object
    Dim ObjBundlePullCylin As Object
    
    
    Dim parExchangerLength  As Double
    Dim parExchangerNeckLength As Double
    Dim parExchangerNeckDiameter As Double
    Dim parExchangerTaperLength As Double
    Dim parExchangerDiameter  As Double
    Dim parBundleFlangeTk  As Double
    Dim parBundleFlangeDia  As Double
    Dim parExchangerFlangeTk1 As Double
    Dim parBundlePullingLength  As Double
    Dim parBotSupportCenFromPP  As Double
    Dim parBottomSupportCentoCen  As Double
    Dim parSupport1Thickness  As Double
    Dim parSupport2Thickness  As Double
    Dim parBottomSupportHeight  As Double
    Dim parSupportLength As Double
    Dim parFrontEndFlangeDia As Double
    Dim parFrontEndFlangeTk1 As Double
    Dim parFrontEndLength1  As Double
    Dim parFrontEndLength2 As Double
    Dim parFrontEndFlangeTk2 As Double
    Dim parFrontEndFlangeTk3 As Double
    Dim parInsulationThickness As Double

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parExchangerLength = arrayOfInputs(2)                           'P1
    parExchangerNeckLength = arrayOfInputs(3)                   'P2
    parExchangerNeckDiameter = arrayOfInputs(4)                 'P3
    parExchangerTaperLength = arrayOfInputs(5)                  'P4
    parExchangerDiameter = arrayOfInputs(6)                         'P5
    parBundleFlangeTk = arrayOfInputs(7)                                'P6
    parBundleFlangeDia = arrayOfInputs(8)                               'P7
    parExchangerFlangeTk1 = arrayOfInputs(9)                    'P8
    parBundlePullingLength = arrayOfInputs(10)                  'P9
    parBotSupportCenFromPP = arrayOfInputs(11)              'P10
    parBottomSupportCentoCen = arrayOfInputs(12)        'P11
    parSupport1Thickness = arrayOfInputs(13)                    'P12
    parSupport2Thickness = arrayOfInputs(14)                    'P13
    parBottomSupportHeight = arrayOfInputs(15)              'P14
    parSupportLength = arrayOfInputs(16)                        'P15
    parFrontEndFlangeDia = arrayOfInputs(17)                'P30
    parFrontEndFlangeTk1 = arrayOfInputs(18)                'P31
    parFrontEndLength1 = arrayOfInputs(19)                  'P32
    parFrontEndLength2 = arrayOfInputs(20)                  'P33
    parFrontEndFlangeTk2 = arrayOfInputs(21)                'P34
    parFrontEndFlangeTk3 = arrayOfInputs(22)                'P35
    parInsulationThickness = arrayOfInputs(23)
    
    iOutput = 0
    
 ' Insert your code for output 12(Bundle Pulling Cylinder)
    Dim startPoint As New AutoMath.DPosition
    Dim endPoint As New AutoMath.DPosition
    'P32
    
    startPoint.Set parFrontEndLength1, CenterPos.y, CenterPos.z
    'P9-P32
    endPoint.Set (parFrontEndLength1 - parBundlePullingLength), _
                            CenterPos.y, CenterPos.z
                            
    Set ObjBundlePullCylin = PlaceCylinder(m_OutputColl, startPoint, endPoint, _
                                                parExchangerNeckDiameter, True)
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBundlePullCylin
    Set ObjBundlePullCylin = Nothing
    Set startPoint = Nothing
    Set endPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
        
End Sub
