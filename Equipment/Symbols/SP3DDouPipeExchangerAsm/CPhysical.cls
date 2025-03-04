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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:03:14 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-06, Intergraph Corporation. All rights reserved.
'
'   CPhysical.cls
'   Author:          MS
'   Creation Date:  Thursday, Apr 25 2002
'   Description:
'    This class module is the place for user to implement graphical part of VBSymbol for this aspect
'    This is E320 Equipment Assembly Double Pipe Exchanger symbol.
'    Physical Aspect consist of Exchanger Body, Two pipes, Block and Two Supports
'    and Two Nozzles.
'    Maintenance Aspect is for Bundle Pulling area.
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO             Made compatible with Smart Occurrence based Equipments
'   11.May.2005     kkk         CR-76070: Removed the Nozzle related parameters and code to create nozzles.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'   27.Sep.2006    RRK  TR-84724  If variables of type double are compared, appropriate CmpDblXXX functions are used.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit

Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()

'''

End Sub


Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    
    Dim iOutput     As Double
    Dim ObjPipe1 As Object
    Dim ObjPipe2 As Object
    Dim ObjBlock As Object
    Dim ObjSupport1 As Object
    Dim ObjSupport2 As Object
    
    Dim m_oGBSFactory  As IJGeneralBusinessObjectsFactory
    Dim m_oControlPoint As IJControlPoint
    
    Dim parBlockEndFromPP As Double 'P1
    Dim parPipeEndFromPP As Double 'P2
    Dim parPipeCentoCen As Double 'P3
    Dim parPipeDiameter As Double 'P4
    Dim parBlockLength As Double 'P5
    Dim parBlockHeight As Double 'P6
    Dim parBlockWidth As Double 'P7
    Dim parBundlePullingLength As Double 'P8
    Dim parSupport1CenFromPP As Double 'P9
    Dim parSupportCentoCen As Double 'P10
    Dim parSupportHeight As Double 'P11
    Dim parSupportThickness As Double 'P12
    Dim parSupportWidth As Double 'P13
    

' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parBlockEndFromPP = arrayOfInputs(2)
    parPipeEndFromPP = arrayOfInputs(3)
    parPipeCentoCen = arrayOfInputs(4)
    parPipeDiameter = arrayOfInputs(5)
    parBlockLength = arrayOfInputs(6)
    parBlockHeight = arrayOfInputs(7)
    parBlockWidth = arrayOfInputs(8)
    parBundlePullingLength = arrayOfInputs(9)
    parSupport1CenFromPP = arrayOfInputs(10)
    parSupportCentoCen = arrayOfInputs(11)
    parSupportHeight = arrayOfInputs(12)
    parSupportThickness = arrayOfInputs(13)
    parSupportWidth = arrayOfInputs(14)
    
    Dim geomFactory     As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
    
    Dim stPoint   As New AutoMath.DPosition
    Dim enPoint   As New AutoMath.DPosition
    
    iOutput = 0
    
 ' Insert your code for output 1(Pipe1)
    stPoint.Set -parPipeEndFromPP, 0, parPipeCentoCen / 2
    enPoint.Set parBlockEndFromPP - parBlockLength, 0, parPipeCentoCen / 2
    
    Set ObjPipe1 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parPipeDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipe1
    Set ObjPipe1 = Nothing
    
 ' Insert your code for output 2(Pipe2)
    stPoint.Set -parPipeEndFromPP, 0, -parPipeCentoCen / 2
    enPoint.Set parBlockEndFromPP - parBlockLength, 0, -parPipeCentoCen / 2
    
    Set ObjPipe2 = PlaceCylinder(m_OutputColl, stPoint, enPoint, parPipeDiameter, True)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPipe2
    Set ObjPipe2 = Nothing
    
 ' Insert your code for output 3(Block)
    stPoint.Set parBlockEndFromPP - parBlockLength, parBlockWidth / 2, parBlockHeight / 2
    enPoint.Set parBlockEndFromPP, -parBlockWidth / 2, -parBlockHeight / 2
    
    Set ObjBlock = PlaceBox(m_OutputColl, stPoint, enPoint)
    
' Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBlock
    Set ObjBlock = Nothing
    
'P12  is defined then place Support

    If (CmpDblGreaterthan(parSupportThickness, LINEAR_TOLERANCE)) Then
        Dim lineStrPoints(0 To 11)  As Double
        Dim ObjDefsur1  As IngrGeom3D.Plane3d
        Dim ObjDefsur2  As IngrGeom3D.Plane3d
        'P9 is defined Place First Support
            If (CmpDblGreaterthan(parSupport1CenFromPP, LINEAR_TOLERANCE)) Then
                stPoint.Set parSupport1CenFromPP - parSupportThickness / 2, _
                            parSupportWidth / 2, _
                            parSupportHeight / 2
                enPoint.Set parSupport1CenFromPP + parSupportThickness / 2, _
                            -parSupportWidth / 2, _
                            -parSupportHeight / 2
    
                Set ObjSupport1 = PlaceBox(m_OutputColl, stPoint, enPoint)
            ' Create  DefaultSurface for Support1
                
                lineStrPoints(0) = parSupport1CenFromPP - parSupportThickness / 2
                lineStrPoints(1) = parSupportWidth / 2
                lineStrPoints(2) = -parSupportHeight / 2
        
                lineStrPoints(3) = parSupport1CenFromPP + parSupportThickness / 2
                lineStrPoints(4) = parSupportWidth / 2
                lineStrPoints(5) = -parSupportHeight / 2
            
                lineStrPoints(6) = parSupport1CenFromPP + parSupportThickness / 2
                lineStrPoints(7) = -parSupportWidth / 2
                lineStrPoints(8) = -parSupportHeight / 2
            
                lineStrPoints(9) = parSupport1CenFromPP - parSupportThickness / 2
                lineStrPoints(10) = -parSupportWidth / 2
                lineStrPoints(11) = -parSupportHeight / 2
        
                 Set ObjDefsur1 = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)
                
                ' Set the output
                m_OutputColl.AddOutput "Support_", ObjSupport1
                m_OutputColl.AddOutput "DefaultSur_", ObjDefsur1
                Set ObjSupport1 = Nothing
                Set ObjDefsur1 = Nothing
            End If              'P9 is Defined
            
            'P9 and P10 are defined Place 2nd Support
            If CmpDblGreaterthan(parSupport1CenFromPP, LINEAR_TOLERANCE) And CmpDblGreaterthan(parSupportCentoCen, LINEAR_TOLERANCE) Then
                stPoint.Set parSupport1CenFromPP + parSupportCentoCen - parSupportThickness / 2, _
                            parSupportWidth / 2, _
                            parSupportHeight / 2
                enPoint.Set parSupport1CenFromPP + parSupportCentoCen + parSupportThickness / 2, _
                            -parSupportWidth / 2, _
                            -parSupportHeight / 2
    
                Set ObjSupport2 = PlaceBox(m_OutputColl, stPoint, enPoint)
                
            ' Create  DefaultSurface for Support2
                lineStrPoints(0) = parSupport1CenFromPP + parSupportCentoCen - parSupportThickness / 2
                lineStrPoints(1) = parSupportWidth / 2
                lineStrPoints(2) = -parSupportHeight / 2
        
                lineStrPoints(3) = parSupport1CenFromPP + parSupportCentoCen + parSupportThickness / 2
                lineStrPoints(4) = parSupportWidth / 2
                lineStrPoints(5) = -parSupportHeight / 2
            
                lineStrPoints(6) = parSupport1CenFromPP + parSupportCentoCen + parSupportThickness / 2
                lineStrPoints(7) = -parSupportWidth / 2
                lineStrPoints(8) = -parSupportHeight / 2
            
                lineStrPoints(9) = parSupport1CenFromPP + parSupportCentoCen - parSupportThickness / 2
                lineStrPoints(10) = -parSupportWidth / 2
                lineStrPoints(11) = -parSupportHeight / 2
        
                 Set ObjDefsur2 = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, lineStrPoints)
                
                ' Set the output
                m_OutputColl.AddOutput "Support_", ObjSupport2
                m_OutputColl.AddOutput "DefaultSur_", ObjDefsur2
                Set ObjSupport2 = Nothing
                Set ObjDefsur2 = Nothing
            End If          'P9 and P10 are defined
        
    End If           'P12 are Defined


    
'==========================================
'Construction of  Control Point
'==========================================
    Set m_oGBSFactory = New GeneralBusinessObjectsFactory
    Set m_oControlPoint = m_oGBSFactory.CreateControlPoint(m_OutputColl.ResourceManager, 0, 0, 0, 0, , , True)

    m_oControlPoint.Type = cpControlPoint
    m_oControlPoint.SubType = cpProcessEquipment
    m_OutputColl.AddOutput "DouPipeExchangerControlPoint", m_oControlPoint
    Set m_oControlPoint = Nothing
    Set m_oGBSFactory = Nothing
    
    
    Set geomFactory = Nothing
    Set stPoint = Nothing
    Set enPoint = Nothing
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    
    
End Sub

