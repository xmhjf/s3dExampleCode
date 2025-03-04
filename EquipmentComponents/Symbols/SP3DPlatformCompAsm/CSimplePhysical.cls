VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:15:02 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003-2004, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          MS
'   Creation Date:  Friday, 2 May 2003
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy   who                change description
'   -----------   -----              ------------------
'   09.Jul.2003   SymbolTeam(India)  Copyright Information, Header  is added.
'   23.Aug.2003   SymbolTeam(India)  Updated Error Handling Code
'   23.Jun.2004   svsmylav           TR-60704: Replaced 'm_OutputColl.ResourceManager' with "Nothing'
'                                    while creating oTLine1, oTArc1, oTLine2 and oTArc2 transient outputs to
'                                    avoid extra curves/lines in the symbol graphics.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'   11.Jul.2006      kkc                    DI 95670-Replaced names with initials in the revision history.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private PI As Double
Private Const MODULE = "Physical:" 'Used for error messages
Private Sub Class_Initialize()
Const METHOD = "Class_Initialize:"
On Error GoTo Errx
    PI = Atn(1) * 4
    Exit Sub

Errx:
    Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub
Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim oPartFclt       As PartFacelets.IJDPart
    Dim iOutput     As Double


    Dim parVesselDiameter As Double
    Dim parPlatformWidth As Double
    Dim parPlatformHeight As Double
    Dim parPlatformAngle As Double
    Dim parClearancefromVessel As Double
    Dim parInsulationThickness As Double
    
' Inputs
    Set oPartFclt = arrayOfInputs(1)
    parVesselDiameter = arrayOfInputs(2)
    parPlatformWidth = arrayOfInputs(3)
    parPlatformHeight = arrayOfInputs(4)
    parPlatformAngle = arrayOfInputs(5)
    parClearancefromVessel = arrayOfInputs(6)
    parInsulationThickness = arrayOfInputs(7)
     
     iOutput = 0

'Case of 360 degrees.
    If (parPlatformAngle = 2 * PI) Then
        parPlatformAngle = 6.2831
    End If
    Dim geomFactory As IngrGeom3D.GeometryFactory
    Set geomFactory = New IngrGeom3D.GeometryFactory
          
'Creation of Bottom Platform
    'Create the lines and arcs of the top surface . use this lines and arcs to create the complexstring.
    'Use the complex string to project the thickness.
    Dim oComplexString     As New IngrGeom3D.ComplexString3d
    Dim oTLine1 As Object
    Dim oTArc2 As Object
    Dim oTLine2 As Object
    Dim oTArc1 As Object
    Dim ObjPlatformBody As Object
       
    Dim oCenterPoint As New AutoMath.DPosition
    Dim arrPt(1 To 4)       As New AutoMath.DPosition
    Dim oNormalVector   As IJDVector
    Dim oCollection          As Collection
    
    Set oNormalVector = New AutoMath.DVector
    Set oCollection = New Collection
    oNormalVector.Set 0, 0, 1
     
    oCenterPoint.Set 0, 0, 0
    Dim dOrigintoPlatform As Double
    dOrigintoPlatform = parVesselDiameter / 2 + parClearancefromVessel + parInsulationThickness
    arrPt(1).Set (dOrigintoPlatform) * Sin(parPlatformAngle / 2), _
                            (dOrigintoPlatform) * Cos(parPlatformAngle / 2), _
                           0
    arrPt(2).Set (dOrigintoPlatform + parPlatformWidth) * Sin(parPlatformAngle / 2), _
                            (dOrigintoPlatform + parPlatformWidth) * Cos(parPlatformAngle / 2), _
                           0
    arrPt(3).Set -(dOrigintoPlatform + parPlatformWidth) * Sin(parPlatformAngle / 2), _
                            (dOrigintoPlatform + parPlatformWidth) * Cos(parPlatformAngle / 2), _
                           0
    arrPt(4).Set -(dOrigintoPlatform) * Sin(parPlatformAngle / 2), _
                            (dOrigintoPlatform) * Cos(parPlatformAngle / 2), _
                           0
    'Top Line 1
    Set oTLine1 = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPt(1).x, arrPt(1).y, arrPt(1).z, arrPt(2).x, arrPt(2).y, arrPt(2).z)
    oCollection.Add oTLine1
    'Top Arc 2
    Set oTArc2 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                          oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                                         oNormalVector.x, oNormalVector.y, oNormalVector.z, _
                                                         arrPt(2).x, arrPt(2).y, arrPt(2).z, _
                                                        arrPt(3).x, arrPt(3).y, arrPt(3).z)
    oCollection.Add oTArc2
    'Top Line 2
    Set oTLine2 = geomFactory.Lines3d.CreateBy2Points(Nothing, arrPt(3).x, arrPt(3).y, arrPt(3).z, arrPt(4).x, arrPt(4).y, arrPt(4).z)
    oCollection.Add oTLine2
    'Top Arc 1
    oNormalVector.Set 0, 0, -1
    Set oTArc1 = geomFactory.Arcs3d.CreateByCtrNormStartEnd(Nothing, _
                                                      oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                                     oNormalVector.x, oNormalVector.y, oNormalVector.z, _
                                                     arrPt(4).x, arrPt(4).y, arrPt(4).z, _
                                                    arrPt(1).x, arrPt(1).y, arrPt(1).z)
    oCollection.Add oTArc1
    
    Set oComplexString = PlaceTrCString(arrPt(1), oCollection)
         
    'Project the complexstring to create the body of the platform
    Set ObjPlatformBody = PlaceProjection(m_OutputColl, oComplexString, oNormalVector, parPlatformHeight, True)
    
    'Set the output
    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPlatformBody
   Set ObjPlatformBody = Nothing
   'Remove the complex string
    Dim Objcurves As IJDObject
    Set Objcurves = oComplexString
    Objcurves.Remove
    Set oComplexString = Nothing
    
    Dim count As Integer
    For count = 1 To oCollection.count
        oCollection.Remove 1
    Next count
    Set oCollection = Nothing
    Set oCenterPoint = Nothing
    Set oNormalVector = Nothing
    Set oTLine1 = Nothing
    Set oTArc2 = Nothing
    Set oTLine2 = Nothing
    Set oTArc1 = Nothing
    For count = 1 To 4
        Set arrPt(count) = Nothing
    Next count
    Set geomFactory = Nothing
    Exit Sub
    
ErrorLabel:
        Err.Raise Err.Number, Err.Source & " " & METHOD, Err.Description, _
       Err.HelpFile, Err.HelpContext
End Sub

