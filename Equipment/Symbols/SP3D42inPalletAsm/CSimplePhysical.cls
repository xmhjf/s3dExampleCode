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
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:00:33 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          HL
'   Creation Date:  Friday, Nov 22 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_OutputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
    
    Dim CenterPos           As IJDPosition
    Dim StartPos            As IJDPosition
    Dim EndPos              As IJDPosition
    
    Dim geomFactory         As New IngrGeom3D.GeometryFactory
    
    Set CenterPos = New DPosition
    Set StartPos = New DPosition
    Set EndPos = New DPosition
    
    CenterPos.Set 0, 0, 0
    
    Dim iOutput             As Double
    Dim iCount              As Integer
    
    Dim ObjTopSlat          As Object
    Dim ObjBottomSlat       As Object
    Dim ObjBrace            As Object
    Dim ObjPoint            As Object
    Dim ObjDefaultSurface   As Object
    
    Dim SlatWidth           As Double
    Dim SlatHeight          As Double
    Dim SlatLength          As Double
    Dim SlatSpacing         As Double
    
    Dim BraceWidth          As Double
    Dim BraceHeight         As Double
    Dim BraceLength         As Double
    Dim BraceSpacing        As Double
    
    Dim xOffset As Double
    Dim yOffset As Double
    Dim zOffset As Double
    
    iOutput = 0

    SlatWidth = 6 * 0.0254
    SlatHeight = 0.5 * 0.0254
    SlatLength = 42 * 0.0254
    SlatSpacing = 1.2 * 0.0254
    xOffset = 21 * 0.0254
    yOffset = 21 * 0.0254
    zOffset = 4.5 * 0.0254
    
' Insert your code for output 1-6(6 Top Slats)
' StartPos left bottom corner
' EndPos right Top corner

    For iCount = 1 To 6
        StartPos.Set 0 + (iCount - 1) * (SlatWidth + SlatSpacing) - xOffset, -yOffset, zOffset
        EndPos.Set StartPos.x + SlatWidth, StartPos.y + SlatLength, StartPos.z + SlatHeight
        Set ObjTopSlat = PlaceBox(m_OutputColl, StartPos, EndPos)
    
' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjTopSlat
     
        Set ObjTopSlat = Nothing
     Next iCount
   
' Insert your code for output 7-12(6 Bottom Slats)
 
    For iCount = 7 To 12
        StartPos.Set 0 + (iCount - 7) * (SlatWidth + SlatSpacing) - xOffset, -yOffset, 0
        EndPos.Set StartPos.x + SlatWidth, StartPos.y + SlatLength, StartPos.z + SlatHeight
        Set ObjBottomSlat = PlaceBox(m_OutputColl, StartPos, EndPos)

' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBottomSlat
    
        Set ObjBottomSlat = Nothing
     Next iCount
        
' Insert your code for output 13-15(3 Vertical Braces)
    BraceWidth = 2 * 0.0254
    BraceHeight = 4 * 0.0254
    BraceLength = 41 * 0.0254
    BraceSpacing = 17.5 * 0.0254
     
   For iCount = 13 To 15
        StartPos.Set 0.5 * 0.0254 - xOffset, 0.5 * 0.0254 + (iCount - 13) * (BraceWidth + BraceSpacing) - yOffset, SlatHeight
        EndPos.Set StartPos.x + BraceLength, StartPos.y + BraceWidth, StartPos.z + BraceHeight
        Set ObjBrace = PlaceBox(m_OutputColl, StartPos, EndPos)
    
    ' Set the output
        iOutput = iOutput + 1
        m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjBrace
       
        Set ObjBrace = Nothing
   Next iCount
     
' Insert your code for output 16(Point)
     
    Set geomFactory = New IngrGeom3D.GeometryFactory
    Set ObjPoint = geomFactory.Points3d.CreateByPoint(m_OutputColl.ResourceManager, CenterPos.x, CenterPos.y, CenterPos.z)
 
' Set the output

    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjPoint
    Set ObjPoint = Nothing


' Insert your code for output 17(Default Surface)
        
    Dim Xmin As Double, Ymin As Double, Zmin As Double
    Dim Xmax As Double, Ymax As Double, Zmax As Double
    Dim Points(0 To 11) As Double

    Xmin = -21 * 0.0254
    Ymin = -21 * 0.0254
    Zmin = 0
    Xmax = 21 * 0.0254
    Ymax = 21 * 0.0254
    Zmax = 0

    Points(0) = Xmin
    Points(1) = Ymin
    Points(2) = Zmin
    Points(3) = Xmin
    Points(4) = Ymax
    Points(5) = Zmax
    Points(6) = Xmax
    Points(7) = Ymax
    Points(8) = Zmax
    Points(9) = Xmax
    Points(10) = Ymin
    Points(11) = Zmin
    Set ObjDefaultSurface = geomFactory.Planes3d.CreateByPoints(m_OutputColl.ResourceManager, 4, Points)
    
' Set the output

    iOutput = iOutput + 1
    m_OutputColl.AddOutput arrayOfOutputs(iOutput), ObjDefaultSurface
    Set ObjDefaultSurface = Nothing
    
    Set geomFactory = Nothing
    Set StartPos = Nothing
    Set EndPos = Nothing
    Set CenterPos = Nothing
  
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub
