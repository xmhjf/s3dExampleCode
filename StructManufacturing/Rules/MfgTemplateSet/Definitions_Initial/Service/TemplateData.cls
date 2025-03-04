VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************************************************
'Copyright (C) 2001, Intergraph Limited. All rights reserved.
'
'Abstract:
'    TemplateData object encapsulates the attributes such as control point, template plane etc. of a Template.
'
'Description:
'History :
'   Ram Kandimalla            08/19/2005      Creation.
'******************************************************************************************************************

Option Explicit

Private Const MODULE = "GSCADStrMfgTemplate.TemplateData"

Public Enum enumTemplateType
    System = 0
    Sketched
    Edge                                        ' Treat edge as special case since intersection between
    SystemTemplateWithModifiedBL
    SketchedTemplateWithModifiedBL
    TemplateWithAddedBL
End Enum                                        ' frame and surface fails 90% ofthe time at the edge


Public ControlPoint As IJDPosition             ' intersection between BaseControlLine and DefiningObject
Public Plane As IJPlane                        ' Templateplane
Public TemplateType As enumTemplateType        ' Frame or Skecthed or on edge
Private m_oDefiningObj As Object                   ' FrameObject or SkecthingLineObject or Edge
Private m_oBottomLine As IUnknown
Private m_eGroupType As enumGroupType

Private Sub Class_Initialize()
    m_eGroupType = Primary 'initialization
    Set ControlPoint = Nothing
    Set Plane = Nothing
    Set DefiningObj = Nothing
    Set BottomLine = Nothing
End Sub

Private Sub Class_Terminate()
    Set ControlPoint = Nothing
    Set Plane = Nothing
    Set DefiningObj = Nothing
    Set BottomLine = Nothing
End Sub

Public Function IsValid() As Boolean
    IsValid = True
    If ControlPoint Is Nothing Then IsValid = False
    If Plane Is Nothing Then IsValid = False
    If DefiningObj Is Nothing Then IsValid = False
    If BottomLine Is Nothing Then IsValid = False
End Function

Public Property Set BottomLine(oBottomLine As IUnknown)
On Error Resume Next
    'user may reset to nothing
    If oBottomLine Is Nothing Then
        Set m_oBottomLine = Nothing
        Exit Property
    End If
    
    Dim oMfgMGHelper As IJMfgMGHelper
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Dim oCS As IJComplexString
    
    Set oMfgMGHelper = New MfgMGHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    
    ' always store as IJComplexString so code is simplified everywhere else
    If TypeOf oBottomLine Is IJComplexString Then
        Set m_oBottomLine = oMfgGeomHelper.ApproximateCurveByStrokeAndFit(oBottomLine)
    ElseIf TypeOf oBottomLine Is IJWireBody Then
        oMfgMGHelper.WireBodyToComplexString oBottomLine, oCS
        Set m_oBottomLine = oMfgGeomHelper.ApproximateCurveByStrokeAndFit(oCS)
    ElseIf TypeOf oBottomLine Is IJCurve Then
        Set m_oBottomLine = oBottomLine
    Else    ' we only have IJComplexString or IJWirebody or IJCurve
        Set m_oBottomLine = Nothing
    End If
       
CleanUp:
    Set oMfgMGHelper = Nothing
    Set oMfgGeomHelper = Nothing
    Set oCS = Nothing
    Exit Property
ErrorHandler:
    GoTo CleanUp
End Property

Public Property Get BottomLine() As IUnknown
    Set BottomLine = m_oBottomLine
End Property


Public Property Set DefiningObj(oDefiningObj As Object)
On Error GoTo ErrorHandler
    'user may reset to nothing
    If oDefiningObj Is Nothing Then
        Set m_oDefiningObj = Nothing
        Exit Property
    End If
    
    Dim oMfgMGHelper As IJMfgMGHelper
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Dim oCS As IJComplexString
    
    Set oMfgMGHelper = New MfgMGHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    If TypeOf oDefiningObj Is IHFrame Then
        Set m_oDefiningObj = oDefiningObj
    ElseIf TypeOf oDefiningObj Is IJPlane Then
        Set m_oDefiningObj = oDefiningObj
    ElseIf TypeOf oDefiningObj Is IJWireBody Then
        ' always store as IJComplexString so code is simplified everywhere else
        oMfgMGHelper.WireBodyToComplexString oDefiningObj, oCS
        Set m_oDefiningObj = oMfgGeomHelper.ApproximateCurveByStrokeAndFit(oCS)
    ElseIf TypeOf oDefiningObj Is IJComplexString Then
        Set m_oDefiningObj = oMfgGeomHelper.ApproximateCurveByStrokeAndFit(oDefiningObj)
    ElseIf TypeOf oDefiningObj Is IJCurve Then
        Set m_oDefiningObj = oDefiningObj
    Else    ' we only have IJComplexString or IJWirebody or IJCurve
        Set m_oDefiningObj = Nothing
    End If
       
CleanUp:
    Set oMfgMGHelper = Nothing
    Set oMfgGeomHelper = Nothing
    Set oCS = Nothing
    Exit Property
ErrorHandler:
    GoTo CleanUp
End Property

Public Property Get DefiningObj() As Object
    Set DefiningObj = m_oDefiningObj
End Property
 
Public Property Get GroupType() As Long
    GroupType = m_eGroupType
End Property

Public Property Let GroupType(ByVal eNewValue As Long)
    m_eGroupType = eNewValue
End Property
