VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SymbolHandling"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-----------------------------------------------------------------------------------------
'  Copyright (C) 1998 - 2002, Intergraph Corporation.  All rights reserved.
'
'  Project: MfgPlateTabHelper
'  File:    SymbolHandling - a GlobalMultiUse class with methods for handling Tab Symbols
'
'  This Class contains global functions and subroutines used when selecting,
'  modifying and joining the PlateFeatureTabs or PlateWeldTabs to the plates
'  outer contour
'  (these Tabs are extra pieces of steel added to the plate to secure
'   a perfect butt-weld start and stop at each end of the plate edge.)
'
'  Author: Paal Fredrik Borsheim
'
'  History:
'   21/03-2002   PFB     Creation date
'   2004.04.22   MJV     Included the correct error handling
'
'-----------------------------------------------------------------------------------------
Option Explicit

Private Const MODULE = "MfgPlateTabHelper.SymbolHandling"

Private Function IsEqual(Var1 As Double, Var2 As Double) As Boolean
Const Tolerance = 0.00001
   If Abs(Var1 - Var2) < Tolerance Then
      IsEqual = True
   Else
      IsEqual = False
   End If
End Function


'Helper function used since we cannot do "Set ... = New MfgGeom2d"
Public Function CreateNewMfgGeom2d(oActConn As IJDAccessMiddle, _
                                    oCurveGeometry As GSCADMfgRuleHelpersDefinitions.IJComplexString, _
                                    GeoType As StrMfgGeometryType) As GSCADMfgRuleHelpersDefinitions.IJMfgGeom2d

   Const METHOD = "MfgPlateTabHelper: CreateNewMfgGeom2d"
   On Error GoTo ErrorHandler

   Dim GeomFactory As New MfgGeom2dFactory
   Dim NewGeom2d As IJMfgGeom2d
   Set NewGeom2d = GeomFactory.Create(oActConn.ResourceManager)
   
   Dim oCloneCS As IJComplexString
   Dim oDummyCurve As IJCurve
   Dim i As Long
   
   Set oCloneCS = New ComplexString3d
   
   For i = 1 To oCurveGeometry.CurveCount
        oCurveGeometry.GetCurve i, oDummyCurve
        oCloneCS.AddCurve oDummyCurve, True
   Next i
   
   NewGeom2d.PutGeometry oCloneCS
   NewGeom2d.PutGeometryType GeoType

   Set CreateNewMfgGeom2d = NewGeom2d


CleanUp:
    Set GeomFactory = Nothing
    Set NewGeom2d = Nothing
    Set oCloneCS = Nothing
    Set oDummyCurve = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp:
End Function


'******************************************************************************
' Routine: GetTabsFromCatalog
'
' Abstract: Find the tabs from CatalogData and return tabs of the given type.
'
' Description: Takes ViewName as Input. Ex:
'******************************************************************************
Public Function GetTabsFromCatalog(sViewName As String) As IJDCollection
    Const METHOD = "GetTabsFromCatalog"
    On Error GoTo ErrorHandler
    
    Dim oMfgCatalogQueryHelper As New MfgCatalogQueryHelper
    Dim sQueryString As String
    
    sQueryString = "SELECT * FROM " + sViewName
    
    Set GetTabsFromCatalog = oMfgCatalogQueryHelper.GetPartCollectionFromDBQuery(sQueryString)
    
CleanUp:
    Set oMfgCatalogQueryHelper = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

'******************************************************************************
' Routine: IsCorrectTabType
'
' Abstract: Check the oCatalogData to see if it's correct type.
'
' Description: TabType 0 - WeldTab
'              TabType 1 - FeatureTab
'******************************************************************************
Public Function IsCorrectTabType(oCatalogData As Object, TabType As Long) As Boolean
    Dim oAttributes As IJDAttributes
    Dim oAttrIID As Variant
    Dim oAttrCol As IJDAttributesCol
    Dim oAttribute As IJDAttribute
    Dim bFound As Boolean
    Dim strTabType As String
    
    Const METHOD = "IsCorrectTabType"
    On Error GoTo ErrorHandler

    'Declare each type of tab.
    If TabType = 0 Then
        strTabType = "Weld"
    Else
        strTabType = "Feature"
    End If
    
    bFound = False
    IsCorrectTabType = False
    Set oAttributes = oCatalogData
    
    For Each oAttrIID In oAttributes
        Set oAttrCol = oAttributes.CollectionOfAttributes(oAttrIID)
        For Each oAttribute In oAttrCol
            'ex: Name will be "FeatureTab2" as defined in excel sheet.
            If oAttribute.AttributeInfo.Name = "Name" Then
                If InStr(1, oAttribute.Value, strTabType, vbTextCompare) > 0 Then
                    IsCorrectTabType = True
                    bFound = True
                    Exit For
                End If
            End If
        Next oAttribute
        If bFound Then
            Exit For
        End If
    Next oAttrIID
    
CleanUp:
    Set oAttributes = Nothing
    Set oAttrIID = Nothing
    Set oAttrCol = Nothing
    Set oAttribute = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

'******************************************************************************
' Routine: GetAnglesFromCatalog
'
' Abstract: Given a CatalogData, gets Min/Max angle attribute from the CatalogData.
'
' Description: This sub is only for WeldTab.
'              Note: Angle's unit is radian not degree.
'******************************************************************************
Public Sub GetAnglesFromCatalog(oCatalogData As Object, dMax As Double, dMin As Double)
    Dim oAttributes As IJDAttributes
    Dim oAttrIID As Variant
    Dim oAttrCol As IJDAttributesCol
    Dim oAttribute As IJDAttribute
    
    Const METHOD = "GetAnglesFromCatalog"
    On Error GoTo ErrorHandler
    
    Set oAttributes = oCatalogData
    
    For Each oAttrIID In oAttributes
        Set oAttrCol = oAttributes.CollectionOfAttributes(oAttrIID)
        For Each oAttribute In oAttrCol
            If oAttribute.AttributeInfo.Name = "MinCornerAngle" Then
                dMin = CDbl(oAttribute.Value)
            End If
            If oAttribute.AttributeInfo.Name = "MaxCornerAngle" Then
                dMax = CDbl(oAttribute.Value)
            End If
        Next oAttribute
    Next oAttrIID
        
CleanUp:
    Set oAttributes = Nothing
    Set oAttrIID = Nothing
    Set oAttrCol = Nothing
    Set oAttribute = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

'******************************************************************************
' Routine: GetSizesFromCatalog
'
' Abstract: Given a CatalogData, gets Min/Max feature size attribute from the CatalogData.
'
' Description: This sub is only for FeatureTab.
'              Note: The unit of size is meter.
'******************************************************************************
Public Sub GetSizesFromCatalog(oCatalogData As Object, dMax As Double, dMin As Double)
    Dim oAttributes As IJDAttributes
    Dim oAttrIID As Variant
    Dim oAttrCol As IJDAttributesCol
    Dim oAttribute As IJDAttribute
    
    Const METHOD = "GetSizesFromCatalog"
    On Error GoTo ErrorHandler
    
    Set oAttributes = oCatalogData
    
    For Each oAttrIID In oAttributes
        Set oAttrCol = oAttributes.CollectionOfAttributes(oAttrIID)
        For Each oAttribute In oAttrCol
            If oAttribute.AttributeInfo.Name = "MinFeatureSize" Then
                dMin = CDbl(oAttribute.Value)
            End If
            If oAttribute.AttributeInfo.Name = "MaxFeatureSize" Then
                dMax = CDbl(oAttribute.Value)
            End If
        Next oAttribute
    Next oAttrIID
    
CleanUp:
    Set oAttributes = Nothing
    Set oAttrIID = Nothing
    Set oAttrCol = Nothing
    Set oAttribute = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

'******************************************************************************
' Routine: GetParametersFromCatalog
'
' Abstract: Given a CatalogData, get length, height, radius and angle from it.
'
' Description: oCatalogData can be for FeatureTab or WeldTab.
'******************************************************************************
Public Sub GetParametersFromCatalog(oCatalogData As Object, dLength As Double, dHeight As Double, _
                                    dRadius As Double, dAngle As Double)
                                    
    Const WELD_TAB = 0
    Const FEAT_TAB = 1
    
    Dim bIsWeldTab As Boolean
    Dim strLength As String
    Dim strHeight As String
    Dim strRadius As String
    Dim strAngle As String
                                    
    Dim oAttributes As IJDAttributes
    Dim oAttrIID As Variant
    Dim oAttrCol As IJDAttributesCol
    Dim oAttribute As IJDAttribute
    
    Const METHOD = "GetParametersFromCatalog"
    On Error GoTo ErrorHandler
                                    
'    'Check if oCatalogData is WeldTab or not.
'    bIsWeldTab = False
'    bIsWeldTab = IsCorrectTabType(oCatalogData, WELD_TAB)
'
'    If bIsWeldTab Then
'        strLength = "WeldTabLength"
'        strHeight = "WeldTabHeight"
'        strRadius = "WeldTabRadius"
'        strAngle = "WeldTabAngle"
'    Else
'        strLength = "FeatureTabLength"
'        strHeight = "FeatureTabHeight"
'        strRadius = "FeatureTabRadius"
'        strAngle = "FeatureTabAngle"
'    End If
    
    strLength = "PlateTabLength"
    strHeight = "PlateTabHeight"
    strRadius = "PlateTabRadius"
    strAngle = "PlateTabAngle"

    
    'Get each parameter.
    Set oAttributes = oCatalogData
    
    For Each oAttrIID In oAttributes
        Set oAttrCol = oAttributes.CollectionOfAttributes(oAttrIID)
        For Each oAttribute In oAttrCol
            If oAttribute.AttributeInfo.Name = strLength Then
                dLength = CDbl(oAttribute.Value)
            ElseIf oAttribute.AttributeInfo.Name = strHeight Then
                dHeight = CDbl(oAttribute.Value)
            ElseIf oAttribute.AttributeInfo.Name = strRadius Then
                dRadius = CDbl(oAttribute.Value)
            ElseIf oAttribute.AttributeInfo.Name = strAngle Then
                dAngle = CDbl(oAttribute.Value)
            End If
        Next oAttribute
    Next oAttrIID
    
CleanUp:
    Set oAttributes = Nothing
    Set oAttrIID = Nothing
    Set oAttrCol = Nothing
    Set oAttribute = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

'******************************************************************************
' Routine: UpdateSymbolWithParameters
'
' Abstract: Given a symbol definition, apply new value to each input's parameter
'           -content.
'
' Description:
'******************************************************************************
Public Sub UpdateSymbolWithParameters(oSymbolDef As IJDSymbolDefinition, dLength As Double, _
                                        dHeight As Double, dRadius As Double, dAngle As Double)
    Const METHOD = "UpdateSymbolWithParameters"
    On Error GoTo ErrorHandler
    
    Dim oInputs As IJDInputs
    Dim oInput As IJDInput
    Dim oParameterContent As IJDParameterContent
    
    Set oInputs = oSymbolDef.IJDInputs
    
    'Apply new value to the each Input.
    For Each oInput In oInputs
        If InStr(1, oInput.Name, "Length", vbTextCompare) > 0 Then
            Set oParameterContent = oInput.DefaultParameterValue
            If Not oParameterContent Is Nothing Then
                oParameterContent.UomValue = dLength
            End If
            Set oParameterContent = Nothing
        ElseIf InStr(1, oInput.Name, "Height", vbTextCompare) > 0 Then
            Set oParameterContent = oInput.DefaultParameterValue
            If Not oParameterContent Is Nothing Then
                oParameterContent.UomValue = dHeight
            End If
            Set oParameterContent = Nothing
        ElseIf InStr(1, oInput.Name, "Radius", vbTextCompare) > 0 Then
            Set oParameterContent = oInput.DefaultParameterValue
            If Not oParameterContent Is Nothing Then
                oParameterContent.UomValue = dRadius
            End If
            Set oParameterContent = Nothing
        ElseIf InStr(1, oInput.Name, "Angle", vbTextCompare) > 0 Then
            Set oParameterContent = oInput.DefaultParameterValue
            If Not oParameterContent Is Nothing Then
                oParameterContent.UomValue = dAngle
            End If
            Set oParameterContent = Nothing
        End If
    Next oInput
    
CleanUp:
    Set oParameterContent = Nothing
    Set oInput = Nothing
    Set oInputs = Nothing
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Sub

'******************************************************************************
' Routine: CalculateFeatureSize
'
' Abstract: Returns the feature size. (oGeometry should support IJRange interface.)
'
' Description: This function calculates the size according to the following logic.
'              First get a box that envelops the geometry, and then return the
'              shortest length of a box. (This logic may be modifiable by shipyard.)
'******************************************************************************
Public Function CalculateFeatureSize(oGeometry As IJComplexString) As Double
    Const METHOD = "CalculateFeatureSize"
    On Error GoTo ErrorHandler
    
    Dim oRuleHelperSupport As IJMfgRuleHelpersSupport
    Dim dMinX As Double, dMinY As Double, dMinZ As Double
    Dim dMaxX As Double, dMaxY As Double, dMaxZ As Double
    Dim dX As Double, dY As Double
    
    'Get Range points.
    Set oRuleHelperSupport = New MfgRuleHelpersSupport
    oRuleHelperSupport.GetRange oGeometry, dMinX, dMinY, dMinZ, dMaxX, dMaxY, dMaxZ
    
    dX = Abs(dMaxX - dMinX)
    dY = Abs(dMaxY - dMinY)
    
    If dX > dY Then
        CalculateFeatureSize = dY
    Else
        CalculateFeatureSize = dX
    End If
    
CleanUp:
    Set oRuleHelperSupport = Nothing
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function
'-----------------------------------------------------------------------------------------
' This function is joining the curves in two separate ComplexStrings,
' It first checks that endpoint of curve1 is equal to startpoint of curve2
' and updates the geometry in curve1 to include curve2
' If the points are not equal, it only returns curve1.
' If curve1 is empty, it should return curve2.
'-----------------------------------------------------------------------------------------
Public Function JoinComplexStrings(CString1 As GSCADMfgRuleHelpersDefinitions.IJComplexString, _
                                   CString2 As GSCADMfgRuleHelpersDefinitions.IJComplexString) As GSCADMfgRuleHelpersDefinitions.IJComplexString

    Const METHOD = "JoinComplexStrings"
    On Error GoTo ErrorHandler
    
   
   Dim NumberOfCurvesToAdd As Long, i As Long
   Dim LastCurve1Index As Long
   Dim FirstCurve1 As IJCurve
   Dim LastCurve1 As IJCurve
   Dim LastCurve2Index As Long
   Dim FirstCurve2 As IJCurve
   Dim LastCurve2 As IJCurve
   Dim NextCurve2 As IJCurve
   Dim Xs1 As Double, Xs2 As Double, Xe1 As Double, Xe2 As Double, xTmp As Double
   Dim Ys1 As Double, Ys2 As Double, Ye1 As Double, Ye2 As Double, yTmp As Double
   Dim Zs1 As Double, Zs2 As Double, Ze1 As Double, Ze2 As Double, zTmp As Double
   
   If CString1 Is Nothing Then
      Set JoinComplexStrings = CString2
      Exit Function
   ElseIf CString2 Is Nothing Then
      Set JoinComplexStrings = CString1
      Exit Function
   End If
   On Error Resume Next    'CurveCount returns error, not zero if empty
   LastCurve1Index = CString1.CurveCount
   LastCurve2Index = CString2.CurveCount
   
   Call CString1.GetCurve(1, FirstCurve1)
   Call CString2.GetCurve(1, FirstCurve2)
   Call CString1.GetCurve(LastCurve1Index, LastCurve1)
   Call CString2.GetCurve(LastCurve2Index, LastCurve2)
   Call FirstCurve1.EndPoints(Xs1, Ys1, Zs1, xTmp, yTmp, zTmp)
   Call LastCurve1.EndPoints(xTmp, yTmp, zTmp, Xe1, Ye1, Ze1)
   Call FirstCurve2.EndPoints(Xs2, Ys2, Zs2, xTmp, yTmp, zTmp)
   Call LastCurve2.EndPoints(xTmp, yTmp, zTmp, Xe2, Ye2, Ze2)
   
    'Copy CString1
    Dim oCloneCS As ComplexString3d
    Dim oDummyCurve As IJCurve
    Set oCloneCS = New ComplexString3d
    For i = 1 To CString1.CurveCount
        CString1.GetCurve i, oDummyCurve
        oCloneCS.AddCurve oDummyCurve, True
    Next i
    Set oDummyCurve = Nothing
   
   If (IsEqual(Xe1, Xs2) And IsEqual(Ye1, Ys2) And IsEqual(Ze1, Zs2)) Then
      'Startpoinbt of curve2 is equal to endpoint of curve1, Join as is
      For i = 1 To LastCurve2Index
         Call CString2.GetCurve(i, NextCurve2)
'         Call CString1.AddCurve(NextCurve2, True)
         oCloneCS.AddCurve NextCurve2, True
         'LastCurve1Index = LastCurve1Index + 1
      Next i
    
   ElseIf (IsEqual(Xs1, Xs2) And IsEqual(Ys1, Ys2) And IsEqual(Zs1, Zs2)) Then
      'Both startpoints are equal, reverse curve1
'      MsgBox "Unfinished coding in JoinComplexStrings, Need to reverse Curve1"
      For i = 1 To LastCurve2Index
        CString2.GetCurve i, NextCurve2
'        CString1.AddCurve NextCurve2, False
        oCloneCS.AddCurve NextCurve2, False
      Next i
            
   ElseIf (IsEqual(Xe1, Xe2) And IsEqual(Ye1, Ye2) And IsEqual(Ze1, Ze2)) Then
      'Both endpoints are equal, reverse curve2
'      MsgBox "Unfinished coding in JoinComplexStrings, Need to reverse Curve2"
      For i = 1 To LastCurve2Index
         Call CString2.GetCurve(i, NextCurve2)
'         Call CString1.AddCurve(NextCurve2, True)
         oCloneCS.AddCurve NextCurve2, True
         'LastCurve1Index = LastCurve1Index + 1
      Next i

   ElseIf (IsEqual(Xs1, Xe2) And IsEqual(Ys1, Ye2) And IsEqual(Zs1, Ze2)) Then
      'Curve2 are coming before curve1, reverse both curves
'      MsgBox "Unfinished coding in JoinComplexStrings, Need to reverse both Curves"
      For i = 1 To LastCurve2Index
        CString2.GetCurve i, NextCurve2
'        CString1.AddCurve NextCurve2, False
        oCloneCS.AddCurve NextCurve2, False
      Next i

   Else
      'Cannot join curves
      'MsgBox "Error during joining of tab and plategeometry, Lines are not connected"
   End If
   
CleanUp:
   Set LastCurve1 = Nothing
   Set NextCurve2 = Nothing
'  Set JoinComplexStrings = CString1
   Set JoinComplexStrings = oCloneCS
   Set oCloneCS = Nothing
   Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

'-------------------------------------------------------------------------------------
'  Returns a size parameter describing some maximum size for a Tab
'  Checks minimum of curve radius, curve length or min distance to next corner.
'  Add length of curves if almost tangent and straight lines. Add whole arclength if
'  it spans a small angle, but not for more than 45 degrees
'-------------------------------------------------------------------------------------
Public Function GetMaxTabHeightForLine(Line As GSCADMfgRuleHelpersDefinitions.IJComplexString, _
                                      Optional bCheckEnd As Boolean = False) As Double
   Const METHOD = "GetMaxTabHeightForLine"
   On Error GoTo ErrorHandler
    
   Dim DistToEndOfLine As Double
   Dim FirstLineLength As Double
   Dim SizeParameter As Double
   Dim ArcRadius As Double
   Dim ArcAngleSpan As Double
   Dim AngleBetwLines As Double
   Dim bContinue As Boolean
   Dim StartIndex As Long, EndIndex As Long, StepIndex As Long, Index As Long
   Dim StartCurve As IJCurve, EndCurve As IJCurve, NextCurve As IJCurve
   Dim StartCurveObj As Object, EndCurveObj As Object, NextCurveObj As Object
   Dim X1 As Double, X2 As Double, X3 As Double, X4 As Double, XDir As Double, XsDir As Double
   Dim Y1 As Double, Y2 As Double, Y3 As Double, Y4 As Double, YDir As Double, YsDir As Double
   Dim Z1 As Double, Z2 As Double, Z3 As Double, Z4 As Double, ZDir As Double, ZsDir As Double
   Dim oLine As IngrGeom3D.Line3d
   Dim oArc As IngrGeom3D.Arc3d
   Dim StartParam As Double, EndParam As Double
   Dim bStopCheck As Boolean
   'Dim oSpline As IngrGeom3D.BSplineCurve3d
   
   Const AcceptAngleDistortion As Double = 0.8
   
   StartIndex = 1
   EndIndex = Line.CurveCount
   StepIndex = 1
   If bCheckEnd = True Then
      StartIndex = EndIndex
      EndIndex = 1
      StepIndex = -1
   End If
   
   'Get length between endpoints
   Call Line.GetCurve(StartIndex, StartCurveObj)
   Call Line.GetCurve(EndIndex, EndCurveObj)
   Set StartCurve = StartCurveObj
   Set EndCurve = EndCurveObj
   Call StartCurve.EndPoints(X1, Y1, Z1, X2, Y2, Z2)
   Call EndCurve.EndPoints(X3, Y3, Z3, X4, Y4, Z4)
   If bCheckEnd Then
      DistToEndOfLine = Sqr((X4 - X1) ^ 2 + (Y4 - Y1) ^ 2)
   Else
      DistToEndOfLine = Sqr((X3 - X2) ^ 2 + (Y3 - Y2) ^ 2)
   End If
   
   SizeParameter = 0
   XsDir = X2 - X1
   YsDir = Y2 - Y1
   ZsDir = 0         'Only considering 2d lines
   bStopCheck = False
   
   'For each curve ...
   'For index = StartIndex + StepIndex To EndIndex Step StepIndex
   For Index = StartIndex To EndIndex Step StepIndex
      Call Line.GetCurve(Index, NextCurveObj)
      'check curvetype:
      'Select Case curvetype
      If TypeOf NextCurveObj Is IngrGeom3D.Line3d Then
         Set oLine = NextCurveObj
         Call oLine.GetDirection(XDir, YDir, ZDir)
         If GetAngleBetweenVectors(XDir, YDir, XsDir, YsDir) < AcceptAngleDistortion Then
            SizeParameter = SizeParameter + oLine.Length
         Else
            bStopCheck = True
         End If
         Set oLine = Nothing
      ElseIf TypeOf NextCurveObj Is IngrGeom3D.Arc3d Then
         Set oArc = NextCurveObj
         Call oArc.GetStartPoint(X1, Y1, Z1)
         Call oArc.GetEndPoint(X2, Y2, Z2)
         XDir = X2 - X1
         YDir = Y2 - Y1
         If GetAngleBetweenVectors(XDir, YDir, XsDir, YsDir) < AcceptAngleDistortion Then
            SizeParameter = SizeParameter + Sqr(XDir ^ 2 + YDir ^ 2)
         ElseIf oArc.SweepAngle > AcceptAngleDistortion Then
            SizeParameter = SizeParameter + oArc.Radius / 2
            bStopCheck = True
         End If
         Set oArc = Nothing
      ElseIf TypeOf NextCurveObj Is IngrGeom3D.BSplineCurve3d Then
         'handle trough the IJCurve interface
         Set NextCurve = NextCurveObj
         Call NextCurve.ParamRange(StartParam, EndParam)
         If bCheckEnd Then
            StartParam = EndParam
         End If
         Call NextCurve.Evaluate(StartParam, X1, Y1, Z1, XDir, YDir, ZDir, X2, Y2, Z2)
         If GetAngleBetweenVectors(XDir, YDir, XsDir, YsDir) < AcceptAngleDistortion Then
            SizeParameter = SizeParameter + NextCurve.Length
         Else
            bStopCheck = True
         End If
         Set NextCurve = Nothing
      Else
         'Some unexpected curve type is used. Do nothing
      End If
      'End Select
   
      If bStopCheck Then Exit For
   Next Index
   
   '.....  UNFINISHED!!!!!
   'Temporary:
   If SizeParameter > 0 Then
      GetMaxTabHeightForLine = SizeParameter
   Else
      GetMaxTabHeightForLine = DistToEndOfLine
   End If
   
CleanUp:
   Set oLine = Nothing
   Set oArc = Nothing
   Set StartCurve = Nothing
   Set EndCurve = Nothing
   Set NextCurve = Nothing
   Set StartCurveObj = Nothing
   Set EndCurveObj = Nothing
   Set NextCurveObj = Nothing

   Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function


'Simple sub to return angle between two planar vectors
Private Function GetAngleBetweenVectors(xDir1 As Double, yDir1 As Double, _
                                        xDir2 As Double, ydir2 As Double) As Double
Const METHOD = "GetAngleBetweenVectors"
On Error GoTo ErrorHandler
       
   Dim Angle1 As Double, Angle2 As Double
   Const Tolerance As Double = 0.00001
   Const pi As Double = 3.14159265358979
   
   If Abs(xDir1) > Tolerance Then
      Angle1 = Atn(yDir1 / xDir1)
   ElseIf yDir1 > 0 Then
      Angle1 = pi / 2
   Else
      Angle1 = -pi / 2
   End If

   If Abs(xDir2) > Tolerance Then
      Angle2 = Atn(ydir2 / xDir2)
   ElseIf ydir2 > 0 Then
      Angle2 = pi / 2
   Else
      Angle2 = -pi / 2
   End If

   GetAngleBetweenVectors = Abs(Angle2 - Angle1)
   
CleanUp:
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

'******************************************************************************
' Routine: ChangeDirection
'
' Abstract: Given a ComplexString, invert each curve's direction that comprise of
'           ComplexString
'
' Description: The contours that are represented by MfgGeom2dCol are counterclockwise,
'              so in some cases, we need to invert the direction of contours that
'              comprise of only Tab part.
'******************************************************************************
Public Function ChangeDirection(oComplexString As IJComplexString) As IJComplexString
    
    Const METHOD = "ChangeDirection"
    On Error GoTo ErrorHandler
    
    Dim oNewComplexString As ComplexString3d
    Dim oCurve As IJCurve
    Dim lCount As Long
    Dim Index As Long
    Dim oLine As Line3d
    Dim oArc As Arc3d
    Dim cx As Double, cy As Double, cz As Double
    Dim sx As Double, sy As Double, sz As Double
    Dim ex As Double, ey As Double, ez As Double
    
    Set oNewComplexString = New ComplexString3d
    
    lCount = oComplexString.CurveCount
    For Index = 1 To lCount
        oComplexString.GetCurve Index, oCurve
        If Not oCurve Is Nothing Then
            If TypeOf oCurve Is Line3d Then
                Set oLine = oCurve
                oLine.GetStartPoint sx, sy, sz
                oLine.GetEndPoint ex, ey, ez
                oLine.DefineBy2Points ex, ey, ez, sx, sy, sz
            ElseIf TypeOf oCurve Is Arc3d Then
                Set oArc = oCurve
                oArc.GetCenterPoint cx, cy, cz
                oArc.GetStartPoint sx, sy, sz
                oArc.GetEndPoint ex, ey, ez
                oArc.DefineByCenterStartEnd cx, cy, cz, ex, ey, ez, sx, sy, sz
            End If
            Set oLine = Nothing
            Set oArc = Nothing
        End If
        oNewComplexString.AddCurve oCurve, True
        Set oCurve = Nothing
    Next Index
    
    Set ChangeDirection = oNewComplexString
    
CleanUp:
    Set oCurve = Nothing
    Set oLine = Nothing
    Set oArc = Nothing
    Set oNewComplexString = Nothing
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
    GoTo CleanUp
End Function

 
