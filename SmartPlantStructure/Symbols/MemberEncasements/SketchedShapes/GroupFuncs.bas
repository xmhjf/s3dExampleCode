Attribute VB_Name = "GroupFuncs"
Option Explicit
'========================================================================
'
'    Copyright 2002 Intergraph Corporation
'    All Rights Reserved
'
'    Including software, file formats, and audio-visual displays;
'    may only be used pursuant to applicable software license
'    agreement; contains confidential and proprietary information of
'    Intergraph and/or third parties which is protected by copyright
'    and trade secret law and may not be provided or otherwise made
'    available without proper authorization.
'
'    Unpublished -- rights reserved under the Copyright Laws of the
'    United States.
'
'    Intergraph Corporation
'    Huntsville, Alabama   35894-0001
'
'========================================================================
'
' Method ReturnGroupsAsOrderedArray
'
'
'   11/26/02    Dehaussy Caroline   Created module and added this function
'   05/07/03    Dehaussy Caroline   Removed the called to the method HandleError which is not defined in this module. (TR 39940)
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const MODULE = "Module GroupFuncs : "


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Method:      ReturnGroupAsOrderedArray
'''
'''    Purpose:   Given a group of end pt connected objects, order the objects so that for every object,
' the next one has either its end or its start point common to it
'
'''
'''    Inputs:
'         The graphic group as a RAD2D.Group, an empty array of objects and a long to
'         return the number of objects in the array.
'
'     Outputs:
'         The array will be initialized and filled with the ordered list of graphic
'         objects.  The number of objects in the list will be returned.  This need not
'         be the same as the number in the group as the group may, by accident, contain
'         none endpt objects, such as dimensions, relationships, text etc.  Any object
'         that does not support endpt, will thus not be added to the array.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReturnGroupAsOrderedArray(ByRef group As RAD2D.group, ByRef aOrderedArray() As Object, lNumReturned As Long)
Const MT = "ReturnGroupAsOrderedArray"
    Dim oObjectColl As New Collection
    Dim tempColl As New Collection
    Dim firstObject As Object

    Dim currentObject As Object
    Dim bFound As Boolean
    Dim bFirstTimeNotFound As Boolean
    Dim ii As Integer
    Dim x As Double
    Dim y As Double

    Dim oTemp As Object

    On Error GoTo ErrorHandler
    lNumReturned = 0    ' something goes wrong this will be 0
'
    For Each currentObject In group
'
      Select Case (currentObject.Type)
        Case igLine2d, igArc2d, igEllipticalArc2d
''        If currentObject.Type = igLine2d Or currentObject.Type = igArc2d Or currentObject.Type = igEllipticalArc2d _
''            Or currentObject.Type = igLineString2d Then
            ' JC 11/17/99: removed Circle and Ellipse because they don't have end points
            oObjectColl.Add currentObject
''        ElseIf currentObject.Type = igCircle2d Or currentObject.Type = igEllipse2d _
''                Or currentObject.Type = igRectangle2d Then
        Case igCircle2d, igEllipse2d, igRectangle2d
            Exit Sub
        Case igLineString2d, igBsplineCurve2d
          If currentObject.GetGeometry.IsClosed = vbTrue Then
            Exit Sub
          Else
            oObjectColl.Add currentObject
          End If
        Case Else   '' ignore unknowns as group is only supposed to be made of the above
'''dbg            MsgBox "unknow obj: " & currentObject.Name & " type: " & currentObject.Type
        End Select
'
    Next currentObject
    'Msgbox "oObjectColl.Count = " & oObjectColl.Count
    
    ' [TR-CP-77035; Jeamis] - Do not process on if we don't have a collection to process.
    If (oObjectColl.Count) Then
        Set currentObject = oObjectColl.Item(1)
        Set firstObject = currentObject
        
        If currentObject.Type = igBsplineCurve2d Or currentObject.Type = igLineString2d Then
            Dim xx As Double, yy As Double
            currentObject.GetGeometry.GetEndPoints xx, yy, x, y
        Else
            currentObject.GetEndPoint x, y
        End If
        oObjectColl.Remove 1
        tempColl.Add currentObject
        
        bFound = True
        bFirstTimeNotFound = True
        
        While oObjectColl.Count > 0 And bFound
            bFound = False
            
            For ii = 1 To oObjectColl.Count
                        
                Set oTemp = oObjectColl.Item(ii)
                
                If PointFitObjectEnd(x, y, oTemp, x, y) Then
                    oObjectColl.Remove ii
                    Set currentObject = oTemp
                    bFound = True
                    Exit For
                End If
            Next ii
        
            If bFound Then
                ' add the object into the list, either before the first element or after it
                If bFirstTimeNotFound Then
                    tempColl.Add currentObject
                Else
                    tempColl.Add currentObject, , 1
                End If
            ElseIf bFirstTimeNotFound Then '<=> Not found and bFirstTimeNotFound
                ' if no end is found, go reverse (for open contour)
                If firstObject.Type = igBsplineCurve2d Or currentObject.Type = igLineString2d Then
                    Dim x2 As Double, y2 As Double
                    firstObject.GetGeometry.GetEndPoints x, y, x2, y2
                Else
                    firstObject.GetStartPoint x, y
                End If
                bFound = True
                bFirstTimeNotFound = False
            End If
        Wend
    
    '    Dim obj As Object
        lNumReturned = tempColl.Count
        If lNumReturned > 0 Then
            Dim i As Long
            ReDim aOrderedArray(lNumReturned)
            i = 0
            For Each currentObject In tempColl
        '        Set obj = currentObject
        '        oColl.Add obj
                Set aOrderedArray(i) = currentObject
                i = i + 1
            Next currentObject
        End If
   
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
End Sub


Public Function PointFitObjectEnd(x As Double, y As Double, obj As Object, newX As Double, newY As Double) As Boolean
    Const MT = "PointFitObjectEnd"

    Dim Sx As Double
    Dim Sy As Double
    Dim Ex As Double
    Dim Ey As Double
    
      ' bsplines can have a special trim object attached that shows the true
      ' endpoints as the user requested.  Need to get this object.
    If obj.Type = igBsplineCurve2d Then
'      Dim oTrimBspline As RAD2D.BSplineCurve2d
'      Set oTrimBspline = GetTrimmedCurveFromBspline(obj)
'      If oTrimBspline Is Nothing Then
        obj.GetGeometry.GetEndPoints Sx, Sy, Ex, Ey
'        obj.GetEndPoint Ex, Ey
'        Else
'            oTrimBspline.GetGeometry.GetEndPoints Sx, Sy, Ex, Ey
'            Set oTrimBspline = Nothing
'        End If
    Else
      obj.GetStartPoint Sx, Sy
      obj.GetEndPoint Ex, Ey
    End If
         
    ' if the coordinates are equal to the object start point
    If (Abs(x - Sx) < 0.000001 And Abs(y - Sy) < 0.000001) Then
        PointFitObjectEnd = True
        newX = Ex
        newY = Ey
    ' if the coordinates are equal to the object end point
    ElseIf (Abs(x - Ex) < 0.000001 And Abs(y - Ey) < 0.000001) Then
        PointFitObjectEnd = True
        newX = Sx
        newY = Sy
    Else
        PointFitObjectEnd = False
        newX = x
        newY = y
    End If
    Exit Function
    
ErrorHandler:
    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
    
End Function

Public Function Convert2DTo3D(RadObject As Object, pConverter As IMSTools2D.GeometryConverter) As Object
Const METHOD = "Convert2DTo3D"
On Error GoTo ErrorHandler

    Dim elesSP3D  As IJElements
    Set elesSP3D = New JObjectCollection
    Dim oSP3DObj As Object
    
    Dim rad2dtype As Long
    
    rad2dtype = RadObject.Type
    
    If rad2dtype = igLine2d Or rad2dtype = igArc2d Or rad2dtype = igEllipticalArc2d Or _
        rad2dtype = igLineString2d Or rad2dtype = igCircle2d Or rad2dtype = igEllipse2d Or _
        rad2dtype = igBsplineCurve2d Or rad2dtype = igRectangle2d Then
             
        pConverter.Convert2DTo3D RadObject, Nothing, Nothing, Nothing, elesSP3D, Nothing

        If elesSP3D.Count > 0 Then
            If TypeOf elesSP3D.Item(1) Is IMoniker Then
                pConverter.FindByMoniker3D elesSP3D.Item(1), oSP3DObj
            Else
                Set oSP3DObj = elesSP3D.Item(1)
            End If
        End If
    End If
    If oSP3DObj Is Nothing Then
        Set Convert2DTo3D = Nothing
    Else
        Set Convert2DTo3D = oSP3DObj
    End If

    elesSP3D.Clear
    Set elesSP3D = Nothing
    Set oSP3DObj = Nothing
    Exit Function

ErrorHandler:
    Err.Raise E_FAIL
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Method:      GetTrimmedCurveFromBspline
'''
'''     Given a Bspline curve see if it has a segmented style attached that changes
'''     the display w.r.t. one or 2 trimming objects.  The visible bspline will be
'''     geometrical shorter than the stored object, so return a bspline corresponding
'''     to the trimmed geometry.
'''     The temporary bspline should be deleted by the caller when finished.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function GetTrimmedCurveFromBspline(ObjBsp As RAD2D.BSplineCurve2d) As Object
'
'    Dim objAssocSegStyle As SEGSTYLELib.AsSegStyle
'    Dim oGeomBSP As RAD2D.BSplineCurve2d
'    Dim oBspCurves As RAD2D.BSplineCurves2d
''
'    Const MT = "GetTrimmedCurveFromBspline"
'
'    On Error GoTo ErrorHandler
'
'    If Not ObjBsp Is Nothing Then
'      Set objAssocSegStyle = ObjBsp.ControllingSegmentedStyle
'    End If
''
'    If Not objAssocSegStyle Is Nothing Then
'        Set oBspCurves = ObjBsp.Document.ActiveSheet.BSplineCurves2d
'        Set oGeomBSP = oBspCurves.AddByGeometry(ObjBsp.GetGeometry)
''''         MsgBox "Before Trim Num nodes: " & Str(oGeomBSP.NodeCount)
'        objAssocSegStyle.SetTrimGeometry oGeomBSP
''''        MsgBox "After Trim Num nodes: " & Str(oGeomBSP.NodeCount)
'        Set GetTrimmedCurveFromBspline = oGeomBSP
''''    Else
''''        MsgBox "No seg style on located bsp"
'    End If
'
'    Exit Function
'
'ErrorHandler:
'    MsgBox Err.Source & ": " & Trim$(Str$(Err.Number)) & " - " & Err.Description
'
'End Function
'
