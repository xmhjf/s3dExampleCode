Attribute VB_Name = "SystemNameRulesUtils"
'*******************************************************************
'  Copyright (C) 2002-2008 Intergraph Corporation.  All rights reserved.
'
'  Project: SystemNameRulesUtils
'
'  Abstract: The file contains utilities used for checking name uniqueness
'   in SystemsAndSpecs
'
'  Author: T. Merchant
'   (The code for the subroutine CountOccurrences was copied from an example
'    in the Microsoft Office 2000/Visual Basic Programmer's Guide section of
'    the MSDN.)
'
' 02/22/02  T. Merchant Created
' 10/10/06  T. Merchant TR 102095 - Create System for support is not working correctly.
'               TestForDuplicateName was failing if existing child objects
'               do not support IJNamedItem.
' 05/21/08  T. Merchant TR 130441 - Systems get renamed on hierarchy import.
'               In NameSubscript, do not interpret commas within numeric strings
'               within parentheses as thousands separators.
'
'******************************************************************

Option Explicit

Option Compare Text  'To make the comparison insensitive to case (CR 4835)

Public Const ernDuplicateName = 4 + vbObjectError + 512
Public Const ernBlankName = 4 + vbObjectError + 513
Public Const E_FAIL = -2147467259

Sub TestForDuplicateName(ByVal oObject As Object, _
    boolIsDuplicate As Boolean, strOrigName As String, strNewName As String)
    'Check to see if the new system has any sibling systems with the same name.
    
    Dim oObjectName As IJNamedItem
    Set oObjectName = oObject

    boolIsDuplicate = False

    Dim oChild As IJDesignChild
    Set oChild = oObject
    Dim oParent As IJDesignParent
    Set oParent = oChild.GetParent

    Dim oChildrenCollection As New JObjectCollection
    Dim oChildren As IJDObjectCollection
    Set oChildren = oChildrenCollection
    oParent.GetChildren oChildren, "SystemChildren"

    Dim oChildObj As Object
    Dim oChildName As IJNamedItem

    strOrigName = oObjectName.Name
    strNewName = strOrigName

    Dim lngCurrSubscript As Long
    Dim lngHighSubscript As Long
    lngHighSubscript = 0

    Dim NamesMatch As Boolean
    NamesMatch = False
    
    For Each oChildObj In oChildren
        If TypeOf oChildObj Is IJNamedItem Then
            Set oChildName = oChildObj
            If Not oChildName Is oObjectName Then
                If oChildName.Name = oObjectName.Name Then
                    NamesMatch = True
                End If
            End If
        Set oChildName = Nothing
        End If
    Next

    If NamesMatch Then
        'Loop through each named child of the target parent system to see if there
        ' is already a child object with the same name.
        For Each oChildObj In oChildren
            If TypeOf oChildObj Is IJNamedItem Then
                Set oChildName = oChildObj
                If Not oChildName Is oObjectName Then
                    lngCurrSubscript = 0
                    CheckNamesMatch oChildName.Name, oObjectName.Name, lngCurrSubscript, NamesMatch
                    If NamesMatch Then
                        boolIsDuplicate = True
                        If lngCurrSubscript > 0 And lngCurrSubscript > lngHighSubscript Then
                                lngHighSubscript = lngCurrSubscript
                        ElseIf lngCurrSubscript = -1 And lngHighSubscript = 0 Then
                            lngHighSubscript = -1
                        End If
                    End If
                End If
            Set oChildName = Nothing
            End If
        Next
    End If

    Set oChildren = Nothing
    Set oChildrenCollection = Nothing
    Set oChild = Nothing
    Set oParent = Nothing

    If lngHighSubscript = -1 Then
        'The name of this system matches the name of at least one other system,
        ' but its subscript is unique.  Leave it as it is.
        boolIsDuplicate = False
    End If

    If boolIsDuplicate Then
        ' The same name already exists
        'Test to see if the new name already has a subscript
        Dim strNameWithoutSubscript As String
        Dim lngTmpNbr As Long
        NameSubscript strOrigName, lngTmpNbr, strNameWithoutSubscript
        If strOrigName = strNameWithoutSubscript Then
            oObjectName.Name = oObjectName.Name & "(" & Format(lngHighSubscript + 1) & ")"
        Else
            oObjectName.Name = strNameWithoutSubscript & "(" & Format(lngHighSubscript + 1) & ")"
        End If
        strNewName = oObjectName.Name
    End If

    Set oObjectName = Nothing

End Sub

Sub CheckNamesMatch(strOldName As String, strNewName As String, _
    lngOutSub As Long, NamesMatch As Boolean)
    'Check to see if the two input strings match each other.

    Dim strOldNameWOSub As String
    Dim strNewNameWOSub As String
    Dim lngNewSub As Long
    Dim lngOldSub As Long

    NamesMatch = False

    NameSubscript strOldName, lngOldSub, strOldNameWOSub
    NameSubscript strNewName, lngNewSub, strNewNameWOSub

    If strNewName = strOldName Then
        NamesMatch = True
        If strOldName = strOldNameWOSub Then
            'The names match and they don't have subscripts.
            lngOutSub = 0
        Else
            'The names match and they do have subscripts.
            lngOutSub = lngOldSub
        End If
    Else
        If strNewName = strOldNameWOSub Then
            'The existing name has a subscript and the new name does
            ' not.  When the subscript is removed from the existing name,
            ' it matches the new name.
            NamesMatch = True
            lngOutSub = lngOldSub
        Else
            If strNewNameWOSub = strOldName Then
                'The new name has a subscript and the existing name
                ' does not.  When the subscript is removed from the
                ' new name, it matches the existing name.
                NamesMatch = True
'                lngOutSub = -1
                If lngOldSub > lngNewSub Then
                    lngOutSub = lngOldSub
                Else
                    lngOutSub = lngNewSub
                End If
            Else
                If strNewNameWOSub = strOldNameWOSub Then
                    NamesMatch = True
                    If lngOldSub = lngNewSub Then
                        'Both the new name and the existing name have
                        ' subscripts.  The names match and the subscripts
                        ' also match.
                        lngOutSub = lngOldSub
                    Else
                        'Both the new name and the existing name have
                        ' subscripts.  The names match but the subscripts
                        ' do not match.
'                        lngOutSub = -1
                        If lngOldSub > lngNewSub Then
                            lngOutSub = lngOldSub
                        Else
                            lngOutSub = lngNewSub
                        End If
                    End If
                Else
                    'The new name is not the same as the existing name.
                    NamesMatch = False
                End If
            End If
        End If
    End If
    If lngOutSub = 0 Then
        lngOutSub = 1
    End If

End Sub

Sub NameSubscript(strOriginalName As String, _
                          lngSubscript As Long, _
                          strStrippedName As String)
    'Find if the name passed in ends with a subscript.  If so, return that subscript
    ' and the name with the subscript removed.

    Dim strSplit() As String
    Dim lngNbrWords As Long
    Dim strWorkingName As String
    Dim strWorkingSubscript As String
    Dim lngNbrOpen As Long
    Dim lngNbrClose As Long
    Dim lngLong As Long
    Dim lngShort As Long
    Dim lngNdx As Long

    lngSubscript = 0
    strStrippedName = strOriginalName
    strWorkingName = strOriginalName

    'Look for an open parenthesis that would indicate that we might have
    ' a subscript
    lngNbrOpen = CountOccurrences(strWorkingName, "(")
    lngNbrClose = CountOccurrences(strWorkingName, ")")

    If lngNbrOpen = 0 Or lngNbrOpen <> lngNbrClose Then
        'We have either no open parenthesis or an uneven number of open and close
        ' parenthesis.  In the first case, there is no subscript.  In the second
        ' case, we don't know what to do so we will treat it as if there is no subscript.
    Else
        strSplit = Split(strWorkingName, "(")
        lngNbrWords = UBound(strSplit) - LBound(strSplit) + 1
        If lngNbrWords > 1 Then
            'We found the open parenthesis and its not the last character in
            ' the string.
            strWorkingName = strSplit(0)
            For lngNdx = 1 To lngNbrWords - 2
                strWorkingName = strWorkingName & "(" & strSplit(lngNdx)
            Next
            ' Let's see if we have a close
            ' parenthesis at the end of the second string.
            strWorkingSubscript = strSplit(lngNbrWords - 1)
            'Verify that there is a closed parenthesis at the end of the
            ' second string
            lngLong = Len(strWorkingSubscript)
            lngNbrClose = CountOccurrences(strWorkingSubscript, ")")
            If lngNbrClose = 1 Then
                'There is a closed parenthesis somewhere in the second string.
                '  Is it at the end of the string?
                strSplit = Split(strWorkingSubscript, ")")
                lngShort = Len(strSplit(0))
                If lngLong = lngShort + 1 Then
                    'We have the index - in string form.  Convert it.
                    If IsNumeric(strSplit(0)) Then
                        'We could still have unwanted commas, periods, dollar
                        ' signs, or other numeric punctuation.  Use the Val
                        ' function to check for them.
                        Dim dValue As String
                        dValue = Val(strSplit(0))
                        Dim strValue As String
                        strValue = Format(dValue, "#")
                        If Len(strSplit(0)) = Len(strValue) Then
                            lngSubscript = CLng(strSplit(0))
                            'Strip the subscript from the name
                            strStrippedName = strWorkingName
                        End If
                    End If
                End If
            End If
        End If
    End If

End Sub

Function CountOccurrences(strText As String, _
                                  strFind As String, _
                                  Optional lngCompare As VbCompareMethod) As Long

   ' Count occurrences of a particular character or characters.
   ' If lngCompare argument is omitted, procedure performs binary comparison.

   Dim lngPos       As Long
   Dim lngTemp      As Long
   Dim lngCount     As Long

   ' Specify a starting position. We don't need it the first
   ' time through the loop, but we'll need it on subsequent passes.
   lngPos = 1
   ' Execute the loop at least once.
   Do
      ' Store position at which strFind first occurs.
      lngPos = InStr(lngPos, strText, strFind, lngCompare)
      ' Store position in a temporary variable.
      lngTemp = lngPos
      ' Check that strFind has been found.
      If lngPos > 0 Then
         ' Increment counter variable.
         lngCount = lngCount + 1
         ' Define a new starting position.
         lngPos = lngPos + Len(strFind)
      End If
   ' Loop until last occurrence has been found.
   Loop Until lngPos = 0
   ' Return the number of occurrences found.
   CountOccurrences = lngCount
End Function


