Option Explicit On
Imports System
Imports Ingr.SP3D.Common.Middle
Imports System.Collections.ObjectModel


''' <summary>
''' Copyright (C) 2007 Intergraph Corporation.  All rights reserved.
''' 
''' Class <c>UserDefinedNameRule</c> provides an example of a VB.NET name rule.
''' This rule examines a system and takes action to assure that its name is
''' unique among the system children sharing the same parent system.
''' 
''' Auther:  T. Merchant
''' Date created:  Aug 3, 2007
''' 
''' History:
''' </summary>
''' <remarks>
''' If the name is blank a new one is generated based on the system type and
''' the number of child systems already present.
''' 
''' If the name is non-blank a search is made for another systems under the same
''' parent system with the same name.  If one is found the name of the system being
''' examined by adding a subscript to the name.
''' 
''' Examples:
''' Original name is blank and there are three other systems under the parent system.
'''     System is of type Area.  New name:  NewAreaSystem4.
''' Original name:  A.  Another system named A already exists.  New name:  A(1).
''' Original name:  A.  Systems named A, A(1), A(2) already exist.  New name:  A(3).
''' Original name:  A(2).  Systems named A, A(1), A(2) already exist.  New name:  A(3).
''' Original name:  A(2).  Systems named A, A(1), A(3) already exist.  New name:  A(4).
''' 
''' To use this name rule, modify CatalogData\Bulkload\DataFiles\GenericNamingRules.xls.
''' Add a new line to assign this name rule to a system class or modify an existing assignment.
''' For example, you could replace the standard VB name rule assigned to a system class with
''' this .Net naming rule.  Do do so, enter:
''' "SystemNameRulesNetVB,SystemNameRulesNetVB.UserDefinedNameRule|SystemNameRulesNetVBCab.Cab"
''' in the SolverProgID column.
''' </remarks>
Public Class UserDefinedNameRule
    Inherits NameRuleBase

    Private Const E_DUPLICATENAME As Integer = 4 + vbObjectError + 512
    Private Const E_BLANKNAME As Integer = 4 + vbObjectError + 513
    Private Const E_FAIL As Integer = -2147467259

    ''' <summary>
    ''' Check the name of this sytem for uniqueness among the systems that share its parent system.
    ''' Create a new name if it has no name or if its name is not unique.
    ''' </summary>
    ''' <param name="oObject">Input.  System object whose name is being checked.</param>
    ''' <param name="oParents">Input.  Parent system of oEntity.</param>
    ''' <param name="oActiveEntity">Input.  Naming active entity.</param>
    ''' <remarks></remarks>
    Public Overrides Sub ComputeName(ByVal oObject As BusinessObject, _
            ByVal oParents As ReadOnlyCollection(Of BusinessObject), _
            ByVal oActiveEntity As BusinessObject)

        Try
            'Check for blank names.
            Dim strNewName As String = ""
            If IsNameBlank(oObject, strNewName) Then
                ' A new default name was created for the system.  Raise an error to
                ' let the user know if this change.
                NotifyUser(E_BLANKNAME, "ComputeName", _
                    "Blank system names are not allowed.  Name has been changed to:  " & strNewName, _
                    "NAMING")

            End If

            'Check for duplicate names.
            Dim strOrigName As String = ""
            If NameIsDuplicate(oObject, strOrigName, strNewName) Then
                ' The name of the system being checked is already in use by another
                ' system that shares the same parent.  A subscripted version of the
                ' original name has been created.  raise an error to let the user know
                ' of this change.
                NotifyUser(E_DUPLICATENAME, "ComputeName", _
                    "Duplicate system name:  " & strOrigName & "  has been changed to:  " & strNewName, _
                    "NAMING")
            End If

        Catch ex As Exception
            If Not Err.Source.Equals("NotifyUser") Then
                Throw New Exception("Unexpected error:  SystemNameRulesNetVB.UserDefinedNameRule.ComputeName")
            End If
        End Try

    End Sub

    ''' <summary>
    ''' Get the parent object of the system being named.
    ''' </summary>
    ''' <param name="oEntity">Input.  System object whose name is being checked.</param>
    ''' <returns>Nothing</returns>
    ''' <remarks></remarks>
    Public Overrides Function GetNamingParents(ByVal oEntity As BusinessObject) As Collection(Of BusinessObject)

        GetNamingParents = Nothing

    End Function

    ''' <summary>
    ''' Examine the name to see if it is blank.  If it is blank, generate a new name
    ''' based on its system type and the number of systems present as children of its
    ''' parent system.
    ''' </summary>
    ''' <param name="oSystem">Input.  System whose name is being checked.</param>
    ''' <param name="strGeneratedName">Output.  Generated new name.</param>
    ''' <returns>Boolean that indicates if the name was blank.</returns>
    ''' <remarks></remarks>
    Private Function IsNameBlank(ByVal oSystem As BusinessObject, ByRef strGeneratedName As String) As Boolean

        IsNameBlank = False

        Try

            strGeneratedName = ""

            Dim strCurName As String
            strCurName = GetName(oSystem)
            If strCurName Is Nothing Then
                strCurName = ""
            End If

            If strCurName.Length = 0 Then
                IsNameBlank = True

                'Get the collection of child systems under the parent system
                Dim oParentSystem As BusinessObject
                oParentSystem = GetParent(HierarchyTypes.System, oSystem)
                Dim oChildren As ReadOnlyCollection(Of BusinessObject)
                oChildren = GetSystemChildren(oParentSystem)

                'Create a starting new name based on the system type and number of children
                ' systems already present under the parent system.
                Dim lCount As Long
                lCount = oChildren.Count
                Dim sType As String
                sType = GetTypeString(oSystem)
                strGeneratedName = "New" + sType + "System" + Str(lCount)

                Dim bComparedAllChildren As Boolean = False
                While Not bComparedAllChildren
                    ' Loop through all the sibling systems to see that the proposed
                    ' name is unique.
                    Dim bFoundMatch As Boolean = False
                    Dim iNdx As Integer
                    For iNdx = 0 To iNdx = oChildren.Count - 1
                        Dim oChild As BusinessObject
                        oChild = oChildren(iNdx)
                        Dim strSysName As String
                        strSysName = GetName(oChild)
                        If strGeneratedName = strSysName Then
                            bFoundMatch = True
                            Exit For
                        End If
                    Next iNdx
                    If bFoundMatch Then
                        ' Found another child system with the same name.  Increment the
                        ' count of the proposed name and compare to all siblings again.
                        lCount = lCount + 1
                        strGeneratedName = "New" + sType + "System" + lCount
                        bFoundMatch = False
                    Else
                        ' The proposed name is unique among the sibling systems.
                        bComparedAllChildren = True
                    End If
                End While
                ' Set "Name" property value on IJNamedItem interface
                oSystem.SetPropertyValue(strGeneratedName, "IJNamedItem", "Name")
            End If

        Catch ex As Exception
            Throw New Exception("Unexpected error:  SystemNameRulesNetVB.UserDefinedNameRule.IsnameBlank")
        End Try

    End Function

    ''' <summary>
    ''' Compare the name of the system whose name is being checked and rename if there is
    ''' another system who shares the same parent system that has the same name.
    ''' </summary>
    ''' <param name="oSystem">Input.  System whose name is being checked.</param>
    ''' <param name="strOrigName">Output.  Original name of the system.</param>
    ''' <param name="strNewName">Output.  Modified name of the system.</param>
    ''' <returns>Boolean that indicates if the name is a duplicate among the
    ''' systems that share the same parent.</returns>
    ''' <remarks></remarks>
    Private Function NameIsDuplicate(ByVal oSystem As BusinessObject, _
            ByRef strOrigName As String, ByRef strNewName As String) As Boolean

        NameIsDuplicate = False

        Try

            Dim oParent As BusinessObject
            oParent = GetParent(HierarchyTypes.System, oSystem)

            Dim oChildren As ReadOnlyCollection(Of BusinessObject)
            oChildren = GetSystemChildren(oParent)

            strOrigName = GetName(oSystem)
            strNewName = strOrigName

            Dim lHighSubscript As Long = 0

            ' Loop through each named child of the target parent system to see if there
            ' is already a child object with the same name.
            'Dim iNdx As Integer
            'For iNdx = 0 To iNdx = oChildren.Count - 1
            '    Dim oChild As BusinessObject
            '    oChild = oChildren(iNdx)
            Dim oChild As BusinessObject
            For Each oChild In oChildren

                Dim strChildName As String
                strChildName = GetName(oChild)

                Dim lCurrSubscript As Long = 0

                ' Be sure not to compare the object with itself.
                If Not oChild.Equals(oSystem) Then
                    If CheckNamesMatch(strChildName, strOrigName, lCurrSubscript) Then
                        NameIsDuplicate = True
                        If lCurrSubscript > 0 And lCurrSubscript > lHighSubscript Then
                            lHighSubscript = lCurrSubscript
                        ElseIf lCurrSubscript = -1 And lHighSubscript = 0 Then
                            lHighSubscript = -1
                        End If
                    End If
                End If
            Next oChild
            'Next iNdx

            If lHighSubscript = -1 Then
                ' The name of this system matches the name of at least one other system,
                ' but its subscript is unique.  Leave it as it is.
                NameIsDuplicate = False
            End If

            If NameIsDuplicate Then
                ' The same name already exists
                ' Test to see if the new name already has a subscript
                Dim lNextSubscript As Long
                lNextSubscript = lHighSubscript + 1
                Dim strNextSubscript As String
                strNextSubscript = lNextSubscript.ToString()

                Dim strNameWithoutSubscript As String = ""
                Dim lTmpNbr As Long
                lTmpNbr = NameSubscript(strOrigName, strNameWithoutSubscript)
                If strOrigName.Equals(strNameWithoutSubscript) Then
                    strNewName = strOrigName + "(" + Format(lNextSubscript) + ")"
                Else
                    strNewName = strNameWithoutSubscript + "(" + Format(lNextSubscript) + ")"
                End If
                oSystem.SetPropertyValue(strNewName, "IJNamedItem", "Name")

            End If

        Catch ex As Exception
            Throw New Exception("Unexpected error:  SystemNameRulesNetVB.UserDefinedNameRule.NameIsDuplicate")
        End Try

    End Function

    ''' <summary>
    ''' Add an error to the errors collection that indicates a unacceptable name has been
    ''' corrected.
    ''' </summary>
    ''' <param name="iErrNum">Input.  Error number that indicates either a blank name or
    ''' a duplicate name has been corrected.</param>
    ''' <param name="strSource">Input.  Name of calling subroutine.</param>
    ''' <param name="strDescription">Input.  Description of the problem.</param>
    ''' <param name="strContext">Input.  Error context.</param>
    ''' <remarks>
    ''' Input parameter strContext should be set to "NAMING" for callers in SystemsAndSpecs
    ''' to be able to handle this correctly - by notifying the user but not raising
    ''' the error any higher.
    ''' </remarks>
    Private Sub NotifyUser(ByVal iErrNum As Integer, ByVal strSource As String, _
            ByVal strDescription As String, ByVal strContext As String)

        ' On initial creation we are waiting for a logging service to be provided by the CommonApp
        ' team.  See DI CP122309 - Create a logger service.

    End Sub

    ''' <summary>
    ''' Get a collection of all the system children of a parent system.
    ''' </summary>
    ''' <param name="oParent">Input.  Parent system.</param>
    ''' <returns>Collection of the system children.</returns>
    ''' <remarks></remarks>
    Private Function GetSystemChildren(ByVal oParent As BusinessObject) As ReadOnlyCollection(Of BusinessObject)
        GetSystemChildren = Nothing
        Try

            ' Get relation collection for "SystemHierarchy" relationship with "SystemChildren" rolename
            Dim oRelationCollection As RelationCollection
            oRelationCollection = oParent.GetRelationship("SystemHierarchy", "SystemChildren")

            If Not oRelationCollection Is Nothing Then
                GetSystemChildren = oRelationCollection.TargetObjects
            End If

        Catch ex As Exception
            Throw New Exception("Unexpected error:  SystemNameRulesNetVB.UserDefinedNameRule.GetSystemChildren")
        End Try

    End Function

    ''' <summary>
    ''' Given two strings, check to see if they are the same.  If they are the same,
    ''' create a new name based on the original name with a subscript appended.
    ''' </summary>
    ''' <param name="strOldName">Input.  Original name of system.</param>
    ''' <param name="strNewName">Output.  New name of the system.</param>
    ''' <param name="lOutSub">Output.  Subscript used in creating the unique name, if needed.</param>
    ''' <returns>Boolean flag indicating if the name of the system was unique.</returns>
    ''' <remarks>
    ''' If dealing with system names with subscripts, two names are determined to be
    ''' equal if the base parts of their names are the same.  So A, A(1), and A(10)
    ''' would all be identified as having the same name.
    ''' 
    ''' The subscript is determined by examining all other sibling systems and identifying
    ''' all with the same root name, findign the highest subscript, and adding one.
    '''</remarks>
    Private Function CheckNamesMatch(ByVal strOldName As String, ByRef strNewName As String, _
            ByRef lOutSub As Long) As Boolean
        CheckNamesMatch = False
        Try

            Dim strOldNameWOSub As String
            strOldNameWOSub = ""
            Dim lOldSub As Long
            lOldSub = NameSubscript(strOldName, strOldNameWOSub)

            Dim strNewNameWOSub As String
            strNewNameWOSub = ""
            Dim lNewSub As Long
            lNewSub = NameSubscript(strNewName, strNewNameWOSub)

            If (strNewName.Equals(strOldName)) Then
                CheckNamesMatch = True
                If strOldName.Equals(strOldNameWOSub) Then
                    ' The names match and they don't have subscripts.
                    lOutSub = 0
                Else
                    ' The names match and they do have subscripts.
                    lOutSub = lOldSub
                End If
            Else
                If strNewName.Equals(strOldNameWOSub) Then
                    ' The existing name has a subscript and the new name does
                    ' not.  When the subscript is removed from the existing name,
                    ' it matches the new name.
                    CheckNamesMatch = True
                    lOutSub = lOldSub
                Else
                    If strNewNameWOSub.Equals(strOldName) Then
                        ' The new name has a subscript and the existing name
                        ' does not.  When the subscript is removed from the
                        ' new name, it matches the existing name.
                        CheckNamesMatch = True
                        If lOldSub > lNewSub Then
                            lOutSub = lOldSub
                        Else
                            lOutSub = lNewSub
                        End If
                    Else
                        If strNewNameWOSub.Equals(strOldNameWOSub) Then
                            CheckNamesMatch = True
                            If lOldSub = lNewSub Then
                                ' Both the new name and the existing name have
                                ' subscripts.  The names match and the subscripts
                                ' also match.
                                lOutSub = lOldSub
                            Else
                                ' Both the new name and the existing name have
                                ' subscripts.  The names match but the subscripts
                                ' do not match.
                                If lOldSub > lNewSub Then
                                    lOutSub = lOldSub
                                Else
                                    lOutSub = lNewSub
                                End If
                            End If
                        Else
                            'The new name is not the same as the existing name.
                            CheckNamesMatch = False
                        End If
                    End If
                End If
            End If

            If lOutSub = 0 Then
                lOutSub = 1
            End If

        Catch ex As Exception
            Throw New Exception("Unexpected error:  SystemNameRulesNetVB.UserDefinedNameRule.ChecknamesMatch")
        End Try

    End Function

    ''' <summary>
    ''' Given a system name, determine if it has a subscript and if so, strip that
    ''' subscript from the original name.
    ''' </summary>
    ''' <param name="strOriginalName">Input.  Original name.</param>
    ''' <param name="strStrippedName">Output.  Original name stripped of its subscript.</param>
    ''' <returns>Numeric value of the subscript.</returns>
    ''' <remarks></remarks>
    Private Function NameSubscript(ByVal strOriginalName As String, ByRef strStrippedName As String) As Long
        NameSubscript = 0

        Try

            strStrippedName = strOriginalName
            Dim strWorkingName As String
            strWorkingName = strOriginalName

            ' Look for an open parenthesis that would indicate that we might have
            ' a subscript
            Dim strTest As String
            strTest = "("
            Dim lNbrOpen As Long
            lNbrOpen = CountOccurrences(strWorkingName, strTest)
            strTest = ")"
            Dim lNbrClose As Long
            lNbrClose = CountOccurrences(strWorkingName, strTest)

            If lNbrOpen = 0 Or lNbrOpen <> lNbrClose Then
                ' We have either no open parenthesis or an uneven number of open and close
                ' parenthesis.  In the first case, there is no subscript.  In the second
                ' case, we don't know what to do so we will treat it as if there is no subscript.
            Else
                Dim cDelim As Char
                cDelim = "("

                Dim strSplit() As String
                strSplit = strWorkingName.Split(cDelim)

                Dim lNbrWords As Long
                lNbrWords = strSplit.GetUpperBound(0) - strSplit.GetLowerBound(0) + 1

                If lNbrWords > 1 Then
                    ' We found the open parenthesis and its not the last character in
                    ' the string.
                    strWorkingName = strSplit(0)
                    Dim lNdx As Long
                    For lNdx = 1 To lNdx = lNbrWords - 2
                        strWorkingName = strWorkingName + "(" + strSplit(lNdx)
                    Next lNdx

                    ' Let's see if we have a close
                    ' parenthesis at the end of the second string.
                    Dim strWorkingSubscript As String
                    strWorkingSubscript = strSplit(lNbrWords - 1)

                    ' Verify that there is a closed parenthesis at the end of the
                    ' second string
                    Dim lLong As Long
                    lLong = strWorkingSubscript.Length
                    strTest = ")"
                    lNbrClose = CountOccurrences(strWorkingSubscript, strTest)
                    If lNbrClose = 1 Then
                        ' There is a closed parenthesis somewhere in the second string.
                        '  Is it at the end of the string?
                        cDelim = ")"
                        strSplit = strWorkingSubscript.Split(cDelim)
                        Dim lShort As Long
                        lShort = strSplit(0).Length
                        If lLong = (lShort + 1) Then
                            ' We have the index - in string form.  Convert it.
                            If IsNumeric(strSplit(0)) Then
                                NameSubscript = Long.Parse(strSplit(0))
                                ' Strip the subscript from the name
                                strStrippedName = strWorkingName
                            End If
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            Throw New Exception("Unexpected error:  SystemNameRulesNetVB.UserDefinedNameRule.NameSubscript")
        End Try

    End Function

    ''' <summary>
    ''' Given a string and a string to find, count how many times the string to
    ''' find is present in the string.
    ''' </summary>
    ''' <param name="strText">Input.  String to be examined.</param>
    ''' <param name="strFind">Input.  Character being searched for.</param>
    ''' <param name="lngCompare">Optional input.  Compare method.</param>
    ''' <returns>Number of times the character being searched for appears in
    ''' the input string.</returns>
    ''' <remarks></remarks>
    Private Function CountOccurrences(ByVal strText As String, ByVal strFind As String, _
    Optional ByVal lngCompare As Microsoft.VisualBasic.CompareMethod = CompareMethod.Binary) As Long
        CountOccurrences = 0

        Try

            ' Count occurrences of a particular character or characters.
            ' If lngCompare argument is omitted, procedure performs binary comparison.

            Dim lngPos As Long
            Dim lngTemp As Long
            Dim lngCount As Long

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

        Catch ex As Exception
            Throw New Exception("Unexpected error:  SystemNameRulesNetVB.UserDefinedNameRule.CountOccurrences")
        End Try

    End Function

End Class

