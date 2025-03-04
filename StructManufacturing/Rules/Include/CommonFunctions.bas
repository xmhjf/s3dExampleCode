Attribute VB_Name = "CommonFunctions"
'***************************************************************************
'  Copyright (C) 2000, Intergraph Corporation.  All rights reserved.
'
'  Project: GSCADStrMfgNameRules
'
'  Description: This modules contains common subroutines and functions that can be
'               used by GSCAD Commands but mainly for the GSCADStrMfgNameRules project.
'               The following is a list of the functions provided by this module.
'
'       GetMainCategoryName - Returns the Main Member Category Name from given string.
'       GetElementName      - Returns the Element name which is created with combination
'                             of the First & Mid token of given string.
'       GetFirstToken       - Returns the first token of string by given delimiter.
'       GetLastToken        - Returns the last token of string by given delimiter.
'       GetMidToken         - Returns the mid token of string by given delimiter.
'
'  History:
'     Yeun-gyu Kim  01/05/00    Initial creation.
'     Yeun-gyu Kim  01/26/00    Made GetFirstToken function more stable.
'     Marcel veldhuizen 04.04.2004 Included the correct error handling
'
'***************************************************************************

Option Explicit

Private Const Module = "CommonFunctions: "

'****************************************************************************************************
'Description
'   According to the following Use Cases, It returns the Main Member Category Name.
'
'       http://www.gscad.com/DATA/Structural_apps/sdet/UseCases/sdet/Naming/SDS-Naming1.htm
'       http://www.gscad.com/DATA/Structural_apps/sdet/UseCases/sdet/Naming/SDS-Naming2.htm
'       http://www.gscad.com/DATA/Structural_apps/sdet/UseCases/sdet/Naming/SDS-Naming3.htm
'
'   The summarized contents of the Use Cases above are as follows:
'
'   The System member name is composed of Main member name(B-CDE) and Secondary member name(FGH).
'
'   B-CDE-FGH
'
'       B : Plate system location name.
'       C : Serial No. to avoid duplicate main member names.
'       D : Main member category name.
'           ex; TTO(tank top), SHE(shell), DEC(deck), STR(stringer), CHA(chain pipe), ........
'       E : Serial No. to avoid duplicate main member when a main member is splitted.
'       F : Serial No. to avoid duplicating secondary member name.
'       G : Secondary member category name.
'           ex; B(bracket), C(collar plate), D(double plate), F(flat bar), S(shell plate), ...
'       H : When Secondary member is splitted, Serial No. is used to avoid duplicating Secondary member.
'
'   Example: "F30-1BHK1-1A"  (BHK will be returned as the Main Member Category Name)
'****************************************************************************************************
Public Function GetMainCategoryName(strInput As String, strDelimiter As String) As String
    Const METHOD = "GetMainCategoryName"
    On Error GoTo ErrorHandler
    
    Dim strMidToken, chLetter As String
    Dim i, j As Integer
    
    'Find the mid token
    strMidToken = GetMidToken(strInput, strDelimiter)
    
    If strMidToken = "" Then
        'There was no Main Member Category Name found from the input string "strInput".
        GetMainCategoryName = ""
    Else
        'Remove the starting and tailing digital number from the strMidToken.
        GetMainCategoryName = ""
        j = Len(strMidToken)
        For i = 1 To j
            chLetter = Left(strMidToken, 1)
            strMidToken = Right(strMidToken, Len(strMidToken) - 1)
            'Asc("1")=49, Asc("9")=57, Asc("A")=65, Asc("Z")=90, Asc("a")=97, Asc("z")=122
            If Asc(chLetter) > 64 Then
                GetMainCategoryName = GetMainCategoryName & chLetter
            End If
        Next
    End If

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'*********************************************************************************************
' Description:
'   Returns the Element name which is created with combination of the First & Mid token
'   of the inputted string "strInput".
'*********************************************************************************************
Public Function GetElementName(strInput As String, strDelimiter As String) As String
    Const METHOD = "GetElementName"
    On Error GoTo ErrorHandler
    
    Dim strFirstToken, chLetter As String
    Dim strLocation, strMainCategoryName As String
    Dim i, j As Integer
    
    'Find the first token.
    strFirstToken = GetFirstToken(strInput, strDelimiter)
    
    'Takes the location number from the System Location Name (i.e. strFirstToken).
    j = Len(strFirstToken)
    For i = 1 To j
        chLetter = Left(strFirstToken, 1)
        'Asc("1")=49, Asc("9")=57, Asc("A")=65, Asc("Z")=90, Asc("a")=97, Asc("z")=122
        If Asc(chLetter) < 58 Then
            strLocation = strFirstToken
            Exit For
        End If
        strFirstToken = Right(strFirstToken, Len(strFirstToken) - 1)
    Next
    
    'Get the Main Member Category Name.
    strMainCategoryName = GetMainCategoryName(strInput, strDelimiter)
    
    If Len(strMainCategoryName) < 3 Then
        'Takes the first character.
        GetElementName = strLocation + Left(strMainCategoryName, 1)
    Else
        'Takes the first & last character.
        GetElementName = strLocation + Left(strMainCategoryName, 1) + Right(strMainCategoryName, 1)
    End If

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function GetFirstToken(strInput As String, strDelimiter As String) As String
    Const METHOD = "GetFirstToken"
    On Error GoTo ErrorHandler
    
    Dim MyPos
    
    MyPos = InStr(strInput, strDelimiter)
    If MyPos = 0 Then
        GetFirstToken = strInput
    Else
        GetFirstToken = Left(strInput, MyPos - 1)
    End If

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function GetLastToken(strInput As String, strDelimiter As String) As String
    Const METHOD = "GetLastToken"
    On Error GoTo ErrorHandler
    
    Dim SearchString, MyPos
    Dim strLength As Integer
'    Dim strTemp, chLetter As String
'    Dim i, j As Integer

    SearchString = strInput
    strLength = Len(SearchString)
    MyPos = InStr(SearchString, strDelimiter)
    Do While MyPos <> 0
        strLength = strLength - MyPos
        SearchString = Right(SearchString, strLength)
        MyPos = InStr(SearchString, strDelimiter)
    Loop
    
    GetLastToken = SearchString

    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'****************************************************************************************************
'Description
'   Returns the mid token of string if found.
'   It is assumed that the input string (strInput) contains just two strDelimeter character.
'   ex; "F30-1BHK1-1A"  (1BHK1 will be returned)
'****************************************************************************************************
Public Function GetMidToken(strInput As String, strDelimiter As String) As String
    Const METHOD = "GetMidToken"
    On Error GoTo ErrorHandler
    
    Dim SearchString, MyPos
    
    SearchString = strInput
    MyPos = InStr(SearchString, strDelimiter)   'Find 1st delimeter.
    
    SearchString = Right(SearchString, Len(strInput) - MyPos)
    MyPos = InStr(SearchString, strDelimiter)   'Find 2nd delimeter.
    
    If MyPos = 0 Then
        GetMidToken = ""    'There was no mid token found.
    Else
        GetMidToken = Left(SearchString, MyPos - 1)
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
 
