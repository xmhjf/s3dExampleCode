'***************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project  : M:\SharedContent\Src\CommonRoute\Rules\RouteCompNameRules\
'
'  Class    : PipeCompNameRule
'
'  Abstract : The file contains implementation of the naming rules for PipeComponents.
'
'***************************************************************************************
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports SystemsAndSpecsMiddle = Ingr.SP3D.Systems.Middle
Imports Ingr.SP3D.Route.Middle
Imports Ingr.SP3D.Common.Middle.Services

Public Class PipeCompNameRule
    Inherits NameRuleBase
    Private Const strCountFormat = "0000"       ' define fixed-width number field for name

    Public Overrides Sub ComputeName(ByVal oEntity As Common.Middle.BusinessObject, ByVal oParents As System.Collections.ObjectModel.ReadOnlyCollection(Of Common.Middle.BusinessObject), ByVal oActiveEntity As Common.Middle.BusinessObject)
        Dim oNamedItem As NamedItemHelper
        Dim strName As String, strSequenceNo As String, strLocation As String, strNameSuffix As String, strNameBasis As String
        Dim lcount As Long
        Dim oLogError As LogError
        oLogError = MiddleServiceProvider.ErrorLogger()
        Dim sMethod = "PipeCompNameRule : ComputeName() "
        strNameSuffix = Nothing

        If oEntity Is Nothing Then
            Exit Sub
        End If

        Try
            oNamedItem = New NamedItemHelper(oEntity)
            'To get the Engineering Tag
            Dim oPart As RoutePart
            Dim oFeatColl As ReadOnlyCollection(Of RouteFeature)
            Dim oPathFeat As IPipePathFeature
            oPart = oEntity
            If oPart.IsBasePart() = True Then
                oFeatColl = oPart.Features
                oPathFeat = oFeatColl.Item(0)
                strNameSuffix = oPathFeat.Tag
                If Len(Trim(strNameSuffix)) <= 0 Then
                    strNameSuffix = oPathFeat.ShortCode
                End If
            Else
                If Not oPart Is Nothing Then
                    Dim oPipeComp As PipeComponent
                    oPipeComp = oPart
                    strNameSuffix = oPipeComp.ShortCode
                End If
            End If

            If Len(Trim(strNameSuffix)) <= 0 Then
                strNameSuffix = GetTypeString(oEntity)
            End If
            strNameBasis = GetNamingParentsString(oActiveEntity)
            If strNameSuffix <> strNameBasis Then
                'Adding Runname as  naming parent string
                SetNamingParentsString(oActiveEntity, strNameSuffix)
                strLocation = vbNullString
                GetCountAndLocationID(strNameSuffix, lcount, strLocation)
                strSequenceNo = Format(lcount, strCountFormat)
                strName = strNameSuffix + "-" + strSequenceNo
                'Setting Name on the Object
                oNamedItem = New NamedItemHelper(oEntity)
                SetName(oEntity, strName)
            End If
        Catch ex As Exception
            Dim sException As String = sMethod + ex.Message
            If (Not oLogError Is Nothing) Then
                oLogError.Log(sException)
            End If
            Throw ex
        End Try
    End Sub

    Public Overrides Function GetNamingParents(ByVal oEntity As BusinessObject) As System.Collections.ObjectModel.Collection(Of BusinessObject)                
        Dim oParentColl As New Collection(Of BusinessObject)
        GetNamingParents = Nothing
        Dim oPart As RoutePart
        Dim oFeatColl As ReadOnlyCollection(Of RouteFeature)
        Dim oPathFeat As IPipePathFeature = Nothing
        If oEntity Is Nothing Then
            Exit Function
        End If
        Try
            oPart = oEntity
            If Not oPart Is Nothing Then
                oFeatColl = oPart.Features
                If oFeatColl.Count >= 1 Then
                    oPathFeat = oFeatColl.Item(0)
                End If
            End If
            If Not oPathFeat Is Nothing Then
                oParentColl.Add(oPathFeat)
            End If
            GetNamingParents = oParentColl
        Catch ex As Exception
            Throw ex
        End Try

    End Function
End Class
