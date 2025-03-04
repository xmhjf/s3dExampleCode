'***************************************************************************************
'  Copyright (C) 2011, Intergraph Corporation.  All rights reserved.
'
'  Project  : M:\SharedContent\Src\CommonRoute\Rules\RouteClampNameRules
'
'  Class    : PipeRunNameRule
'
'  Abstract : The file contains implementation of the naming rules for PipeRuns.
'
'***************************************************************************************
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Systems.Middle
Imports Ingr.SP3D.Route.Middle
Imports Ingr.SP3D.Common.Middle.Services


Public Class PipeClampNameRule
    Inherits NameRuleBase
    Private Const strCountFormat = "0000"

    ''' <summary>
    '''  Creates a name for the object passed in. The name is based on the parents
    '''  name and object name.The Naming Parents are added in AddNamingParents().
    '''  Both these methods are called from naming rule semantic.
    ''' </summary>
    ''' <param name="oEntity">Child object that needs to have the naming rule naming.</param>
    ''' <param name="oParents">Naming parents collection.</param>
    ''' <param name="oActiveEntity">Naming rules active entity on which the NamingParentsString is stored.</param>
    Public Overrides Sub ComputeName(ByVal oEntity As BusinessObject, ByVal oParents As ReadOnlyCollection(Of BusinessObject), ByVal oActiveEntity As BusinessObject)
        Dim strLocation As String, strSequenceNumber As String, strConnItemName As String, strConnItemOwnerName As String, strName As String, strParent As String
        Dim lcount As Long
        Dim oConnItemOwnerNamedItem As NamedItemHelper, oConnItemNamedItem As NamedItemHelper
        Dim oLogError As LogError
        oLogError = MiddleServiceProvider.ErrorLogger()
        Dim sMethod = "PipeClampNameRule : ComputeName() "

        If oEntity Is Nothing Or oParents Is Nothing Then
            Exit Sub
        End If

        Try
            oConnItemOwnerNamedItem = New NamedItemHelper(oParents.Item(0))
            strConnItemOwnerName = oConnItemOwnerNamedItem.Name        

            oConnItemNamedItem = New NamedItemHelper(oEntity)
            strConnItemName = GetTypeString(oEntity)            

            strParent = GetNamingParentsString(oActiveEntity)

            If strConnItemName <> strParent Then
                SetNamingParentsString(oActiveEntity, strConnItemName)
                strLocation = vbNullString
                GetCountAndLocationID(strConnItemName, lcount, strLocation)
                strSequenceNumber = Format(lcount, strCountFormat)
                strName = strConnItemOwnerName + "-" + "Clamp" + "-" + strSequenceNumber
                SetName(oEntity, strName)
            End If
        Catch oEx As Exception
            Dim sException As String = sMethod + oEx.Message
            If (Not oLogError Is Nothing) Then
                oLogError.Log(sException)
            End If
            Throw oEx
        End Try
    End Sub
    ''' <summary>
    ''' All the Naming Parents that need to participate in an objects naming are added here to the
    ''' Collection(Of BusinessObject). The parents added here are used in computing the name of the object in
    ''' ComputeName(). Both these methods are called from naming rule semantic.
    ''' </summary>
    ''' <param name="oEntity">Child object that needs to have the naming rule naming.</param>
    ''' <returns> Collection of parents that participate </returns>
    Public Overrides Function GetNamingParents(ByVal oEntity As BusinessObject) As Collection(Of BusinessObject)
        Dim oParent As BusinessObject
        Dim oRelationColl As RelationCollection
        Dim oColl As ReadOnlyCollection(Of BusinessObject)
        Dim oParentColl As New Collection(Of BusinessObject)
        GetNamingParents = Nothing
        If oEntity Is Nothing Then
            Exit Function
        End If
        Try

            oRelationColl = oEntity.GetRelationship("OwnsImpliedItems", "Owner")
            oColl = oRelationColl.TargetObjects()
            oParent = oColl.Item(0)

            If Not oParent Is Nothing Then
                oParentColl.Add(oParent)        ' Add parent to collection
            End If
            GetNamingParents = oParentColl
        Catch ex As CmnException
            Throw ex
        End Try
    End Function
End Class
