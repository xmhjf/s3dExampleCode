'***************************************************************************************
'  Copyright (C) 2010, Intergraph Corporation.  All rights reserved.
'
'  Project  : M:\SharedContent\Src\CommonRoute\Rules\RouteRunNameRules\3DAPIRouteRunNameRules\
'
'  Class    : ConduitRunNameRule
'
'  Abstract : The file contains implementation of the naming rules for ConduitRun.
'
'***************************************************************************************
Imports System.Collections.ObjectModel
Imports Ingr.SP3D.Common.Middle
Imports SystemsAndSpecsMiddle = Ingr.SP3D.Systems.Middle
Imports Ingr.SP3D.Route.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.ReferenceData.Middle

Public Class ConduitRunNameRule
    Inherits NameRuleBase
    Private Const CountFormat = "0000"
    Private Const CONDUIT = "EC"

    ''' <summary>
    '''  Creates a name for the object passed in. The name is based on the parents
    '''  name and object name.The Naming Parents are added in AddNamingParents().
    '''  Both these methods are called from naming rule semantic.
    ''' </summary>
    ''' <param name="oEntity">Child object that needs to have the naming rule naming.</param>
    ''' <param name="oParents">Naming parents collection.</param>
    ''' <param name="oActiveEntity">Naming rules active entity on which the NamingParentsString is stored.</param>

    Public Overrides Sub ComputeName(ByVal oEntity As BusinessObject, ByVal oParents As ReadOnlyCollection(Of BusinessObject), ByVal oActiveEntity As BusinessObject)

        Dim sRunNewName, sParentSystemName As String, sNameBasis As String = vbNullString
        Dim oParentSystem As SystemsAndSpecsMiddle.System
        Dim oLogError As LogError
        oLogError = MiddleServiceProvider.ErrorLogger()
        Dim sMethod = "ConduitRunNameRule : ComputeName() "
        Try
            oParentSystem = oParents.Item(0)
            sParentSystemName = oParentSystem.Name
            sNameBasis = GetNamingParentsString(oActiveEntity)

            If Not (StrComp(sParentSystemName, sNameBasis, vbTextCompare) = 0) Then

                Dim lCount As Long, sSeqNo As String, sLocationID As String = vbNullString

                GetCountAndLocationID(sParentSystemName, lCount, sLocationID)

                sSeqNo = Format(lCount, CountFormat)

                sRunNewName = sParentSystemName + "-" + CONDUIT + "-" + "-" + sSeqNo

                SetNamingParentsString(oActiveEntity, sRunNewName)
                SetName(oEntity, sRunNewName)

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

        Dim oParentsColl As New Collection(Of BusinessObject)
        Dim oParent As BusinessObject

        GetNamingParents = Nothing
        Try
            oParent = GetParent(HierarchyTypes.System, oEntity) ' Get the System Parent
            If (Not oParent Is Nothing) Then
                oParentsColl.Add(oParent) 'Add the Parent to the ParentColl
            End If
            GetNamingParents = oParentsColl
        Catch ex As Exception
            Throw ex
        End Try

    End Function

End Class
