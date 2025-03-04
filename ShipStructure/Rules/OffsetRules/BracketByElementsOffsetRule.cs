//------------------------------------------------------------------------------------------------------------------------------
//Copyright 2014 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  BracketByElementsOffsetRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\ShipStructure\Rules\Release\SMTrippingRules.dll
//  Original Class Name: ‘ConnectBracket’ in VB content
//
//Abstract:
//  BracketByElementsOffsetRule is a .NET bracket by elements offset rule which decides the bracket by elements offset data.
//  This class subclasses from OffsetRuleBase.
//
//Change History:
//  dd.mmm.yyyy    who    change description
//------------------------------------------------------------------------------------------------------------------------------
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// BracketByElementsOffsetRule is a .NET bracket by elements offset rule which decides the bracket by elements offset data on the given support.
    /// </summary>
    [DefaultLocalizer(ResourceIdentifiers.DEFAULT_RESOURCE, ResourceIdentifiers.DEFAULT_ASSEMBLY)]
    public class BracketByElementsOffsetRule : OffsetRuleBase
    {
        //============================================================================================================
        //DefinitionName/ProgID of this rule is "OffsetRules,Ingr.SP3D.Content.Structure.BracketByElementsOffsetRule"
        //============================================================================================================
        #region Public override methods

        /// <summary>
        /// Evaluates the offset data at given connected object based on "Angle of Approach"
        ///  •	Refreshes the cached items if necessary
        ///  •	Selects mounting face
        ///  •	Determines the range that the angle falls within
        ///  •	If the angle is below the cross section (the rule writer must determine which angles are above and below the cross section by analyzing the mounting face):
        ///      •	Set the attachment load point = mounting point of connected element
        ///  •	Else, sets the load point (and u and v offset, if necessary) based on the angle and mounting point of the connected element.
        /// </summary>
        /// <param name="supports">The bracket by elements supports.</param>
        /// <param name="supportIndex">The bracket support index at which the offset data is requested.</param>
        /// <param name="connectedStiffenerCrossSection">The support stiffener CrossSection.</param>
        /// <param name="approachAngle">The angle (degrees) of the 2D vector that is between the mouse cursor and the intersection point of the supporting object.</param>
        /// <param name="supportStiffenerLoadPoint">The support stiffener load point.</param>
        /// <param name="supportStiffenerMountingFaceName">The support stiffener mounting face name.</param>
        /// <param name="isConnectedObjectStiffenerSystem">Flag to determine whether connected object is StiffenerSystem or not.</param>
        /// <param name="isConnectionLapped">True in case of lapped connection (bracket plane parallel to the lateral face of the supporting stiffener), 
        ///                            <para>false for attached case (bracket plane intersects to the lateral face of the supporting stiffener).</para></param>
        /// <returns>Point offset data for the bracket at requested connected stiffener.</returns>
        public override PointOffsetData GetBracketPointOffsetData(Collection<BracketSupportDefinition> supports,
            SupportIndex supportIndex, CrossSection connectedStiffenerCrossSection, double approachAngle, int supportStiffenerLoadPoint,
            int supportStiffenerMountingFaceName, bool isConnectedObjectStiffenerSystem, bool isConnectionLapped)
        {
            //initialize the offset data with default values
            PointOffsetData pointOffsetData = new PointOffsetData((int)LoadPoint.Bottom, 0.0, 0.0);

            //cache the section parameters
            base.UpdateCrossSectionParameters(connectedStiffenerCrossSection);

            bool isOtherSupportEdgeReinforcement = false;

            // connected object is stiffener
            if (isConnectedObjectStiffenerSystem)
            {
                //get the corresponding support profiles (either EdgeReinforcementSystem or StiffenerSystem) using index
                StiffenerSystemBase support1 = null;
                StiffenerSystemBase support2 = null;

                //get the first two supports in the collection if available.
                foreach (BracketSupportDefinition supportDefinition in supports)
                {
                    switch (supportDefinition.Index)
                    {
                        case SupportIndex.First:
                            support1 = supportDefinition.Support as StiffenerSystemBase;
                            break;
                        case SupportIndex.Second:
                            support2 = supportDefinition.Support as StiffenerSystemBase;
                            break;
                        case SupportIndex.Third://we are not interested in third support right now
                            break;
                        default:
                            this.ErrorStatus = new OffsetRuleErrorStatus(OffsetRuleErrorType.InvalidSectionType, base.GetString(ResourceIdentifiers.ErrInvalidSupportIndex, "Invalid support index") + " " + supportDefinition.Index.ToString());
                            //error is created so stop processing
                            return pointOffsetData;
                    }
                }

                //attached connection
                if (!isConnectionLapped)
                {
                    isOtherSupportEdgeReinforcement = base.IsOtherSupportEdgeReinforcement(supportIndex, support1, support2);
                    pointOffsetData = GetOffsetDataAtNonLappedConnection(base.ConnectedStiffenerCrossSectionParameters, approachAngle, supportStiffenerLoadPoint, supportStiffenerMountingFaceName, isOtherSupportEdgeReinforcement);
                }
                else //lapped case
                {
                    pointOffsetData = GetOffsetDataAtLappedConnection(base.ConnectedStiffenerCrossSectionParameters, approachAngle, supportStiffenerLoadPoint, supportStiffenerMountingFaceName);
                }

                //if other support is EdgeReinforcementSystem then, set the vOffset to 0.0
                if (supportIndex == SupportIndex.First)
                {
                    if (support1 is StiffenerSystem && support2 is EdgeReinforcementSystem)
                    {
                        pointOffsetData.VOffset = 0.0;
                    }
                }
                else if (supportIndex == SupportIndex.Second)
                {
                    if (support2 is StiffenerSystem && support1 is EdgeReinforcementSystem)
                    {
                        pointOffsetData.VOffset = 0.0;
                    }
                }
            }
            else
            {
                object connectedSupport = null;

                //get the first two supports in the collection if available.
                foreach (BracketSupportDefinition supportDefinition in supports)
                {
                    if (supportIndex == supportDefinition.Index)
                    {
                        connectedSupport = supportDefinition.Support;
                    }
                }

                // If conneceted support is member system 
                if (connectedSupport != null && connectedSupport is MemberSystem)
                {
                    MemberSystem memberSupport = connectedSupport as MemberSystem;
                    isOtherSupportEdgeReinforcement = false;
                    supportStiffenerMountingFaceName = 0;
                    if (!isConnectionLapped)
                    {
                        pointOffsetData = GetOffsetDataAtNonLappedConnection(base.ConnectedStiffenerCrossSectionParameters, approachAngle, supportStiffenerLoadPoint, supportStiffenerMountingFaceName, isOtherSupportEdgeReinforcement, memberSupport);
                    }
                    else //lapped case
                    {
                        pointOffsetData = GetOffsetDataAtLappedConnection(base.ConnectedStiffenerCrossSectionParameters, approachAngle, supportStiffenerLoadPoint, supportStiffenerMountingFaceName, memberSupport);
                    }
                }
                else
                {
                    this.ErrorStatus = new OffsetRuleErrorStatus(OffsetRuleErrorType.InvalidSectionType, base.GetString(ResourceIdentifiers.ErrInvalidSupportIndex, "Invalid support"));
                    //error is created so stop processing
                    return pointOffsetData;
                }
            }
            return pointOffsetData;
        }

        #endregion Public override methods
    }
}