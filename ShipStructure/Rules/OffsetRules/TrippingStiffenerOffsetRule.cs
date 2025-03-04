//-------------------------------------------------------------------------------------------------------------------------------------
//Copyright 2014 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  TrippingStiffenerOffsetRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\ShipStructure\Rules\Release\SMTrippingRules.dll
//  Original Class Name: ‘ConnectStiffener’ in VB content
//
//Abstract:
//  TrippingStiffenerOffsetRule is a .NET tripping stiffener offset rule which decides the tripping stiffener offset data.
//  This class subclasses from OffsetRuleBase.
//
//Change History:
//  dd.mmm.yyyy    who    change description
//-------------------------------------------------------------------------------------------------------------------------------------
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// TrippingStiffenerOffsetRule is a .NET tripping stiffener offset rule which decides the tripping stiffener key (load) point and offset data on the connecting stiffener.
    /// </summary>
    [DefaultLocalizer(ResourceIdentifiers.DEFAULT_RESOURCE, ResourceIdentifiers.DEFAULT_ASSEMBLY)]
    public class TrippingStiffenerOffsetRule : OffsetRuleBase
    {
        //============================================================================================================
        //DefinitionName/ProgID of this rule is "OffsetRules,Ingr.SP3D.Content.Structure.TrippingStiffenerOffsetRule"
        //============================================================================================================
        #region Public override methods

        /// <summary>
        /// Gets the point offset data, the load point on the connecting stiffener and the offsets (u and v) from that load point for a tripping stiffener.
        ///  •	Refreshes the cached items if necessary
        ///  •	Selects mounting face
        ///  •	Determines the range that the angle falls within
        ///  •	If the angle is below the cross section (the rule writer must determine which angles are above and below the cross section by analyzing the mounting face):
        ///      •	Set the load point = load point of connecting stiffener
        ///  •	Else, sets the load point (and u and v offset, if necessary) based on the angle and load point of the connecting stiffener.
        /// </summary>
        /// <param name="trippingStiffenerCrossSection">The tripping stiffener CrossSection.</param>
        /// <param name="connectedStiffenerCrossSection">The connected stiffener CrossSection.</param>
        /// <param name="approachAngle">The approach angle (degrees).</param>
        /// <param name="connectedStiffenerMountingFaceName">The connected stiffener mounting face name.</param>
        /// <param name="connectedStiffenerLoadPoint">The connected stiffener load point.</param>
        /// <param name="attachmentMethod">The attachment method to the connecting stiffener at the start/end of tripping stiffener.</param>
        /// <param name="isConnectedObjectStiffenerSystem">Flag to determine whether connected object is StiffenerSystem or not.</param>
        /// <returns>Point offset data for the tripping stiffener at requested connected stiffener.</returns>
        public override PointOffsetData GetStiffenerPointOffsetData(CrossSection trippingStiffenerCrossSection, CrossSection connectedStiffenerCrossSection,
            double approachAngle, int connectedStiffenerMountingFaceName, int connectedStiffenerLoadPoint,
            EndAttachmentMethod attachmentMethod, bool isConnectedObjectStiffenerSystem)
        {
            //initialize the offset data with default values
            PointOffsetData pointOffsetData = new PointOffsetData((int)LoadPoint.Bottom, 0.0, 0.0);

            //cache the section parameters
            base.UpdateCrossSectionParameters(trippingStiffenerCrossSection, connectedStiffenerCrossSection);

            //attached connection
            if (isConnectedObjectStiffenerSystem && (attachmentMethod == EndAttachmentMethod.Connected || attachmentMethod == EndAttachmentMethod.Twisted))
            {
                pointOffsetData = base.GetOffsetDataAtNonLappedConnection(base.ConnectedStiffenerCrossSectionParameters, approachAngle, connectedStiffenerLoadPoint, connectedStiffenerMountingFaceName, false);
            }
            else if (attachmentMethod == EndAttachmentMethod.Lapped) //lapped connection
            {
                pointOffsetData = base.GetOffsetDataAtLappedConnectionWithOffset(base.CrossSectionParameters, base.ConnectedStiffenerCrossSectionParameters, approachAngle, connectedStiffenerLoadPoint, connectedStiffenerMountingFaceName);
            }

            return pointOffsetData;
        }

        #endregion Public override methods
    }
}
