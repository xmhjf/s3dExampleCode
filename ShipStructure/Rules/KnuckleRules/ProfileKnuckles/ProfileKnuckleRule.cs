//----------------------------------------------------------------------------------------------------------------
//Copyright 2014 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  ProfileKnuckleRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\ShipStructure\Rules\Release\ProfileKnuckleRules.dll
//  Original Class Name: ‘ProfileKnuckleInit’ in VB content
//
//Abstract:
//  ProfileKnuckleRule is a .NET ProfileKnuckle rule which decides the ProfileKnuckle’s treatment type and inner radius.
//  This class subclasses from ProfileKnuckleRuleBase.
//
//Change History:
//  dd.mmm.yyyy    who    change description
//----------------------------------------------------------------------------------------------------------------
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// ProfileKnuckle rule which decides the ProfileKnuckle’s treatment type and inner radius.
    /// </summary>
    public class ProfileKnuckleRule : ProfileKnuckleRuleBase
    {
        //====================================================================================================
        //DefinitionName/ProgID of this rule is "KnuckleRules,Ingr.SP3D.Content.Structure.ProfileKnuckleRule"
        //====================================================================================================

        #region Instance variables

        /// <summary>
        /// Constant for section type FB.
        /// </summary>
        private const string FlatBar = "FB";

        /// <summary>
        /// Constant for section type UA.
        /// </summary>
        private const string UnequalAngle = "UA";

        /// <summary>
        /// Constant for section type EA.
        /// </summary>
        private const string EqualAngle = "EA";

        /// <summary>
        /// Constant for section type B.
        /// </summary>
        private const string Bulb = "B";

        #endregion Instance variables

        #region Public override methods

        /// <summary>
        /// Gets the knuckle edge feature root class.
        /// It should match the class name in the bulkload file.
        /// </summary>
        /// <returns>The edge feature root class string.</returns>        
        public override string EdgeFeatureRootSelector
        {
            get
            {
                return "RootEdge";
            }
        }

        /// <summary>
        /// Gets the edge port for knuckle edge feature. This method returns the edge port to create kunckle edge feature.
        /// </summary>
        /// <param name="profileKnuckle">The profile knuckle.</param>   
        /// <returns>The edge port in which to create the edge feature.</returns>
        public override TopologyPort GetEdgePortForKnuckleEdgeFeature(ProfileKnuckle profileKnuckle)
        {
            //This method returns the edge port to create kunckle edge feature
            //The edge port is on profile trim geometry
            TopologyPort edgePort = null;

            //if there is not a profile part associated yet then return
            ProfilePart knuckleProfilePart = profileKnuckle.ProfilePart;
            if (knuckleProfilePart != null)
            {
                SectionFaceType face1 = SectionFaceType.Unknown;
                SectionFaceType face2 = SectionFaceType.Unknown;

                //if the profile knuckle is due to a plate knuckle then the corner feature is on the bottom of the web
                //this is assuming that the stiffener is mounted on its bottom face (this is true 99.99% of the time)
                if (profileKnuckle.ReferenceCurve != null)
                {
                    face1 = SectionFaceType.Web_Left;
                    face2 = SectionFaceType.Bottom;
                }
                else
                {
                    string profilePartSectionType = knuckleProfilePart.SectionType;

                    //the knuckle is due to a landing curve path
                    //There is no face for a FlatBar on which a feature can be place
                    //A T section is currently not supported.  This would require a feature on both sides of the web.
                    switch (profilePartSectionType)
                    {
                        //again the assumption is that the mounting face is bottom
                        case ProfileKnuckleRule.UnequalAngle:
                        case ProfileKnuckleRule.EqualAngle:
                            face1 = SectionFaceType.Top;
                            face2 = SectionFaceType.Top_Flange_Right;
                            break;
                        case Bulb:
                            face1 = SectionFaceType.Top;
                            face2 = SectionFaceType.Top_Flange_Right_Top_Corner;
                            break;
                        default:
                            break;
                    }
                }

                if (face1 != SectionFaceType.Unknown && face2 != SectionFaceType.Unknown)
                {
                    edgePort = knuckleProfilePart.GetLateralEdgePort(GeometryOperationTypes.PartFinalTrim, face1, face2, false);
                }
            }

            return edgePort;
        }

        /// <summary>
        /// Gets the profile knuckle treatment type.  'Split' is returned unless the profile knuckle is associated to a plate knuckle
        /// and that plate knuckle treatment type is 'Bend'.  In such case, 'Bend' is returned as the profile knuckle treatment type.
        /// </summary>
        /// <param name="profile">The profile.</param>
        /// <param name="profileKnuckle">The profile knuckle.</param>
        /// <returns>The profile knuckle treatment type.</returns>        
        public override ProfileKnuckleTreatmentType GetTreatmentType(Profile profile, ProfileKnuckle profileKnuckle)
        {
            //the GetTreatmentType on the base class returns Split as default implementation.                          
            return base.GetTreatmentType(profile, profileKnuckle);
        }

        /// <summary>
        /// Gets the inner radius (meters) of a profile knuckle.
        /// </summary>
        /// <param name="profileKnuckle">The profile knuckle.</param>
        /// <returns>The profile knuckle inner radius.</returns>        
        public override double GetInnerRadius(ProfileKnuckle profileKnuckle)
        {
            double innerRadius = 0;

            ProfilePart profilePart = profileKnuckle.ProfilePart;

            if (profilePart != null)
            {
                CrossSection profilePartCrossSection = profilePart.CrossSection;
                string sectionType = profilePart.SectionType;

                //for now it should always be false according to the comment on the VB6 rule unless the profile knuckle treamemnt type is 
                //feature at bend, then we ask the profile knuckle if it is convex or not.
                bool profileKnuckleConvex = false;

                if (this.GetTreatmentType(profilePart, profileKnuckle) == ProfileKnuckleTreatmentType.FeatureAtBend)
                {
                    profileKnuckleConvex = profileKnuckle.IsConvex;
                }

                //only bend at web.  May need to add condition in the future for flange.                
                if (sectionType == ProfileKnuckleRule.FlatBar || !profileKnuckleConvex)
                {
                    //get the web thickness
                    double webThickness = StructHelper.GetDoubleProperty(profilePartCrossSection, "IJUAXSectionWeb", "WebThickness");

                    //if the web thickness is less than 10mm then the inner radius is 40mm
                    //else if the web thickness is less than 20mm ( 10mm < thickness < 20mm) then the inner radius is 80mm
                    //else if the web thickness is less than 30mm ( 20mm < thickness < 30mm) then the inner radius is 120mm
                    //else (thickness > 30mm) then the inner radius is 160mm
                    if (webThickness < 0.01)
                    {
                        innerRadius = 0.04;
                    }
                    else if (webThickness < 0.02)
                    {
                        innerRadius = 0.08;
                    }
                    else if (webThickness < 0.03)
                    {
                        innerRadius = 0.12;
                    }
                    else
                    {
                        innerRadius = 0.16;
                    }
                }
                else
                {
                    //for all other cross sections the inner radius is 4 times the cross section height (depth)
                    double profileHeight = profilePartCrossSection.Depth;

                    if (profileHeight < 0.0001 && profilePartCrossSection.SupportsInterface("IJUAXSectionWeb"))
                    {
                        profileHeight = StructHelper.GetDoubleProperty(profilePartCrossSection, "IJUAXSectionWeb", "WebLength");
                    }

                    innerRadius = profileHeight * 4;
                }
            }
            else
            {
                //the inner radius returned from the base class is 0.001 as this is the default when no profile part is available
                innerRadius = base.GetInnerRadius(profileKnuckle);
            }

            return innerRadius;
        }

        /// <summary>
        /// Determines whether the knuckle type is allowed for the profile knuckle for the given treatment type.
        /// Assumptions are:
        /// 'Bend' is only allowed for profiles with 'FlatBar' cross-section. 'Ignore' or 'Split' treatment is always allowed.
        /// 'Split&Extend' treatment is only allowed for manual Knuckle.
        /// For other treatment types, uses knuckle angle and manual knuckle check to decide if knuckle type is allowed.
        /// </summary>
        /// <param name="profileKnuckle">The profile knuckle.</param>
        /// <param name="profileKnuckleTreatmentType">The profile knuckle treatment type.</param>
        /// <returns>true if the knuckle type allowed, otherwise false.</returns>        
        public override bool IsKnuckleTypeAllowed(ProfileKnuckle profileKnuckle, ProfileKnuckleTreatmentType profileKnuckleTreatmentType)
        {
            bool isKnuckleTypeAllowed = false;

            //if the base returns not allowed then return false, that means that the knuckle is not related to any part yet
            //or that the profile knuckle does not implements IJProfileKnuckleMfg
            //if the base returns that it is allowed then continue with the rule checks to determine if it is indeed allowed
            if (!base.IsKnuckleTypeAllowed(profileKnuckle, profileKnuckleTreatmentType)) { return isKnuckleTypeAllowed; }

            bool isProfileKnuckleAssociatedWithPlateKnuckle = false;

            //if the profile knuckle is associated to a PlateKnuckle then the reference curve will not be null
            if (profileKnuckle.ReferenceCurve != null)
            {
                isProfileKnuckleAssociatedWithPlateKnuckle = true;
            }

            bool isProfileKnuckleManual = profileKnuckle.IsManual; //created using manual profile knuckle command

            ProfilePart profileKnucklePart = profileKnuckle.ProfilePart;
            string crossSectionType = profileKnucklePart.SectionType;

            switch (profileKnuckleTreatmentType)
            {
                case ProfileKnuckleTreatmentType.Bend:
                    if (crossSectionType == ProfileKnuckleRule.FlatBar) { isKnuckleTypeAllowed = true; }
                    break;
                case ProfileKnuckleTreatmentType.Extend:
                    isKnuckleTypeAllowed = this.IsExtendTreatmentTypeAllowed(profileKnuckle, isProfileKnuckleManual, isProfileKnuckleAssociatedWithPlateKnuckle);
                    break;
                case ProfileKnuckleTreatmentType.FeatureAtBend:
                    double knuckleAngle = Math3d.Deg(profileKnuckle.Angle);  //convert from radians to degrees
                    if (!isProfileKnuckleManual) { isKnuckleTypeAllowed = this.IsFeatureAtBendTreamentTypeAllowed(profileKnuckle, crossSectionType, knuckleAngle, isProfileKnuckleAssociatedWithPlateKnuckle); }
                    break;
                case ProfileKnuckleTreatmentType.Ignore:
                case ProfileKnuckleTreatmentType.Split:
                    isKnuckleTypeAllowed = true;
                    break;
                case ProfileKnuckleTreatmentType.SplitAndExtend:
                    isKnuckleTypeAllowed = isProfileKnuckleManual;
                    break;
                default:
                    break;
            }

            return isKnuckleTypeAllowed;
        }

        #endregion Public override methods

        #region Private methods

        /// <summary>
        /// Determines if 'Extend' treatment type is allowed for the profile knuckle. 
        /// Assumptions: 'Extend' is allowed for Manual Profile Knuckle or Profile Knuckle associated with a Plate Knuckle which is at the
        /// start or end of the profile part.
        /// </summary>
        /// <param name="profileKnuckle">The profile knuckle.</param>
        /// <param name="isManual">The profile knuckle is created using the manual profile knuckle command or not.</param>
        /// <param name="isAssociatedWithPlateKnuckle">The profile knuckle is associated with plate knuckle or not.</param>
        /// <returns>True if ProfileKnuckle type is allowed; otherwise, false.</returns>
        private bool IsExtendTreatmentTypeAllowed(ProfileKnuckle profileKnuckle, bool isManual, bool isAssociatedWithPlateKnuckle)
        {
            bool isTypeAllowed = false;

            if ((isManual || isAssociatedWithPlateKnuckle) && profileKnuckle.AtStartOrEnd)
            {
                isTypeAllowed = true;
            }

            return isTypeAllowed;
        }

        /// <summary>
        /// Determines if 'FeatureAtBend' treatment type is allowed for the profile knuckle. 
        /// Assumptions:  'FeatureAtBend' is allowed for certain knuckle angle and profile height if 
        /// the Profile Knuckle is not associated with a Plate Knuckle, it is not Convex, and it's cross-section type is an 'Angle'.
        /// </summary>
        /// <param name="profileKnuckle">The profile knuckle.</param>
        /// <param name="crossSectionType">Type of the cross section.</param>
        /// <param name="angle">The angle.</param>
        /// <param name="isAssociatedWithPlateKnuckle">The profile knuckle is associated with plate knuckle or not.</param>
        /// <returns>True if ProfileKnuckle type is allowed; otherwise, false.</returns>
        private bool IsFeatureAtBendTreamentTypeAllowed(ProfileKnuckle profileKnuckle, string crossSectionType, double angle, bool isAssociatedWithPlateKnuckle)
        {
            bool isTypeAllowed = false;
            bool isConvexKnuckle = profileKnuckle.IsConvex;

            if (!isConvexKnuckle && !isAssociatedWithPlateKnuckle)
            {
                CrossSection crossSection = profileKnuckle.ProfilePart.CrossSection;
                double profileHeight = 0.0;
                if (crossSection.SupportsInterface("IStructCrossSectionDimensions"))
                {
                    profileHeight = crossSection.Depth;
                }

                //If Depth is not defined, then use WebLength as height
                if (profileHeight < 0.0001 && crossSection.SupportsInterface("IJUAXSectionWeb"))
                {
                    profileHeight = StructHelper.GetDoubleProperty(crossSection, "IJUAXSectionWeb", "WebLength");
                }

                //bend and insert at flang
                if (crossSectionType == ProfileKnuckleRule.UnequalAngle || crossSectionType == ProfileKnuckleRule.EqualAngle)
                {
                    //if the knuckle angle is greater than 90 degrees then depending of the profile height it must be slightly
                    //less than straight (180 degrees)
                    if (angle > 90)
                    {
                        if (profileHeight >= 0.2 && profileHeight < 0.3)
                        {
                            if (angle <= 179)
                            {
                                isTypeAllowed = true;
                            }
                        }
                        else if (profileHeight >= 0.3)
                        {
                            if (angle <= 178)
                            {
                                isTypeAllowed = true;
                            }
                        }
                    }
                }
            }

            return isTypeAllowed;
        }

        #endregion Private methods
    }
}