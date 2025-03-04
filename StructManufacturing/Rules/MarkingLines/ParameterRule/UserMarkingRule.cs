//----------------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Purpose: Sample User Marking Rule which provides the required information to create the Markings by the user.
//
//
//      Author:  Suma Mallena
//
//      History:
//      August 24th, 2015   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// User Marking Rule that provides the required information to create the Markings by the user.
    /// </summary>
    public class UserMarkingRule : UserMarkingRuleBase
    {
        /// <summary>
        /// Returns the list of marking types to create the User Marking.
        /// </summary>
        /// <param name="userMarkingInformation">The user marking information for getting the configuration data.</param>
        /// <exception cref="CmnNullArgumentException">Raised when userMarkingInformation is null.</exception>
        public override Collection<int> GetAllowableTypes(UserMarkingInformation userMarkingInformation)
        {
            Collection<int> markingTypes = new Collection<int>(); ;

            try
            {
                if (userMarkingInformation == null)
                    throw new CmnNullArgumentException("userMarkingInformation");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                Marking userMarking = null;
                if (userMarkingInformation.ManufacturingPart != null)
                {
                    userMarking = (Marking)userMarkingInformation.ManufacturingPart;
                }

                if (userMarking == null)
                    return null;

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                BusinessObject parentPart = userMarkingInformation.ManufacturingParent;

                // Get the marking types from the configuration XML.
                Dictionary<int, object> markingTypesFromCofiguration = null;
                if (parentPart is PlatePartBase)
                {
                    markingTypesFromCofiguration  = userMarkingInformation.GetArguments("PlateMarkingTypes");                 
                }
                else if ( parentPart is StiffenerPartBase)
                {
                    markingTypesFromCofiguration = userMarkingInformation.GetArguments("ProfileMarkingTypes"); 
                }
                else if (parentPart is MemberPart)
                {
                    markingTypesFromCofiguration = userMarkingInformation.GetArguments("MemberMarkingTypes");
                }
                else if (parentPart is PinJig)
                {
                    markingTypesFromCofiguration = userMarkingInformation.GetArguments("PinJigMarkingTypes"); 
                }
                else
                {
                    return null;
                }

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                foreach (object type in markingTypesFromCofiguration.Values)
                {
                    markingTypes.Add(Convert.ToInt32(type));
                }
                
                #endregion Set Outputs

            }
            catch (Exception e)
            {
                LogForToDoList(e, 4034, "Call to User Marking Rule failed with the error to get the marking types." + e.Message);
            }

            return markingTypes;
        }

        /// <summary>
        /// Returns the information of the related entity of marking.
        /// </summary>
        /// <param name="userMarkingInformation">The user marking information for getting the configuration data.</param>
        /// <param name="referenceEntity">The reference entity of the marking.</param>
        /// <param name="relatedEntity">The related entity of the marking.</param>
        /// <param name="markingType">The user marking type.</param>
        /// <exception cref="CmnNullArgumentException">Raised when userMarkingInformation is null.</exception>
        public override MarkingRelatedEntityInformation GetMarkingRelatedEntityInformation(UserMarkingInformation userMarkingInformation, BusinessObject referenceEntity, BusinessObject relatedEntity, int markingType)
        {
            MarkingRelatedEntityInformation relatedEntityInputs = null;
            try
            {
                if (userMarkingInformation == null)
                    throw new CmnNullArgumentException("userMarkingInformation");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                Marking userMarking = null;
                if (userMarkingInformation.ManufacturingPart != null)
                {
                    userMarking = (Marking)userMarkingInformation.ManufacturingPart;
                }

                if (userMarking == null)
                    return null;

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                relatedEntityInputs = new MarkingRelatedEntityInformation();

                switch (markingType)
                {
                    case (int)ManufacturingGeometryType.PlateLocationMark:
                    case (int)ManufacturingGeometryType.ProfileLocationMark:
                        {
                            relatedEntityInputs.ReferenceName = referenceEntity.ToString();
                            break;
                        }
                    case (int)ManufacturingGeometryType.MarginMark:
                        {
                            relatedEntityInputs.MaximumAssemblyMarginValue = 0.02;
                            relatedEntityInputs.MaximumFabricationMarginValue = 0.03;
                            relatedEntityInputs.MaximumCustomMarginValue = 0.04;
                            break;
                        } 
                }

                #endregion Processing

            }
            catch (Exception e)
            {
                LogForToDoList(e, 4034, "Call to User Marking Rule failed with the error to get the marking related entity attributes." + e.Message);
            }

            return relatedEntityInputs;
        }

        /// <summary>
        /// Returns the filter criteria for the marking step based on the type.
        /// Depending upon the command step (Reference Curve step or Related Part step), the filter criteria can be controlled for markings of various types.
        /// </summary>
        /// <param name="userMarkingInformation">The user marking information for getting the configuration data.</param>
        /// <param name="filterStep">Indicates the marking command step.</param>
        /// <param name="markingType">The user marking type.</param>
        /// <exception cref="CmnNullArgumentException">Raised when userMarkingInformation is null.</exception>
        public override string GetLocateFilter(UserMarkingInformation userMarkingInformation, int filterStep, int markingType)
        {
            string locateCriteria = null;

            const string IID_IJPlateSystem = "{E0B23CD4-7CEB-11d3-B351-0050040EFC17}";
            const string IID_IJStiffenerSystem = "{E0B23CD5-7CEB-11d3-B351-0050040EFC17}";
            const string IID_IJPlatePart = "{780F26C2-82E9-11D2-B339-080036024603}";
            const string IID_IJProfilePart = "{69F3E7BF-40A0-11D2-B324-080036024603}";
            const string IID_IJRefCurveOnSurface = "{EBE54C96-77B1-11D5-8A1D-00C04F79B54E}";
            const string IID_IJSeam = "{02C1327F-2C31-11D2-8329-0800367F3D03}";

            try
            {
                if (userMarkingInformation == null)
                    throw new CmnNullArgumentException("userMarkingInformation");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                // Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                switch (filterStep)
                {
                    case 1: //ReferenceStep
                        {
                            switch ((ManufacturingGeometryType)markingType)
                            {
                                case ManufacturingGeometryType.PlateLocationMark:
                                case ManufacturingGeometryType.ProfileLocationMark:
                                case ManufacturingGeometryType.MarginMark:
                                    {
                                        locateCriteria = IID_IJRefCurveOnSurface +
                                                         " OR " + "[StrMfgMarkingLinesCmd.CMarkingFilter,IsLogicalConnection]" +
                                                         " OR " + "[StrMfgMarkingLinesCmd.CMarkingFilter,IsEdgePort]";
                                        break;
                                    }
                                case ManufacturingGeometryType.PinJigMark:
                                    {
                                        locateCriteria = "[StrMfgMarkingLinesCmd.CMarkingFilter,IsLogicalConnection]";
                                        break;
                                    }
                                default:
                                    {
                                        locateCriteria = IID_IJRefCurveOnSurface +
                                                           " OR " + "[StrMfgMarkingLinesCmd.CMarkingFilter,IsLogicalConnection]" +
                                                           " OR " + "[StrMfgMarkingLinesCmd.CMarkingFilter,IsEdgePort]";
                                        break;
                                    }
                            }

                            break;
                        }

                    case 2: //RelatedPartStep
                        {
                            switch ((ManufacturingGeometryType)markingType)
                            {
                                case ManufacturingGeometryType.PlateLocationMark:
                                    {
                                        locateCriteria = IID_IJPlateSystem + " OR " + IID_IJPlatePart;
                                        break;
                                    }
                                case ManufacturingGeometryType.ProfileLocationMark:
                                    {
                                        locateCriteria = IID_IJStiffenerSystem + " OR " + IID_IJProfilePart;
                                        break;
                                    }
                                case ManufacturingGeometryType.PinJigMark:
                                    {
                                        locateCriteria = IID_IJPlateSystem + " OR " + IID_IJPlatePart + " OR " +
                                                         IID_IJStiffenerSystem + " OR " + IID_IJProfilePart + " OR " +
                                                         IID_IJRefCurveOnSurface + " OR " + IID_IJSeam;
                                        break;
                                    }
                                default:
                                    {
                                        locateCriteria = IID_IJPlateSystem + " OR " + IID_IJPlatePart + " OR " +
                                                         IID_IJStiffenerSystem + " OR " + IID_IJProfilePart + " OR " +
                                                         IID_IJRefCurveOnSurface + " OR " + IID_IJSeam +
                                                         " OR " + "[StrMfgMarkingLinesCmd.CMarkingFilter,IsLogicalConnection]" +
                                                         " OR " + "[StrMfgMarkingLinesCmd.CMarkingFilter,IsEdgePort]";
                                        break;
                                    }
                            }
                            break;
                        }
                }
                #endregion

            }
            catch (Exception e)
            {
                LogForToDoList(e, 4034, "Call to User Marking Rule failed with the error to get the marking locate criteria." + e.Message);
            }

            return locateCriteria;
        }
    }
}
