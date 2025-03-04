//----------------------------------------------------------------------------------
//      Copyright (C) 2015 Intergraph Corporation.  All rights reserved.
//
//      Purpose: Sample PinJig Remarking Rule. This serves as an example for customizing the rule with different implementation.
//               It provides the following :
//               - Remarking Surface
//               - Entities and Geometries that create a Remarking Line of the PinJig of given type.
//               - Remarking types that satisfy a particular purpose.
//               - Filter criteria for the PinJig Remarking step based on the remarking type.
//
//      Author:  Suma Mallena
//
//      History:
//      March 1st, 2015   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Grids.Middle;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Sample PinJig Remarking Rule.
    /// </summary>
    public class RemarkRuleSample : PinJigRemarkRuleBase
    {
        /// <summary>
        /// Provides the remarking surface of the PinJig.
        /// </summary>
        /// <param name="pinJigInformation">The pinjig info.</param>
        public override ISurface GetRemarkingSurface(PinJigInformation pinJigInformation)
        {
            ISurface remarkingSurface = null;
            try
            {
                if (pinJigInformation == null)
                    throw new ArgumentNullException("Input pinJigInformation is empty.");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                PinJig mfgPinJig = null;
                if (pinJigInformation.ManufacturingPart != null)
                {
                    mfgPinJig = (PinJig)pinJigInformation.ManufacturingPart;
                }

                if (mfgPinJig == null)
                    return null;

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                // Get Remarking Surface.
                int referenceSurfaceType = Convert.ToInt32(pinJigInformation.GetArguments("ReferenceSurfaceType", "RemarkingSurface").FirstOrDefault().Value);
                int offsetType = Convert.ToInt32(pinJigInformation.GetArguments("OffsetType", "RemarkingSurface").FirstOrDefault().Value);
                      
                remarkingSurface = base.GetOffsetSurface(mfgPinJig, (OffsetType)offsetType, 0.0, (ReferenceSurfaceType)referenceSurfaceType);

                #endregion Processing

            }
            catch (Exception e)
            {
                LogForToDoList(e, "SMCustomWarningMessages", 5022, "Call to PinJig Remarking Surface Rule failed with the error" + e.Message);
            }

            return remarkingSurface;

        }

        /// <summary>
        /// Returns the entities that create a remarking of the pin jig of given type.
        /// </summary>
        /// <param name="pinJigInformation">Data class that pin jig rules use for getting the inputs and also information from configuration.</param>
        /// <param name="attributeName">The pin jig remarking attribute name for which the entities have to be returned.</param>
        public override ReadOnlyCollection<BusinessObject> GetRemarkingEntities(PinJigInformation pinJigInformation, string attributeName)
        {
            List<BusinessObject> remarkingEntities = null;
            try
            {
                if (pinJigInformation == null)
                    throw new ArgumentNullException("Input pinJigInformation is empty.");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                PinJig mfgPinJig = null;
                if (pinJigInformation.ManufacturingPart != null)
                {
                    mfgPinJig = (PinJig)pinJigInformation.ManufacturingPart;
                }

                if (mfgPinJig == null)
                    return null;

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                // Get Remarking Entities                

                //ConnectionType:
                // 1 = Logical
                // 2 = Physical
                int connectionType;

                switch (attributeName)
                {
                    case "SeamRemark":
                        connectionType = Convert.ToInt32(pinJigInformation.GetArguments("ConnectionType", "SeamRemark").FirstOrDefault().Value);
                        remarkingEntities = base.GetSeamRemarkEntities(mfgPinJig, (ConnectionType)connectionType).ToList<BusinessObject>();
                        break;

                    case "PlateRemark":
                        connectionType = Convert.ToInt32(pinJigInformation.GetArguments("ConnectionType", "PlateRemark").FirstOrDefault().Value);
                        remarkingEntities = base.GetPlateRemarkEntities(mfgPinJig, (ConnectionType)connectionType).ToList<BusinessObject>();
                        break;

                    case "ProfileRemark":
                        connectionType = Convert.ToInt32(pinJigInformation.GetArguments("ConnectionType", "ProfileRemark").FirstOrDefault().Value);
                        remarkingEntities = base.GetProfileRemarkEntities(mfgPinJig, (ConnectionType)connectionType).ToList<BusinessObject>();
                        break;

                    case "GridLineX_Remark":
                        if (mfgPinJig.CoordinateSystemFromParent != null)
                        {
                            remarkingEntities = base.GetGridPlanes(mfgPinJig, AxisType.X).ToList<BusinessObject>();
                        }
                        break;

                    case "GridLineY_Remark":
                        if (mfgPinJig.CoordinateSystemFromParent != null)
                        {
                            remarkingEntities = base.GetGridPlanes(mfgPinJig, AxisType.Y).ToList<BusinessObject>();
                        }
                        break;

                    case "GridLineZ_Remark":
                        if (mfgPinJig.CoordinateSystemFromParent != null)
                        {
                            remarkingEntities = base.GetGridPlanes(mfgPinJig, AxisType.Z).ToList<BusinessObject>();
                        }
                        break;

                    case "RefCurveRemark":
                        remarkingEntities = base.GetReferenceCurveLines(mfgPinJig, PlateType.Hull).ToList<BusinessObject>();
                        break;

                    case "UserRemark":
                        remarkingEntities = new List<BusinessObject>();
                        remarkingEntities.AddRange(base.GetUserMarkings(mfgPinJig, ManufacturingGeometryType.PinJigDiagonalMark).ToList<BusinessObject>());
                        remarkingEntities.AddRange(base.GetUserMarkings(mfgPinJig, ManufacturingGeometryType.PinJigMark).ToList<BusinessObject>());

                        break;

                    case "UserExtend":
                        remarkingEntities = base.GetUserMarkings(mfgPinJig, ManufacturingGeometryType.ExtendedPinJigMark).ToList<BusinessObject>();
                        break;

                    default:
                        ReadOnlyCollection<BusinessObject> unidentifiedEntities = GetUnidentifiedRemarkingEntities(pinJigInformation);
                        if (unidentifiedEntities != null)
                        {
                            remarkingEntities = unidentifiedEntities.ToList<BusinessObject>();
                        }
                        break;
                }


                #endregion Processing

            }
            catch (Exception e)
            {
                LogForToDoList(e, "SMCustomWarningMessages", 5022, "Call to PinJig Remarking Entities Rule failed with the error" + e.Message);
            }

            if (remarkingEntities != null)
            {
                return new ReadOnlyCollection<BusinessObject>(remarkingEntities);
            }
            else
            {
                return null;
            }

        }

        /// <summary>
        /// Returns the geometries of a remarking of the pin jig of given type.
        /// </summary>
        /// <param name="pinJigInformation">Data class that pin jig rules use for getting the inputs and also information from configuration.</param>
        /// <param name="attributeName">The pin jig remarking attribute name for which the entities have to be returned.</param>
        /// <param name="entitiesToRemark">The entities for which the remarkings need to be created.</param>
        public override ReadOnlyCollection<ManufacturingGeometry> GetRemarkingGeometries(PinJigInformation pinJigInformation, string attributeName, IEnumerable<BusinessObject> entitiesToRemark)
        {
            ReadOnlyCollection<ManufacturingGeometry> remarkingGeometries = null;

            try
            {
                if (pinJigInformation == null)
                    throw new ArgumentNullException("Input pinJigInformation is empty.");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                PinJig mfgPinJig = null;
                if (pinJigInformation.ManufacturingPart != null)
                {
                    mfgPinJig = (PinJig)pinJigInformation.ManufacturingPart;
                }

                if (mfgPinJig == null)
                    return null;

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                // Get Remarking Geometries               

                //ConnectionType:
                // 1 = Logical
                // 2 = Physical
                int connectionType = -1;

                switch (attributeName)
                {
                    case "SeamRemark":
                        connectionType = Convert.ToInt32(pinJigInformation.GetArguments("ConnectionType", "SeamRemark").FirstOrDefault().Value);
                        remarkingGeometries = base.GetSeamRemarkings(mfgPinJig, entitiesToRemark, (ConnectionType)connectionType);
                        break;

                    case "PlateRemark":
                        connectionType = Convert.ToInt32(pinJigInformation.GetArguments("ConnectionType", "PlateRemark").FirstOrDefault().Value);

                        int attributeValue = mfgPinJig.RemarkingSettings.PlateRemarking;
                        switch (attributeValue)
                        {
                            case 1: //Apply
                                remarkingGeometries = base.GetPlateRemarkings(mfgPinJig, entitiesToRemark, (ConnectionType)connectionType, 1, PlateType.Unspecified);
                                break;
                            case 2: // Both Sides
                                remarkingGeometries = base.GetPlateRemarkings(mfgPinJig, entitiesToRemark, (ConnectionType)connectionType, 2, PlateType.DeckPlate);
                                break;
                            case 3: // Both sides with LC
                                remarkingGeometries = base.GetPlateRemarkings(mfgPinJig, entitiesToRemark, (ConnectionType)connectionType, 3, PlateType.DeckPlate);
                                break;
                        }

                        break;

                    case "ProfileRemark":
                        connectionType = Convert.ToInt32(pinJigInformation.GetArguments("ConnectionType", "ProfileRemark").FirstOrDefault().Value);
                        remarkingGeometries = base.GetProfileRemarkings(mfgPinJig, entitiesToRemark, (ConnectionType)connectionType);
                        break;

                    case "GridLineX_Remark":
                        remarkingGeometries = base.GetGridRemarkings(mfgPinJig, entitiesToRemark, ManufacturingSubGeometryType.PinJigXFrameRemarking);
                        break;

                    case "GridLineY_Remark":
                        remarkingGeometries = base.GetGridRemarkings(mfgPinJig, entitiesToRemark, ManufacturingSubGeometryType.PinJigYFrameRemarking);
                        break;

                    case "GridLineZ_Remark":
                        remarkingGeometries = base.GetGridRemarkings(mfgPinJig, entitiesToRemark, ManufacturingSubGeometryType.PinJigZFrameRemarking);
                        break;

                    case "RefCurveRemark":
                        remarkingGeometries = base.GetReferenceCurveRemarkings(mfgPinJig, entitiesToRemark, PlateType.Hull);
                        break;

                    case "UserRemark":
                        remarkingGeometries = base.GetUserRemarkings(mfgPinJig, entitiesToRemark);
                        break;

                    case "UserExtend":
                        remarkingGeometries = base.GetUserExtendRemarkings(mfgPinJig);
                        break;

                    case "SeamControlRemark":
                        remarkingGeometries = null; //TODO
                        break;

                    default:
                        remarkingGeometries = GetUnidentifiedRemarkingGeometries(pinJigInformation);
                        break;

                }

                #endregion Processing

            }
            catch (Exception e)
            {
                LogForToDoList(e, "SMCustomWarningMessages", 5022, "Call to PinJig Remarking Geometry Rule failed with the error" + e.Message);
            }

            return remarkingGeometries;

        }

        /// <summary>
        /// Returns the list of Remarking types that satisfy a particular purpose.
        /// </summary>
        /// <param name="pinJigInformation">Data class that pin jig rules use for getting the inputs and also information from configuration.</param>
        /// <param name="purposeOfRemarkingTypes">Specifies the purpose of the remarking type.</param>
        public override Collection<int> GetRemarkingTypes(PinJigInformation pinJigInformation, RemarkingPurposeType purposeOfRemarkingTypes)
        {
            Collection<int> remarkingTypes = null;
            try
            {
                if (pinJigInformation == null)
                    throw new ArgumentNullException("Input pinJigInformation is empty.");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                // Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                switch (purposeOfRemarkingTypes)
                {
                    case RemarkingPurposeType.IncludeInMarkingCommand:
                        //Specify the types of remarking lines that should be displayed
                        //within the RAD 2D environment of the marking command.
                        remarkingTypes = new Collection<int>();
                        remarkingTypes.Add((int)ManufacturingSubGeometryType.PinJigPlateRemarking);
                        remarkingTypes.Add((int)ManufacturingSubGeometryType.PinJigProfileRemarking);

                        break;

                    case RemarkingPurposeType.ExcludeFromIntersectionPointCreation:
                        //Specify the types of remarking lines that should be excluded
                        // from participating in the intersection point creation process.
                        remarkingTypes = new Collection<int>();
                        remarkingTypes.Add((int)ManufacturingGeometryType.NavalArchLine);
                        remarkingTypes.Add((int)ManufacturingSubGeometryType.PinJigReferenceCurveRemarking);
                        remarkingTypes.Add((int)ManufacturingGeometryType.PinJigDiagonalMark);

                        break;

                    case RemarkingPurposeType.RemarkingLinePriority:
                        // Specify rank (order of descending priority) of remarking line types.
                        // When two remarking lines are identical (they completely overlap each other)
                        // the line whose type has higher rank (lower index) will be used for
                        remarkingTypes = new Collection<int>();
                        remarkingTypes.Add((int)ManufacturingGeometryType.PinJigFloorContourLine);
                        remarkingTypes.Add((int)ManufacturingSubGeometryType.PinJigFrameRemarking);
                        remarkingTypes.Add((int)ManufacturingSubGeometryType.PinJigSeamRemarking);

                        // Above example, Contour lines will be preferred over Frames, and
                        // frames will be preferred over seams.  Types not listed above will
                        // "lose" to types listed above.  If two lines overlap, and neither of
                        // their types are listed above, one will be arbitrarily picked.

                        break;
                }

                #endregion Processing
 
            }
            catch (Exception e)
            {
                LogForToDoList(e, "SMCustomWarningMessages", 5023, "Call to PinJig Remarking Types Rule failed with the error" + e.Message);
            }

            return remarkingTypes;
        }

        /// <summary>
        /// Returns the filter criteria for the pin jig remarking step based on the remarking type.
        /// </summary>
        /// <param name="pinJigInformation">Data class that pin jig rules use for getting the inputs and also information from configuration.</param>
        /// <param name="attributeName">The pin jig remarking attribute name for which the locate filter have to be returned.</param>
        public override string GetLocateFilter(PinJigInformation pinJigInformation, string attributeName)
        {
            string remarkingFilter = null;
            try
            {
                if (pinJigInformation == null)
                    throw new ArgumentNullException("Input pinJigInformation is empty.");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                PinJig mfgPinJig = null;
                if (pinJigInformation.ManufacturingPart != null)
                {
                    mfgPinJig = (PinJig)pinJigInformation.ManufacturingPart;
                }

                if (mfgPinJig == null)
                    return null;

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                // Get Remarking Filter                

                switch (attributeName)
                {
                    case "SeamRemark":
                        remarkingFilter = base.GetDefaultLocateFilter(RemarkingEntityType.DesignSeam);
                        break;

                    case "PlateRemark":
                        remarkingFilter = base.GetDefaultLocateFilter(RemarkingEntityType.ConnectedPlatePart);
                        break;

                    case "ProfileRemark":
                        remarkingFilter = base.GetDefaultLocateFilter(RemarkingEntityType.ConnectedProfilePart);
                        break;

                    case "GridLineX_Remark":
                        remarkingFilter = base.GetDefaultLocateFilter(RemarkingEntityType.XPlane);
                        break;

                    case "GridLineY_Remark":
                        remarkingFilter = base.GetDefaultLocateFilter(RemarkingEntityType.YPlane);
                        break;

                    case "GridLineZ_Remark":
                        remarkingFilter = base.GetDefaultLocateFilter(RemarkingEntityType.ZPlane);
                        break;

                    case "RefCurveRemark":
                        remarkingFilter = base.GetDefaultLocateFilter(RemarkingEntityType.NavalArchLine);
                        break;

                    case "UserRemark":
                        remarkingFilter = base.GetDefaultLocateFilter(RemarkingEntityType.UserPinJigMark);
                        break;

                    case "UserExtend":
                        remarkingFilter = base.GetDefaultLocateFilter(RemarkingEntityType.ExtendPinJigIntersectionMark);
                        break;

                    case "SeamControlRemark":
                        remarkingFilter = base.GetDefaultLocateFilter(RemarkingEntityType.DesignSeam);
                        break;
                }

                #endregion Processing

            }
            catch (Exception e)
            {
                LogForToDoList(e, "SMCustomWarningMessages", 5022, "Call to PinJig Remarking Filter Rule failed with the error" + e.Message);
            }

            return remarkingFilter;

        }

     }
}
