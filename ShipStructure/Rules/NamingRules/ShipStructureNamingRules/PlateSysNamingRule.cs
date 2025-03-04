using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Common.Middle.Services.Hidden;
using Ingr.SP3D.Content.Structure;
using System.Collections.ObjectModel;
using Ingr.SP3D.Grids.Middle;

using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// This is an example class from which a new plate system naming rule can be derived.  It
    /// implements a single required interface, INamingSolverBase, which will be used by the
    /// relevant naming semantics to name a Molded Forms business object
    /// </summary>
    /// 

    // The WrapperProgID is a required attribute and should always be set to "MarineRuleWrappers.NamingSolverWrapper"
    [WrapperProgID("MarineRuleWrappers.NamingSolverWrapper")]
    public class PlateSysNamingRule : INamingSolverBase
    {
        #region INamingSolverBase Members

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oEntity">Business object to be named</param>
        /// <param name="oActiveEntity">Naming active entity associated with the business object 
        /// to be named.  If no active entity exists, this will contain the design parent of
        /// the the business object to be named</param>
        /// <param name="strIdx">String index (1-based) uniquely identifying the business object based
        /// within its associated naming active entity or its design parent.  This could be optionally
        /// incorporated into the returned string</param>
        /// <returns>A string value used to name the business object</returns>
        virtual public string ComputeName(Ingr.SP3D.Common.Middle.BusinessObject oEntity, Ingr.SP3D.Common.Middle.BusinessObject oActiveEntity, string strIdx)
        {
            try
            {
                // This naming rule only works on bracket plate systems or root plate systems
                // which both inherit from PlateSystemBase.  Retrieve the PlateSystemBaseObject
                // by casting oEntity.
                PlateSystemBase oPlateSysBase = (PlateSystemBase)oEntity;

                //Leaf and root plate systems are named in the same naming rule, but we name
                //them in a slightly different manner.
                if (oPlateSysBase.IsRoot == true)
                {
                    // Root plate systems are named in the form:
                    // <refplane>-<index><platetype>-<location id>
                    // 
                    // refplane: The name of the coincident reference plane defining the plate or
                    // "A" if it is not applicable (such as with plane by offset, plane by elements,
                    // etc.
                    // index: The naming index as determined by the naming rule wrapper
                    // platetype: A string representation of the plate type.  See GetPlateTypeString
                    // for the string returned for each plate type
                    // location id: The location ID for the site as specified in Project Management
                    return GetNamingPlaneString(oPlateSysBase) + "-" + strIdx +
                        GetPlateTypeString(oPlateSysBase) + "-" +
                        MiddleServiceProvider.SiteMgr.ActiveSite.ActiveLocation.NamingString;
                }
                else
                {
                    // Leaf plate systems are simpler.  They inherit the root plate system name
                    // and append the naming index to it
                    return ((INamedItem)oPlateSysBase.SystemParent).Name + "-" + strIdx;
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("PlateSysNamingRule.ComputeName: Error encountered (" + e.Message + ")");
            }

            return "Error"; // On the occurrence of an error, return "Error"
        }

        /// <summary>
        /// When a naming rule references one or more objects in order to calculate
        /// the name of another object, these should be specified via the GetNamingParents method.
        /// 
        /// In this rule, a root plate system has its defining plane as its only naming parent.  A
        /// leaf plate system will return its design parent
        /// </summary>
        /// <param name="oEntity">The business object to be named</param>
        /// <returns>A collection of naming parents for which an active entity will be created
        /// and associated</returns>
        virtual public System.Collections.ObjectModel.Collection<Ingr.SP3D.Common.Middle.BusinessObject> GetNamingParents(Ingr.SP3D.Common.Middle.BusinessObject oEntity)
        {
            Collection<BusinessObject> oRetColl = new Collection<BusinessObject>();
            PlateSystemBase oPlateSysBase = (PlateSystemBase)oEntity;

            try
            {
                // if it is a root plate system and not a hull, we want to add the coincident plane
                // as a naming parent.
                // A hull plate system has no naming parent
                if (oPlateSysBase.IsRoot == true)
                {
                    if (oPlateSysBase.Type != PlateType.Hull)
                    {
                        BusinessObject oNamingPlane = GetNamingPlane(oPlateSysBase);
                        
                        // Don't attempt to add a null naming parent.
                        if (oNamingPlane != null) oRetColl.Add(oNamingPlane);
                    }
                }
                else
                {
                    // leaf plate system.  Return design parent as naming parent
                    oRetColl.Add((BusinessObject)oPlateSysBase.SystemParent);
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("PlateSysNamingRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }
            return oRetColl;
        }
        #endregion


        protected string GetNamingPlaneString(PlateSystemBase oPlateSysBase)
        {
            string strRetVal = "A"; // default value is "A"

            BusinessObject oNamingPlane = GetNamingPlane(oPlateSysBase);

            if (oNamingPlane != null)
            {
                strRetVal = oNamingPlane.ToString();
            }
            else
            {
                strRetVal = GetPlaneGlobalPositionString(oPlateSysBase);
            }

            return strRetVal;
        }

        /// <summary>
        /// Private method creates the reference plane part of the string used by ComputeName.
        /// </summary>
        /// <param name="oPlateSysBase">Plate/bracket system object to be named</param>
        /// <returns>String representation of the plane defining the plate system or bracket</returns>
        protected BusinessObject GetNamingPlane(PlateSystemBase oPlateSysBase)
        {
            BusinessObject oNamingPlane = null;
            PlaneDefinition oPlaneDef = null;
            try
            {
                oPlaneDef = oPlateSysBase.PlaneDefinition;
                if (oPlaneDef != null)
                {
                    switch (oPlaneDef.Type)
                    {
                        // it is possible to specifiy strings for any type of plane 
                        // definition type available
                        case PlaneDefinitionType.Coincident:
                            oPlateSysBase.PlaneDefinition.GetInputs(out oNamingPlane);
                            break;
                        case PlaneDefinitionType.Offset:
                            double offset = 0.0;
                            oPlateSysBase.PlaneDefinition.GetInputs(out oNamingPlane, out offset);
                            break;
                    }
                    if (!(oNamingPlane is GridPlaneBase))
                        oNamingPlane = null;
                }
            }
            // Plane by elements is not currently supported.  Swallow the exception.
            catch (Exception)
            {
                oNamingPlane = null;
            }
            return oNamingPlane;
        }

        /// <summary>
        /// This creates the plate type part of the string used by ComputeName
        /// </summary>
        /// <param name="oPlateSysBase">Plate/bracket system object to be named</param>
        /// <returns>String representation of the plane defining the plate system or bracket</returns>
        protected string GetPlateTypeString(PlateSystemBase oPlateSysBase)
        {
            string strPlateType = "Unknown";    // default value is "Unknown"

            switch (oPlateSysBase.Type)
            {                   
                case PlateType.DeckPlate:           // deck plates
                    strPlateType = "DCK";
                    break;
                case PlateType.TBulkheadPlate:      // transverse bulkhead plates
                    strPlateType = "TBH";
                    break;
                case PlateType.LBulkheadPlate:      // longitudinal bulkhead plates
                    strPlateType = "LBH";
                    break;
                case PlateType.LongitudinalTube:    // longitudinal tube plates
                    strPlateType = "LT";
                    break;
                case PlateType.TransverseTube:      // transverse tube plates
                    strPlateType = "TT";
                    break;
                case PlateType.VerticalTube:        // vertical tube plates
                    strPlateType = "VT";
                    break;
                case PlateType.TubePlate:           // tube plates
                    strPlateType = "TB";
                    break;
                case PlateType.FlangePlate:         // flange plates
                    strPlateType = "FL";
                    break;
                case PlateType.WebPlate:            // web plates
                    strPlateType = "WB";
                    break;
                case PlateType.GeneralPlate:        // general plates
                    strPlateType = "GP";
                    break;
            }

            return strPlateType;
        }
        /// <summary>
        /// This creates the planar plate global position string used by ComputeName
        /// </summary>
        /// <param name="oPlateSysBase">Plate/bracket system object to be named</param>
        /// <returns>String representation of planar plate global position</returns>
        protected string GetPlaneGlobalPositionString(PlateSystemBase oPlateSysBase)
        {
            string strPositionString = "A"; //Arbitrary position
            try
            {
                PlaneDefinition oPlaneDef = null;
                try
                {
                    oPlaneDef = oPlateSysBase.PlaneDefinition;
                }
                // some plane constructions are not currently supported.  Swallow the exception.
                catch (Exception)
                {
                    oPlaneDef = null;
                }
                //no plane definition it is not created by any existing plae definitions.
                if (oPlaneDef == null)
                {
                    SurfaceGeometryType plateSurfaceGeometryType = oPlateSysBase.SurfaceGeometryType;
                    //Still it is a planar means there is no reference when the plate is created
                    //use global positioning naming pattern
                    if (plateSurfaceGeometryType == SurfaceGeometryType.Planar)
                    {
                        Ingr.SP3D.Common.Middle.IPlane oPlane = null;
                        oPlane = (Ingr.SP3D.Common.Middle.IPlane)((PlateSystem)oPlateSysBase).GetPorts(PortType.Face)[1];                        
                        if (oPlane != null)
                        {
                            Vector oPlateNormalVector = oPlane.Normal;
                            Position oPlateRootPoint = oPlane.RootPoint;

                            Vector oXaxis = new Vector(1, 0, 0);
                            Vector oYaxis = new Vector(0, 1, 0);
                            Vector oZaxis = new Vector(0, 0, 1);

                            double dDotX = Math.Abs(oPlateNormalVector % oXaxis);
                            double dDotY = Math.Abs(oPlateNormalVector % oYaxis);
                            double dDotZ = Math.Abs(oPlateNormalVector % oZaxis);

                            int iPosition = 0;
                            if (Math.Abs(dDotX - 1) <= Math3d.DistanceTolerance)
                            {
                                //roundoff the X position to nearest mm
                                iPosition = (int)Math.Round(oPlateRootPoint.X * 1000);
                                strPositionString = "X" + iPosition.ToString();
                            }
                            else if (Math.Abs(dDotY - 1) <= Math3d.DistanceTolerance)
                            {
                                //roundoff the Y position to nearest mm
                                iPosition = (int)Math.Round(oPlateRootPoint.Y * 1000);
                                strPositionString = "Y" + iPosition.ToString();
                            }
                            else if (Math.Abs(dDotZ - 1) <= Math3d.DistanceTolerance)
                            {
                                //roundoff the Z position to nearest mm
                                iPosition = (int)Math.Round(oPlateRootPoint.Z * 1000);
                                strPositionString = "Z" + iPosition.ToString();
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                strPositionString = "A"; //Arbitrary position
            }
            return strPositionString;
        }
    }
}
