//-----------------------------------------------------------------------------
//      Copyright (C) 2011-16 Intergraph Corporation.  All rights reserved.
//
//      Component: This SlotMappingRule returns section alias and ports of Penetrating
//                 Object for the Plate/Plate or Plate/Stiffener Comination  
//
//      Author:  
//
//      History:
//      OCt 19, 2011       BSLee                  Created
//      Jun 25, 2013       svsmylav               TR-232503: for plate penetrating top and bottom ports of
//                                                a member, origin of sketching plane is adjusted to member mid-height.                                              
//      Feb 14, 2014       GH -                 TR-CP-248074- Modified GetExternalLateralEdgePortCol() method to return the lateral Edges Ports which are connected.                  
//      Apr 04, 2014       Hema                   TR-CP-250621 :Modified GetCrossSectionCoordSys() method supporting for curved deck plate
//      Apr 23, 2015       Kameswari              TR-CP-264747  Slot and collars are not created for some of the plates in customer model.  
//	    Sep 16, 2105	   RPK                    CR-279479 Updated Slot Mapping Rule to Look at Naming Category in Addition to Plate type
//      Nov 26,2015        pkakula                CR-274363 Updated slot mapping rule to use plate sub-type to drive plate-thru-plate slot behaviour.
//      Nov 19, 2015       mchandak               DI-275514 and DI-275249 Updated Slot Mapping Rule to avoid interop calls for GetNormalFromPostion() and IsExternalWire()    
//		Nov 19, 2106	   RPK			  DI-CP-275254 Replced enumConnectableTransientPorts/GetPersistentPort with GetPorts(.Net API)
//      December 11, 2015       hgajula                DI-CP-284051  Fix coverity defects stated in November 6, 2015 report
//      December 11, 2015       knukala                DI-CP-275525  Replace PlaceIntersectionObject method in SlotMapng rule with .Net API 
//      February 3, 2016        PYK                     DI-CP-287121  Fix coverity defects stated in January 15, 2016 report
//      March 11, 2016     svsmylav               TR-288446 Modified 'GetSectionAlias' method incase penetrated part is not bound to the base plate.
//      March    28,2016        HG                     DI-CP-291260  Fix Coverity issues stated in March 11, 2016
//      April 1, 2016      svsmylav               TR-291159 Modified 'GetEmulatedSectionAlias' method to use flange seam start position
//                                                to determine ER web right port outward normal.
//      April 12, 2016     dsmamidi               TR-236727 Modified GetSectionAlias() to check for seam between penetrating and penetrated.
//      April 14, 2016     svsmylav               TR-288446 Modified 'GetConnectedPenetratingParts' method to avoid Hull plate using plate system base
//                                                  (earlier type of plate part is attempted leading to errors).
//      May 16, 2016       RPK               TR-CP-294261 	Modified the penetrated object type returned from GetRootSystemFromPart() to PlateSystemBase 
//-----------------------------------------------------------------------------

using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Configuration;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Common.Middle.Services.Hidden;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Content.Structure.EmulatedPortMappings;


namespace Ingr.SP3D.Content.Structure
{

    /// <summary>
    /// Configuration of Penetrating Object
    /// </summary>
    internal enum PenetratingConfig
    {
        /// <summary>
        /// Cannot determine the Penetrating Configuration
        /// </summary>
        Unknown = -1,
        /// <summary>
        /// Plate + Plate
        /// </summary>
        Plate_Plate = 0,
        /// <summary>
        /// Plate + ER
        /// </summary>
        Plate_ER = 1,
        /// <summary>
        /// Plate + Stiffener    
        /// </summary>
        Plate_Stiffener = 2,
        /// <summary>
        /// Plate Only, It is FlatBar    
        /// </summary>
        Plate_Only = 3,
        /// <summary>
        /// Profile_Only, It is just Stiffener    
        /// </summary>
        Stiffener_Only = 4,
        /// <summary>
        /// Bend Plate only, It is just Bend Plate in some cross section form
        /// </summary>
        BendPlate_Only = 5
    }

    //SectionAlias
    //BUT,BUTL2, BUTL3,FB,EA,UA

    /// <summary>
    /// This rule is intended to map ports from connected plates and stiffeners that penetrate a common object to a standard stiffener cross section type.  
    /// The goal is reuse of existing slot and collar symbols and rules.  Only a few simple cases need to be mapped
    /// </summary>    
    public class SlotMappingRule : SlotMappingRuleBase, ICustomSlotMappingRule
    {
        private double tolerance = 0.0001; //0.1 mm 
        private double toleranceOneMilliMeter = 0.001; //1.0 mm

        private Seam webSeamObj = null;
        private Seam flangeSeamObj = null;
        private object webObj = null;
        private object flangeObj = null;
        private string secAlias = "";

        private Dictionary<int, IPort> mappedPortCollection = null;
        private IPort basePlatePort = null;

        private IPoint sketchPlaneOrigin = null;
        private Vector sketchPlaneUDir = null;
        private Vector sketchPlaneVDir = null;


        #region ImplementICustomSlotMappingRule

        /// <summary>
        /// Gets the SectionAlias which emulate Standard Profile for the Plate/Plate or Plate/Stiffener combination 
        /// </summary>
        /// <param name="penetratingPart">The Penetrating Object.</param>
        /// <param name="penetratedPart">The Penetrated Object.</param>
        /// <param name="sectionAlias">The section alias.</param>
        /// <param name="web">A web of penetrating Object </param>
        /// <param name="flange">A Flange of penetrating Object </param>

        public void GetSectionAlias(object penetratingPart, object penetratedPart, out string sectionAlias, out object web, out object flange, out object secondWeb, out object secondFlange)
        {

            if (penetratingPart == null)
            {
                throw new ArgumentNullException("PenetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("PenetratedPart");
            }

            sectionAlias = "UnknownAlias";
            web = null;
            flange = null;
            secondWeb = null;
            secondFlange = null;

            PenetratingConfig penetratingStatus = PenetratingConfig.Unknown;
            Collection<BusinessObject> connectedPenetratingParts;

            CommonFuncs.GetSeamBtwObjects(penetratingPart, penetratedPart, out this.webSeamObj);

            if (webSeamObj == null)
            {
                return;
            }

            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            //If the penetrating object is a stiffener then check the stiffened plate to determine if it is the
            //web plate.  If the stiffened plate is not the web plate, then return UnknownAlias.  If the stiffened
            //plate is the web plate, then return 'FlangeObject' alias so the AC Rule knows not to create an AC.
            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            if (penetratingPart is StiffenerPart)
            {
                StiffenerSystem stiffenerLeafSystem = (StiffenerSystem)CommonFuncs.GetLeafSystemFromPart(penetratingPart);
                PlateSystem stiffenedPlateRootSystem = (PlateSystem)stiffenerLeafSystem.PlateToStiffen;

                if ((stiffenedPlateRootSystem.Type != PlateType.WebPlate) && (stiffenedPlateRootSystem.NamingCategory != (int)PlateNamingCategory.WebPlate)
                    && (GetPlateSubTypePropertyValue(stiffenedPlateRootSystem) != (int)PlateSubType.Web))
                {
                    //Stiffened Plate is not the Web Plate - Stiffener is Not Part of a Penetrating Combination
                    return;
                }
                else
                {
                    //Stiffened Plate is a Web Plate - Stiffener is the Flange
                    GetConnectedPenetratingParts(penetratingPart, penetratedPart, out connectedPenetratingParts);
                    if (connectedPenetratingParts != null)
                    {
                        if (connectedPenetratingParts.Count > 0)
                        {
                            web = connectedPenetratingParts[0];
                        }
                    }

                    flange = penetratingPart;
                    return;
                }
            }
            else if (penetratingPart is EdgeReinforcementPart)
            {
                EdgeReinforcementSystem ERRootSystem = (EdgeReinforcementSystem)CommonFuncs.GetRootSystemFromPart(penetratingPart);

                if (ERRootSystem.EdgeToReinforce.Connectable is PlateSystem)
                {
                    PlateSystem plateSystem = (PlateSystem)ERRootSystem.EdgeToReinforce.Connectable;
                    if ((plateSystem.Type != PlateType.WebPlate) && (plateSystem.NamingCategory != (int)PlateNamingCategory.WebPlate)
                        && (GetPlateSubTypePropertyValue(plateSystem) != (int)PlateSubType.Web))
                    {
                        //Plate is not the Web Plate - ER is Not Part of a Penetrating Combination
                        return;
                    }
                    else
                    {
                        //Plate is a Web Plate - ER is the Flange
                        GetConnectedPenetratingParts(penetratingPart, penetratedPart, out connectedPenetratingParts);
                        if (connectedPenetratingParts != null)
                        {
                            if (connectedPenetratingParts.Count > 0)
                            {
                                web = connectedPenetratingParts[0];
                            }
                        }

                        flange = penetratingPart;
                        return;
                    }
                }
                else
                {
                    return;
                }
            }
            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            //If the penetrating object is a plate then check if it is the web plate.  If it is not the web plate
            //then return 'FlangeObject' alias so the AC Rule knows not to create an AC.
            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            else if (penetratingPart is PlatePart)
            {   

                PlatePart PenetratingPlate = (PlatePart)penetratingPart;

                if ((PenetratingPlate.Type != PlateType.WebPlate) && (PenetratingPlate.NamingCategory != (int)PlateNamingCategory.WebPlate)
                     && (GetPlateSubTypePropertyValue(PenetratingPlate) != (int)PlateSubType.Web))
                {
                    if (PenetratingPlate.Type == PlateType.FlangePlate || (GetPlateSubTypePropertyValue(PenetratingPlate) == (int)PlateSubType.Flange))
                    {
                        //Plate is the Flange Plate
                        GetConnectedPenetratingParts(penetratingPart, penetratedPart, out connectedPenetratingParts);
                        if (connectedPenetratingParts != null)
                        {
                            if (connectedPenetratingParts.Count > 0)
                            {
                                web = connectedPenetratingParts[0];
                            }
                        }

                        flange = penetratingPart;

                        return;
                    }
                    else
                    {
                        //Plate is not the Web Plate - Plate is not the Flange Plate
                        return;
                    }
                }
                else if ((PenetratingPlate.Type == PlateType.WebPlate) || (PenetratingPlate.NamingCategory == (int)PlateNamingCategory.WebPlate)
                         || (GetPlateSubTypePropertyValue(PenetratingPlate) == (int)PlateSubType.Web))
                {
                    try
                    {
                        this.basePlatePort = GetBasePlatePort((PlatePartBase)penetratingPart, (IStructConnectable)penetratedPart);
                        if (this.basePlatePort == null)
                        {
                            sectionAlias = "UnknownAlias";
                            return;
                        }
                    }
                    catch
                    {
                        //Incase exception is observed during 'GetBasePlatePort' method call...
                        sectionAlias = "UnknownAlias";
                        return;
                    }
                }
            }

            // Only get here when penetrating part is WebPlate
            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            //Get the Connected Penetrating Flange and Check Configuration
            //Recognizes plate/ER, plate/Stiffener, and plate/plate combinations (two objects Combination)
            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            Collection<BusinessObject> connectedPenetratingFlangePartCol = null;

            GetConnectedPenetratingFlangePart(penetratingPart,
                                              penetratedPart,
                                              out penetratingStatus,
                                              out connectedPenetratingFlangePartCol);

            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            //Get the SectionAlias, Web, and Flange, Web and Flange Seams 
            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            if (penetratingStatus == PenetratingConfig.Unknown)
            {
                return;
            }
            else
            {
                if (connectedPenetratingFlangePartCol.Count == 0)
                {
                    web = penetratingPart;
                    CommonFuncs.GetSeamBtwObjects(penetratingPart, penetratedPart, out this.webSeamObj);
                }

                else if (connectedPenetratingFlangePartCol.Count == 1)
                {
                    web = penetratingPart;
                    CommonFuncs.GetSeamBtwObjects(web, penetratedPart, out this.webSeamObj);

                    flange = connectedPenetratingFlangePartCol[0];
                    CommonFuncs.GetSeamBtwObjects(flange, penetratedPart, out this.flangeSeamObj);
                }

                else
                {
                    //We don't consider this case any more 
                    return;
                }
            }

            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            //GetSectionAlias from Web and Flange (Seam Object) depends on PenetrationStatus 
            //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
           
            if (web != null)
            {
                if (flange == null)
                {
                    if (penetratingStatus == PenetratingConfig.Plate_Only)
                    {
                        sectionAlias = "FB";

                    }
                    else if (penetratingStatus == PenetratingConfig.BendPlate_Only)
                    {
                        sectionAlias = "UA";

                    }
                   
                }
                else
                {

                    GetEmulatedSectionAlias(web, flange, penetratedPart, penetratingStatus, out sectionAlias);
                }
            }

            // This member variables are used in GetSketchingPlane method 
            this.webObj = web;
            this.flangeObj = flange;
            this.secAlias = sectionAlias;
        }

        public int GetPlateSubTypePropertyValue(Plate plate)
        {
            return ((PropertyValueCodelist)plate.GetPropertyValue("IJPlateSubType", "SubType")).PropValue;
        }
        /// <summary>
        /// Get the Ports on the penetrating parts depends on SectionAlias
        /// If SectionAlias is UnknownAlias then, return empty Collection. 
        /// </summary>
        /// <param name="penetratingPart">The Penetrating Object.</param>
        /// <param name="penetratedPart">The Penetrated Object.</param>
        /// <param name="basePlatePort">The Port of BasePlate for Slot and Collar</param>
        /// <returns>The Ports on Penetrating parts. Ports that don't apply are omitted </returns>
        public Dictionary<int, IPort> GetEmulatedPorts(object penetratingPart, object penetratedPart, out IPort basePlatePort)
        {

            if (penetratingPart == null)
            {
                throw new ArgumentNullException("PenetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("PenetratedPart");
            }


            basePlatePort = null;
            Dictionary<int, IPort> mappedPorts = null;

            MemberPart oMbrPart = penetratedPart as MemberPart;
            if (oMbrPart != null)
            {
                PlatePart oPlatePart = penetratingPart as PlatePart;
                ICurve oMbrAxis = oMbrPart.Axis as ICurve;
                if (oPlatePart != null)
                {
                    //Member Insert Plate Case

                    TopologyPort MbrTopPort = null;
                    ReadOnlyCollection<TopologyPort> MbrPortsCollection = oMbrPart.GetPorts(TopologyGeometryType.Face, ContextTypes.Lateral, GeometryStage.Initial);
                    foreach (TopologyPort oMbrPort in MbrPortsCollection)
                    {
                        if (oMbrPort.SectionId == 514)
                        {
                            MbrTopPort = oMbrPort;
                            break;
                        }
                    }

                    IPort oPlateBasePort = oPlatePart.GetPort(TopologyGeometryType.Face, -1, -1, ContextTypes.Base, -1, false);
                    IPort oPlateOffsetPort = oPlatePart.GetPort(TopologyGeometryType.Face, -1, -1, ContextTypes.Offset, -1, false);
                    ISurface oplatesurface;
                    oplatesurface = (ISurface)oPlateBasePort;


                    // Calling this method to get insertPlate type
                    PenetrationType PenetrationType = CommonFuncs.GetInsertPlateType(penetratingPart, penetratedPart);

                    GeometryServices oGeomOpr = new GeometryServices();
                    Position MbrStPos = new Position();
                    Position MbrEndPos = new Position();
                    Vector BaseNormal = new Vector();
                    SurfaceScopeType ScopeType;

                    ISurface PlateBaseSurface = (ISurface)oPlateBasePort;
                    if (PlateBaseSurface != null)
                    {
                        PlateBaseSurface.ScopeNormal(out ScopeType, out BaseNormal);
                    }

                    oMbrAxis.EndPoints(out MbrStPos, out MbrEndPos);

                    Collection<ICurve> curves = new Collection<ICurve>();
                    curves.Add(oMbrAxis);
                    ComplexString3d MbrAxisCmplxStr = new ComplexString3d(curves);
                    Position PosOnSurface;
                    Double OffsetDistance = 0;

                    Matrix4X4 oMbrMatrix = oMbrPart.GetMatrixAtPosition(MbrStPos);
                    Vector MbrU_Vector = new Vector(oMbrMatrix.GetIndexValue(4), oMbrMatrix.GetIndexValue(5), oMbrMatrix.GetIndexValue(6));
                    Vector MbrV_Vector = new Vector(oMbrMatrix.GetIndexValue(8), oMbrMatrix.GetIndexValue(9), oMbrMatrix.GetIndexValue(10));

                    MbrU_Vector.Length = 1;
                    MbrV_Vector.Length = 1;
                    if (PenetrationType == PenetrationType.Mbrflngpenetration)
                    {
                        Position BasePlateCentroid = PlateBaseSurface.Centroid();
                        PosOnSurface = BasePlateCentroid;
                        OffsetDistance = oMbrAxis.ProjectPoint(BasePlateCentroid).DistanceToPoint(BasePlateCentroid);
                    }
                    else
                    {
                        CrossSectionServices MbrPartHlpr = new CrossSectionServices();
                        double uDistance = 0;
                        double vDistance = 0;
                        double Tolerance = 0.001; // Tolerance value is the distance between Plate and the member port. The value is 1mm 
                        double Midthicknessval;
                        if (PenetrationType == PenetrationType.TopPenetration)
                        {
                            MbrPartHlpr.GetCardinalPointDelta((ProfilePart)oMbrPart, oMbrPart.CardinalPoint, 8, out uDistance, out vDistance);
                            //Checking whether vDistance is positive or negative and accordingly adding the Tolearnce value
                            if (vDistance > 0 || ((vDistance).EqualTo(0) == true))
                            {
                                Midthicknessval = -(Tolerance);
                            }
                            else
                            {
                                Midthicknessval = (Tolerance);
                            }
                        }
                        else if (PenetrationType == PenetrationType.BtmPenetration)
                        {
                            MbrPartHlpr.GetCardinalPointDelta((ProfilePart)oMbrPart, oMbrPart.CardinalPoint, 2, out uDistance, out vDistance);
                            if (vDistance < 0 || (vDistance.EqualTo(0) == true))
                            {
                                Midthicknessval = (Tolerance);
                            }
                            else
                            {
                                Midthicknessval = -(Tolerance);
                            }
                        }
                        else
                        {
                            //These cases are not handled
                            return mappedPorts;
                        }

                        Vector uVector = new Vector(MbrU_Vector);
                        Vector vVector = new Vector(MbrV_Vector);
                        uVector.Length = -uDistance;
                        vVector.Length = vDistance + Midthicknessval; //The vDistance is reduced or increased by given Tolerance*10 value

                        Vector OffsetDirection = new Vector(uVector + vVector);

                        PosOnSurface = MbrStPos.Offset(OffsetDirection);

                        OffsetDistance = PosOnSurface.DistanceToPoint(MbrStPos);
                    }

                    ComplexString3d oProjAxis = oGeomOpr.GetCurveByOffset(MbrAxisCmplxStr, PosOnSurface, OffsetDistance);

                    ICurve ProjectedAxis = (ICurve)oProjAxis;

                    ReadOnlyCollection<TopologyPort> CollOfPorts = oPlatePart.GetPorts(TopologyGeometryType.Face, ContextTypes.Lateral, GeometryStage.Initial);

                    double mindist = 10000;
                    double maxdist = -1;

                    IPort oWebLeftPort = null;
                    IPort oWebRightPort = null;
                    IPort oTopPort = null;
                    IPort oBottomPort = null;

                    Position posStart;
                    Position posEnd;

                    ProjectedAxis.EndPoints(out posStart, out posEnd);

                    foreach (TopologyPort oPort in CollOfPorts)
                    {
                        Collection<BusinessObject> IntersectionCollection;
                        GeometryIntersectionType IntersectionType;

                        ISurface oPortSurface = (ISurface)oPort;
                        IPlane oPlane = (IPlane)oPort;

                        //filter out non planar ports
                        //in some cases non planar ports is also supporitng the Planar Interface
                        //but its root point is undefined based on that(as a kludge) filtering out the non planar cases.
                        if ((oPlane != null) && (oPlane.RootPoint != null))
                        {
                            oPortSurface.Intersect(ProjectedAxis, out IntersectionCollection, out IntersectionType);

                            if (IntersectionCollection != null)
                            {
                                if (IntersectionCollection.Count > 0)
                                {

                                    IPoint oPoint = IntersectionCollection[0] as IPoint;
                                    if (oPoint != null)
                                    {
                                        IPoint pointSt = new Point3d(posStart);
                                        double dist = oPoint.DistanceFromPoint(pointSt);
                                        if (dist > maxdist)
                                        {
                                            maxdist = dist;
                                            oWebRightPort = (IPort)oPort;
                                        }

                                        if (dist < mindist)
                                        {
                                            mindist = dist;
                                            oWebLeftPort = (IPort)oPort;
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if ((oWebLeftPort == null) || (oWebRightPort == null))
                    {
                        // WebLeft And WebRight ports not mapped yet
                        //return empty mapped Ports
                        mappedPorts = new Dictionary<int, IPort>();
                        return mappedPorts;
                    }

                    if (oWebLeftPort == oWebRightPort)
                    {
                        //Improper Mapping -- Mapped Ports cant be equal
                        //return empty mapped Ports
                        mappedPorts = new Dictionary<int, IPort>();
                        return mappedPorts;
                    }



                    if (PenetrationType == PenetrationType.Mbrflngpenetration)
                    {
                        TopologyPort MbrWLPort = null;
                        foreach (TopologyPort oMbrPort in MbrPortsCollection)
                        {
                            if (oMbrPort.SectionId == 257)
                            {
                                MbrWLPort = oMbrPort;
                                break;
                            }
                        }

                        ISurface MbrWebLeftSurface = (ISurface)MbrWLPort; // oMbrPart.GetPort(TopologyGeometryType.Face, 0, 257, ContextTypes.Lateral, 257, false);
                        Vector WebLeftNormal = new Vector();
                        SurfaceScopeType WebLeftScopeType;
                        if (MbrWebLeftSurface != null)
                            MbrWebLeftSurface.ScopeNormal(out WebLeftScopeType, out WebLeftNormal);

                        basePlatePort = (IPort)MbrWLPort;
                        if (WebLeftNormal.Dot(BaseNormal) >= 0)
                        {
                            oTopPort = oPlateBasePort;
                            oBottomPort = oPlateOffsetPort;
                        }
                        else
                        {
                            oTopPort = oPlateOffsetPort;
                            oBottomPort = oPlateBasePort;
                        }
                    }
                    else
                    {

                        TopologyPort MbrBtmPort = null;
                        foreach (TopologyPort oMbrPort in MbrPortsCollection)
                        {
                            if (oMbrPort.SectionId == 513)
                            {
                                MbrBtmPort = oMbrPort;
                                break;
                            }
                        }

                        if (MbrV_Vector.Dot(BaseNormal) >= 0)
                        {
                            if (PenetrationType == PenetrationType.TopPenetration)
                            {
                                //Base --> bottom , Offset --> Top
                                oTopPort = oPlateOffsetPort;
                                oBottomPort = oPlateBasePort;
                                basePlatePort = (IPort)MbrTopPort;
                            }
                            else if (PenetrationType == PenetrationType.BtmPenetration)
                            {
                                //Base --> Top, Offset --> Bottom
                                oTopPort = oPlateBasePort;
                                oBottomPort = oPlateOffsetPort;
                                basePlatePort = (IPort)MbrBtmPort;
                            }
                            else
                            {
                                // These cases are not handled
                                return mappedPorts;
                            }
                        }
                        else
                        {
                            if (PenetrationType == PenetrationType.TopPenetration)
                            {
                                //Base --> Top, Offset --> Bottom
                                oTopPort = oPlateBasePort;
                                oBottomPort = oPlateOffsetPort;
                                basePlatePort = (IPort)MbrTopPort;
                            }
                            else if (PenetrationType == PenetrationType.BtmPenetration)
                            {
                                //Base --> Bottom, Offset --> Top
                                oTopPort = oPlateOffsetPort;
                                oBottomPort = oPlateBasePort;
                                basePlatePort = (IPort)MbrBtmPort;
                            }
                            else
                            {
                                // These cases are not handled
                                return mappedPorts;
                            }
                        }
                    }


                    mappedPorts = new Dictionary<int, IPort>();
                    mappedPorts.Add((int)SectionFaceType.Top, oTopPort);
                    mappedPorts.Add((int)SectionFaceType.Bottom, oBottomPort);
                    mappedPorts.Add((int)SectionFaceType.Web_Left, oWebLeftPort);
                    mappedPorts.Add((int)SectionFaceType.Web_Right, oWebRightPort);

                    return mappedPorts;

                }

            }
            string sectionAlias = "UnknownAlias";
            object web = null;
            object flange = null;
            object secondWeb = null;
            object secondFlange = null;

            //1. Get SectionAlias
            if ((this.webObj == null) || (this.flangeObj == null) || (this.secAlias == ""))
            {
                //Need to Call GetSectionAlias to get the Section Alias Information
                GetSectionAlias(penetratingPart, penetratedPart, out sectionAlias, out web, out flange, out secondWeb, out secondFlange);
            }
            else
            {
                //Use the Member Variables
                sectionAlias = this.secAlias;
                web = this.webObj;
                flange = this.flangeObj;
            }

            if (sectionAlias == "UnknownAlias")
            {
                mappedPorts = new Dictionary<int, IPort>();
                return mappedPorts;

            }
            if (web != null)
            {
                this.basePlatePort = GetBasePlatePort((PlatePartBase)penetratingPart, (IStructConnectable)penetratedPart);

                if (this.basePlatePort != null)
                {

                    IEmulatedPortMap emulatedPortMap = null;
                    switch (sectionAlias)
                    {

                        case "FB":

                            emulatedPortMap = new SectionFBPortMap();
                            mappedPorts = emulatedPortMap.GetEmulatedPortsMap(penetratingPart, penetratedPart, this.basePlatePort, sectionAlias, web, flange, secondWeb, secondFlange, out basePlatePort);
                            break;

                        case "BUT":

                            emulatedPortMap = new SectionBUTPortMap();
                            mappedPorts = emulatedPortMap.GetEmulatedPortsMap(penetratingPart, penetratedPart, this.basePlatePort, sectionAlias, web, flange, secondWeb, secondFlange, out basePlatePort);
                            break;

                        case "BUTL2":

                            emulatedPortMap = new SectionBUTL2PortMap();
                            mappedPorts = emulatedPortMap.GetEmulatedPortsMap(penetratingPart, penetratedPart, this.basePlatePort, sectionAlias, web, flange, secondWeb, secondFlange, out basePlatePort);
                            break;

                        case "BUTL3":

                            emulatedPortMap = new SectionBUTL3PortMap();
                            mappedPorts = emulatedPortMap.GetEmulatedPortsMap(penetratingPart, penetratedPart, this.basePlatePort, sectionAlias, web, flange, secondWeb, secondFlange, out basePlatePort);
                            break;

                        case "UA":
                        case "EA": // EA,UA - Same Result 

                            emulatedPortMap = new SectionUAPortMap();
                            mappedPorts = emulatedPortMap.GetEmulatedPortsMap(penetratingPart, penetratedPart, this.basePlatePort, sectionAlias, web, flange, secondWeb, secondFlange, out basePlatePort);
                            break;

                        default:
                            break;

                    }
                }
            }
            //Set the member variables
            this.mappedPortCollection = mappedPorts;
            basePlatePort = this.basePlatePort;

            return mappedPorts;
        }

        /// <summary>
        /// Get the Sketching Plane to create Slot.
        /// 
        /// </summary>
        /// <param name="penetratingPart">The Penetrating Object.</param>
        /// <param name="penetratedPart">The Penetrated Object.</param>
        /// <param name="origin">Origin Point</param>
        /// <param name="uDirection">Origin Point</param>
        /// <param name="vDirection">Origin Point</param>  
        public void GetSketchingPlane(object penetratingPart, object penetratedPart, out IPoint origin, out Vector uDirection, out Vector vDirection)
        {

            // Checking inputs
            if (penetratingPart == null)
            {
                throw new ArgumentNullException("Input oPenetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("Input oPenetratedPart");
            }

            origin = null;
            uDirection = null;
            vDirection = null;

            // check if Web and Penetrated Object is detailed or not. 


            Dictionary<int, IPort> mappedPorts = null;
            IPort basePlatePort = null;

            if (penetratedPart is MemberPart)
            {
                MemberPart oMbrPart = penetratedPart as MemberPart;
                if (penetratingPart is PlatePart)
                {
                    //as of now this is a InsertPlate case
                    PlatePart oPlatePart = penetratingPart as PlatePart;

                    TopologyPort oTopPort = null;
                    TopologyPort oBtmPort = null;

                    ReadOnlyCollection<TopologyPort> MbrPortsCollection = oMbrPart.GetPorts(TopologyGeometryType.Face, ContextTypes.Lateral, GeometryStage.Initial);
                    foreach (TopologyPort oMbrPort in MbrPortsCollection)
                    {
                        if (oMbrPort.SectionId == (int)SectionFaceType.Top)
                        {
                            oTopPort = oMbrPort;
                            break;
                        }
                    }
                    foreach (TopologyPort oMbrPort in MbrPortsCollection)
                    {
                        if (oMbrPort.SectionId == (int)SectionFaceType.Bottom)
                        {
                            oBtmPort = oMbrPort;
                            break;
                        }
                    }

                    // Calling this method to get insertPlate type
                    PenetrationType PenetrationType = CommonFuncs.GetInsertPlateType(penetratingPart, penetratedPart);

                    Position MbrStPos = new Position();
                    Position MbrEndPos = new Position();
                    ICurve oMbrAxis = oMbrPart.Axis as ICurve;
                    oMbrAxis.EndPoints(out MbrStPos, out MbrEndPos);

                    Vector uVector = new Vector(MbrEndPos.X - MbrStPos.X, MbrEndPos.Y - MbrStPos.Y, MbrEndPos.Z - MbrStPos.Z);
                    Vector vVector = new Vector();

                    ISurface oPlateBaseSurface = (ISurface)oPlatePart.GetPort(TopologyGeometryType.Face, -1, -1, ContextTypes.Base, -1, false);

                    //Assumption Plate Base Surface is Planar
                    Position oCentrePos;
                    if (oPlateBaseSurface != null)
                    {
                        oCentrePos = oPlateBaseSurface.Centroid();

                        if (PenetrationType == PenetrationType.Mbrflngpenetration)
                        {
                            //Plate Penetrating Flange of a member i.e. Flange Plane is Sketch Plane
                            TopologyPort TopFlangePort = oTopPort;
                            IPlane oFlangePlane = TopFlangePort as IPlane;
                            if (oFlangePlane != null)
                            {

                                TopologyPort oWLPort = null;
                                foreach (TopologyPort oMbrPort in MbrPortsCollection)
                                {
                                    if (oMbrPort.SectionId == (int)SectionFaceType.Web_Left)
                                    {
                                        oWLPort = oMbrPort;
                                        break;
                                    }
                                }
                                ISurface WebLeftSurface = (ISurface)oWLPort;

                                SurfaceScopeType eScopeType;
                                WebLeftSurface.ScopeNormal(out eScopeType, out vVector);

                            }
                            else
                            {
                                //currently only planar surfaces are supported
                                //need to enhance this method for non planar surfaces.
                                return;
                            }
                        }
                        else
                        {
                            //Plate Penetrating Web of a member i.e. Web Plane is Sketch Plane
                            TopologyPort oWLPort = null;
                            foreach (TopologyPort oMbrPort in MbrPortsCollection)
                            {
                                if (oMbrPort.SectionId == (int)SectionFaceType.Web_Left)
                                {
                                    oWLPort = oMbrPort;
                                    break;
                                }
                            }

                            TopologyPort WebLeftPort = oWLPort; //oMbrPart.GetPort(TopologyGeometryType.Face, 0, 257, ContextTypes.Lateral, 257, false);
                            IPlane oWebPlane = WebLeftPort as IPlane;
                            if (oWebPlane != null)
                            {
                                ISurface TopPortSurface = (ISurface)oTopPort; //oMbrPart.GetPort(TopologyGeometryType.Face, 0, 514, ContextTypes.Lateral, 514, false);
                                SurfaceScopeType eScopeType;
                                TopPortSurface.ScopeNormal(out eScopeType, out vVector);
                                if (PenetrationType == PenetrationType.TopPenetration)
                                {
                                    vVector.Length = -1;
                                }
                            }
                            else
                            {
                                //currently only planar surfaces are supported
                                //need to enhance this method for non planar surfaces.
                                return;
                            }
                        }

                        vVector.Length = 1;
                        uVector.Length = 1;

                        uDirection = uVector;
                        vDirection = vVector;
                        origin = new Point3d(oCentrePos);

                        //For case where both top-and-bottom ports of member are penetrated, modify origin
                        if (PenetrationType == PenetrationType.Mbrflngpenetration)
                        {
                            //Adjust the origin of the sketching plane to be at mid-height of the bounded member 
                            Point3d PltCenter = new Point3d(oCentrePos);
                            double dMinDistance = 0;
                            Position PosOnTopPort = null;
                            ISurface TopPortSurf = (ISurface)oTopPort;
                            TopPortSurf.DistanceBetween(PltCenter, out dMinDistance, out PosOnTopPort);

                            Position PosOnBtmPort = null;
                            ISurface BtmPortSurf = (ISurface)oBtmPort;
                            BtmPortSurf.DistanceBetween(PltCenter, out dMinDistance, out PosOnBtmPort);

                            origin.X = (PosOnTopPort.X + PosOnBtmPort.X) / 2;
                            origin.Y = (PosOnTopPort.Y + PosOnBtmPort.Y) / 2;
                            origin.Z = (PosOnTopPort.Z + PosOnBtmPort.Z) / 2;
                        }

                        return;
                    }
                }
            }

            if ((this.mappedPortCollection == null) || (this.basePlatePort == null))
            {
                mappedPorts = GetEmulatedPorts(penetratingPart, penetratedPart, out basePlatePort);
            }
            else
            {
                mappedPorts = this.mappedPortCollection;
                basePlatePort = this.basePlatePort;
            }

            if (mappedPorts != null && mappedPorts.Count != 0)
            {

                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                //1. Calculate Origin Point 
                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                TopologyPort webLftPort = null;
                TopologyPort btmport = null;

                if (mappedPorts.ContainsKey((int)SectionFaceType.Web_Left))
                    webLftPort = (TopologyPort)mappedPorts[(int)SectionFaceType.Web_Left];

                if (mappedPorts.ContainsKey((int)SectionFaceType.Bottom))
                    btmport = (TopologyPort)mappedPorts[(int)SectionFaceType.Bottom];

                if (webLftPort == null || btmport == null)
                    return;

                ISurface webLeftPort = webLftPort;
                ISurface bottomPort = btmport;
                TopologySurface webLftSurface = (TopologySurface)webLftPort.Geometry;

                Collection<ICurve> colCurves;
                GeometryIntersectionType eCode;
                Position positionOnIntersecionCurve;
                Position positionOnPenetrated;
                double distance;
                bool firstPointFound = false;
                bool secondPointFound = false;
                ReadOnlyCollection<TopologyPort> penetratedPartPorts = null;
                Position firstPositionOnCrv = null;
                Position secondPositionOnCrv = null;
                //Get the intersection curve collection between web left port and web bottom port. Iterate the loop for tghe collection of curves. 
                //If the distance between the intersection curve to the penetrated base port is zero, then get the position on intersection curve. 
                //Same is done with the penetrated offset port also. Origin is the mid point of those two points on base and offset ports.
                webLftSurface.Intersect((ISurface)btmport.Geometry, out colCurves, out eCode);

                foreach (ICurve intersectionCurveBtwWebLeftAndBottom in colCurves)
                {

                    if (penetratedPart is PlatePart)
                    {
                        PlatePart penetratedPlatePart = (PlatePart)penetratedPart;
                        penetratedPartPorts = penetratedPlatePart.GetPorts(TopologyGeometryType.Face, ContextTypes.Base, GeometryStage.Initial);
                        if (penetratedPartPorts != null)
                        {
                            foreach (TopologyPort port in penetratedPartPorts)
                                if (port.OperationId > (int)0)
                                {
                                    port.DistanceBetween(intersectionCurveBtwWebLeftAndBottom, out distance, out positionOnPenetrated, out positionOnIntersecionCurve);
                                    if (positionOnIntersecionCurve != null)
                                    {
                                        if (distance < tolerance)
                                        {
                                            firstPositionOnCrv = positionOnIntersecionCurve;
                                            firstPointFound = true;
                                            break;
                                        }
                                        else if (distance < toleranceOneMilliMeter)
                                            firstPositionOnCrv = positionOnIntersecionCurve; //second approach
                                    }
                                }
                            penetratedPartPorts = penetratedPlatePart.GetPorts(TopologyGeometryType.Face, ContextTypes.Offset, GeometryStage.Initial);
                            if (penetratedPartPorts != null)
                            {
                                foreach (TopologyPort port in penetratedPartPorts)
                                    if (port.OperationId > (int)0)
                                    {
                                        port.DistanceBetween(intersectionCurveBtwWebLeftAndBottom, out distance, out positionOnPenetrated, out positionOnIntersecionCurve);
                                        if (positionOnIntersecionCurve != null)
                                        {
                                            if (distance < tolerance)
                                            {
                                                secondPositionOnCrv = positionOnIntersecionCurve;
                                                secondPointFound = true;
                                                break;
                                            }
                                            else if (distance < toleranceOneMilliMeter)
                                                secondPositionOnCrv = positionOnIntersecionCurve;
                                        }
                                    }
                            }
                        }
                        else if (penetratedPart is StiffenerPart)
                        {
                            StiffenerPart stiffenrPart = (StiffenerPart)penetratedPart;
                            penetratedPartPorts = stiffenrPart.GetPorts(TopologyGeometryType.Face, ContextTypes.Lateral, GeometryStage.Initial);
                            if (penetratedPartPorts != null)
                            {
                                foreach (TopologyPort port in penetratedPartPorts)
                                {
                                    if (port.SectionId == (int)SectionFaceType.Web_Left)
                                    {
                                        port.DistanceBetween(intersectionCurveBtwWebLeftAndBottom, out distance, out positionOnPenetrated, out positionOnIntersecionCurve);
                                        if (distance < this.tolerance)
                                            firstPositionOnCrv = positionOnIntersecionCurve;
                                    }
                                    else if (port.SectionId == (int)SectionFaceType.Web_Right)
                                    {
                                        port.DistanceBetween(intersectionCurveBtwWebLeftAndBottom, out distance, out positionOnPenetrated, out positionOnIntersecionCurve);
                                        if (distance < this.tolerance)
                                            secondPositionOnCrv = positionOnIntersecionCurve;
                                    }
                                    else if (firstPositionOnCrv != null && secondPositionOnCrv != null)
                                        break;
                                }
                            }
                        }
                        if (firstPositionOnCrv != null && secondPositionOnCrv != null)
                        {
                            if (firstPointFound == false && secondPointFound == false)
                            {
                                //Likely this is a curved deck case for which 'DistanceBetween' method returns less accurate distance value.
                                //try another approach
                                double distanceFirstPos = intersectionCurveBtwWebLeftAndBottom.GetDistanceAlongCurve(firstPositionOnCrv);
                                double distanceSecondPos = intersectionCurveBtwWebLeftAndBottom.GetDistanceAlongCurve(secondPositionOnCrv);
                                distance = (distanceFirstPos + distanceSecondPos) / 2;
                                Position originPos = intersectionCurveBtwWebLeftAndBottom.PointAtDistanceAlong(distance);
                                origin = new Point3d(originPos.X, originPos.Y, originPos.Z);
                            }
                        }
                    }

                    if (firstPointFound == true && secondPointFound == true && firstPositionOnCrv != null && secondPositionOnCrv != null)
                    {
                        origin = new Point3d(firstPositionOnCrv.X / 2 + secondPositionOnCrv.X / 2,
                                         firstPositionOnCrv.Y / 2 + secondPositionOnCrv.Y / 2,
                                         firstPositionOnCrv.Z / 2 + secondPositionOnCrv.Z / 2);
                        break;
                    }
                }

                TopologyPort basePortofPenetreatedPart = null;

                //If laternal port is connnected to Base Plate, it is BottomPort of Web

                if (penetratedPart is PlatePart)
                {
                    basePortofPenetreatedPart = (TopologyPort)CommonFuncs.GetBaseOrOffsetPortOfPlatePart(penetratedPart, false, penetratingPart);
                }
                else if (penetratedPart is StiffenerPart)
                {

                    StiffenerPart stiffenrPart = (StiffenerPart)penetratedPart;
                    ReadOnlyCollection<TopologyPort> sectionFacesCol = null;
                    sectionFacesCol = stiffenrPart.GetPorts(TopologyGeometryType.Face, GeometryStage.Initial);
                    if (sectionFacesCol != null)
                    {
                        foreach (TopologyPort port in sectionFacesCol)
                        {

                            if (port.SectionId == (int)SectionFaceType.Web_Right)
                            {
                                basePortofPenetreatedPart = (TopologyPort)port;
                            }
                            else
                            {

                            }
                        }
                    }
                }
                else
                {
                }


                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                //2. Calculate V Direction
                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                //Get the normal Vector of WebRight : Nwr
                TopologyPort topoPortOfWebRight = (TopologyPort)mappedPorts[(int)SectionFaceType.Web_Right];
                if (basePortofPenetreatedPart != null)
                {
                    Vector normalOfWebRight = CommonFuncs.GetNormalVectorOfTopologyPort(topoPortOfWebRight, basePortofPenetreatedPart);

                    //Get the normla Vector of Top : Nt
                    TopologyPort topoTopPort = (TopologyPort)mappedPorts[(int)SectionFaceType.Top];
                    Vector normalOfTopPort = CommonFuncs.GetNormalVectorOfTopologyPort(topoTopPort, basePortofPenetreatedPart);

                    //Get the Normal of Penetrated object : Np 
                    // Assume that : Origin point is on the neutral surface of Penetrated object.
                    // So Get the normal vector from the intersecting point between basePort Of Penetrated Object and web Left.
                    // it might be a tiny error according to the thickness of Penetrated object
                    TopologyPort topoPortOfWebLeft = (TopologyPort)webLeftPort;
                    Vector normalOfPenetratedObj = CommonFuncs.GetNormalVectorOfTopologyPort(basePortofPenetreatedPart, topoPortOfWebLeft); //oTopoPortOfWebLeft.Normal;

                    //V Vector = NPxNwr, negate if VxNt is negative 
                    Vector vVector = normalOfPenetratedObj.Cross(normalOfWebRight);

                    if (vVector.Dot(normalOfTopPort) < 0)
                    {
                        vVector.X = -vVector.X;
                        vVector.Y = -vVector.Y;
                        vVector.Z = -vVector.Z;
                    }
                    vDirection = vVector;


                    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    //3. Calculate U Direction
                    //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    Vector uVector = vDirection.Cross(normalOfPenetratedObj);

                    uDirection = uVector;

                    // Set the Member Variables
                    this.sketchPlaneOrigin = origin;
                    this.sketchPlaneUDir = uDirection;
                    this.sketchPlaneVDir = vDirection;
                }
            }
            else
            {

            }


        }

        /// <summary>
        /// Get the cross section dimensions of the penetrating section alias.
        /// 
        /// </summary>
        /// <param name="penetratingPart">The Penetrating Object.</param>
        /// <param name="penetratedPart">The Penetrated Object.</param>
        /// <param name="Depth">The Depth of the Penetrating Cross Section Alias</param>
        /// <param name="Width">The Width of the Penetrating Cross Section Alias</param>
        /// <param name="webThickness">The Web Thickness of the Penetrating Cross Section Alias</param>
        /// <param name="flangeThickness">The Flange Thickness of the Penetrating Cross Section Alias</param>
        public void GetSectionDimensions(object penetratingPart, object penetratedPart, out double Depth, out double Width, out double webThickness, out double flangeThickness)
        {
            Depth = 0;
            Width = 0;
            webThickness = 0;
            flangeThickness = 0;

            if (penetratingPart == null)
            {
                throw new ArgumentNullException("PenetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("PenetratedPart");
            }

            // Get the Section Alias
            string sectionAlias = "UnknownAlias";
            object web = null;
            object flange = null;
            object secondWeb = null;
            object secondFlange = null;

            if ((this.webObj == null) || (this.flangeObj == null) || (this.secAlias == ""))
            {
                //Need to Call GetSectionAlias to get the Section Alias Information
                GetSectionAlias(penetratingPart, penetratedPart, out sectionAlias, out web, out flange, out secondWeb, out secondFlange);
            }
            else
            {
                //Use the Member Variables
                sectionAlias = this.secAlias;
                web = this.webObj;
                flange = this.flangeObj;
            }

            // Get the Slot Mapping
            Dictionary<int, IPort> mappedPorts = null;
            IPort basePlate = null;
            if ((this.mappedPortCollection == null) || (this.basePlatePort == null))
            {
                mappedPorts = GetEmulatedPorts(penetratingPart, penetratedPart, out basePlate);
            }
            else
            {
                mappedPorts = this.mappedPortCollection;
                basePlate = this.basePlatePort;
            }
            TopologyPort basePlatePort = (TopologyPort)basePlate;

            // Get the sketching plane origin
            IPoint sketchingPlaneOrigin;
            Vector sketchingPlaneU;
            Vector sketchingPlaneV;
            if ((this.sketchPlaneOrigin == null) || (this.sketchPlaneUDir == null) || (this.sketchPlaneVDir == null))
            {
                GetSketchingPlane(penetratingPart, penetratedPart, out sketchingPlaneOrigin, out sketchingPlaneU, out sketchingPlaneV);
            }
            else
            {
                sketchingPlaneOrigin = this.sketchPlaneOrigin;
                sketchingPlaneU = this.sketchPlaneUDir;
                sketchingPlaneV = this.sketchPlaneVDir;
            }
            if (sketchingPlaneOrigin != null)
            {
                Position origin = sketchingPlaneOrigin.Position;

                // Section Alias Ports
                IPort top;
                IPort bottom;
                IPort webLeft;
                IPort webRight;
                IPort topFlangeLeft;
                IPort topFlangeRight;
                IPort topFlangeRightBottom;
                IPort topFlangeRightTop;

                // Section Alias Topology Ports
                TopologyPort topPort;
                TopologyPort bottomPort;
                TopologyPort webLeftPort;
                TopologyPort webRightPort;
                TopologyPort topFlangeLeftPort;
                TopologyPort topFlangeRightPort;
                TopologyPort topFlangeRightBottomPort;
                TopologyPort topFlangeRightTopPort;

                Position webLeftPoint;
                Position topFlangeLeftPoint;
                Position topFlangeRightPoint;

                Position body1Point;
                Position body2Point;

                // Get the Ports Common to All Cross Section Aliases
                mappedPorts.TryGetValue((int)SectionFaceType.Top, out top);
                mappedPorts.TryGetValue((int)SectionFaceType.Bottom, out bottom);
                mappedPorts.TryGetValue((int)SectionFaceType.Web_Left, out webLeft);
                mappedPorts.TryGetValue((int)SectionFaceType.Web_Right, out webRight);

                topPort = (TopologyPort)top;
                bottomPort = (TopologyPort)bottom;
                webLeftPort = (TopologyPort)webLeft;
                webRightPort = (TopologyPort)webRight;

                // Get the Cross Section Coordinate System
                Vector sectionW;
                Vector sectionU;
                Vector sectionV;
                if (basePlatePort == null)
                {
                    throw new ArgumentNullException("Input basePlatePort");
                }
                else
                {

                    GetCrossSectionCoordSys(penetratedPart, webRightPort, basePlatePort, out sectionW, out sectionU, out sectionV);

                    // Get the Web Center Plane and Cross Section Planes
                    Plane3d webCenterPlane = new Plane3d(origin, sectionU);
                    Plane3d crossSectionPlane = new Plane3d(origin, sectionW);

                    webRightPort.DistanceBetween((ISurface)webLeftPort, out webThickness, out body1Point, out body2Point);

                    //Get the seam between penetrated and penetrating ports and get the end points on that seam. 
                    PlateSystemBase penetratedPlateSystem = null;
                    PlateSystem penetratingPlateSystem = null;
                    ProfileSystem penetratedProfileSystem = null;

                    ISplit splitBetweenPenetartingAndPenetrated = null;
                    penetratingPlateSystem = (PlateSystem)CommonFuncs.GetRootSystemFromPart(penetratingPart);
                    if (penetratedPart is PlatePart)
                    {
                        penetratedPlateSystem = (PlateSystemBase)CommonFuncs.GetRootSystemFromPart(penetratedPart);
                        splitBetweenPenetartingAndPenetrated = penetratedPlateSystem.GetSplit((BusinessObject)penetratingPlateSystem);
                    }
                    else if (penetratedPart is ProfilePart)
                    {
                        penetratedProfileSystem = (ProfileSystem)CommonFuncs.GetRootSystemFromPart(penetratedPart);
                        splitBetweenPenetartingAndPenetrated = penetratedProfileSystem.GetSplit((BusinessObject)penetratingPlateSystem);
                    }

                    Seam seamBetweenPenetartingAndPenetrated = null;
                    if (splitBetweenPenetartingAndPenetrated != null)
                    {
                        seamBetweenPenetartingAndPenetrated = (Seam)splitBetweenPenetartingAndPenetrated;
                    }
                    else
                    {
                        throw new Exception("unable to get the split between penetarting and penetrated");

                    }
                    Position posStart = null;
                    Position posEnd = null;
                    if (seamBetweenPenetartingAndPenetrated != null)
                    {
                        seamBetweenPenetartingAndPenetrated.EndPoints(out posStart, out posEnd);
                    }
                    else
                    {
                        throw new Exception("unable to get the seam between penetarting and penetrated");

                    }
                    //Get closest point on the webcenter plane with respect to first(any) position on the seam. 
                    //Use this closest point to get a position on web bottom port which is bottom position. Use bottom position to get top position on penetrating top port. 
                    double minimumDistance = 0;
                    Position topPos = null;
                    Position btmPos = null;
                    Point3d startPosPoint3d = new Point3d(posStart.X, posStart.Y, posStart.Z);

                    Position startPosOnMidPlane = null;

                    webCenterPlane.DistanceBetween(startPosPoint3d, out minimumDistance, out startPosOnMidPlane);

                    Point3d startPtOnMidPlane = new Point3d(startPosOnMidPlane);

                    TopologySurface topSurface = (TopologySurface)topPort.Geometry;
                    TopologySurface bottomSurface = (TopologySurface)bottomPort.Geometry;

                    bottomSurface.DistanceBetween(startPtOnMidPlane, out minimumDistance, out btmPos);

                    Point3d btmPosPoint = new Point3d(btmPos);
                    topSurface.DistanceBetween(btmPosPoint, out minimumDistance, out topPos);

                    Point3d topPosPoint = new Point3d(topPos);
                    Position topPosOnMidPlane = null;
                    webCenterPlane.DistanceBetween(topPosPoint, out minimumDistance, out topPosOnMidPlane);
                    Depth = topPosOnMidPlane.DistanceToPoint(btmPos);

                    Vector widthVector;

                    switch (sectionAlias)
                    {
                        case "FB":

                            // Get the Flange Thickness
                            flangeThickness = 0;

                            // Get the Width
                            Width = webThickness;

                            break;

                        case "BUTL2":
                        case "BUT":

                            // Get the Mapped Ports
                            mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Left, out topFlangeLeft);
                            mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right, out topFlangeRight);
                            mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right_Bottom, out topFlangeRightBottom);

                            topFlangeLeftPort = (TopologyPort)topFlangeLeft;
                            topFlangeRightPort = (TopologyPort)topFlangeRight;
                            topFlangeRightBottomPort = (TopologyPort)topFlangeRightBottom;

                            // Get the Flange Thickness
                            topPort.DistanceBetween((ISurface)topFlangeRightBottomPort, out flangeThickness, out body1Point, out body2Point);

                            // Get the Width
                            topFlangeLeftPoint = topFlangeLeftPort.ProjectPoint(origin);
                            topFlangeRightPoint = topFlangeRightPort.ProjectPoint(origin);

                            widthVector = topFlangeRightPoint.Subtract(topFlangeLeftPoint);

                            Width = Math.Abs(widthVector.Dot(sectionU));

                            break;

                        case "BUTL3":

                            // Get the mapped ports
                            mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right, out topFlangeRight);
                            mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right_Bottom, out topFlangeRightBottom);
                            mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right_Top, out topFlangeRightTop);

                            topFlangeRightPort = (TopologyPort)topFlangeRight;
                            topFlangeRightBottomPort = (TopologyPort)topFlangeRightBottom;
                            topFlangeRightTopPort = (TopologyPort)topFlangeRightTop;

                            // Get the Flange Thickness
                            topFlangeRightTopPort.DistanceBetween((ISurface)topFlangeRightBottomPort, out flangeThickness, out body1Point, out body2Point);

                            // Get the Width
                            webLeftPoint = webLeftPort.ProjectPoint(origin);
                            topFlangeRightPoint = topFlangeRightPort.ProjectPoint(origin);

                            widthVector = topFlangeRightPoint.Subtract(webLeftPoint);

                            Width = Math.Abs(widthVector.Dot(sectionU));

                            break;
                        case "UA":
                        case "EA":

                            // Get the mapped ports
                            mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right, out topFlangeRight);
                            mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right_Bottom, out topFlangeRightBottom);

                            topFlangeRightPort = (TopologyPort)topFlangeRight;
                            topFlangeRightBottomPort = (TopologyPort)topFlangeRightBottom;

                            // Get the Flange Thickness
                            topPort.DistanceBetween((ISurface)topFlangeRightBottomPort, out flangeThickness, out body1Point, out body2Point);

                            // Get the Width
                            webLeftPoint = webLeftPort.ProjectPoint(origin);
                            topFlangeRightPoint = topFlangeRightPort.ProjectPoint(origin);

                            widthVector = topFlangeRightPoint.Subtract(webLeftPoint);

                            Width = Math.Abs(widthVector.Dot(sectionU));

                            break;

                        default:
                            break;
                    }

                }
            }
            else
            {
                throw new ArgumentNullException("Input sketchingPlaneOrigin");
            }
        }

        /// <summary>
        /// Returns the Depth of the penetrating cross section alias.
        /// 
        /// </summary>
        /// <param name="penetratingPart">The Penetrating Object.</param>
        /// <param name="penetratedPart">The Penetrated Object.</param>
        public double GetSectionDepth(object penetratingPart, object penetratedPart)
        {
            double Depth = 0;

            if (penetratingPart == null)
            {
                throw new ArgumentNullException("PenetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("PenetratedPart");
            }

            // Get the Slot Mapping
            Dictionary<int, IPort> mappedPorts = null;
            IPort basePlate = null;
            if ((this.mappedPortCollection == null) || (this.basePlatePort == null))
            {
                mappedPorts = GetEmulatedPorts(penetratingPart, penetratedPart, out basePlate);
            }
            else
            {
                mappedPorts = this.mappedPortCollection;
                basePlate = this.basePlatePort;
            }

            TopologyPort basePlatePort = (TopologyPort)basePlate;

            // Get the sketching plane origin
            IPoint sketchingPlaneOrigin;
            Vector sketchingPlaneU;
            Vector sketchingPlaneV;
            if ((this.sketchPlaneOrigin == null) || (this.sketchPlaneUDir == null) || (this.sketchPlaneVDir == null))
            {
                GetSketchingPlane(penetratingPart, penetratedPart, out sketchingPlaneOrigin, out sketchingPlaneU, out sketchingPlaneV);
            }
            else
            {
                sketchingPlaneOrigin = this.sketchPlaneOrigin;
                sketchingPlaneU = this.sketchPlaneUDir;
                sketchingPlaneV = this.sketchPlaneVDir;
            }
            PlateSystemBase penetratedPlateSystem = null;
            PlateSystem penetratingPlateSystem = null;
            ProfileSystem penetratedProfileSystem = null;

            //Get the seam between penetrated and penetrating ports and get the end points on that seam. 
            if (sketchingPlaneOrigin != null)
            {
                ISplit splitBetweenPenetartingAndPenetrated=null;
                penetratingPlateSystem = (PlateSystem)CommonFuncs.GetRootSystemFromPart(penetratingPart);
                if (penetratedPart is PlatePart)
                {
                    penetratedPlateSystem = (PlateSystemBase)CommonFuncs.GetRootSystemFromPart(penetratedPart);
                    splitBetweenPenetartingAndPenetrated = penetratedPlateSystem.GetSplit((BusinessObject)penetratingPlateSystem);
                }
                else if (penetratedPart is ProfilePart)
                {
                    penetratedProfileSystem = (ProfileSystem)CommonFuncs.GetRootSystemFromPart(penetratedPart);
                    splitBetweenPenetartingAndPenetrated = penetratedProfileSystem.GetSplit((BusinessObject)penetratingPlateSystem);
                }
                Position posStart = null;
                Position posEnd = null;
                Seam seamBetweenPenetartingAndPenetrated = null;
                if (splitBetweenPenetartingAndPenetrated != null)
                {
                    seamBetweenPenetartingAndPenetrated = (Seam)splitBetweenPenetartingAndPenetrated;

                    seamBetweenPenetartingAndPenetrated.EndPoints(out posStart, out posEnd);
                }
                else
                {
                    throw new Exception("Unable to get the seam between penetarting and penetrated");

                }


                double minimumDistance;
                Position topPos = null;
                Position btmPos = null;
                Point3d startPosPoint3d = new Point3d(posStart.X, posStart.Y, posStart.Z);

                Position origin = sketchingPlaneOrigin.Position;

                // Section Alias Ports
                IPort top;
                IPort bottom;
                IPort webRight;

                // Section Alias Topology Ports
                TopologyPort topPort;
                TopologyPort bottomPort;
                TopologyPort webRightPort;

                // Get the Top and Bottom Ports
                mappedPorts.TryGetValue((int)SectionFaceType.Top, out top);
                mappedPorts.TryGetValue((int)SectionFaceType.Bottom, out bottom);
                mappedPorts.TryGetValue((int)SectionFaceType.Web_Right, out webRight);

                topPort = (TopologyPort)top;
                bottomPort = (TopologyPort)bottom;
                webRightPort = (TopologyPort)webRight;

                // Get the Cross Section Coordinate System
                Vector sectionW;
                Vector sectionU;
                Vector sectionV;
                if (basePlatePort != null)
                {

                    GetCrossSectionCoordSys(penetratedPart, webRightPort, basePlatePort, out sectionW, out sectionU, out sectionV);

                    // Get the Web Center Plane and Cross Section Planes
                    Plane3d webCenterPlane = new Plane3d(origin, sectionU);
                    Plane3d crossSectionPlane = new Plane3d(origin, sectionW);
                    Position startPosOnMidPlane = null;

                    webCenterPlane.DistanceBetween(startPosPoint3d, out minimumDistance, out startPosOnMidPlane);
                    //Get closest point on the webcenter plane with respect to first(any) position on the seam. 
                    //Use this closest point to get a position on web bottom port which is bottom position. Use bottom position to get top position on penetrating top port. 
                    Point3d startPtOnMidPlane = new Point3d(startPosOnMidPlane);

                    TopologySurface topSurface = (TopologySurface)topPort.Geometry;
                    TopologySurface bottomSurface = (TopologySurface)bottomPort.Geometry;

                    bottomSurface.DistanceBetween(startPtOnMidPlane, out minimumDistance, out btmPos);

                    Point3d btmPosPoint = new Point3d(btmPos);
                    topSurface.DistanceBetween(btmPosPoint, out minimumDistance, out topPos);

                    Point3d topPosPoint = new Point3d(topPos);
                    Position topPosOnMidPlane = null;
                    webCenterPlane.DistanceBetween(topPosPoint, out minimumDistance, out topPosOnMidPlane);
                    Depth = topPosOnMidPlane.DistanceToPoint(btmPos);
                    return Depth;
                }

                else
                {
                    throw new ArgumentNullException("Input basePlatePort");
                }
            }
            else
            {
                throw new ArgumentNullException("Input sketchingPlaneOrigin");
            }
        }

        /// <summary>
        /// Returns the width of the penetrating cross section alias.
        /// 
        /// </summary>
        /// <param name="penetratingPart">The Penetrating Object.</param>
        /// <param name="penetratedPart">The Penetrated Object.</param>
        public double GetSectionWidth(object penetratingPart, object penetratedPart)
        {
            double Width = 0;

            if (penetratingPart == null)
            {
                throw new ArgumentNullException("PenetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("PenetratedPart");
            }

            // Get the Section Alias
            string sectionAlias = "UnknownAlias";
            object web = null;
            object flange = null;
            object secondWeb = null;
            object secondFlange = null;

            if ((this.webObj == null) || (this.flangeObj == null) || (this.secAlias == ""))
            {
                //Need to Call GetSectionAlias to get the Section Alias Information
                GetSectionAlias(penetratingPart, penetratedPart, out sectionAlias, out web, out flange, out secondWeb, out secondFlange);
            }
            else
            {
                //Use the Member Variables
                sectionAlias = this.secAlias;
                web = this.webObj;
                flange = this.flangeObj;
            }

            // Get the Slot Mapping
            Dictionary<int, IPort> mappedPorts = null;
            IPort basePlate = null;
            if ((this.mappedPortCollection == null) || (this.basePlatePort == null))
            {
                mappedPorts = GetEmulatedPorts(penetratingPart, penetratedPart, out basePlate);
            }
            else
            {
                mappedPorts = this.mappedPortCollection;
                basePlate = this.basePlatePort;
            }
            TopologyPort basePlatePort = (TopologyPort)basePlate;

            // Get the sketching plane origin
            IPoint sketchingPlaneOrigin;
            Vector sketchingPlaneU;
            Vector sketchingPlaneV;
            if ((this.sketchPlaneOrigin == null) || (this.sketchPlaneUDir == null) || (this.sketchPlaneVDir == null))
            {
                GetSketchingPlane(penetratingPart, penetratedPart, out sketchingPlaneOrigin, out sketchingPlaneU, out sketchingPlaneV);
            }
            else
            {
                sketchingPlaneOrigin = this.sketchPlaneOrigin;
                sketchingPlaneU = this.sketchPlaneUDir;
                sketchingPlaneV = this.sketchPlaneVDir;
            }
            if (sketchingPlaneOrigin != null)
            {
                Position origin = sketchingPlaneOrigin.Position;

                // Section Alias Ports
                IPort webLeft;
                IPort webRight;
                IPort topFlangeLeft;
                IPort topFlangeRight;

                // Section Alias Topology Ports
                TopologyPort webLeftPort;
                TopologyPort webRightPort;
                TopologyPort topFlangeLeftPort;
                TopologyPort topFlangeRightPort;

                Position webLeftPoint;
                Position topFlangeLeftPoint;
                Position topFlangeRightPoint;

                // Get the Ports Common to All Cross Section Aliases
                mappedPorts.TryGetValue((int)SectionFaceType.Web_Left, out webLeft);
                mappedPorts.TryGetValue((int)SectionFaceType.Web_Right, out webRight);

                webLeftPort = (TopologyPort)webLeft;
                webRightPort = (TopologyPort)webRight;

                // Get the Cross Section Coordinate System
                Vector sectionW;
                Vector sectionU;
                Vector sectionV;
                if (basePlatePort != null)
                {

                    GetCrossSectionCoordSys(penetratedPart, webRightPort, basePlatePort, out sectionW, out sectionU, out sectionV);

                    Vector widthVector;

                    switch (sectionAlias)
                    {
                        case "FB":

                            // Get the Width

                            Position body1Point;
                            Position body2Point;
                            // Get the Web Thickness as the minimum distance between the Web Left and Web Right Ports
                            webLeftPort.DistanceBetween((ISurface)webRightPort, out Width, out body1Point, out body2Point);

                            break;

                        case "BUTL2":
                        case "BUT":

                            // Get the Mapped Ports
                            mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Left, out topFlangeLeft);
                            mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right, out topFlangeRight);

                            topFlangeLeftPort = (TopologyPort)topFlangeLeft;
                            topFlangeRightPort = (TopologyPort)topFlangeRight;

                            // Get the Width
                            topFlangeLeftPoint = topFlangeLeftPort.ProjectPoint(origin);
                            topFlangeRightPoint = topFlangeRightPort.ProjectPoint(origin);

                            widthVector = topFlangeRightPoint.Subtract(topFlangeLeftPoint);

                            Width = Math.Abs(widthVector.Dot(sectionU));

                            break;

                        case "BUTL3":
                        case "UA":
                        case "EA":

                            // Get the mapped ports
                            mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right, out topFlangeRight);

                            topFlangeRightPort = (TopologyPort)topFlangeRight;

                            // Get the Width
                            webLeftPoint = webLeftPort.ProjectPoint(origin);
                            topFlangeRightPoint = topFlangeRightPort.ProjectPoint(origin);

                            widthVector = topFlangeRightPoint.Subtract(webLeftPoint);

                            Width = Math.Abs(widthVector.Dot(sectionU));

                            break;

                        default:
                            break;
                    }
                }
                else
                {
                    throw new ArgumentNullException("Input basePlatePort");
                }
            }
            else
            {
                throw new ArgumentNullException("Input sketchingPlaneOrigin");
            }
            return Width;
        }

        /// <summary>
        /// Returns the width of the penetrating cross section alias.
        /// 
        /// </summary>
        /// <param name="penetratingPart">The Penetrating Object.</param>
        /// <param name="penetratedPart">The Penetrated Object.</param>
        public double GetSectionWebThickness(object penetratingPart, object penetratedPart)
        {
            double webThickness = 0;

            if (penetratingPart == null)
            {
                throw new ArgumentNullException("PenetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("PenetratedPart");
            }

            // Get the Slot Mapping
            Dictionary<int, IPort> mappedPorts = null;
            IPort basePlate = null;
            if ((this.mappedPortCollection == null) || (this.basePlatePort == null))
            {
                mappedPorts = GetEmulatedPorts(penetratingPart, penetratedPart, out basePlate);
            }
            else
            {
                mappedPorts = this.mappedPortCollection;
                basePlate = this.basePlatePort;
            }
            TopologyPort basePlatePort = (TopologyPort)basePlate;

            // Section Alias Ports
            IPort webRight;
            IPort webLeft;

            // Section Alias Topology Ports
            TopologyPort webRightPort;
            TopologyPort webLeftPort;

            Position webLeftPoint = null;
            Position webRightPoint = null;

            // Get the Web Left and Web Right Ports
            mappedPorts.TryGetValue((int)SectionFaceType.Web_Right, out webRight);
            mappedPorts.TryGetValue((int)SectionFaceType.Web_Left, out webLeft);

            webRightPort = (TopologyPort)webRight;
            webLeftPort = (TopologyPort)webLeft;

            // Get the Web Thickness as the minimum distance between the Web Left and Web Right Ports
            webLeftPort.DistanceBetween((ISurface)webRightPort, out webThickness, out webLeftPoint, out webRightPoint);

            return webThickness;
        }

        /// <summary>
        /// Returns the flange thickness of the penetrating cross section alias.
        /// 
        /// </summary>
        /// <param name="penetratingPart">The Penetrating Object.</param>
        /// <param name="penetratedPart">The Penetrated Object.</param>
        public double GetSectionFlangeThickness(object penetratingPart, object penetratedPart)
        {
            double flangeThickness = 0;

            if (penetratingPart == null)
            {
                throw new ArgumentNullException("PenetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("PenetratedPart");
            }

            // Get the Section Alias
            string sectionAlias = "UnknownAlias";
            object web = null;
            object flange = null;
            object secondWeb = null;
            object secondFlange = null;

            if ((this.webObj == null) || (this.flangeObj == null) || (this.secAlias == ""))
            {
                //Need to Call GetSectionAlias to get the Section Alias Information
                GetSectionAlias(penetratingPart, penetratedPart, out sectionAlias, out web, out flange, out secondWeb, out secondFlange);
            }
            else
            {
                //Use the Member Variables
                sectionAlias = this.secAlias;
                web = this.webObj;
                flange = this.flangeObj;
            }

            // Get the Slot Mapping
            Dictionary<int, IPort> mappedPorts = null;
            IPort basePlate = null;
            if ((this.mappedPortCollection == null) || (this.basePlatePort == null))
            {
                mappedPorts = GetEmulatedPorts(penetratingPart, penetratedPart, out basePlate);
            }
            else
            {
                mappedPorts = this.mappedPortCollection;
                basePlate = this.basePlatePort;
            }
            TopologyPort basePlatePort = (TopologyPort)basePlate;

            // Section Alias Ports
            IPort top;
            IPort topFlangeRightBottom;
            IPort topFlangeRightTop;

            // Section Alias Topology Ports
            TopologyPort topPort;
            TopologyPort topFlangeRightBottomPort;
            TopologyPort topFlangeRightTopPort;

            Position topPoint;
            Position topFlangeRightBottomPoint;
            Position topFlangeRightTopPoint;


            switch (sectionAlias)
            {
                case "FB":

                    // Get the Flange Thickness
                    flangeThickness = 0;

                    break;

                case "BUTL2":
                case "BUT":

                    // Get the Mapped Ports
                    mappedPorts.TryGetValue((int)SectionFaceType.Top, out top);
                    mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right_Bottom, out topFlangeRightBottom);

                    topPort = (TopologyPort)top;
                    topFlangeRightBottomPort = (TopologyPort)topFlangeRightBottom;

                    // Get the Flange Thickness
                    topPort.DistanceBetween((ISurface)topFlangeRightBottomPort, out flangeThickness, out topPoint, out topFlangeRightBottomPoint);

                    break;

                case "BUTL3":

                    // Get the mapped ports
                    mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right_Bottom, out topFlangeRightBottom);
                    mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right_Top, out topFlangeRightTop);

                    topFlangeRightBottomPort = (TopologyPort)topFlangeRightBottom;
                    topFlangeRightTopPort = (TopologyPort)topFlangeRightTop;

                    // Get the Flange Thickness
                    topFlangeRightBottomPort.DistanceBetween((ISurface)topFlangeRightTopPort, out flangeThickness, out topFlangeRightTopPoint, out topFlangeRightBottomPoint);

                    break;
                case "UA":
                case "EA":

                    // Get the mapped ports
                    mappedPorts.TryGetValue((int)SectionFaceType.Top, out top);
                    mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right_Bottom, out topFlangeRightBottom);

                    topPort = (TopologyPort)top;
                    topFlangeRightBottomPort = (TopologyPort)topFlangeRightBottom;

                    topPort.DistanceBetween((ISurface)topFlangeRightBottomPort, out flangeThickness, out topPoint, out topFlangeRightBottomPoint);

                    break;

                default:
                    break;
            }
            return flangeThickness;
        }

        /// <summary>
        /// Returns the flange offset of the penetrating cross section alias.
        /// 
        /// </summary>
        /// <param name="penetratingPart">The Penetrating Object.</param>
        /// <param name="penetratedPart">The Penetrated Object.</param>
        public double GetSectionFlangeOffset(object penetratingPart, object penetratedPart)
        {
            double FlangeOffset = 0;

            if (penetratingPart == null)
            {
                throw new ArgumentNullException("PenetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("PenetratedPart");
            }

            // Get the Section Alias
            string sectionAlias = "UnknownAlias";
            object web = null;
            object flange = null;
            object secondWeb = null;
            object secondFlange = null;

            if ((this.webObj == null) || (this.flangeObj == null) || (this.secAlias == ""))
            {
                //Need to Call GetSectionAlias to get the Section Alias Information
                GetSectionAlias(penetratingPart, penetratedPart, out sectionAlias, out web, out flange, out secondWeb, out secondFlange);
            }
            else
            {
                //Use the Member Variables
                sectionAlias = this.secAlias;
                web = this.webObj;
                flange = this.flangeObj;
            }

            // Get the Slot Mapping
            Dictionary<int, IPort> mappedPorts = null;
            IPort basePlate = null;
            if ((this.mappedPortCollection == null) || (this.basePlatePort == null))
            {
                mappedPorts = GetEmulatedPorts(penetratingPart, penetratedPart, out basePlate);
            }
            else
            {
                mappedPorts = this.mappedPortCollection;
                basePlate = this.basePlatePort;
            }
            TopologyPort basePlatePort = (TopologyPort)basePlate;

            // Get the sketching plane origin
            IPoint sketchingPlaneOrigin;
            Vector sketchingPlaneU;
            Vector sketchingPlaneV;
            if ((this.sketchPlaneOrigin == null) || (this.sketchPlaneUDir == null) || (this.sketchPlaneVDir == null))
            {
                GetSketchingPlane(penetratingPart, penetratedPart, out sketchingPlaneOrigin, out sketchingPlaneU, out sketchingPlaneV);
            }
            else
            {
                sketchingPlaneOrigin = this.sketchPlaneOrigin;
                sketchingPlaneU = this.sketchPlaneUDir;
                sketchingPlaneV = this.sketchPlaneVDir;
            }
            if (sketchingPlaneOrigin != null)
            {
                Position origin = sketchingPlaneOrigin.Position;

                // Section Alias Ports
                IPort top;
                IPort webLeft;
                IPort webRight;
                IPort topFlangeLeft;
                IPort topFlangeRightTop;

                // Section Alias Topology Ports
                TopologyPort topPort;
                TopologyPort webLeftPort;
                TopologyPort webRightPort;
                TopologyPort topFlangeLeftPort;
                TopologyPort topFlangeRightTopPort;

                Position topPoint;
                Position webLeftPoint;
                Position topFlangeLeftPoint;
                Position topFlangeRightTopPoint;

                // Get the Ports Common to All Cross Section Aliases
                mappedPorts.TryGetValue((int)SectionFaceType.Web_Right, out webRight);

                webRightPort = (TopologyPort)webRight;

                // Get the Cross Section Coordinate System
                Vector sectionW;
                Vector sectionU;
                Vector sectionV;
                if (basePlatePort != null)
                {
                    GetCrossSectionCoordSys(penetratedPart, webRightPort, basePlatePort, out sectionW, out sectionU, out sectionV);

                    Vector offsetVector;

                    switch (sectionAlias)
                    {
                        case "FB":
                        case "UA":
                        case "EA":

                            FlangeOffset = 0;
                            break;

                        case "BUTL2":
                        case "BUT":

                            // Get the Mapped Ports
                            mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Left, out topFlangeLeft);
                            mappedPorts.TryGetValue((int)SectionFaceType.Web_Left, out webLeft);

                            topFlangeLeftPort = (TopologyPort)topFlangeLeft;
                            webLeftPort = (TopologyPort)webLeft;

                            // Get the Flange Offset
                            topFlangeLeftPoint = topFlangeLeftPort.ProjectPoint(origin);
                            webLeftPoint = webLeftPort.ProjectPoint(origin);

                            offsetVector = webLeftPoint.Subtract(topFlangeLeftPoint);

                            FlangeOffset = Math.Abs(offsetVector.Dot(sectionU));

                            break;

                        case "BUTL3":

                            // Get the Mapped Ports
                            mappedPorts.TryGetValue((int)SectionFaceType.Top, out top);
                            mappedPorts.TryGetValue((int)SectionFaceType.Top_Flange_Right_Top, out topFlangeRightTop);

                            topPort = (TopologyPort)top;
                            topFlangeRightTopPort = (TopologyPort)topFlangeRightTop;

                            // Get the Flange Offset
                            topPoint = topPort.ProjectPoint(origin);
                            topFlangeRightTopPoint = topFlangeRightTopPort.ProjectPoint(origin);

                            offsetVector = topPoint.Subtract(topFlangeRightTopPoint);

                            FlangeOffset = Math.Abs(offsetVector.Dot(sectionV));

                            break;

                        default:
                            break;
                    }
                }
                else
                {
                    throw new ArgumentNullException("Input basePlatePort ");
                }
            }
            else
            {
                throw new ArgumentNullException("Input sketchingPlaneOrigin ");
            }
            return FlangeOffset;
        }


        #endregion


        #region Private Methods

        /// <summary>
        /// GetEmulatedSectionAlias returns sectionAlias according to Penetrating Configuration
        /// If Plate-Plate Combination 
        /// 1. Get the intersection point between Web Seam and Flange Seam 
        /// 2. Get Start and End Point of Flange Seam 
        /// 3. Measure distances between Start point, End Point  and intersection point 
        /// 4. Get the ratio which determins BUT, BUT2, UA : Web thickness / FlangeWidth  --> this is customizable. User can use other measurement value 
        /// 5. Get the MeasurementVlaue = Abs(DistanceFormEndPos - DistnaceFromStartPos) / FlangeWidth 
        /// 
        /// If It is Tee  - Web is bounded by Flange 
        ///     if (Math.Abs(dMeasurementValue - dRatio) less than tolerance) --> BUT
        ///     else if dMeasurementValue  less than 1- dRatio) --> BUTL2
        ///     else --> UA 
        /// if It is not Tee 
        ///      if (Math.Abs(oFlangePlateSystem.ThicknessOffset) > 0) --> BUTL3
        ///      else
        ///         if dMeasurementValue less than 1- dRatio  --> BUTL3
        ///         else --> UA 
        /// </summary>
        /// <param name="web">Input WebPart.</param>
        /// <param name="flange">Input FlangePart.</param>
        /// <param name="penetratedObject">Input PenetratedObject .</param>
        /// <param name="config">Input PenetratingConfiguration .</param>
        /// <param name="sectionAlias"> Output - SectionAlias- FB, BUT, BUTL2, BUTL3, UA (In case of EA, it returns UA) .</param>
        private void GetEmulatedSectionAlias(object web, object flange, object penetratedObject, PenetratingConfig config, out string sectionAlias)
        {

            if (web == null)
            {
                throw new ArgumentNullException("Input oWeb");
            }
            if (flange == null)
            {
                throw new ArgumentNullException("Input flange");
            }

            sectionAlias = "UnknownAlias";

            ICurve webSeamCurve = (ICurve)this.webSeamObj;

            if (webSeamCurve==null)
            {
                throw new Exception("WebSeamCurve is null");
            }

            ICurve flangeSeamCurve = (ICurve)this.flangeSeamObj;

            Position startPos = null;
            Position endPos = null;

            if (flangeSeamCurve != null)
            {
                flangeSeamCurve.EndPoints(out startPos, out endPos);
            }

            Double distnaceFromStartPos = 0;
            Double distanceFormEndPos = 0;
            GeometryIntersectionType interSectionType = GeometryIntersectionType.Unknown;
            PlateSystem webPlateSystem = (PlateSystem)CommonFuncs.GetRootSystemFromPart(web);

            switch (config)
            {
                case PenetratingConfig.Plate_Only:

                    sectionAlias = "FB";
                    break;
                case PenetratingConfig.Plate_Plate:

                    PlateSystem flangePlateSystem = null;

                    flangePlateSystem = (PlateSystem)CommonFuncs.GetRootSystemFromPart(flange);



                    //If there is mutual bounding case, skip it. 
                    if (CommonFuncs.IsMutualBounding(webPlateSystem, flangePlateSystem) == true)
                    {
                        break;
                    }

                    //if Web is bounded by Flange. --> It can be BUT, BUTL2, or UA
                    bool isTee = false;
                    Collection<Object> webBoundaryCol = webPlateSystem.Boundaries;
                    if (webBoundaryCol != null)
                    {
                        if (webBoundaryCol.Contains((BusinessObject)flangePlateSystem))
                        {
                            isTee = true;
                        }
                    }

                    PlatePart webPlate = (PlatePart)web;
                    PlatePart flangePlate = (PlatePart)flange;

                    // Get the Base & Offset ports of the web plate
                    TopologyPort BasePort;
                    Plane3d BasePlane;
                    TopologyPort OffsetPort;
                    Plane3d OffsetPlane;
                    Position pointOnPort;
                    Vector normalOnPort;
                    Position pointOnPlane;
                    Point3d point;
                    ReadOnlyCollection<TopologyPort> PortColl = null;
                    double startToBase, startToOffset, endToBase, endToOffset;
                    double startOffset, endOffset;

                    if (isTee == true)
                    {
                        PortColl = webPlate.GetPorts(TopologyGeometryType.Face, ContextTypes.Base, GeometryStage.Initial);
                        BasePort = PortColl[0];
                        PortColl = webPlate.GetPorts(TopologyGeometryType.Face, ContextTypes.Offset, GeometryStage.Initial);
                        OffsetPort = PortColl[0];

                        if (startPos != null && endPos != null)
                        {
                            pointOnPort = BasePort.ProjectPoint(startPos);
                            normalOnPort = BasePort.OutwardNormalAtPoint(pointOnPort);
                            BasePlane = new Plane3d(pointOnPort, normalOnPort);

                            pointOnPort = OffsetPort.ProjectPoint(startPos);
                            normalOnPort = OffsetPort.OutwardNormalAtPoint(pointOnPort);
                            OffsetPlane = new Plane3d(pointOnPort, normalOnPort);

                            point = new Point3d(startPos.X, startPos.Y, startPos.Z);
                            BasePlane.DistanceBetween(point, out startToBase, out pointOnPlane);
                            OffsetPlane.DistanceBetween(point, out startToOffset, out pointOnPlane);

                            point = new Point3d(endPos.X, endPos.Y, endPos.Z);
                            BasePlane.DistanceBetween(point, out endToBase, out pointOnPlane);
                            OffsetPlane.DistanceBetween(point, out endToOffset, out pointOnPlane);

                            if (startToBase < startToOffset)
                            {
                                startOffset = startToBase;
                            }
                            else
                            {
                                startOffset = startToOffset;
                            }

                            if (endToBase < endToOffset)
                            {
                                endOffset = endToBase;
                            }
                            else
                            {
                                endOffset = endToOffset;
                            }

                            if ((startOffset <= this.tolerance) || (endOffset <= this.tolerance))
                            {
                                sectionAlias = "UA";
                            }
                            // For Now - Assume if either is less then or equal to 15mm then assumes its a BUTL2
                            else if ((startToOffset <= 0.015 + this.tolerance) || (endOffset <= 0.015 + this.tolerance))
                            {
                                sectionAlias = "BUTL2";
                            }
                            else
                            {
                                sectionAlias = "BUT";
                            }

                        }
                        else
                        {
                            throw new ArgumentNullException("startPos and endPos are not valid");
                        }
                    }
                    else
                    {
                        webSeamCurve.EndPoints(out startPos, out endPos);

                        PortColl = flangePlate.GetPorts(TopologyGeometryType.Face, ContextTypes.Base, GeometryStage.Initial);
                        BasePort = PortColl[0];
                        PortColl = flangePlate.GetPorts(TopologyGeometryType.Face, ContextTypes.Offset, GeometryStage.Initial);
                        OffsetPort = PortColl[0];

                        pointOnPort = BasePort.ProjectPoint(startPos);
                        normalOnPort = BasePort.OutwardNormalAtPoint(pointOnPort);
                        BasePlane = new Plane3d(pointOnPort, normalOnPort);

                        pointOnPort = OffsetPort.ProjectPoint(startPos);
                        normalOnPort = OffsetPort.OutwardNormalAtPoint(pointOnPort);
                        OffsetPlane = new Plane3d(pointOnPort, normalOnPort);

                        point = new Point3d(startPos.X, startPos.Y, startPos.Z);
                        BasePlane.DistanceBetween(point, out startToBase, out pointOnPlane);
                        OffsetPlane.DistanceBetween(point, out startToOffset, out pointOnPlane);

                        point = new Point3d(endPos.X, endPos.Y, endPos.Z);
                        BasePlane.DistanceBetween(point, out endToBase, out pointOnPlane);
                        OffsetPlane.DistanceBetween(point, out endToOffset, out pointOnPlane);

                        if (startToBase < startToOffset)
                        {
                            startOffset = startToBase;
                        }
                        else
                        {
                            startOffset = startToOffset;
                        }

                        if (endToBase < endToOffset)
                        {
                            endOffset = endToBase;
                        }
                        else
                        {
                            endOffset = endToOffset;
                        }

                        if ((startOffset <= this.tolerance) || (endOffset <= this.tolerance))
                        {
                            sectionAlias = "UA";
                        }
                        else
                        {
                            sectionAlias = "BUTL3";
                        }
                    }
                    break;

                case PenetratingConfig.Plate_ER:


                    if (flange is EdgeReinforcementPart)
                    {
                        EdgeReinforcementPart erFlange = (EdgeReinforcementPart)flange;

                        //Get the WebLength 
                        CrossSection xSection = erFlange.CrossSection;
                        double flangeWebLength = 0;
                        if (xSection != null)
                        {
                            PropertyValueDouble propWebLength = (PropertyValueDouble)xSection.GetPropertyValue("IJUAXSectionWeb", "WebLength");

                            if (propWebLength != null)
                            {
                                flangeWebLength = (double)propWebLength.PropValue;
                            }
                        }
                        else
                        {
                            throw new NotImplementedException();
                        }

                        //Face 
                        if (erFlange.EdgeReinforcementPosition == EdgeReinforcementPosition.OnFace)
                        {
                            //There is offset 
                            if (Math.Abs((double)erFlange.OffsetValue) > this.tolerance)
                            {
                                sectionAlias = "BUTL3";
                            }
                            else
                            {
                                sectionAlias = "UA";
                            }

                        }
                        //Edge 
                        else if (erFlange.EdgeReinforcementPosition == EdgeReinforcementPosition.OnEdgeOffset)
                        {

                            //   LoadPoint.WebLef = Centerend
                            //             
                            //   -----------------
                            //   |               |
                            //   --------*--------
                            //
                            PlatePart plateWeb = (PlatePart)web;

                            TopologyPort erWebRightPort = null;
                            TopologyPort erTopPort = null;
                            TopologyPort erBottomPort = null;
                            ReadOnlyCollection<TopologyPort> profilePortCol = null;
                            profilePortCol = erFlange.GetPorts(TopologyGeometryType.Face, GeometryStage.Initial);

                            // Get web Right port on the Profile (Flange)
                            foreach (TopologyPort port in profilePortCol)
                            {
                                if (port.SectionId == (int)SectionFaceType.Web_Right)
                                {
                                    erWebRightPort = port;
                                }
                                else if (port.SectionId == (int)SectionFaceType.Top)
                                {
                                    erTopPort = port;
                                }
                                else if (port.SectionId == (int)SectionFaceType.Bottom)
                                {
                                    erBottomPort = port;
                                }
                            }

                            if (erWebRightPort == null && erTopPort == null && erBottomPort == null)
                            {
                                throw new Exception("Check whether erWebRightPort or erTopPort or erBottomPort is null");
                            }

                            // Get the Base port of the Plate (Web)
                            TopologyPort plateBasePort;
                            TopologyPort plateOffsetPort;
                            ReadOnlyCollection<TopologyPort> platePortCol = null;
                            platePortCol = plateWeb.GetPorts(TopologyGeometryType.Face, ContextTypes.Base, GeometryStage.Initial);
                            plateBasePort = platePortCol[0];
                            platePortCol = plateWeb.GetPorts(TopologyGeometryType.Face, ContextTypes.Offset, GeometryStage.Initial);
                            plateOffsetPort = platePortCol[0];

                            // Get the plate base port and the edge reinforcement web right port normals
                            Vector webNormal = plateBasePort.Normal;

                            // Use Web Normal and Flange Normal to get the cross section plane normal
                            Vector XSectionNormal = new Vector();
                            IStructConnectable structConnectable = (IStructConnectable)penetratedObject;
                            ReadOnlyCollection<TopologyPort> penetratedPortCol = null;
                            TopologyPort penetratedPort = null;

                            Position positionOnPenetrated;
                            double distance = 0;

                            if (penetratedObject is PlatePart)
                            {
                                penetratedPortCol = structConnectable.GetPorts(TopologyGeometryType.Face, ContextTypes.Base, GeometryStage.Initial);
                                penetratedPort = penetratedPortCol[0];

                                Position positionOnERWebRight = new Position();

                                //Flange seam position to be projected onto the ER web right port to find normal
                                Vector flangeNormal = erWebRightPort.OutwardNormalAtPoint(startPos);
                                XSectionNormal = webNormal.Cross(flangeNormal);

                            }
                            else if ((penetratedObject is ProfilePart) || (penetratedObject is MemberPart))
                            {
                                TopologyPort webRight = null;
                                TopologyPort top = null;
                                penetratedPortCol = structConnectable.GetPorts(TopologyGeometryType.Face, ContextTypes.Lateral, GeometryStage.Initial);

                                foreach (TopologyPort port in penetratedPortCol)
                                {
                                    if (port.SectionId == (int)SectionFaceType.Web_Right)
                                    {
                                        webRight = port;
                                    }
                                    else if (port.SectionId == (int)SectionFaceType.Top)
                                    {
                                        top = port;
                                    }
                                }

                            }

                            // Intersect the Plate Base Port with the Penetrated Port
                            Position positionOnplateBasePort;
                            plateBasePort.DistanceBetween((ISurface)penetratedPort, out distance, out positionOnplateBasePort, out positionOnPenetrated);

                            // Create the approximate cross section plane
                            Plane3d XSectionPlane = null;
                            if (distance < this.tolerance)
                                XSectionPlane = new Plane3d(positionOnplateBasePort, XSectionNormal);

                            // Get a Vertex on the Intersection and use as an approximate point for the cross section plane
                            // Intersect the Penetrating Ports with the Cross Section Plane
                            ICurve plateBaseWire = null;
                            ICurve plateOffsetWire = null;
                            ICurve erTopWire = null;
                            ICurve erBottomWire = null;
                            Collection<ICurve> wireCol;
                            GeometryIntersectionType IntersectionType;
                            if (XSectionPlane != null)
                            {
                                XSectionPlane.Intersect((ISurface)plateBasePort, out wireCol, out IntersectionType);
                                plateBaseWire = wireCol[0];
                                XSectionPlane.Intersect((ISurface)plateOffsetPort, out wireCol, out IntersectionType);
                                plateOffsetWire = wireCol[0];
                                XSectionPlane.Intersect((ISurface)erTopPort, out wireCol, out IntersectionType);
                                erTopWire = wireCol[0];
                                XSectionPlane.Intersect((ISurface)erBottomPort, out wireCol, out IntersectionType);
                                erBottomWire = wireCol[0];
                            }
                            else
                            {
                                throw new Exception("XSectionPlane Should not be null");

                            }
                            // Get the distances between the wires to determine the section alias
                            double erTopToPlateBase, erTopToPlateOffset, erBotToPlateBase, erBotToPlateOffset;
                            Position pos1, pos2;
                            erTopWire.DistanceBetween(plateBaseWire, out erTopToPlateBase, out pos1, out pos2);
                            erTopWire.DistanceBetween(plateOffsetWire, out erTopToPlateOffset, out pos1, out pos2);
                            erBottomWire.DistanceBetween(plateBaseWire, out erBotToPlateBase, out pos1, out pos2);
                            erBottomWire.DistanceBetween(plateOffsetWire, out erBotToPlateOffset, out pos1, out pos2);

                            double webTh;
                            plateOffsetWire.DistanceBetween(plateBaseWire, out webTh, out pos1, out pos2);

                            if ((erTopToPlateBase <= this.tolerance) || (erTopToPlateOffset <= this.tolerance) || (erBotToPlateBase <= this.tolerance) || (erBotToPlateOffset <= this.tolerance))
                            {
                                sectionAlias = "UA";
                            }
                            else if ((erTopToPlateOffset <= webTh + this.tolerance) && (erTopToPlateBase <= webTh + this.tolerance))
                            {
                                sectionAlias = "UA";
                            }
                            else if ((erBotToPlateOffset <= webTh + this.tolerance) && (erBotToPlateBase <= webTh + this.tolerance))
                            {
                                sectionAlias = "UA";
                            }
                            else if ((erTopToPlateOffset < 0.015 + this.tolerance) || (erTopToPlateBase < 0.015 + this.tolerance))
                            {
                                sectionAlias = "BUTL2";
                            }
                            else if ((erBotToPlateOffset < 0.015 + this.tolerance) || (erBotToPlateBase < 0.015 + this.tolerance))
                            {
                                sectionAlias = "BUTL2";
                            }
                            else
                            {
                                sectionAlias = "BUT";
                            }
                        }
                        else if (erFlange.EdgeReinforcementPosition == EdgeReinforcementPosition.OnEdgeCentered)
                        {
                            sectionAlias = "BUT";
                        }
                    }
                    break;
                case PenetratingConfig.Plate_Stiffener:
                    sectionAlias = "BUTL3";

                    webSeamCurve.EndPoints(out startPos, out endPos);


                    Collection<Position> intersectPointOnWebSeam = null;
                    Collection<Position> overlapPointOnWebSeam = null;
                    try
                    {
                        if (flangeSeamCurve==null)
                        {
                            throw new Exception("flangeSeamCurve is null");
                        }
                        webSeamCurve.Intersect(flangeSeamCurve, out intersectPointOnWebSeam, out overlapPointOnWebSeam, out interSectionType);
                    }
                    catch (Exception)
                    {
                        throw new SlotMappingException("There is no Intersection Point between Web Seam and Flange Seam");
                    }

                    if (intersectPointOnWebSeam.Count == 1)
                    {

                        StiffenerPart stiffenerPart = null;

                        stiffenerPart = (StiffenerPart)flange;


                        CrossSection xSection = stiffenerPart.CrossSection;
                        double webThickness = 0;
                        PropertyValueDouble propWebThickness = (PropertyValueDouble)xSection.GetPropertyValue("IJUAXSectionWeb", "WebThickness");

                        if (propWebThickness != null)
                        {
                            webThickness = (double)propWebThickness.PropValue;
                        }

                        distnaceFromStartPos = startPos.DistanceToPoint(intersectPointOnWebSeam[0]);
                        distanceFormEndPos = endPos.DistanceToPoint(intersectPointOnWebSeam[0]);
                        Double minimumDistance = Math.Min(distnaceFromStartPos, distanceFormEndPos);

                        if (stiffenerPart.LoadPoint == (int)LoadPoint.Bottom_Flange_Left_Bottom_Corner)
                        {
                            if (minimumDistance <= this.tolerance)
                            {
                                sectionAlias = "UA";
                            }
                        }
                        else if (stiffenerPart.LoadPoint == (int)LoadPoint.Bottom)
                        {
                            if (Math.Abs(minimumDistance - webThickness / 2) <= this.tolerance)
                            {
                                sectionAlias = "UA";
                            }
                        }
                        else if (stiffenerPart.LoadPoint == (int)LoadPoint.Bottom_Flange_Right_Bottom_Corner)
                        {
                            if (Math.Abs(minimumDistance - webThickness) <= this.tolerance || minimumDistance <= this.tolerance)
                            {
                                sectionAlias = "UA";
                            }

                        }
                        else
                        {
                            // TBD
                        }
                    }
                    else
                    {
                        // TBD
                    }


                    break;
                case PenetratingConfig.Stiffener_Only:
                    break;

                case PenetratingConfig.BendPlate_Only:
                    {
                        sectionAlias = "UA";
                    }
                    break;
                default:
                    break;

            }
        }

        /// <summary>
        /// This method check Penetrating Configuration 
        /// Plate Only, Plate-Plate, Plate-ER, Plate-Stiffener, Unknown
        /// If there is only one stiffenr, it returns Unknown 
        /// </summary>
        /// <param name="penetratingObject">Input PenetratingPart.</param>
        /// <param name="otherPenetratingObject">Input Other PenetratingPart .</param>
        /// <returns>Returns PenetratingConfiguration - Plate Only, Plate-Plate, Plate-ER, Plate-Stiffener, Unknown  </returns>
        private PenetratingConfig GetPenetratingConfiguration(object penetratingObject, object otherPenetratingObject)
        {

            if (penetratingObject == null)
            {
                throw new ArgumentNullException("Input PenetratingPart");
            }

            if (otherPenetratingObject == null)
            {
                if (penetratingObject is PlatePart)
                {
                    return PenetratingConfig.Plate_Only;
                }
                else if (penetratingObject is StiffenerPart)
                {
                    //If there is only one stiffener, just return Unknown to return UnknownAlias. 
                    //In the rule, check the Alias and don't call this mapping rule 
                    return PenetratingConfig.Unknown;
                }
                else
                {
                    return PenetratingConfig.Unknown;
                }
            }
            else
            {
                if (penetratingObject is PlatePart && otherPenetratingObject is PlatePart)
                {
                    return PenetratingConfig.Plate_Plate;
                }
                else if ((penetratingObject is PlatePart && otherPenetratingObject is EdgeReinforcementPart))
                {
                    return PenetratingConfig.Plate_ER;
                }
                else if ((penetratingObject is PlatePart && otherPenetratingObject is StiffenerPart))
                {
                    return PenetratingConfig.Plate_Stiffener;
                }
                else if ((penetratingObject is EdgeReinforcementPart && otherPenetratingObject is PlatePart))
                {
                    return PenetratingConfig.Plate_ER;
                }
                else if (penetratingObject is StiffenerPart && otherPenetratingObject is PlatePart)
                {
                    return PenetratingConfig.Plate_Stiffener;
                }

            }
            return PenetratingConfig.Unknown;

        }

        /// <summary>
        /// Check if this object is Intersecting seam and partial splitter  
        /// This method is needed for following reason
        /// Base Plate is also connected object to the penetrating object and it penetrates penetrated object fully 
        /// To get other penetrating object, this base plate should be excluded in the GetSectionAlias method
        /// **********************************
        /// *                                *
        /// *      Partial Penetration       *   
        /// *            ***                 *  
        /// *                                *
        /// *      Full Penetration          *
        /// **********************************
        /// 
        /// Caution :
        ///     In this method, assign m_oWebSeam and m_oFlangeSeam
        ///     This variable will be used in GetEmulatedPorts method. 
        /// </summary>
        /// <param name="penetratedPart">Input oPenetratedPart .</param>
        /// <param name="inputRootSystem">Input oInputRootSystem which penetrates a penetrated object .</param>
        /// 
        private bool IsPartialSplitter(object penetratedPart, object inputLeafSystem)
        {

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("Input penetratedPart");
            }

            if (inputLeafSystem == null)
            {
                throw new ArgumentNullException("Input inputRootSystem");
            }

            bool intersectingSeam = false;
            PlateSystemBase rootPenetratedPlateSystem = null;
            PlateSystemBase leafPenetratedPlateSystem = null;
            StiffenerSystem rootPenetratedStiffenerSystem = null;
            StiffenerSystem leafPenetratedStiffenerSystem = null;
            BusinessObject inputRootSystem = null;

            ISystemChild systemChild = (ISystemChild)inputLeafSystem;
            inputRootSystem = (BusinessObject)systemChild.SystemParent;

            ISplit split = null;
            Seam seam = null;

            if (penetratedPart is PlatePart)
            {
                leafPenetratedPlateSystem = (PlateSystemBase)CommonFuncs.GetLeafSystemFromPart(penetratedPart);
                rootPenetratedPlateSystem = (PlateSystemBase)CommonFuncs.GetRootSystemFromPart(penetratedPart);

                split = rootPenetratedPlateSystem.GetSplit(inputRootSystem);
                seam = (Seam)split;

                if ((seam != null) && (seam.SeamType == SeamType.Intersection))
                {
                    intersectingSeam = true;
                }
            }
            else if (penetratedPart is StiffenerPart)
            {
                leafPenetratedStiffenerSystem = (StiffenerSystem)CommonFuncs.GetLeafSystemFromPart(penetratedPart);
                rootPenetratedStiffenerSystem = (StiffenerSystem)CommonFuncs.GetRootSystemFromPart(penetratedPart);

                split = rootPenetratedStiffenerSystem.GetSplit(inputRootSystem);
                seam = (Seam)split;

                if ((seam != null) && (seam.SeamType == SeamType.Intersection))
                {
                    intersectingSeam = true;
                }
            }


            //1. Check if inputRootSystem is intersecting the penetrated system or not 
            if (intersectingSeam == false)
            {
                //We don't need to check it is partial penetration or not 
                return false;
            }
            else
            {
                if (inputLeafSystem is PlateSystem)
                {
                    if (((PlateSystem)inputLeafSystem).Type == PlateType.FlangePlate || (GetPlateSubTypePropertyValue((PlateSystem)inputLeafSystem) == (int)PlateSubType.Flange))
                    {
                        //Plate is a Flange Plate - Return True
                        return true;
                    }
                }
                else if (inputLeafSystem is ProfileSystem)
                {
                    //Stiffener is connected to web plate - It is the Flange
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// GetConnectedPenetratingParts get parts that:
        ///    Connected to input penetrating part and
        ///    Penetrate input penetrated part
        /// </summary>
        /// <param name="penetratingPart">Input Penetrating Object.</param>
        /// <param name="penetratedPart">Input Penetrated Object.</param>
        /// <param name="penetratingConnectedParts">Connected Penetrating Parts Collection .</param> 
        /// 
        // Get penetrating leaf system
        // Get leaf systems connected to penetrating leaf
        // For each leaf system connected to penetrating leaf
        //    Get its root system
        //    Check if connected root penetrates penetrated root(intersection seam exist)
        //    If yes, keep leaf system first part
        //
        private void GetConnectedPenetratingParts(object penetratingPart,
                                                 object penetratedPart,
                                                 out Collection<BusinessObject> penetratingConnectedParts)
        {
            if (penetratingPart == null)
                throw new ArgumentNullException("Input penetratingPart");

            if (penetratedPart == null)
                throw new ArgumentNullException("Input penetratedPart");

            object penetratingRootSys = CommonFuncs.GetRootSystemFromPart(penetratingPart);
            object penetratedRootSys = CommonFuncs.GetRootSystemFromPart(penetratedPart);

            ISplit splitAtRootSystemsOfPtgPtd = null;
            Seam seamAtRootSystemsOfPtgPtd = null;

            PlateSystemBase penetratedRootPlateSys = null;
            StiffenerSystem penetratedRootStiffenerSys = null;

            //Get seam between penetrating and penetrated systems
            if (penetratedPart is PlatePart)
            {
                penetratedRootPlateSys = (PlateSystemBase)penetratedRootSys;
                splitAtRootSystemsOfPtgPtd = penetratedRootPlateSys.GetSplit((BusinessObject)penetratingRootSys);
            }
            else if (penetratedPart is ProfilePart)
            {
                penetratedRootStiffenerSys = (StiffenerSystem)penetratedRootSys;
                splitAtRootSystemsOfPtgPtd = penetratedRootStiffenerSys.GetSplit((BusinessObject)penetratingRootSys);
            }
            if (splitAtRootSystemsOfPtgPtd != null)
            {
                seamAtRootSystemsOfPtgPtd = (Seam)splitAtRootSystemsOfPtgPtd;
            }
            else
            {
                throw new Exception("unable to get the seam between penetarting and penetrated");

            }
            // Check the input object type and get the other object 
            object penetratingLeaf = CommonFuncs.GetLeafSystemFromPart(penetratingPart);
            object penetratedLeaf = CommonFuncs.GetLeafSystemFromPart(penetratedPart);
            ReadOnlyCollection<BusinessObject> connectedLeafSystems = null;

            penetratingConnectedParts = new Collection<BusinessObject>();
            if (penetratingPart is PlatePart)
            {
                //Get Connected Objects From Plate Part
                Plate plateLeafSystem = (Plate)penetratingLeaf;
                connectedLeafSystems = plateLeafSystem.GetConnectedObjects();
            }
            else if (penetratingPart is ProfilePart)
            {
                //Get Connected Objects From Profile Part
                Profile ProfileLeafSystem = (Profile)penetratingLeaf;
                connectedLeafSystems = ProfileLeafSystem.GetConnectedObjects();
            }
            BusinessObject connectedRootObject = null;
            ISystemChild systemChild;
            Seam seam = null;
            ISplit split;

            if (connectedLeafSystems == null)
            {
                throw new ArgumentNullException("connectedLeafSystems Collection");
            }
            //Get the seam range box and connected objects' range box and check if both overlap.
            RangeBox rangBoxofSeamPtgPtd = null;

            rangBoxofSeamPtgPtd = seamAtRootSystemsOfPtgPtd.Range;
            double seamLength = seamAtRootSystemsOfPtgPtd.Length;


            Math3d math3d = new Math3d();
            foreach (BusinessObject ConnectedObject in connectedLeafSystems)
            {
                if (ConnectedObject.Equals(penetratedLeaf))
                {
                    //Ignore the Penetrated Part
                    continue;
                }
                else if (ConnectedObject is Plate)
                {
                    PlateSystemBase leafPlateSystem = (PlateSystemBase)ConnectedObject;
                    if (leafPlateSystem.Type == PlateType.Hull)
                    {
                        //Ignore the hull
                        continue;
                    }
                }
                Collection<BusinessObject> connectedObjectBOCol = new Collection<BusinessObject>();
                connectedObjectBOCol.Add(ConnectedObject);
                OrientedRangeBox rangeBoxOfConnObject = math3d.GetOrientedRangeBox(connectedObjectBOCol);
                RangeBoxIntersectionType connObjToPtgPtdSeam = rangeBoxOfConnObject.Intersects(rangBoxofSeamPtgPtd);

                if (connObjToPtgPtdSeam != RangeBoxIntersectionType.Outside)
                {
                    // Check if connected root system penetrates penetrated root system
                    systemChild = (ISystemChild)ConnectedObject;
                    connectedRootObject = (BusinessObject)systemChild.SystemParent;
                    //"seam' is the between penetarted and connected penetrating part
                    if (penetratedPart is PlatePart)
                    {
                        split = penetratedRootPlateSys.GetSplit(connectedRootObject);
                        seam = (Seam)split;
                    }
                    else if (penetratedPart is StiffenerPart)
                    {
                        split = penetratedRootStiffenerSys.GetSplit(connectedRootObject);
                        seam = (Seam)split;
                    }

                    if (seamAtRootSystemsOfPtgPtd.Equals(seam))
                        continue; //to avoid leaf ER systems on penetrating web part

                    if ((seam != null) && (seam.SeamType == SeamType.Intersection))
                    {
                        ISystem tempLeafParentSys = (ISystem)ConnectedObject;
                        object penetratingRoot = CommonFuncs.GetRootSystemFromPart(penetratingPart);

                        if (penetratingPart is PlatePart)
                        {
                            PlateSystem penetratingRootPlateSys = null;
                            penetratingRootPlateSys = (PlateSystem)penetratingRoot;
                            split = penetratingRootPlateSys.GetSplit(connectedRootObject);
                            Seam seamBtwWebAndFlange;
                            //Seam2 is the seam between penetrating part and connected objects to penetrating part
                            //If connected object is penetrating through penetrating part, then that part need not be considered
                            seamBtwWebAndFlange = (Seam)split;
                            if (seamBtwWebAndFlange != null)
                            {
                                //Do not consider this part
                            }
                            else
                            {
                                double DisBetSeams;
                                Position pos1, pos2;
                                //distance between flange seam and web seam should be zero
                                seamAtRootSystemsOfPtgPtd.DistanceBetween((ICurve)seam, out DisBetSeams, out pos1, out pos2);
                                if (DisBetSeams <= tolerance)
                                {
                                    penetratingConnectedParts.Add((BusinessObject)tempLeafParentSys.SystemChildren[0]);
                                }
                                else
                                {
                                    //We are not going To add the part in the collection
                                }
                            }
                        }
                        else
                        {
                            penetratingConnectedParts.Add((BusinessObject)tempLeafParentSys.SystemChildren[0]);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Get GetConnectedPenetratingFlangePart to know the Section Alias 
        /// </summary>
        /// <param name="penetratingPart">Input Penetrating Object.</param>
        /// <param name="penetratedPart">Input Penetrated Object.</param>
        /// <param name="config">Internal configuration.</param>
        /// <param name="connectedPenetratingFlangePartCol">Connected Penetrating Flange Part Collection .</param> 

        private void GetConnectedPenetratingFlangePart(object penetratingPart,
                                                       object penetratedPart,
                                                       out PenetratingConfig config,
                                                       out Collection<BusinessObject> connectedPenetratingFlangePartCol)
        {
            if (penetratingPart == null)
            {
                throw new ArgumentNullException("Input penetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("Input penetratedPart");
            }

            config = PenetratingConfig.Unknown;
            connectedPenetratingFlangePartCol = null;

            // Check the input object type and get the other object 
            connectedPenetratingFlangePartCol = new Collection<BusinessObject>();
            Collection<BusinessObject> connectedPenetratingParts = null;

            this.GetConnectedPenetratingParts(penetratingPart, penetratedPart, out connectedPenetratingParts);
            foreach (BusinessObject connectedPart in connectedPenetratingParts)
            {
                if (connectedPart is PlatePart)
                {
                    if (((PlatePart)connectedPart).Type == PlateType.FlangePlate || (GetPlateSubTypePropertyValue((PlatePart)connectedPart) == (int)PlateSubType.Flange))
                    {
                        //Plate is a Flange Plate - Return True
                        connectedPenetratingFlangePartCol.Add(connectedPart);
                    }
                }
                else if (connectedPart is ProfilePart)
                {
                    //Stiffener is connected to web plate - It is the Flange
                    connectedPenetratingFlangePartCol.Add(connectedPart);
                }
            }

            // For now, the number of ConnectedPenetratingRootSystem is zero or one. 
            // Return unknown for other case 
            if (connectedPenetratingFlangePartCol.Count == 0)
            {
                Collection<TopologyPort> WLCollection = null;
                Collection<TopologyPort> WRCollection = null;
                Collection<TopologyPort> LateralCollection = null;

                if (CommonFuncs.IsBendPlate(penetratingPart, penetratedPart, false, out WLCollection, out WRCollection, out LateralCollection))
                {
                    config = PenetratingConfig.BendPlate_Only;
                }
                else
                {
                    config = GetPenetratingConfiguration(penetratingPart, null);
                }
            }
            else if (connectedPenetratingFlangePartCol.Count == 1)
            {
                config = GetPenetratingConfiguration(penetratingPart, connectedPenetratingFlangePartCol[0]);
            }
            else
            {
                config = PenetratingConfig.Unknown;
            }

        }

        /// <summary>
        /// Get the coordinate system of the cross section
        /// </summary>
        /// <param name="penetratingPart">The Penetrating Object.</param>
        /// <param name="penetratedPart">The Penetrated Object.</param>
        /// <param name="webRightPort">Port mapped to Web Right of the Section Alias.</param>
        /// <param name="basePlatePort">Port of the Base Plate.</param>
        /// <param name="sectionW">W Vector of the Section Coordinate System.</param>
        /// <param name="sectionU">U Vector of the Section Coordinate System.</param>
        /// <param name="sectionV">V Vector of the Section Coordinate System.</param> 

        private void GetCrossSectionCoordSys(object penetratedPart, TopologyPort webRightPort, TopologyPort basePlatePort, out Vector sectionW, out Vector sectionU, out Vector sectionV)
        {
            sectionU = null;
            sectionV = null;
            sectionW = null;

            Vector basePlateNormal = null;
            Position origin;
            TopologySurface basePlateSur = (TopologySurface)basePlatePort.Geometry;
            bool basePlatePlanar = basePlatePort.SupportsInterface(CommonFuncs.planeCOMinterface);
            bool webPlatePlanar = webRightPort.SupportsInterface(CommonFuncs.planeCOMinterface);

            if (basePlatePlanar && webPlatePlanar)
            {
                basePlateNormal = basePlatePort.Normal;
                sectionU = new Vector(webRightPort.Normal);
            }
            else //If base plate or penetrating object is curved
            {
                ISurface basePlateSurface = (ISurface)basePlatePort;
                Seam penetratingSeam;

                CommonFuncs.GetSeamBtwObjects(webRightPort.Connectable, penetratedPart, out penetratingSeam);
                if (penetratingSeam != null)
                {
                    //Get seam end points and choose nearer 
                    Position StartPos = null;
                    Position EndPos = null;
                    penetratingSeam.EndPoints(out StartPos, out EndPos);

                    Point3d seamPoint = new Point3d(StartPos.X, StartPos.Y, StartPos.Z);
                    double distance;
                    Position positionOnbasePlateSurface;
                    basePlateSur.DistanceBetween(seamPoint, out distance, out positionOnbasePlateSurface);

                    if (positionOnbasePlateSurface != null && positionOnbasePlateSurface.DistanceToPoint(EndPos) - distance > this.tolerance)
                    {
                        //StartPos is closest to the base plate
                    }
                    else
                    {
                        //check using EndPos
                        seamPoint = new Point3d(EndPos.X, EndPos.Y, EndPos.Z);
                        basePlateSur.DistanceBetween(seamPoint, out distance, out positionOnbasePlateSurface);
                    }

                    if (positionOnbasePlateSurface != null)
                    {
                        origin = positionOnbasePlateSurface;
                        basePlateNormal = basePlateSurface.OutwardNormalAtPoint(origin);
                        Vector webRightPortNormal = webRightPort.OutwardNormalAtPoint(origin);
                        sectionU = new Vector(webRightPortNormal);
                    }
                    else
                    {
                        throw new ArgumentNullException("There is no intersection between seam and base plate");
                    }
                }
            }
            if (sectionU != null && basePlateNormal != null)
                sectionW = sectionU.Cross(basePlateNormal);
            if (sectionW != null)
            {
                sectionV = sectionW.Cross(sectionU);

                sectionU.Length = 1.0;
                sectionW.Length = 1.0;
                sectionV.Length = 1.0;
            }
        }

        #endregion // Private Members
    }

}
