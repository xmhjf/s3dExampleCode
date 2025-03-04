//-----------------------------------------------------------------------------
//      Copyright (C) 2011-16 Intergraph Corporation.  All rights reserved.
//
//      Component:  SectionBUTL3PortMap to be implemented by all classes
//                  that are returning a EmulatedPort map according to the sectionAlias.
//
//      Author:  
//
//      History:
//      November 13, 2011       BSLee                  Created
//      Nov 19, 2015       mchandak               DI-275514 and DI-275249 Updated Slot Mapping Rule to avoid interop calls for GetNormalFromPostion() and IsExternalWire()
//      December 11, 2015       hgajula                DI-CP-284051  Fix coverity defects stated in November 6, 2015 report
//      December 11, 2015       knukala                DI-CP-275525  Replace PlaceIntersectionObject method in SlotMapng rule with .Net API 
//      April 1, 2016      svsmylav               TR-291159 Modified 'GetEmulatedPortsMap' method to use geometry 
//                                                 of the flange port (earlier port is used in 'Intersect' method call).
//-----------------------------------------------------------------------------

using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Configuration;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services.Hidden;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;



namespace Ingr.SP3D.Content.Structure.EmulatedPortMappings
{
    /// <summary>
    /// The SectionBUTL3PortMap class is designed to return emulated port maps based on BUTL3 SectionAlias
    /// </summary>
    internal class SectionBUTL3PortMap : IEmulatedPortMap
    {
        #region IEmulatedPortMap Members
        private double tolerance = 0.0001;

        /// <summary>
        /// Get the Ports on the penetrating parts depends on SectionAlias with SectionFaceType (ex) Web_Left..)
        /// </summary>
        /// <param name="penetratingPart">The Penetrating Object.</param>
        /// <param name="penetratedPart">The Penetrated Object.</param>
        /// <param name="basePltPort">BasePlate Port.</param>
        /// <param name="sectionAlias">SectionAlias .</param>
        /// <param name="web">The Web Object.</param>
        /// <param name="flange">The Flange Object.</param>
        /// <param name="secondWeb">The 2ndWeb Object.</param>
        /// <param name="secondFlange">The 2ndFlange Object.</param>
        /// <param name="basePlatePort">Base Plate Port.</param>
        /// <returns>The Ports on Penetrating parts. Ports that don't apply are omitted </returns>

        public Dictionary<int, IPort> GetEmulatedPortsMap(object penetratingPart, object penetratedPart, IPort basePltPort, string sectionAlias, object web, object flange, object secondWeb, object secondFlange, out IPort basePlatePort)
        {

            if (penetratingPart == null)
            {
                throw new ArgumentNullException("Input PenetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("Input PenetratedPart");
            }

            if (basePltPort == null)
            {
                throw new ArgumentNullException("Input PenetratedLateralEdgePort");
            }

            if (web == null)
            {
                throw new ArgumentNullException("Input Web");
            }


            basePlatePort = null;
            Dictionary<int, IPort> mappedPorts = null;

            mappedPorts = new Dictionary<int, IPort>();
            IPort basePortOfWeb = null;
            IPort offsetPortOfWeb = null;
            IPort bottomPortOfWeb = null;
            IPort topPortOfWeb = null;
            object rootBasePlateSystem = null;

            // Get Web and Base plate ports
            CommonFuncs.GetWebAndBasePlatePorts(penetratingPart, penetratedPart, basePltPort, web, flange, secondWeb, secondFlange, out basePlatePort,
                out basePortOfWeb, out offsetPortOfWeb, out bottomPortOfWeb, out topPortOfWeb, out rootBasePlateSystem);


            //                                Top
            //                               ******
            //                               *    * WebRightTop         TopFlangeRightTop
            //                               *    ****************************
            //                               *                               * TopFlangeRight
            //                               *    ****************************  
            //                               *    *                     TopFlangeRightBottom
            //                       WebLeft *    *  WebRight    
            //                               *    *  
            //                               *    *
            //                               ****** 
            //                               Bottom
            //
            //     
            //     In this case, TopFlangeLeftBottom , TopFlangeRightBottom is same mounting face 
            //                   WebLeft can be base or offset port of web depnds on location of TopFlange
            //                   Other ports depends on Flange. Flange can be a plate or ER


            IPort basePortOfFlange = null;

            if (flange is PlatePart)
            {
                basePortOfFlange = CommonFuncs.GetBaseOrOffsetPortOfPlatePart(flange, true, penetratedPart);

            }
            else if (flange is EdgeReinforcementPart)
            {

                EdgeReinforcementPart erFlangePart = (EdgeReinforcementPart)flange;
                ReadOnlyCollection<TopologyPort> sectionFacesCol = null;
                sectionFacesCol = erFlangePart.GetPorts(TopologyGeometryType.Face, GeometryStage.Initial);
                if (sectionFacesCol != null)
                {
                    foreach (TopologyPort port in sectionFacesCol)
                    {

                        if (port.SectionId == (int)SectionFaceType.Web_Left)
                        {
                            basePortOfFlange = (IPort)port;
                        }
                        else
                        {

                        }
                    }
                }
            }
            else if (flange is StiffenerPart)
            {

                StiffenerPart stiffenrPart = (StiffenerPart)flange;
                ReadOnlyCollection<TopologyPort> sectionFacesCol = null;
                sectionFacesCol = stiffenrPart.GetPorts(TopologyGeometryType.Face, GeometryStage.Initial);
                if (sectionFacesCol != null)
                {
                    foreach (TopologyPort port in sectionFacesCol)
                    {

                        if (port.SectionId == (int)SectionFaceType.Web_Left)
                        {
                            basePortOfFlange = (IPort)port;
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

            Position startPos = null;
            Position endPos = null;
            Position middlePosition = new Position();

            TopologyPort basePortofPenetreatedPart = null;

            //If laternal port is connnected to Base Plate, it is BottomPort of Web

            if (penetratedPart is PlatePart)
            {
                basePortofPenetreatedPart = (TopologyPort)CommonFuncs.GetBaseOrOffsetPortOfPlatePart(penetratedPart, true, penetratingPart);
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

                        if (port.SectionId == (int)SectionFaceType.Web_Left)
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



            // Get Intersection Curve between BasePortOfFlange and BasePortOfPenetratedPlate
            
            if (basePortofPenetreatedPart != null && basePortOfFlange != null)
            {
                ICurve intersectionCurve = null;
                Collection<ICurve> curveCollection;
                GeometryIntersectionType intersectionType;
                TopologySurface baseSurfaceofPenetreatedPart = (TopologySurface)basePortofPenetreatedPart.Geometry;
                TopologyPort baseTopologyPortOfFlange = (TopologyPort)basePortOfFlange;
                TopologySurface baseSurfaceofFlange = (TopologySurface)baseTopologyPortOfFlange.Geometry;
                baseSurfaceofPenetreatedPart.Intersect((ISurface)baseSurfaceofFlange, out curveCollection, out intersectionType);
                if (curveCollection != null)
                {
                    if (curveCollection.Count != 0)
                    {
                        intersectionCurve = curveCollection[0];
                        if (curveCollection.Count > 1)
                        {
                            for (int icount = 1; icount < curveCollection.Count; icount++)
                                if (curveCollection[icount].Length > intersectionCurve.Length)
                                {
                                    intersectionCurve = curveCollection[icount]; //longer curve is considered
                                }
                        }

                        intersectionCurve.EndPoints(out startPos, out endPos);
                    }
                }
            }
            else
            {
                throw new ArgumentNullException("Input basePortofPenetreatedPart or basePortOfFlange");
            }

            if (startPos != null && endPos != null)
            {
                middlePosition.X = (startPos.X + endPos.X) / 2;
                middlePosition.Y = (startPos.Y + endPos.Y) / 2;
                middlePosition.Z = (startPos.Z + endPos.Z) / 2;
            }
            else
            {
                throw new ArgumentNullException("Input startPos or endPos");
            }

           
            TopologyPort baseTopoPortOfWeb = (TopologyPort)basePortOfWeb;
            TopologyPort offsetToPoPortOfWeb = (TopologyPort)offsetPortOfWeb;


            double distBaseAndMiddlePoint = 0;
            double distOffsetAndMiddlePoint = 0;

            Position comPos1;
            Position comPos2;

            Point3d point = new Point3d(middlePosition.X, middlePosition.Y, middlePosition.Z);
            baseTopoPortOfWeb.DistanceBetween(point, out distBaseAndMiddlePoint, out comPos1);
            offsetToPoPortOfWeb.DistanceBetween(point, out distOffsetAndMiddlePoint, out comPos2);
         

            if (distBaseAndMiddlePoint > distOffsetAndMiddlePoint)
            {

                mappedPorts.Add((int)SectionFaceType.Web_Left, basePortOfWeb);
                mappedPorts.Add((int)SectionFaceType.Web_Right, offsetPortOfWeb);
                mappedPorts.Add((int)SectionFaceType.Web_Right_Top, offsetPortOfWeb);


            }
            if (distBaseAndMiddlePoint < distOffsetAndMiddlePoint)
            {
                mappedPorts.Add((int)SectionFaceType.Web_Left, offsetPortOfWeb);
                mappedPorts.Add((int)SectionFaceType.Web_Right, basePortOfWeb);
                mappedPorts.Add((int)SectionFaceType.Web_Right_Top, basePortOfWeb);
            }
            else
            {

            }

            // Web Ports 
            mappedPorts.Add((int)SectionFaceType.Bottom, bottomPortOfWeb);
            mappedPorts.Add((int)SectionFaceType.Top, topPortOfWeb);


            if (flange is PlatePart)
            {
                IPort basePortOfFlangePlate = null;
                IPort offsetPortOfFlangePlate = null;

                basePortOfFlangePlate = CommonFuncs.GetBaseOrOffsetPortOfPlatePart(flange, true, penetratedPart);
                offsetPortOfFlangePlate = CommonFuncs.GetBaseOrOffsetPortOfPlatePart(flange, false, penetratedPart);

                TopologyPort baseTopoPortOfFlangePlate = (TopologyPort)basePortOfFlangePlate;
                TopologyPort offsetTopoPortOfFlangePlate = (TopologyPort)offsetPortOfFlangePlate;
                TopologyPort bottomTopoPortOfWeb = (TopologyPort)bottomPortOfWeb;

                //Get Normal 
                Vector bottomNormal = null, basePortNormal = null, offsetPortNormal = null;
                //Get Normal 
                if (bottomTopoPortOfWeb != null)
                {
                    bottomNormal = CommonFuncs.GetNormalVectorOfTopologyPort(bottomTopoPortOfWeb, (TopologyPort)basePortOfWeb); //oBottomTopoPortOfWeb.Normal;
                }
                else
                {
                    throw new ArgumentNullException("Input bottomNormal");
                }
                if (baseTopoPortOfFlangePlate != null)
                {
                    basePortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(baseTopoPortOfFlangePlate, (TopologyPort)basePortOfWeb); //oBaseTopoPortOfFlangePlate.Normal;
                }
                else
                {
                    throw new ArgumentNullException("Input basePortNormal");
                }
                if (offsetTopoPortOfFlangePlate != null)
                {
                    offsetPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(offsetTopoPortOfFlangePlate, (TopologyPort)basePortOfWeb); //oOffsetTopoPortOfFlangePlate.Normal;
                }
                else
                {
                    throw new ArgumentNullException("Input offsetPortNormal");
                }

                double dotBottomAndBase = bottomNormal.Dot(basePortNormal);
                double dotBottomAndOffset = bottomNormal.Dot(offsetPortNormal);


                if (dotBottomAndBase * dotBottomAndOffset < 0)
                {
                    if (dotBottomAndBase > 0)
                    {
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Top, offsetPortOfFlangePlate);
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Bottom, basePortOfFlangePlate);
                    }
                    else
                    {
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Top, basePortOfFlangePlate);
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Bottom, offsetPortOfFlangePlate);
                    }
                }
                else if (dotBottomAndBase * dotBottomAndOffset > 0)
                {
                    // No case
                }
                else if ((dotBottomAndBase * dotBottomAndOffset).EqualTo(0) == true)
                {
                    // No case
                }
                else
                {

                }

                //TopFlangeRight 
                //Opposit to Web_Left Port 
                TopologyPort webLeftTopoPortOfWeb = (TopologyPort)mappedPorts[(int)SectionFaceType.Web_Left];
                Vector webLeftNormal = CommonFuncs.GetNormalVectorOfTopologyPort(webLeftTopoPortOfWeb, basePortofPenetreatedPart); //oWebLeftTopoPortOfWeb.Normal;

                // TopFlangeRightBottom and TopFlangeLeftBottom 
                // Bottom Port && Top Port

                PlatePart flangePlatePart;
                ReadOnlyCollection<TopologyPort> oFlangeLateralPortsElems;

                flangePlatePart = (PlatePart)flange;

                oFlangeLateralPortsElems = flangePlatePart.GetPorts(TopologyGeometryType.Face,ContextTypes.Lateral , GeometryStage.Initial);
                

                //Get BasePort of PenetratedPlate to check Intersect
                if (oFlangeLateralPortsElems != null && oFlangeLateralPortsElems.Count > 0)
                {
                    for (int i = 0; i <= oFlangeLateralPortsElems.Count-1; i++)
                    {

                        object flangePort = null;
                        flangePort = oFlangeLateralPortsElems[i];


                        if (flangePort is TopologyPort)
                        {
                            TopologyPort port = (TopologyPort)flangePort;
                            //Top_Flange_Right
                            double minimumDistance = 0;
                            Position posSrcPos = null;
                            Position PosInPos = null;
                            basePortofPenetreatedPart.DistanceBetween((ISurface)port.Geometry, out minimumDistance, out posSrcPos, out PosInPos);

                            if (minimumDistance < tolerance)
                            {

                                Vector portNormal = CommonFuncs.GetNormalVectorOfTopologyPort(port, basePortofPenetreatedPart); //oPort.Normal;
                                if (webLeftNormal.Dot(portNormal) < 0)
                                {
                                    mappedPorts.Add((int)SectionFaceType.Top_Flange_Right, port);
                                }

                                else
                                {

                                }
                            }
                        }
                    }
                }


            }
            else if (flange is StiffenerPart || flange is EdgeReinforcementPart)
            {
                IPort webRightPortOfFlange = null;
                IPort webLeftPortOfFlange = null;
                IPort bottomPortOfFlange = null;
                IPort topPortOfFlange = null;

                ReadOnlyCollection<TopologyPort> sectionFacesCol = null;
                if (flange is StiffenerPart)
                {
                    StiffenerPart stiffenerFlangePart = (StiffenerPart)flange;
                    sectionFacesCol = stiffenerFlangePart.GetPorts(TopologyGeometryType.Face, GeometryStage.Initial);
                }
                else if (flange is EdgeReinforcementPart)
                {
                    EdgeReinforcementPart edgeReinForcementPart = (EdgeReinforcementPart)flange;
                    sectionFacesCol = edgeReinForcementPart.GetPorts(TopologyGeometryType.Face, GeometryStage.Initial);
                }

                if (sectionFacesCol != null)
                {
                    foreach (TopologyPort port in sectionFacesCol)
                    {
                        if (port.SectionId == (int)SectionFaceType.Top)
                        {
                            topPortOfFlange = (IPort)port;
                        }
                        else if (port.SectionId == (int)SectionFaceType.Bottom)
                        {
                            bottomPortOfFlange = (IPort)port;
                        }
                        else if (port.SectionId == (int)SectionFaceType.Web_Right)
                        {
                            webRightPortOfFlange = (IPort)port;
                        }
                        else if (port.SectionId == (int)SectionFaceType.Web_Left)
                        {
                            webLeftPortOfFlange = (IPort)port;
                        }
                        else
                        {

                        }
                    }
                }

                TopologyPort bottomTopoPortWeb = (TopologyPort)mappedPorts[(int)SectionFaceType.Bottom];
                TopologyPort webRightTopoPortOfFlange = (TopologyPort)webRightPortOfFlange;
                TopologyPort webLeftTopoPortOfFlange = (TopologyPort)webLeftPortOfFlange;

                Vector bottomNormal = CommonFuncs.GetNormalVectorOfTopologyPort(bottomTopoPortWeb, (TopologyPort)basePortOfWeb);
                Vector webRightPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(webRightTopoPortOfFlange, basePortofPenetreatedPart); //oWebRightTopoPortOfFlange.Normal;
                Vector webLefttPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(webLeftTopoPortOfFlange, basePortofPenetreatedPart); //oWebLeftTopoPortOfFlange.Normal;

                double dotBottomAndWebRight = bottomNormal.Dot(webRightPortNormal);
                double dotBottomAndWebLeft = bottomNormal.Dot(webLefttPortNormal);


                if (dotBottomAndWebRight * dotBottomAndWebLeft < 0)
                {
                    if (dotBottomAndWebRight > 0)
                    {
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Top, webLeftPortOfFlange);
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Bottom, webRightPortOfFlange);
                    }
                    else
                    {
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Top, webRightPortOfFlange);
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Bottom, webLeftPortOfFlange);
                    }
                }
                else if (dotBottomAndWebRight * dotBottomAndWebLeft > 0)
                {
                    // TBD
                }
                else if ((dotBottomAndWebRight * dotBottomAndWebLeft).EqualTo(0) == true)
                {
                    // TBD
                }
                else
                {

                }

                //Opposit to Web_Left Port 
                TopologyPort webLeftTopoPortOfWeb = (TopologyPort)mappedPorts[(int)SectionFaceType.Web_Left];
                TopologyPort topTopoPortOfFlange = (TopologyPort)topPortOfFlange;
                TopologyPort bottomTopoPortOfFlange = (TopologyPort)bottomPortOfFlange;

                Vector webLeftOfWebNormal = CommonFuncs.GetNormalVectorOfTopologyPort(webLeftTopoPortOfWeb, basePortofPenetreatedPart); //oWebLeftTopoPortOfWeb.Normal;
                Vector flangeTopPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(topTopoPortOfFlange, basePortofPenetreatedPart); //oTopTopoPortOfFlange.Normal;
                Vector flangeBottomPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(bottomTopoPortOfFlange, basePortofPenetreatedPart); //oBottomTopoPortOfFlange.Normal;

                double dotWebLeftAndTop = webLeftOfWebNormal.Dot(flangeTopPortNormal);
                double dotWebLeftAndBottom = webLeftOfWebNormal.Dot(flangeBottomPortNormal);


                if (dotWebLeftAndTop * dotWebLeftAndBottom < 0)
                {
                    if (dotWebLeftAndTop > 0)
                    {
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right, bottomPortOfFlange);
                    }
                    else
                    {
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right, topPortOfFlange);
                    }
                }
                else if (dotWebLeftAndTop * dotWebLeftAndBottom > 0)
                {
                    // TBD
                }
                else if ((dotWebLeftAndTop * dotWebLeftAndBottom).EqualTo(0) == true)
                {
                    // TBD 
                }
                else
                {
                    // TBD
                }
            }
            else
            {

            }


            return mappedPorts;
        }

        #endregion

    }
}
