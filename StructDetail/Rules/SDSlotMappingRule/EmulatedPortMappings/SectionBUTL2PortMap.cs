//-----------------------------------------------------------------------------
//      Copyright (C) 2011-16 Intergraph Corporation.  All rights reserved.
//
//      Component:  SectionBUTL2PortMap to be implemented by all classes
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
    /// The SectionBUTL2PortMap class is designed to return emulated port maps based on BUTL2 SectionAlias
    /// </summary>
    internal class SectionBUTL2PortMap : IEmulatedPortMap
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

            //flange can be null object

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


            //    
            //                                  Top
            //                          ******************
            //             TopFlangeLeft*                * TopFlangeRight
            //                          ******************         
            //          TopFlangeLeftBottom  *    *  TopFlangeRighttBottom
            //                               *    *  
            //                       WebLeft *    *  WebRight
            //                       (Base)  *    *  (Offset)
            //                               *    *
            //                               ****** 
            //                               Bottom
            //
            //     
            //     In this case, TopFlangeLeftBottom , TopFlangeRightBottom is same mounting face 
            //                   WebLeft can be base or offset port of web depnds on location of TopFlange
            //                   Other ports depends on Flange. Flange can be a plate or ER


            //Check the Distace between a middle point of seam and Base or Offset port of web 
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
                basePortofPenetreatedPart = (TopologyPort)CommonFuncs.GetBaseOrOffsetPortOfPlatePart(penetratedPart, true, penetratingPart );
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
            ICurve intersectionCurve = null;
            if (basePortofPenetreatedPart != null && basePortOfFlange != null)
            {
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
            //Measure the distance from middleposition to  base port or offset port of 
         
            TopologyPort baseTopoPortOfWeb = (TopologyPort)basePortOfWeb;
            TopologyPort offsetToPologyPortOfWeb = (TopologyPort)offsetPortOfWeb;

            Position comPos1;
            Position comPos2;

            double distBaseAndMiddlePoint = 0;
            double distOffsetAndMiddlePoint = 0;
          
            Point3d point = new Point3d(middlePosition.X, middlePosition.Y, middlePosition.Z);
            baseTopoPortOfWeb.DistanceBetween(point, out distBaseAndMiddlePoint, out comPos1);
            offsetToPologyPortOfWeb.DistanceBetween(point, out distOffsetAndMiddlePoint, out comPos2);
      

            if (distBaseAndMiddlePoint > distOffsetAndMiddlePoint)
            {

                mappedPorts.Add((int)SectionFaceType.Web_Left, basePortOfWeb);
                mappedPorts.Add((int)SectionFaceType.Web_Right, offsetPortOfWeb);


            }
            if (distBaseAndMiddlePoint < distOffsetAndMiddlePoint)
            {
                mappedPorts.Add((int)SectionFaceType.Web_Left, offsetPortOfWeb);
                mappedPorts.Add((int)SectionFaceType.Web_Right, basePortOfWeb);
            }
            else
            {

            }

            // Web Ports 
            mappedPorts.Add((int)SectionFaceType.Bottom, bottomPortOfWeb);


            // Flange Ports 
            if (flange is EdgeReinforcementPart)
            {
                IPort webRightPortOfER = null;
                IPort webLeftPortOfER = null;
                IPort bottomPortOfER = null;
                IPort topPortOfER = null;

                EdgeReinforcementPart erFlangePart = (EdgeReinforcementPart)flange;
                ReadOnlyCollection<TopologyPort> sectionFacesCol = null;
                sectionFacesCol = erFlangePart.GetPorts(TopologyGeometryType.Face, GeometryStage.Initial);
                if (sectionFacesCol != null)
                {
                    foreach (TopologyPort port in sectionFacesCol)
                    {
                        if (port.SectionId == (int)SectionFaceType.Top)
                        {
                            topPortOfER = (IPort)port;
                        }
                        else if (port.SectionId == (int)SectionFaceType.Bottom)
                        {
                            bottomPortOfER = (IPort)port;
                        }
                        else if (port.SectionId == (int)SectionFaceType.Web_Right)
                        {
                            webRightPortOfER = (IPort)port;
                        }
                        else if (port.SectionId == (int)SectionFaceType.Web_Left)
                        {
                            webLeftPortOfER = (IPort)port;
                        }
                        else
                        {

                        }
                    }
                }

                //TopFlangeRightBottom and TopFlangeLeftBottom are same port

                TopologyPort webLeftTopoPortofER = (TopologyPort)webLeftPortOfER;
                TopologyPort webRightTopoPortofER = (TopologyPort)webRightPortOfER;
                TopologyPort topTopoPortOfWeb = (TopologyPort)topPortOfWeb;


                //Get Normal Vector 
                Vector topPortNormal = null, webLeftPortNormal = null, webRightPortNormal = null;
                if (topTopoPortOfWeb != null)
                {
                    topPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(topTopoPortOfWeb, (TopologyPort)basePortOfWeb);//oTopTopoPortOfWeb.Normal;
                }
                else
                {
                    throw new ArgumentNullException("Input topPortNormal");
                }
                if (webLeftTopoPortofER != null)
                {
                    webLeftPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(webLeftTopoPortofER, basePortofPenetreatedPart);//oWebLeftTopoPortofER.Normal;
                }
                else
                {
                    throw new ArgumentNullException("Input webLeftTopoPortofER");
                }
                if (webRightTopoPortofER != null)
                {
                    webRightPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(webRightTopoPortofER, basePortofPenetreatedPart);//oWebRightTopoPortofER.Normal;
                }
                else
                {
                    throw new ArgumentNullException("Input webRightTopoPortofER");
                }

                double dotTopAndWebLeft = topPortNormal.Dot(webLeftPortNormal);
                double dotTopAndWebRight = topPortNormal.Dot(webRightPortNormal);


                if (dotTopAndWebLeft * dotTopAndWebRight < 0)
                {
                    if (dotTopAndWebLeft > 0)
                    {
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Bottom, webRightPortOfER);
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Left_Bottom, webRightPortOfER);
                        mappedPorts.Add((int)SectionFaceType.Top, webLeftPortOfER);
                    }
                    else
                    {
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Bottom, webLeftPortOfER);
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Left_Bottom, webLeftPortOfER);
                        mappedPorts.Add((int)SectionFaceType.Top, webRightPortOfER);
                    }
                }
                else if (dotTopAndWebLeft * dotTopAndWebRight > 0)
                {
                    // No case
                }
                else if ((dotTopAndWebLeft * dotTopAndWebRight).EqualTo(0) == true)
                {
                    // No case
                }
                else
                {

                }

                // TopFlangeRight and TopFlangeLeft
                // The normal of TopFlangeLeft is most closest direction of BaePort

                TopologyPort webLeftTopoPortOfWeb = (TopologyPort)mappedPorts[(int)SectionFaceType.Web_Left];
                TopologyPort topTopoPortOfFlange = (TopologyPort)topPortOfER;
                TopologyPort bottomTopoPortOfFlange = (TopologyPort)bottomPortOfER;


                Vector webLeftNormal = CommonFuncs.GetNormalVectorOfTopologyPort(webLeftTopoPortOfWeb, basePortofPenetreatedPart); //oWebLeftTopoPortOfWeb.Normal;
                Vector topNormal = CommonFuncs.GetNormalVectorOfTopologyPort(topTopoPortOfFlange, basePortofPenetreatedPart); //oTopTopoPortOfFlange.Normal;
                Vector bottomNormal = CommonFuncs.GetNormalVectorOfTopologyPort(bottomTopoPortOfFlange, basePortofPenetreatedPart); //oBottomTopoPortOfFlange.Normal;

                double dotWebLeftTop = webLeftNormal.Dot(topNormal);
                double dotWebLeftBottom = webLeftNormal.Dot(bottomNormal);


                if (dotWebLeftTop > 0 && dotWebLeftBottom < 0)
                {
                    mappedPorts.Add((int)SectionFaceType.Top_Flange_Left, topPortOfER);
                    mappedPorts.Add((int)SectionFaceType.Top_Flange_Right, bottomPortOfER);

                }
                else if (dotWebLeftTop < 0 && dotWebLeftBottom > 0)
                {
                    mappedPorts.Add((int)SectionFaceType.Top_Flange_Left, bottomPortOfER);
                    mappedPorts.Add((int)SectionFaceType.Top_Flange_Right, topPortOfER);
                }

                else if (dotWebLeftTop * dotWebLeftBottom > 0)
                {
                    throw new NotImplementedException();
                }

                else if ((dotWebLeftTop * dotWebLeftBottom).EqualTo(0) == true)
                {
                    throw new NotImplementedException();
                }
                else
                {
                    throw new NotImplementedException();
                }
            }
            else if (flange is PlatePart)
            {
                IPort basePortOfFlangePlate = null;
                IPort offsetPortOfFlangePlate = null;

                basePortOfFlangePlate = CommonFuncs.GetBaseOrOffsetPortOfPlatePart(flange, true, penetratedPart);
                offsetPortOfFlangePlate = CommonFuncs.GetBaseOrOffsetPortOfPlatePart(flange, false, penetratedPart);

                TopologyPort baseTopoPortofFlangePlate = (TopologyPort)basePortOfFlangePlate;
                TopologyPort offsetTopoPortofFlangePlate = (TopologyPort)offsetPortOfFlangePlate;
                TopologyPort topTopoPortOfWeb = (TopologyPort)topPortOfWeb;

                //Get Normal 
                Vector topPortNormal = null, basePortNormal = null, offsetPortNormal = null;
                if (topTopoPortOfWeb != null)
                {
                    topPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(topTopoPortOfWeb, basePortofPenetreatedPart); //oTopTopoPortOfWeb.Normal;
                }
                else
                {
                    throw new ArgumentNullException("Input topPortNormal");
                }
                if (baseTopoPortofFlangePlate != null)
                {
                    basePortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(baseTopoPortofFlangePlate, basePortofPenetreatedPart); //oBaseTopoPortofFlangePlate.Normal;
                }
                else
                {
                    throw new ArgumentNullException("Input basePortNormal");
                }
                if (offsetTopoPortofFlangePlate != null)
                {
                    offsetPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(offsetTopoPortofFlangePlate, basePortofPenetreatedPart); //oOffsetTopoPortofFlangePlate.Normal;
                }
                else
                {
                    throw new ArgumentNullException("Input offsetPortNormal");
                }

                double dotTopAndBase = topPortNormal.Dot(basePortNormal);
                double dotTopAndOffset = topPortNormal.Dot(offsetPortNormal);


                if (dotTopAndBase * dotTopAndOffset < 0)
                {
                    if (dotTopAndBase > 0)
                    {
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Bottom, offsetPortOfFlangePlate);
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Left_Bottom, offsetPortOfFlangePlate);
                        mappedPorts.Add((int)SectionFaceType.Top, basePortOfFlangePlate);
                    }
                    else
                    {
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Bottom, basePortOfFlangePlate);
                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Left_Bottom, basePortOfFlangePlate);
                        mappedPorts.Add((int)SectionFaceType.Top, offsetPortOfFlangePlate);
                    }
                }
                else if (dotTopAndBase * dotTopAndOffset > 0)
                {
                    // No case
                }
                else if ((dotTopAndBase * dotTopAndOffset).EqualTo(0) == true)
                {
                    // No case
                }
                else
                {

                }
                
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
                            //Top Port
                            double minimumDistance = 0;
                            Position posSrcPos = null;
                            Position PosInPos = null;
                            basePortofPenetreatedPart.DistanceBetween((ISurface)port.Geometry, out minimumDistance, out posSrcPos, out PosInPos);

                            if (minimumDistance < tolerance)
                            {

                                Vector portNormal = CommonFuncs.GetNormalVectorOfTopologyPort(port, (TopologyPort)basePortofPenetreatedPart); //oPort.Normal;
                                if (webLeftNormal.Dot(portNormal) > 0)
                                {
                                    mappedPorts.Add((int)SectionFaceType.Top_Flange_Left, port);
                                }
                                else if (webLeftNormal.Dot(portNormal) < 0)
                                {
                                    mappedPorts.Add((int)SectionFaceType.Top_Flange_Right, port);
                                }
                                else
                                {
                                    throw new NotImplementedException();
                                }

                            }
                        }
                    }
                }
            }


            return mappedPorts;
        }

        #endregion


  
    }
}
