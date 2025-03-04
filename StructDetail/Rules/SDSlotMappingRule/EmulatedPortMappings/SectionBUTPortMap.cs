//-----------------------------------------------------------------------------
//      Copyright (C) 2011-15 Intergraph Corporation.  All rights reserved.
//
//      Component:  SectionBUTPortMap to be implemented by all classes
//                  that are returning a EmulatedPort map according to the sectionAlias.
//
//      Author:  
//
//      History:
//      November 13, 2011       BSLee                  Created
//      Nov 19, 2015       mchandak               DI-275514 and DI-275249 Updated Slot Mapping Rule to avoid interop calls for GetNormalFromPostion() and IsExternalWire()
//      December 11, 2015       hgajula                DI-CP-284051  Fix coverity defects stated in November 6, 2015 report
//      December 11, 2015       knukala                DI-CP-275525  Replace PlaceIntersectionObject method in SlotMapng rule with .Net API 
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
    /// The SectionBUTPortMap class is designed to return emulated port maps based on BUT SectionAlias
    /// </summary>
    internal class SectionBUTPortMap : IEmulatedPortMap
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

            
            mappedPorts.Add((int)SectionFaceType.Bottom, bottomPortOfWeb);
            mappedPorts.Add((int)SectionFaceType.Web_Left, basePortOfWeb);
            mappedPorts.Add((int)SectionFaceType.Web_Right, offsetPortOfWeb);


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

            //Check Flange Web Port 

            if (flange is EdgeReinforcementPart)
            {
                IPort webRightPortOfER = null;
                IPort webLeftPortOfER = null;
                IPort bottomPortOfER = null;
                IPort topPortOfER = null;

                EdgeReinforcementPart erFlangePart = (EdgeReinforcementPart)flange;
                ReadOnlyCollection<TopologyPort> sectionFacesCol = null;
                sectionFacesCol = erFlangePart.GetPorts(TopologyGeometryType.Face, GeometryStage.Initial );
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


                TopologyPort webLeftTopoPortofER = (TopologyPort)webLeftPortOfER; 
                TopologyPort webRightTopoPortofER = (TopologyPort)webRightPortOfER;
                TopologyPort topTopoPortOfWeb = (TopologyPort)topPortOfWeb;

                Vector topPortNormal = null;

                if (topTopoPortOfWeb != null)
                {
                    topPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(topTopoPortOfWeb, (TopologyPort)basePortOfWeb); //oTopTopoPortOfWeb.Normal;
                }
                else
                {
                    throw new ArgumentNullException("Input topPortNormal");
                }

                if (basePortofPenetreatedPart == null)
                {
                    throw new ArgumentNullException("Input basePortofPenetreatedPart");
                }
                else
                {


                    if (webLeftTopoPortofER == null)
                    {
                        throw new ArgumentNullException("Input webLeftTopoPortofER");
                    }
                    else if (webRightTopoPortofER == null)
                    {
                        throw new ArgumentNullException("Input webRightTopoPortofER");
                    }
                    else
                    {

                        Vector webLeftPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(webLeftTopoPortofER, basePortofPenetreatedPart); //oWebLeftTopoPortofER.Normal;
                        Vector webRightPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(webRightTopoPortofER, basePortofPenetreatedPart); //oWebRightTopoPortofER.Normal;

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
                        TopologyPort baseTopoPortOfWeb = (TopologyPort)basePortOfWeb;
                        TopologyPort topTopoPortOfFlange = (TopologyPort)topPortOfER;
                        TopologyPort bottomTopoPortOfFlange = (TopologyPort)bottomPortOfER;


                        Vector baseNormal = CommonFuncs.GetNormalVectorOfTopologyPort(baseTopoPortOfWeb, basePortofPenetreatedPart); //oBaseTopoPortOfWeb.Normal;
                        Vector topNormal = CommonFuncs.GetNormalVectorOfTopologyPort(topTopoPortOfFlange, basePortofPenetreatedPart); //oTopTopoPortOfFlange.Normal;
                        Vector bottomNormal = CommonFuncs.GetNormalVectorOfTopologyPort(bottomTopoPortOfFlange, basePortofPenetreatedPart); //oBottomTopoPortOfFlange.Normal;

                        double dotBaseTop = baseNormal.Dot(topNormal);
                        double dotBaseBottom = baseNormal.Dot(bottomNormal);


                        if (dotBaseTop > 0 && dotBaseBottom < 0)
                        {
                            mappedPorts.Add((int)SectionFaceType.Top_Flange_Left, topPortOfER);
                            mappedPorts.Add((int)SectionFaceType.Top_Flange_Right, bottomPortOfER);

                        }
                        else if (dotBaseTop < 0 && dotBaseBottom > 0)
                        {
                            mappedPorts.Add((int)SectionFaceType.Top_Flange_Left, bottomPortOfER);
                            mappedPorts.Add((int)SectionFaceType.Top_Flange_Right, topPortOfER);
                        }

                        else if (dotBaseTop * dotBaseBottom > 0)
                        {
                            throw new NotImplementedException();
                        }

                        else if ((baseNormal.Dot(topNormal) * baseNormal.Dot(bottomNormal)).EqualTo(0) == true)
                        {
                              throw new NotImplementedException();
                        }
                        else
                        {
                            throw new NotImplementedException();

                        }
                    }
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

                Vector topPortNormal = null;

                if (topTopoPortOfWeb != null)
                {
                    topPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(topTopoPortOfWeb, (TopologyPort)basePortOfWeb); //oTopTopoPortOfWeb.Normal;
                }
                else
                {
                    throw new ArgumentNullException("Input topPortNormal");
                }

                if (basePortofPenetreatedPart == null)
                {
                    throw new ArgumentNullException("Input basePortofPenetreatedPart");
                }
                else
                {
                    if (baseTopoPortofFlangePlate == null)
                    {
                        throw new ArgumentNullException("Input baseTopoPortofFlangePlate");
                    }
                    else if (offsetTopoPortofFlangePlate == null)
                    {
                        throw new ArgumentNullException("Input offsetTopoPortofFlangePlate");
                    }
                    Vector basePortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(baseTopoPortofFlangePlate, basePortofPenetreatedPart); //oBaseTopoPortofFlangePlate.Normal;
                    Vector offsetPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(offsetTopoPortofFlangePlate, basePortofPenetreatedPart); //oOffsetTopoPortofFlangePlate.Normal;

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

                    TopologyPort baseTopoPortOfWeb = (TopologyPort)basePortOfWeb;
                    Vector baseNormal = CommonFuncs.GetNormalVectorOfTopologyPort(baseTopoPortOfWeb, basePortofPenetreatedPart); //oBaseTopoPortOfWeb.Normal;

                    // TopFlangeRightBottom and TopFlangeLeftBottom 
                    // Bottom Port && Top Port

                    PlatePart flangePlatePart;
                    ReadOnlyCollection<TopologyPort> flangeLateralPortsElems;

                    flangePlatePart = (PlatePart)flange;

                    flangeLateralPortsElems = flangePlatePart.GetPorts(TopologyGeometryType.Face,ContextTypes.Lateral , GeometryStage.Initial);
                   

                    if (flangeLateralPortsElems != null && flangeLateralPortsElems.Count > 0)
                    {
                        for (int i = 0; i <= flangeLateralPortsElems.Count-1; i++)
                        {
                            object flangePort = null;
                            flangePort = flangeLateralPortsElems[i];

                            if (flangePort is TopologyPort)
                            {
                                TopologyPort port = (TopologyPort)flangePort;
                                
                                double minimumDistance = 0;
                                Position posSrcPos = null;
                                Position PosInPos = null;
                                basePortofPenetreatedPart.DistanceBetween((ISurface)port.Geometry, out minimumDistance, out posSrcPos, out PosInPos);

                                if (minimumDistance < tolerance)
                                {

                                    Vector portNormal = CommonFuncs.GetNormalVectorOfTopologyPort(port, basePortofPenetreatedPart);//oPort.Normal;
                                    if (baseNormal.Dot(portNormal) > 0)
                                    {
                                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Left, port);
                                    }
                                    else if (baseNormal.Dot(portNormal) < 0)
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
            }


            return mappedPorts;
        }

        #endregion

    }
}