//-----------------------------------------------------------------------------
//      Copyright (C) 2011-16 Intergraph Corporation.  All rights reserved.
//
//      Component:  SectionUAPortMap to be implemented by all classes
//                  that are returning a EmulatedPort map according to the sectionAlias.
//
//      Author:  
//
//      History:
//      November 13, 2011       BSLee                  Created
//      Nov 19, 2015       mchandak               DI-275514 and DI-275249 Updated Slot Mapping Rule to avoid interop calls for GetNormalFromPostion() and IsExternalWire()
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
    /// The SectionUAPortMap class is designed to return emulated port maps based on UA and EA SectionAlias
    /// </summary>
    internal class SectionUAPortMap : IEmulatedPortMap
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
                throw new ArgumentNullException("Input oPenetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("Input oPenetratedPart");
            }

            if (basePltPort == null)
            {
                throw new ArgumentNullException("Input oPenetratedLateralEdgePort");
            }

            if (web == null)
            {
                throw new ArgumentNullException("Input oWeb");
            }


            basePlatePort = null;
            Dictionary<int, IPort> mappedPorts = null;


            mappedPorts = new Dictionary<int, IPort>();
            IPort basePortOfWeb = null;
            IPort offsetPortOfWeb = null;
            IPort bottomPortOfWeb = null;
            IPort topPortOfWeb = null;
            object rootBasePlateSystem = null;
            Collection<TopologyPort> MappedWLSidePortCollection = null;
            Collection<TopologyPort> MappedWRSidePortCollection = null;
            Collection<TopologyPort> MappedLateralSidePortCollection = null;

            //if it is Bend Plate case and Bend Plate has no Flange part connected 
            if ((flange == null) && (CommonFuncs.IsBendPlate(penetratingPart, penetratedPart, true, out MappedWLSidePortCollection, out MappedWRSidePortCollection, out MappedLateralSidePortCollection)))
            {
                basePlatePort = basePltPort;

                Position positionSrc;
                Position positionIn;
                IPort oMappedWLPort = null;
                IPort oMappedWRPort = null;
                IPort oMappedTopPort = null;
                IPort oMappedBtmPort = null;
                IPort oMappedTFRPort = null;
                IPort oMappedTFRBPort = null;

                //Based on distance from Base Plate to that of Mapped WebLeft Port Coll figure out which port is Web Left port and Top Port
                //nearmost one is Web Left and Farther one should be Top Port
                if (MappedWLSidePortCollection != null)
                {
                    double dMinDist = 10000;
                    double dMaxDist = -1;
                    foreach (TopologyPort oPort in MappedWLSidePortCollection)
                    {
                        double dPortDist = 0;
                        IPlane oPlane = (IPlane)oPort;
                        if ((oPlane != null) && (oPlane.RootPoint != null))
                        {
                            ISurface PortSurface = (ISurface)oPort;
                            PortSurface.DistanceBetween((ISurface)basePlatePort, out dPortDist, out positionSrc, out positionIn);

                            if (dPortDist < dMinDist)
                            {
                                dMinDist = dPortDist;
                                oMappedWLPort = oPort;
                            }

                            if (dPortDist > dMaxDist)
                            {
                                dMaxDist = dPortDist;
                                oMappedTopPort = oPort;
                            }
                        }
                    }
                }

                //Based on distance from Base Plate to that of Mapped WebRight Port Coll figure out which port is Web Right port and Top Flange Right Bottom
                //nearmost one is Web Right and Farther one should be Top Flange Right Bottom
                if (MappedWRSidePortCollection != null)
                {
                    double dMinDist = 10000;
                    double dMaxDist = -1;
                    foreach (TopologyPort oPort in MappedWRSidePortCollection)
                    {
                        double dPortDist = 0;
                        IPlane oPlane = (IPlane)oPort;
                        if ((oPlane != null) && (oPlane.RootPoint != null))
                        {
                            ISurface PortSurface = (ISurface)oPort;
                            PortSurface.DistanceBetween((ISurface)basePlatePort, out dPortDist, out positionSrc, out positionIn);
                            if (dPortDist < dMinDist)
                            {
                                dMinDist = dPortDist;
                                oMappedWRPort = oPort;
                            }

                            if (dPortDist > dMaxDist)
                            {
                                dMaxDist = dPortDist;
                                oMappedTFRBPort = oPort;
                            }
                        }
                    }
                }

                //Based on distance from Base Plate to that of Mapped Lateral Port Coll figure out which port is Top to that TopFlangeRight
                //nearmost one is Bottom and Farther one should be TopFlangeRight
                if (MappedLateralSidePortCollection != null)
                {
                    double dMinDist = 10000;
                    double dMaxDist = -1;
                    foreach (TopologyPort oPort in MappedLateralSidePortCollection)
                    {
                        double dPortDist = 0;
                        IPlane oPlane = (IPlane)oPort;
                        if ((oPlane != null) && (oPlane.RootPoint != null))
                        {
                            ISurface PortSurface = (ISurface)oPort;
                            PortSurface.DistanceBetween((ISurface)basePlatePort, out dPortDist, out positionSrc, out positionIn);
                            if (dPortDist < dMinDist)
                            {
                                dMinDist = dPortDist;
                                oMappedBtmPort = oPort;
                            }

                            if (dPortDist > dMaxDist)
                            {
                                dMaxDist = dPortDist;
                                oMappedTFRPort = oPort;
                            }
                        }
                    }
                }

                mappedPorts.Add((int)SectionFaceType.Web_Left, oMappedWLPort);
                mappedPorts.Add((int)SectionFaceType.Web_Right, oMappedWRPort);
                mappedPorts.Add((int)SectionFaceType.Top, oMappedTopPort);
                mappedPorts.Add((int)SectionFaceType.Bottom, oMappedBtmPort);
                mappedPorts.Add((int)SectionFaceType.Top_Flange_Right, oMappedTFRPort);
                mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Bottom, oMappedTFRBPort);
            }

            else
            {
                // Get Web and Base plate ports 
                CommonFuncs.GetWebAndBasePlatePorts(penetratingPart, penetratedPart, basePltPort, web, flange, secondWeb, secondFlange, out basePlatePort,
                    out basePortOfWeb, out offsetPortOfWeb, out bottomPortOfWeb, out topPortOfWeb, out rootBasePlateSystem);

                //    
                //                                  Top
                //                               ******************
                //                               *                * TopFlangeRight 
                //                               *    *************  
                //                               *    *  TopFlangeRightBottom
                //                       WebLeft *    *  
                //                               *    * WebRight 
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
                //Get BasePort of PenetratedPart to check Intersect

                TopologyPort basePortofPenetreatedPart = null;

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
                
                if (basePortOfFlange != null && basePortofPenetreatedPart != null)
                {
                    Collection<ICurve> curveCollection;
                    GeometryIntersectionType intersectionType;
                    TopologySurface baseSurfaceofPenetreatedPart = (TopologySurface)basePortofPenetreatedPart.Geometry;

                    ICurve intersectionCurve = null;
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
                    throw new ArgumentNullException("Input basePortOfFlange or basePortofPenetreatedPart");
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

     
                TopologyPort baseTopologyPortOfWeb = (TopologyPort)basePortOfWeb;
                TopologyPort offsetToPologyPortOfWeb = (TopologyPort)offsetPortOfWeb;

                double distBaseAndMiddlePoint = 0;
                double distOffsetAndMiddlePoint = 0;


                Position comPos1;
                Position comPos2;
                               
                Point3d point = new Point3d(middlePosition.X, middlePosition.Y, middlePosition.Z);
                baseTopologyPortOfWeb.DistanceBetween(point, out distBaseAndMiddlePoint, out comPos1);
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

                if (flange is PlatePart)
                {
                    IPort basePortOfFlangePlate = null;
                    IPort offsetPortOfFlangePlate = null;

                    basePortOfFlangePlate = CommonFuncs.GetBaseOrOffsetPortOfPlatePart(flange, true, penetratedPart);
                    offsetPortOfFlangePlate = CommonFuncs.GetBaseOrOffsetPortOfPlatePart(flange, false, penetratedPart);

                    TopologyPort baseTopoPortOfFlangePlate = (TopologyPort)basePortOfFlangePlate;
                    TopologyPort offsetTopoPortOfFlangePlate = (TopologyPort)offsetPortOfFlangePlate;
                    TopologyPort bottomTopoPortOfWeb = (TopologyPort)bottomPortOfWeb;

                    Vector bottomNormal = null, basePortNormal = null, offsetPortNormal = null;
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
                        basePortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(baseTopoPortOfFlangePlate, basePortofPenetreatedPart); //oBaseTopoPortOfFlangePlate.Normal;
                    }
                    else
                    {
                        throw new ArgumentNullException("Input basePortNormal");
                    }
                    if (offsetTopoPortOfFlangePlate != null)
                    {
                        offsetPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(offsetTopoPortOfFlangePlate, basePortofPenetreatedPart); //oOffsetTopoPortOfFlangePlate.Normal;
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
                            mappedPorts.Add((int)SectionFaceType.Top, offsetPortOfFlangePlate);
                            mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Bottom, basePortOfFlangePlate);
                        }
                        else
                        {
                            mappedPorts.Add((int)SectionFaceType.Top, basePortOfFlangePlate);
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
                    ReadOnlyCollection<TopologyPort> flangeLateralPortsElems;

                    flangePlatePart = (PlatePart)flange;

                    flangeLateralPortsElems = flangePlatePart.GetPorts(TopologyGeometryType.Face,ContextTypes.Lateral, GeometryStage.Initial);
                    

                    //Get BasePort of PenetratedPlate to check Intersect
                    if (flangeLateralPortsElems != null && flangeLateralPortsElems.Count > 0)
                    {
                        for (int i = 0; i <= flangeLateralPortsElems.Count-1; i++)
                        {
                            object flangePort = null;
                            flangePort = flangeLateralPortsElems[i];

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

                                    Vector portNormal = CommonFuncs.GetNormalVectorOfTopologyPort(port, basePortofPenetreatedPart);//oPort.Normal;
                                    if (webLeftNormal.Dot(portNormal) < 0)
                                    {
                                        mappedPorts.Add((int)SectionFaceType.Top_Flange_Right, port);
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
                        foreach (TopologyPort oPort in sectionFacesCol)
                        {
                            if (oPort.SectionId == (int)SectionFaceType.Top)
                            {
                                topPortOfFlange = (IPort)oPort;
                            }
                            else if (oPort.SectionId == (int)SectionFaceType.Bottom)
                            {
                                bottomPortOfFlange = (IPort)oPort;
                            }
                            else if (oPort.SectionId == (int)SectionFaceType.Web_Right)
                            {
                                webRightPortOfFlange = (IPort)oPort;
                            }
                            else if (oPort.SectionId == (int)SectionFaceType.Web_Left)
                            {
                                webLeftPortOfFlange = (IPort)oPort;
                            }
                            else
                            {

                            }
                        }
                    }

                    TopologyPort bottomTopoPortWeb = (TopologyPort)mappedPorts[(int)SectionFaceType.Bottom];
                    TopologyPort webRightTopoPortOfFlange = (TopologyPort)webRightPortOfFlange;
                    TopologyPort webLeftTopoPortOfFlange = (TopologyPort)webLeftPortOfFlange;

                Vector bottomNormal = CommonFuncs.GetNormalVectorOfTopologyPort(bottomTopoPortWeb, (TopologyPort)basePortOfWeb); //oBottomTopoPortWeb.Normal;
                Vector webRightPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(webRightTopoPortOfFlange,basePortofPenetreatedPart); //oWebRightTopoPortOfFlange.Normal;
                Vector webLefttPortNormal = CommonFuncs.GetNormalVectorOfTopologyPort(webLeftTopoPortOfFlange, basePortofPenetreatedPart); //oWebLeftTopoPortOfFlange.Normal;

                    double dotBottomAndWebRight = bottomNormal.Dot(webRightPortNormal);
                    double dotBottomAndWebLeft = bottomNormal.Dot(webLefttPortNormal);


                    if (dotBottomAndWebRight * dotBottomAndWebLeft < 0)
                    {
                        if (dotBottomAndWebRight > 0)
                        {
                            mappedPorts.Add((int)SectionFaceType.Top, webLeftPortOfFlange);
                            mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Bottom, webRightPortOfFlange);
                        }
                        else
                        {
                            mappedPorts.Add((int)SectionFaceType.Top, webRightPortOfFlange);
                            mappedPorts.Add((int)SectionFaceType.Top_Flange_Right_Bottom, webLeftPortOfFlange);
                        }
                    }
                    else if (dotBottomAndWebRight * dotBottomAndWebLeft > 0)
                    {
                        // TBD
                    }
                    else if ((dotBottomAndWebRight * dotBottomAndWebLeft).EqualTo(0) == true)
                    {
                        
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

                    }
                }
                else
                {

                }
            }

            return mappedPorts;
        }

        #endregion

    }
}
