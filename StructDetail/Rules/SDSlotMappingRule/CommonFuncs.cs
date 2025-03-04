//-----------------------------------------------------------------------------
//      Copyright (C) 2010-15 Intergraph Corporation.  All rights reserved.
//
//      Component:  CommonFuncs has methods which can be shared in common use to get 
//                  SectionAlias, EmulatedPorts, SketchingPlane     
//
//      Author:  
//
//      History:
//      November 16, 2011       BSLee                  Created
//	    September16, 2015	    RPK		               Added an enum to check for naming category
//      December 11, 2015       hgajula                DI-CP-284051  Fix coverity defects stated in November 6, 2015 report
//      December 11, 2015       knukala                DI-CP-275525  Replace PlaceIntersectionObject method in SlotMapng rule with .Net API 
//      Feburary 3, 2016        PYK                     DI-CP-287121  Fix coverity defects stated in January 15, 2016 report
//      April 12, 2016        dsmamidi                 TR-CP-236727   Modified GetWebAndBasePlatePorts() to support invert logical connections.
//      May 16, 2016            RPK                   TR-CP-294261   Modified the penetrated object type returned from GetRootSystemFromPart() to PlateSystemBase
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


namespace Ingr.SP3D.Content.Structure
{

    internal enum PenetrationType
    {
        UnknownPenetration = 0,
        TopPenetration = 1,
        BtmPenetration = 2,
        Mbrflngpenetration = 3,
    }
    internal enum PlateNamingCategory
    {
        WebPlate = 76,
    }
    internal enum PlateSubType
    {
        None = 0,
        Web = 1,
        Flange = 2,
        Corrugated = 3,
        Swedge = 4,

    }

    /// <summary>
    /// CommonFuncs has methods which can be shared in common use to get 
    /// SectionAlias, EmulatedPorts, SketchingPlane
    /// </summary>
    internal static class CommonFuncs
    {
        internal const string planeCOMinterface = "IJPlane";
        internal const int CTX_LATERAL_LFACE = 132;
        internal const double tolerance = 0.0001; //0.1 mm         
        /// <summary>
        /// Gets the Base or Offset Port of PlatePart
        /// </summary>
        /// <param name="platePartObj">Input Plate Part.</param>
        /// <param name="bBase">true means Base Port, false means Offset Port.</param>
        /// <param name="ConnectedObj">Connected to Input Plate to get port at connected Object</param>
        /// <returns>Returns Base or Offset Port according to bBase parameter.</returns>
        internal static IPort GetBaseOrOffsetPortOfPlatePart(object platePartObj, bool bBase, object ConnectedObj)
        {

            if (platePartObj == null)
            {
                throw new ArgumentNullException("Input PlatePartObj");
            }

            ReadOnlyCollection<TopologyPort> portCol = null;
            PlatePart platePart = (PlatePart)platePartObj;

            try
            {
                if (bBase == true)
                {
                    portCol = platePart.GetPorts(TopologyGeometryType.Face, ContextTypes.Base, GeometryStage.Initial);
                }
                else
                {
                    portCol = platePart.GetPorts(TopologyGeometryType.Face, ContextTypes.Offset,GeometryStage.Initial);

                }
            }
            catch(Exception)
            {
                return null;
            }

            if (portCol != null && portCol.Count == 1)
            {
                if (portCol[0] is IPort)
                {
                    return (IPort)portCol[0];
                }
            }
            else if (portCol != null && portCol.Count > 1 && ConnectedObj !=null)
            {
                object ConnectedSystem = null;

                if (ConnectedObj is PlatePart || ConnectedObj is ProfilePart )
                {
                    //Get Leaf System if connected is Part
                    ConnectedSystem = CommonFuncs.GetLeafSystemFromPart(ConnectedObj);
                }
                else if (ConnectedObj is ISystem)
                {
                    ConnectedSystem = ConnectedObj;
                }
                bool intersection = false;
                //Get the port Connected to Connected Sytem
                foreach (IPort port in portCol)
                {
                    //Get each port                   
                    TopologyPort platePort = (TopologyPort)port;

                    if (platePort.SectionId != -1)
                    {
                        if (ConnectedSystem != null)
                        {
                            //Get Intersection Between Port and Connected System
                            ReadOnlyCollection<TopologyPort> connectedObjBaseportCol = null;
                            if (ConnectedSystem is PlateSystem)
                            {
                                PlatePart connectedPlate = (PlatePart)CommonFuncs.GetPartFromRootSystem((ISystem)ConnectedSystem);
                                connectedObjBaseportCol = connectedPlate.GetPorts(TopologyGeometryType.Face, ContextTypes.Base, GeometryStage.Initial);
                            }
                            else if (ConnectedSystem is ProfileSystem)
                            {
                                ProfilePart connectedPlate = (ProfilePart)CommonFuncs.GetPartFromRootSystem((ISystem)ConnectedSystem);
                                connectedObjBaseportCol = connectedPlate.GetPorts(TopologyGeometryType.Face, ContextTypes.Lateral, GeometryStage.Initial);
                            }

                            double distance;
                            Position posSrcPos;
                            Position posInpos;
                            TopologyPort baseSinglePort;
                            if (connectedObjBaseportCol!= null && connectedObjBaseportCol.Count == 1)
                            {
                                baseSinglePort = connectedObjBaseportCol[0];
                                baseSinglePort.DistanceBetween((ISurface)platePort.Geometry, out distance, out posSrcPos, out posInpos);
                                if (distance < tolerance)
                                {
                                    intersection = true ;
                                }
                            }
                            else if (connectedObjBaseportCol != null && connectedObjBaseportCol.Count > 1)
                            {
                                foreach (TopologyPort basePort in connectedObjBaseportCol)
                                {
                                    //Avoid global port
                                    if (basePort.SectionId != -1)
                                    {
                                        basePort.DistanceBetween((ISurface)platePort.Geometry, out distance, out posSrcPos, out posInpos);
                                        if (distance < tolerance)
                                        {
                                            intersection = true ;
                                            break;
                                        }
                                    }
                                }
                            }

                            if (intersection == true)
                            {
                                // Return the port which is connected to Connected System and if not global port
                                return (IPort)port;
                            }
                        }
                    }
                }
            }
            else
            {
                // TBD

            }

            return null;
        }
        /// <summary>
        /// GetPartFromRootSystem, It returns only PlatePart, StiffenerPart, ER Part. 
        /// It returns only first Part of the SystemChildren collection 
        /// </summary>
        /// <param name="rootSystem">Root System  .</param>
        internal static object GetPartFromRootSystem(ISystem rootSystem)
        {
            if (rootSystem == null)
            {
                throw new ArgumentNullException("Input rootSystem");
            }

            ReadOnlyCollection<ISystemChild> childCollection = rootSystem.SystemChildren;
            List<PlatePart> plateParts = new List<PlatePart>();

            //this collection is defined because we are adding elements dynamically 
            Collection<ISystemChild> childColl = new Collection<ISystemChild>();

            if (childCollection != null)
            {
                if (childCollection.Count > 0)
                {
                    //adding ReadOnly collection elements to the ordinary collection
                    for (int j = 0; j < childCollection.Count; j++)
                    {
                        childColl.Add(childCollection[j]);
                    }

                    //Traversing the each object in the collection
                    for (int i = 0; i < childColl.Count; i++)
                    {
                        //check if the item at this index is PlatePart or not
                        if (typeof(PlatePart) == childColl[i].GetType() || typeof(StiffenerPart) == childColl[i].GetType() || typeof(EdgeReinforcementPart) == childColl[i].GetType())
                        {
                            //If it is PlatePart then add to the PlatePart collecton
                            return ((object)childColl[i]);
                        }
                        //under Root system,we have leaf plate system and also chance of having another Child System 
                        else if (typeof(PlateSystem) == childColl[i].GetType() || typeof(StiffenerSystem) == childColl[i].GetType() || typeof(EdgeReinforcementSystem) == childColl[i].GetType())
                        {
                            //if it is a leaf or another system get the children of it
                            ISystem leafSystem = (ISystem)childColl[i];
                            ReadOnlyCollection<ISystemChild> nextLevelChildColl = leafSystem.SystemChildren;
                            if (nextLevelChildColl != null)
                            {
                                if (nextLevelChildColl.Count > 0)
                                {
                                    //If it is LeafSystem then add to the collecton
                                    for (int j = 0; j < nextLevelChildColl.Count; j++)
                                    {
                                        //add this collection to the main collection
                                        childColl.Add(nextLevelChildColl[j]);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return null;
        }
        /// <summary>
        /// Check if this is detailed part of not 
        /// </summary>
        /// <param name="part">Input Part.</param>
        /// <returns>Returns true, false .</returns>
        internal static bool IsDetailedPart(object part)
        {
            if (part == null)
            {
                throw new ArgumentNullException("Input Part");
            }

            bool bIsDetailedPart = false;

            if (part is PlatePart)
            {
                if (((PlatePart)part).PartGeometryState == PartGeometryStateType.DetailedPart)
                {
                    bIsDetailedPart = true;
                }
            }
            else if (part is StiffenerPart)
            {
                if (((StiffenerPart)part).PartGeometryState == PartGeometryStateType.DetailedPart)
                {
                    bIsDetailedPart = true;
                }
            }
            else if (part is EdgeReinforcementPart)
            {
                if (((EdgeReinforcementPart)part).PartGeometryState == PartGeometryStateType.DetailedPart)
                {
                    bIsDetailedPart = true;
                }
            }
            else
            {
            }

            return bIsDetailedPart;
        }


        /// <summary>
        /// GetNormalVectorOfTopologyPort
        /// In this method, Check it is Planar or not and get the Normal Vector
        /// </summary>
        /// <param name="inputSurface">Input Surface.</param>
        /// <param name="intersectingPort">Neighbor Port which intersect input surface.</param>
        /// <returns>Returns true, false .</returns>
        internal static Vector GetNormalVectorOfTopologyPort(TopologyPort inputSurface, TopologyPort intersectingPort)
        {

            if (inputSurface == null)
            {
                throw new ArgumentNullException("Input inputSurface");
            }

            if (intersectingPort == null)
            {
                throw new ArgumentNullException("Input intersectingPort");
            }

            bool inputPlanarSurface = inputSurface.SupportsInterface(planeCOMinterface);

            if (inputPlanarSurface)
            {
                return inputSurface.Normal;
            }
            else
            {
                Position posOnSurface = CommonFuncs.GetPosOnIntersectCurve(inputSurface, intersectingPort);
                if (posOnSurface != null)
                {
                    return inputSurface.OutwardNormalAtPoint(posOnSurface);
                }

            }
            return null;
        }

        /// <summary>
        /// GetPosOnIntersectCurve. This Function returns middle position of intersected curve  
        /// </summary>
        /// <param name="surface1">Input surface1 .</param>
        /// <param name="surface2">Input surface2 .</param> 
        /// <returns> Point of Intersected curve .</returns>
        internal static Position GetPosOnIntersectCurve(object surface1, object surface2)
        {

            if (surface1 == null)
            {
                throw new ArgumentNullException("Input surface1");
            }

            if (surface2 == null)
            {
                throw new ArgumentNullException("Input surface2");
            }
            ICurve intersectionCurve = null;
            TopologyPort tempPort1 = (TopologyPort)surface1;
            TopologyPort tempPort2 = (TopologyPort)surface2;
            ISurface tempSurface1 = (ISurface)tempPort1.Geometry;
            Collection<ICurve> curveCollection;
            GeometryIntersectionType intersectionType;
            Position startPos = null;
            Position endPos = null;
            tempSurface1.Intersect((ISurface)tempPort2.Geometry, out curveCollection, out intersectionType);
            if (curveCollection != null)
            {
                if (curveCollection.Count != 0)
                {
                    intersectionCurve = (ICurve)curveCollection[0];
                    intersectionCurve.EndPoints(out startPos, out endPos);
                    return startPos;
                }
            }
            return null;
        }

        /// <summary>
        /// GetAugmentedRangeBoxOfIntersectCurve. This Function returns Rangebox of intersection curve augmented by a distance  
        /// </summary>
        /// <param name="surface1">Input surface1 .</param>
        /// <param name="surface2">Input surface2 .</param> 
        /// <param name="surface2">Input distance </param>  
        /// <returns> Rangebox of intersection curve augmented by distance value along tangents at the two ends.</returns>
        internal static RangeBox GetAugmentedRangeBoxOfIntersectCurve(object surface1, object surface2, double distance)
        {
            RangeBox intersectRangeBox = null;
            if (surface1 == null)
            {
                throw new ArgumentNullException("Input surface1");
            }

            if (surface2 == null)
            {
                throw new ArgumentNullException("Input surface2");
            }
            TopologyPort tempPort1 = (TopologyPort)surface1;
            TopologyPort tempPort2 = (TopologyPort)surface2;
            ISurface tempSurface1 = (ISurface)tempPort1.Geometry;
            Collection<ICurve> curveCollection;
            GeometryIntersectionType intersectionType;
            tempSurface1.Intersect((ISurface)tempPort2.Geometry, out curveCollection, out intersectionType);
            if (curveCollection != null)
            {
                TopologyCurve compositeCurve = new TopologyCurve(curveCollection);
                intersectRangeBox = compositeCurve.Range;
                if (distance > tolerance)
                {
                    //Augment the intersection RangeBox by distance value along the curve tangents
                    Position startPos = null;
                    Position endPos = null;

                    compositeCurve.EndPoints(out startPos, out endPos);
                    Vector tangentVecAtStart = compositeCurve.TangentAtPoint(startPos);
                    tangentVecAtStart.Length = -distance;
                    startPos = startPos.Offset(tangentVecAtStart);

                    Vector tangentVecAtEnd = compositeCurve.TangentAtPoint(endPos);
                    tangentVecAtEnd.Length = distance;
                    endPos = endPos.Offset(tangentVecAtEnd);

                    RangeBox augmentedRangeBox = (new RangeBox(startPos, endPos)) + intersectRangeBox;

                    return augmentedRangeBox;
                }
                else
                {
                    return intersectRangeBox;
                }
                
                
            }
            else
            {
                return null;
            }

        }


        /// <summary>
        /// GetLogicalConnections on RootPlateSystem 
        /// </summary>
        /// <param name="parentSystem">Input Plate System .</param>
        /// <returns> ReadOnlyCollection<IConnection> which contains Logical Connections .</returns>
        internal static ReadOnlyCollection<IConnection> GetLogicalConnections(ISystem parentSystem)
        {

            if (parentSystem == null)
            {
                throw new ArgumentNullException("Input parentSystem");
            }
            
            List<IConnection> listLCs = new List<IConnection>();
            try
            {
                //go through the list of children on this plate and
                //build a collection of the LC objects contained in 
                //this plate so we can return that collection to the user
                foreach (ISystemChild child in parentSystem.SystemChildren)
                {
                    if (child.GetType() == typeof(LogicalConnection))
                    {
                        listLCs.Add((IConnection)child);
                    }
                }
            }
            catch (Exception)
            {
            }

            return new ReadOnlyCollection<IConnection>(listLCs);
        }


        // This method is for deciding whether the Plate is inserted into Top flange or Bottom Flange or both the flanges
        // Enum values are given for identifying the above mentioned three cases
        internal static PenetrationType GetInsertPlateType(object penetratingPart, object penetratedPart)
        {
            if (penetratingPart == null)
            {
                throw new ArgumentNullException("PenetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("PenetratedPart");
            }

            MemberPart oMbrPart = penetratedPart as MemberPart;
            PenetrationType PlateIntersection;
            if (oMbrPart != null)
            {
                PlatePart oPlatePart = penetratingPart as PlatePart;
                if (oPlatePart != null)
                {                    
                    //Member Insert Plate Case
                    bool topFlangeResultantIntersection = false;
                    bool bottomFlangeResultantIntersection = false;

                    TopologyPort MbrTopPort = null;
                    ReadOnlyCollection<TopologyPort> MbrPortsCollection = oMbrPart.GetPorts(TopologyGeometryType.Face, ContextTypes.Lateral, GeometryStage.Initial);

                    //Tolerance value is the distance between Plate and the member port
                    Double dTolerance = 0.001;
                    //Member Top Port
                    foreach (TopologyPort oMbrPort in MbrPortsCollection)
                    {
                        if (oMbrPort.SectionId == (int) SectionFaceType.Top)
                        {
                            MbrTopPort = oMbrPort;
                            break;
                        }
                    }           

                    if (MbrTopPort != null)
                    {
                        TopologyPort oPlateBasePort = oPlatePart.GetPort(TopologyGeometryType.Face, -1, -1, ContextTypes.Base,  -1, false);

                        // Check if the distance between the plate port and member is less than tolerance value
                        double dBaseMinDist;
                        Position positionOnMbrPort;
                        Position positionOnPlate;
                        if (oPlateBasePort != null)
                        {
                            //Get distance between plate base port and member top port
                            MbrTopPort.DistanceBetween((ISurface)oPlateBasePort, 
                                            out dBaseMinDist, out positionOnMbrPort, out positionOnPlate);

                            //Check the distance is within the tolerance and plate is parallel to member port
                            if ((dBaseMinDist < dTolerance || System.Math.Abs(dBaseMinDist - oPlatePart.Thickness) <  dTolerance))                                    
                            {
                                Vector mbrTopPortNormal = MbrTopPort.OutwardNormalAtPoint(positionOnMbrPort);
                                Vector plateBasePortNormal = oPlateBasePort.OutwardNormalAtPoint(positionOnPlate);
                                double dotProductAbs = System.Math.Abs(mbrTopPortNormal.Dot(plateBasePortNormal));
                                if (Math3DExtensions.EqualTo(dotProductAbs, 1.0, dTolerance))
                                    topFlangeResultantIntersection = true;
                            }
                        }

                        //If the penetrating plate intersects both the Flanges(Top and Bottom), then it is a MbrFlangePenetration case
                        //else if it intersects only with Top then TopInsertPlate case otherwise BottomInsertPlate case
                        if (topFlangeResultantIntersection == false )
                        {
                            PlateIntersection = PenetrationType.BtmPenetration;
                        }
                        else
                        {
                            TopologyPort MbrBtmPort = null;
                            foreach (TopologyPort oMbrPort in MbrPortsCollection)
                            {
                                if (oMbrPort.SectionId == (int)SectionFaceType.Bottom)
                                {
                                    MbrBtmPort = oMbrPort;
                                    break;
                                }
                            }

                            if (MbrBtmPort != null)
                            {
                                //Get distance between plate base port and member btm port
                                MbrBtmPort.DistanceBetween((ISurface)oPlateBasePort,
                                                out dBaseMinDist, out positionOnMbrPort, out positionOnPlate);

                                //Check the distance is within the tolerance and plate is parallel to member port
                                if ((dBaseMinDist < dTolerance || System.Math.Abs(dBaseMinDist - oPlatePart.Thickness) < dTolerance))
                                {
                                    Vector mbrBtmPortNormal = MbrBtmPort.OutwardNormalAtPoint(positionOnMbrPort);
                                    Vector plateBasePortNormal = oPlateBasePort.OutwardNormalAtPoint(positionOnPlate);
                                    double dotProductAbs = System.Math.Abs(mbrBtmPortNormal.Dot(plateBasePortNormal));
                                    if (Math3DExtensions.EqualTo(dotProductAbs, 1.0, dTolerance))
                                        bottomFlangeResultantIntersection = true;
                                }

                                if (bottomFlangeResultantIntersection == false)
                                {
                                    PlateIntersection =  PenetrationType.TopPenetration;
                                }
                                else
                                {
                                    PlateIntersection = PenetrationType.Mbrflngpenetration;
                                }
                            }
                            else
                            {
                                PlateIntersection = PenetrationType.UnknownPenetration;
                            }
                        }
                    }
                    else
                    {
                        throw new ArgumentNullException("Input MbrTopPort ");
                    }                  
                }
                else 
                { 
                    PlateIntersection = PenetrationType.UnknownPenetration;
                }
            }
            else
            { 
                PlateIntersection =  PenetrationType.UnknownPenetration;
            }
            return PlateIntersection;
        }
 
        /// <summary>
        /// Get RootSystem from Part. PlatePart or StiffenerPart..  
        /// </summary>
        /// <param name="part">Input Part.</param>
        /// <returns> Returns RootSystem Object. </returns>
        internal static object GetRootSystemFromPart(object part)
        {
            if (part == null)
            {
                throw new ArgumentNullException("Input Part");
            }
            
            ISystemChild leafSystem = null;
            object obj = null;

            try
            {
                ISystemChild partChild = (ISystemChild)part;
                leafSystem = (ISystemChild)partChild.SystemParent;
             
                if (leafSystem != null)
                {
                    obj= (object)leafSystem.SystemParent;
                }
            }
            catch (Exception)
            {

            }
            return obj;

        }

        
        /// <summary>
        /// Get LeafSystem from Part. PlatePart or StiffenerPart.
        /// </summary>
        /// <param name="part">Input Part.</param>
        /// <returns> Returns LeafSystem Object. </returns>
        internal static object GetLeafSystemFromPart(object part)
        {
            if (part == null)
            {
                throw new ArgumentNullException("Input Part");
            }
            
            ISystemChild partChild = null;
            try
            {
                partChild = (ISystemChild)part;
            }
            catch (Exception)
            {
            }
            return (object)partChild.SystemParent;

        }

        /// <summary>
        /// Check Mutual Bounding case and return true or false. Inputs should be PlateSystem 
        /// </summary>
        /// <param name="plateRootSystem1">1stPlateRootSystem.</param>
        /// <param name="plateRootSystem2">2ndPlateRootSystem.</param>
        /// <returns> Returns LeafSystem Object. </returns>
        internal static bool IsMutualBounding(PlateSystem plateRootSystem1, PlateSystem plateRootSystem2)
        {

            if (plateRootSystem1 == null)
            {
                throw new ArgumentNullException("Input 1stPlateRootSystem");
            }


            if (plateRootSystem2 == null)
            {
                throw new ArgumentNullException("Input 2ndPlateRootSystem");
            }

            Collection<Object> boundaryCol1 = null;
            Collection<Object> boundaryCol2 = null;

            try
            {
                boundaryCol1 = plateRootSystem1.Boundaries;
                boundaryCol2 = plateRootSystem2.Boundaries;
            }
            catch (Exception)
            {
                return false;
            }
            

            if (boundaryCol1.Contains((BusinessObject)plateRootSystem2) == true && boundaryCol2.Contains((BusinessObject)plateRootSystem1) == true)
            {
                return true;
            }
            else
            {
                return false;
            }

        }



        /// <summary>
        /// Get Intersecting Seam between the PenetratedPart and PenetratingPart   
        /// </summary>
        /// <param name="penetratingPart">Input PnetratingPlatePart .</param>
        /// <param name="penetratedtPart">Input PenetratedPlatePart .</param>
        /// <param name="outputSeam">Output Seam .</param>
        internal static void GetSeamBtwObjects(object penetratingPart, object penetratedPart, out Seam outputSeam)
        {

            if (penetratingPart == null)
            {
                throw new ArgumentNullException("Input PenetratingPart ");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("Input PenetratedPart ");
            }

            outputSeam = null;


            //Make this as Function 
            object rootPenetratingSystem = GetRootSystemFromPart(penetratingPart);
            ISplit split = null;
            Seam seam = null;

            if (penetratedPart is PlatePart)
            {
                PlateSystemBase rootPenetratedPlateSystem = (PlateSystemBase)GetRootSystemFromPart(penetratedPart);
                if (rootPenetratedPlateSystem != null)
                {
                    split = rootPenetratedPlateSystem.GetSplit((BusinessObject)rootPenetratingSystem);
                    seam = (Seam)split;

                    if ((seam != null) && (seam.SeamType == SeamType.Intersection))
                    {
                        outputSeam = seam;
                    }
                }
            }
            else if (penetratedPart is StiffenerPart)
            {
                StiffenerSystem rootPenetratedStiffenerSystem = (StiffenerSystem)GetRootSystemFromPart(penetratedPart);
                if (rootPenetratedStiffenerSystem != null)
                {
                    split = rootPenetratedStiffenerSystem.GetSplit((BusinessObject)rootPenetratingSystem);
                    seam = (Seam)split;

                    if ((seam != null) && (seam.SeamType == SeamType.Intersection))
                    {
                        outputSeam = seam;
                    }
                }
            }

            else
            {

            }

        }
 
        /// <summary>        
        ///1.Get BasePlateRootSystem    
        /// 2.Get Base, Offset, Top, Bottom Port, Web PlatePart
        ///      - Bottom Port is a port which connected to BasePlate
        /// </summary>
        /// <param name="penetratingPart">Input PenetratingPart.</param>
        /// <param name="penetratedPart">Input Other PenetratingPart .</param>
        /// <param name="basePltPort">Input BasePlatePort.</param>
        /// <param name="web">Input PenetratingPart.</param>
        /// <param name="flange">Input Other PenetratingPart .</param>
        /// <param name="secondWeb">Input PenetratedPart.</param>
        /// <param name="secondFlange">Input PenetratingPart.</param>
        /// <param name="basePlatePort">Input Other PenetratingPart .</param>
        /// <param name="basePortOfWeb">Input PenetratedPart.</param>
        /// <param name="offsetPortOfWeb">Input PenetratingConfiguration .</param>
        /// <param name="bottomPortOfWeb">Output WebPart .</param>
        /// <param name="topPortOfWeb">Input PenetratingConfiguration .</param>
        /// <param name="rootBasePlateSystem">Output WebPart .</param>

        internal static void GetWebAndBasePlatePorts(object penetratingPart, object penetratedPart, IPort basePltPort, object web, object flange, object secondWeb, object secondFlange,
            out IPort basePlatePort, out IPort basePortOfWeb, out IPort offsetPortOfWeb, out IPort bottomPortOfWeb, out IPort topPortOfWeb, out object rootBasePlateSystem)
        {

            if (penetratingPart == null)
            {
                throw new ArgumentNullException("Input penetratingPart ");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("Input penetratedPart ");
            }

            if (basePltPort == null)
            {
                throw new ArgumentNullException("Input penetratedLateralEdgePort ");
            }

            if (web == null)
            {
                throw new ArgumentNullException("Input web ");
            }


            basePlatePort = null;
            basePortOfWeb = null;
            offsetPortOfWeb = null;
            bottomPortOfWeb = null;
            topPortOfWeb = null;
            rootBasePlateSystem = null;
            
            PlatePart webPlatePart = (PlatePart)web;

            //Assign basePlatePort input argument to output argument
            basePlatePort = basePltPort;

            //Get RootBasePlateSystem
            rootBasePlateSystem = GetRootSystemFromPart((object)basePltPort.Connectable);
   
            ReadOnlyCollection<TopologyPort> lateralPortsElems;

            lateralPortsElems = webPlatePart.GetPorts(TopologyGeometryType.Face,ContextTypes.Lateral, GeometryStage.Initial);

            // BasePort
            basePortOfWeb = CommonFuncs.GetBaseOrOffsetPortOfPlatePart(webPlatePart, true, penetratedPart);
            PlateSystem rootWebSystem = (PlateSystem)GetRootSystemFromPart(web);

            //Offset Port
            offsetPortOfWeb = CommonFuncs.GetBaseOrOffsetPortOfPlatePart(webPlatePart, false, penetratedPart);
                                        
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
            //Get the intersection range box between web base port and penetrated base port.
            RangeBox intersectionRangeBox = null;
            double flangePlateThickness = 0;
            if (flange != null &&  flange is PlatePart)
            {
                PlatePart flangePlate = (PlatePart)flange;
                flangePlateThickness = flangePlate.Thickness;
            }

            if (basePortOfWeb!=null && basePortofPenetreatedPart !=null)
            {
                intersectionRangeBox = GetAugmentedRangeBoxOfIntersectCurve(basePortOfWeb, basePortofPenetreatedPart, flangePlateThickness);
            }
            else
            {
                throw new Exception("basePortOfWeb and basePortofPenetratedPart should not be null");

            }

            PlatePart basePlatePart = (PlatePart)basePltPort.Connectable;

            Collection<TopologyPort> secondCheckPortsCol = new Collection<TopologyPort>();
            foreach (TopologyPort port in lateralPortsElems)
            {
                bool consideredThisPort = false;
                if (port.OperationId != -1) //Check to avoid global lateral port of web plate 
                {
                    TopologyPort portToBeConsidered = port.GetRelatedPort(GeometryOperationTypes.PartFinalTrim, GraphPosition.After, false);
                    if (portToBeConsidered == null)
                    {
                        portToBeConsidered = port;
                    }

                    RangeBox webPortRange = portToBeConsidered.Range;
                  
                        //Avoid, if there is no overlap between current webport range box and intersection range box
                        if (webPortRange.Intersects(intersectionRangeBox) == false)
                            continue;
                    
                    //overlap exists, so get the web port connections and check if the other connectable part is base plate
                        ReadOnlyCollection<IConnection> connectionCol = portToBeConsidered.Connections;
                    foreach (Connection leafConn in connectionCol)
                    {
                        ConnectionData webPlateConnData = new ConnectionData(leafConn, webPlatePart);
                        IConnectable connectedToWebPart = webPlateConnData.ToConnectable;
                        if (bottomPortOfWeb == null && basePlatePart.Equals(connectedToWebPart))
                        {
                            //base plate found, it is bottom port of web
                            bottomPortOfWeb = portToBeConsidered;
                            consideredThisPort = true;
                            if (flange == null)
                                break;
                        }
                        else if (topPortOfWeb == null && flange != null && flange.Equals(connectedToWebPart))
                        {
                            //other connectable is flange,it is top port of web
                            topPortOfWeb = portToBeConsidered;
                            consideredThisPort = true;
                        }
                        else
                        {
                            //For BUTL3 cross-section, top port of web is not connected to flange part, so second check is needed to find the distance
                            secondCheckPortsCol.Add(portToBeConsidered);
                            consideredThisPort = true;
                        }

                        if (bottomPortOfWeb != null && topPortOfWeb != null)
                            break;
                    } 

                    if (topPortOfWeb != null && flange == null)
                        break;
                    if (consideredThisPort == false)
                        secondCheckPortsCol.Add(portToBeConsidered);
                }
            }
            
            if (topPortOfWeb == null || bottomPortOfWeb == null)
            {
                double platePartThickness = basePlatePart.Thickness;
                Position positionOnWebPort;
                Position positionOnPenetreatedPart;
                double distance;
                foreach (TopologyPort port in secondCheckPortsCol)
                {
                    if (bottomPortOfWeb == null)
                    {
                        //If base plate is not detailed, then bottom port will not be filled up to this point
                        port.DistanceBetween((ISurface)basePlatePort, out distance, out  positionOnWebPort, out positionOnPenetreatedPart);
                        if (distance < tolerance || distance.EqualTo(platePartThickness))
                        {
                            bottomPortOfWeb = port;
                            continue;
                        }
                    }
                    if (topPortOfWeb == null)
                    {
                        port.DistanceBetween((ISurface)basePortofPenetreatedPart, out distance, out  positionOnWebPort, out positionOnPenetreatedPart);
                        if (distance < tolerance)
                        {
                            topPortOfWeb = port; 
                        } 
                    }
                    if (topPortOfWeb != null && bottomPortOfWeb != null)
                    {
                        break;
                    }
                }
            } 
        }

        /// <summary>
        /// IsBendPlate() --> Determines whether the input penetrating part is a Bend Plate case
        ///                   If it is Bend Plate case (for e.g. a Linear Extruded plate having mother curve defintion
        ///                   eqvivalent to UA shape )
        ///                   In such cases send out Collection of Plate Ports on the Web Left side, and also on Web right side
        ///                   only if asked for via flag bNeedMappedCol
        /// </summary>
        /// <param name="penetratingPart">Input PnetratingPlatePart .</param>
        /// <param name="penetratedtPart">Input PenetratedPlatePart .</param>
        /// <param name="bNeedMappedCol">Input bNeedMappedCol .</param>
        /// <param name="MappedWLSidePortColl">Ouput Collection of Outer side ports of Linear Extruded Plate(Bend Plate).</param>
        /// <param name="MappedWRSidePortColl">Ouput Collection of Inner side ports of Linear Extruded Plate(Bend Plate) .</param>
        /// <param name="MappedLateralPortColl">Ouput Collection of Lateral ports of Linear Extruded Plate(Bend Plate).</param>
        /// <returns> Returns Boolean value. </returns>
        internal static bool IsBendPlate(object penetratingPart, object penetratedPart, bool bNeedMappedCol, out Collection<TopologyPort> MappedWLSidePortColl, out Collection<TopologyPort> MappedWRSidePortColl, out Collection<TopologyPort> MappedLateralPortColl)
        {

            if (penetratingPart == null)
            {
                throw new ArgumentNullException("Input penetratingPart ");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("Input penetratedPart ");
            }

            bool IsBendPlate = false;
            MappedLateralPortColl = null;
            MappedWRSidePortColl = null;
            MappedWLSidePortColl = null;

            if ((penetratingPart is PlatePart) && (penetratedPart is PlatePart))
            {
                
                //Get All the Base Port Collection of the Penetrating Part
                PlatePart PenetratingPart = (PlatePart)penetratingPart;
                PlatePart PenetratedPart = (PlatePart)penetratedPart;
                ReadOnlyCollection<TopologyPort> CollBasePort = null;
                Collection<TopologyPort> CollofIntersectedBasePorts = new Collection<TopologyPort>();
                Collection<TopologyPort> CollofIntersectedOffsetPorts = new Collection<TopologyPort>();
                Collection<TopologyPort> CollofIntersectedLateralPorts = new Collection<TopologyPort>();
                Collection<ICurve> CollofBaseNonLinearCurves = new Collection<ICurve>();

                Collection<Vector> CollOfNormalVectors = new Collection<Vector>();
                Plane3d ApproxSketchingPlane = null;

                CollBasePort = PenetratingPart.GetPorts(TopologyGeometryType.Face, ContextTypes.Base, GeometryStage.Initial);

                //Get Approx Sketch Plane
                ApproxSketchingPlane = GetApproxSketchPlane(penetratingPart, penetratedPart);

                int LinearBaseCurveCount=0;
                int NonLinearBaseCurveCount = 0;
                
                foreach (TopologyPort oPort in CollBasePort)
                {   
                    GeometryIntersectionType eIntersectionType;
                    Collection<ICurve> colCurves;
                    ISurface oPortSurface = (ISurface)oPort;

                    ApproxSketchingPlane.Intersect(oPortSurface, out colCurves, out eIntersectionType);

                    SurfaceScopeType eScope;
                    Vector vecNormal;
                    oPortSurface.ScopeNormal(out eScope, out vecNormal);
                    
                    //Ignore global Port cases based on section id's as we need only individual sub ports intersecting with sketching plane
                    if ((colCurves != null) && (colCurves.Count > 0) && (oPort.SectionId != -1))
                    {
                        CollofIntersectedBasePorts.Add(oPort);

                        if (eScope == SurfaceScopeType.Planar)
                        {
                            LinearBaseCurveCount++;
                            CollOfNormalVectors.Add(vecNormal);
                        }
                        else if (colCurves[0] != null)
                        {
                            CollofBaseNonLinearCurves.Add(colCurves[0]);
                            NonLinearBaseCurveCount++;
                        }
                    }
                }

                //Based on Curves count we can decide what type of cross section bend Plate could be
                // for example for UA bend plate case has 2 linear and 1 non linear.
                // if we have to support new section alias (for exapmle "C"), new else case needs
                // to be added based on curve count
                if (LinearBaseCurveCount == 2 && NonLinearBaseCurveCount == 1)
                {
                    // It is a BendPlate Case with UA as cross section(section alias)
                    // Currently as of now only Bend of around 45 to 135 degreees is supported.

                    Vector vecFirstLine = new Vector(CollOfNormalVectors[0]);
                    Vector vecSecondLine = new Vector(CollOfNormalVectors[1]);
                    Double PI = Math.PI;

                    //Get the angle between two vectors
                    double AngleBetweenVectors = vecFirstLine.Angle(vecSecondLine, vecFirstLine.Cross(vecSecondLine));
                    if (AngleBetweenVectors > PI)
                    {
                        AngleBetweenVectors = (2 * PI) - AngleBetweenVectors;
                    }
                    double AngleBetweenLinearCurves = Math.Round(PI - AngleBetweenVectors,4);

                    //Check for angular Tolerance
                    if ((AngleBetweenLinearCurves >= (PI / 4)) || (AngleBetweenLinearCurves <= (3 * (PI / 4))))
                    {
                        IsBendPlate = true;

                        if ((IsBendPlate) && (bNeedMappedCol))
                        {
                            //In this case we have we are sure it is bend plate of UA cross section
                            //and method needs mapped collection. So we try to get only needed mapped
                            //Offset Ports and Lateral ports and return them
                            ReadOnlyCollection<TopologyPort> CollOffsetPort = null;
                            CollOffsetPort = PenetratingPart.GetPorts(TopologyGeometryType.Face, ContextTypes.Offset, GeometryStage.Initial);
                            ICurve OffsetNonLinearCurve = null;

                           
                            foreach (TopologyPort oPort in CollOffsetPort)
                            {
                                GeometryIntersectionType eIntersectionType;
                                Collection<ICurve> colCurves;
                                ISurface oPortSurface = (ISurface)oPort;

                                ApproxSketchingPlane.Intersect(oPortSurface, out colCurves, out eIntersectionType);

                                SurfaceScopeType eScope;
                                Vector vecNormal;
                                oPortSurface.ScopeNormal(out eScope, out vecNormal);

                                //Ignore global Port cases based on section id's as we need only individual sub ports intersecting with sketching plane
                                if ((colCurves != null) && (colCurves.Count > 0) && (oPort.SectionId != -1))
                                {
                                    CollofIntersectedOffsetPorts.Add(oPort);
                                    
                                    if (!(eScope == SurfaceScopeType.Planar) && (colCurves.Count ==1)) 
                                    {   
                                        //We strictly expect only one Non linear curve to be found(for UA) in this for loop
                                        //if found more than one this code has to be enhaneced.
                                        OffsetNonLinearCurve = colCurves[0];
                                    }
                                }
                            }

                            //Similar to the above Offset Ports also we need to mapp Lateral Ports and return

                            //Based on Non Linear Curve Length we can decide if Base Ports is on Web Left side
                            //Or if the OffsetPorts is on Web left
                            ReadOnlyCollection<TopologyPort> CollLateralPort = null;
                            CollLateralPort = PenetratingPart.GetPorts(TopologyGeometryType.Face, ContextTypes.Lateral, GeometryStage.Initial);

                            
                            foreach (TopologyPort oPort in CollLateralPort)
                            {
                                GeometryIntersectionType eIntersectionType;
                                Collection<ICurve> colCurves;
                                ISurface oPortSurface = (ISurface)oPort;

                                ApproxSketchingPlane.Intersect(oPortSurface, out colCurves, out eIntersectionType);

                                SurfaceScopeType eScope;
                                Vector vecNormal;
                                oPortSurface.ScopeNormal(out eScope, out vecNormal);

                                //Ignore global Port cases based on section id's/Operator Id as we need only individual sub ports intersecting with sketching plane
                                if ((colCurves != null) && (colCurves.Count == 1) && (oPort.OperatorId != -1))
                                {
                                    CollofIntersectedLateralPorts.Add(oPort);
                                }
                            }


                            if (Math.Round(OffsetNonLinearCurve.Length,4) > Math.Round(CollofBaseNonLinearCurves[0].Length,4))
                            {
                                MappedWLSidePortColl = CollofIntersectedOffsetPorts;
                                MappedWRSidePortColl = CollofIntersectedBasePorts;
                                MappedLateralPortColl = CollofIntersectedLateralPorts;

                            }
                            else
                            {
                                MappedWLSidePortColl = CollofIntersectedBasePorts;
                                MappedWRSidePortColl = CollofIntersectedOffsetPorts;
                                MappedLateralPortColl = CollofIntersectedLateralPorts;
                            }

                        }                      
                    }
                }
            }

            return IsBendPlate;
        }

        /// <summary>
        /// GetApproxSketchPlane() --> Gets approximated Sketch plane
        /// </summary>
        /// <param name="penetratingPart">Input PnetratingPlatePart .</param>
        /// <param name="penetratedtPart">Input PenetratedPlatePart .</param>
        /// <returns> Returns Plane3d  Sketch Plane (approximation). </returns>
        internal static Plane3d GetApproxSketchPlane(object penetratingPart, object penetratedPart)
        {

            if (penetratingPart == null)
            {
                throw new ArgumentNullException("Input penetratingPart ");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("Input penetratedPart ");
            }

            Plane3d ApproxPlane=null;

            if (penetratedPart is PlatePart)
            {
                PlatePart PenetratedPart = (PlatePart)penetratedPart;
                TopologyPort PtdBasePort;

                PtdBasePort = PenetratedPart.GetPort(TopologyGeometryType.Face, ContextTypes.Base);
                IPlane oBasePlane = PtdBasePort as IPlane;

                if ((oBasePlane != null) && (oBasePlane.RootPoint != null))
                {
                    //For Planar Surface
                    ApproxPlane = new Plane3d(oBasePlane.RootPoint, oBasePlane.Normal);
                }
                else
                {
                    //For Non Planar Surface, get a point(using seam) near the penetrated part and create a plane

                    Seam webSeamObj = null;
                    CommonFuncs.GetSeamBtwObjects(penetratingPart, penetratedPart, out webSeamObj);
                    Position StartPos = null;
                    Position EndPos = null;
                    if (webSeamObj != null)
                    {
                        webSeamObj.EndPoints(out StartPos, out EndPos);

                        ISurface PtdBaseSurface = (ISurface)PtdBasePort;

                        Position PosOnBaseSurface = null;
                        PosOnBaseSurface = PtdBaseSurface.ProjectPoint(StartPos);

                        Vector VecBaseNormal = PtdBaseSurface.OutwardNormalAtPoint(PosOnBaseSurface);
                        ApproxPlane = new Plane3d(PosOnBaseSurface, VecBaseNormal);
                    }
                }
                    
            }

            return ApproxPlane;
        }
    }
}
