//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   FLSampleSupportServices.cs
//   Author       : Rajeswari
//   Creation Date: 03-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 03-Sep-2013  Rajeswari,Hema CR-CP-224478 Convert FlSample_Supports to C# .Net
// 28-Apr-2015      PVK	       Resolve Coverity issues found in April
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{

    public class FLSampleSupportServices
    {
        /// <summary>
        /// This method retrieves the index and the pipe diameter of the pipe furthest from structure
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="insulation">insulation value</param>
        /// <returns>index of route, diameter of pipe</returns>
        /// <code>
        /// GetOutsideRouteProps(this, 1)
        /// </code>
        public static object[] GetOutsideRouteProps(CustomSupportDefinition customSupportDefinition, int insulation)
        {
            try
            {
                Double pipeDiameter = 0, insulationThickness = 0, tempLength = 0;
                string outsideRoutePort = string.Empty;
                int index, idxOutside = 0;
                PipeObjectInfo pipeInfo;
                double min = 10000000;

                for (index = 1; index <= customSupportDefinition.SupportHelper.SupportedObjects.Count; index++)
                {
                    pipeInfo = (PipeObjectInfo)customSupportDefinition.SupportedHelper.SupportedObjectInfo(index);
                    if (index == 1)
                    {
                        tempLength = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                        pipeDiameter = pipeInfo.OutsideDiameter;
                        insulationThickness = pipeInfo.InsulationThickness;
                        if (insulation == 1)
                            pipeDiameter = pipeDiameter + 2 * insulationThickness;
                    }
                    else
                    {
                        tempLength = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route_" + index.ToString(), "Structure", PortDistanceType.Horizontal);
                        pipeDiameter = pipeInfo.OutsideDiameter;
                        insulationThickness = pipeInfo.InsulationThickness;
                        if (insulation == 1)
                            pipeDiameter = pipeDiameter + 2 * insulationThickness;
                    }

                    if (HgrCompareDoubleService.cmpdbl( min , 10000000) == true)
                    {
                        min = tempLength + pipeDiameter / 2;
                        idxOutside = index;
                    }
                    else
                    {
                        if (tempLength + pipeDiameter / 2 > min)
                        {
                            min = tempLength;
                            idxOutside = index;
                        }
                    }
                }
                pipeInfo = (PipeObjectInfo)customSupportDefinition.SupportedHelper.SupportedObjectInfo(idxOutside);

                if (idxOutside == 1)
                    outsideRoutePort = "Route";
                else
                    outsideRoutePort = "Route_" + idxOutside.ToString();

                object[] array = { idxOutside, pipeDiameter, insulationThickness, outsideRoutePort };
                return array;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetOutsideRouteProps." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method retrieves the longest distance between all the pipes
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <returns>Longest distance between pipes</returns>
        /// <code>
        /// GetWidestDistance(this)
        /// </code>
        public static Double GetWidestDistance(CustomSupportDefinition customSupportDefinition)
        {
            try
            {
                double tempLength = 0, newDistance = 0;
                for (int index = 1; index <= customSupportDefinition.SupportHelper.SupportedObjects.Count; index++)
                {
                    if (index == customSupportDefinition.SupportHelper.SupportedObjects.Count)
                        tempLength = 0;
                    else
                    {
                        for (int index2 = 1; index <= customSupportDefinition.SupportHelper.SupportedObjects.Count; index2++)
                        {
                            newDistance = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route_" + index.ToString(), "Route_" + index2.ToString(), PortDistanceType.Horizontal);

                            if (newDistance > tempLength)
                                tempLength = newDistance;
                        }
                    }
                }
                return tempLength;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetWidestDistance." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method gives horizontal distance between the ports
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="port1">The First Port</param>
        /// <param name="port2">The Second Port</param>
        /// <returns></returns>
        /// <code>
        /// GetHorizontalDistanceBetweenPorts(this, port1, port2)
        /// </code>
        public static Double GetHorizontalDistanceBetweenPorts(CustomSupportDefinition customSupportDefinition, string port1, string port2)
        {
            try
            {
                double distanceBetweenPorts = 0;

                double directDistance = customSupportDefinition.RefPortHelper.DistanceBetweenPorts(port1, port2, PortDistanceType.Direct);
                double horizontaldistance = customSupportDefinition.RefPortHelper.DistanceBetweenPorts(port1, port2, PortDistanceType.Horizontal_Perpendicular);
                return distanceBetweenPorts = Math.Sqrt((directDistance * directDistance) - (horizontaldistance * horizontaldistance));
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetHorizontalDistanceBetweenPorts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method checks if there are vertical pipes present
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <code>
        /// CheckPipeOrientation(this)
        /// </code>
        public static object[] CheckPipeOrientation(CustomSupportDefinition customSupportDefinition)
        {
            try
            {
                int iVertical = 0, iHorizontal = 0, iBoth = 0;
                double pipeAngle = 0;
                for (int index = 1; index <= customSupportDefinition.SupportHelper.SupportedObjects.Count; index++)
                {
                    if (index == 1)
                    {
                        pipeAngle = (customSupportDefinition.RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z) * 180 / Math.PI);
                        if ((pipeAngle >= 0 && pipeAngle < 0.00001) || (pipeAngle > 179.999999 && pipeAngle < 180.00001))
                            iVertical = 1;
                        else if ((pipeAngle >= 89.999999 && pipeAngle < 90.00001) || (pipeAngle > 269.999999 && pipeAngle < 270.00001))
                            iHorizontal = 1;
                    }
                    else
                    {
                        pipeAngle = (customSupportDefinition.RefPortHelper.AngleBetweenPorts("Route_" + index.ToString(), PortAxisType.X, OrientationAlong.Global_Z) * 180 / Math.PI);
                        if ((pipeAngle >= 0 && pipeAngle < 0.00001) || (pipeAngle > 179.999999 && pipeAngle < 180.00001))
                            iVertical = 1;
                        else if ((pipeAngle >= 89.999999 && pipeAngle < 90.00001) || (pipeAngle > 269.999999 && pipeAngle < 270.00001))
                            iHorizontal = 1;
                    }
                }

                if (iHorizontal == 1 && iVertical == 1)
                    iBoth = 1;

                object[] array = { iVertical, iHorizontal, iBoth };
                return array;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in CheckPipeOrientation." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method retrieves distance and port for the vertical route
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="outerMostPort">The outer most port</param>
        /// <returns></returns>
        /// <code>
        /// GetDistanceAndPortForVerticalRoute(this, port)
        /// </code>
        public static object[] GetDistanceAndPortForVerticalRoute(CustomSupportDefinition customSupportDefinition, string outerMostPort)
        {
            try
            {
                Double distStructToOuterRoute = 0, distStructToRoute = 0, distOuterRouteToRoute = 0, tempDistance = 0;
                string outsideRoutePort = string.Empty;
                int idxOutside = 0;
                Boolean isVerticalRouteOtherside = false;

                distStructToOuterRoute = FLSampleSupportServices.GetHorizontalDistanceBetweenPorts(customSupportDefinition, outerMostPort, "Structure");
                for (int index = 1; index <= customSupportDefinition.SupportHelper.SupportedObjects.Count; index++)
                {
                    if (index == 1)
                        distOuterRouteToRoute = FLSampleSupportServices.GetHorizontalDistanceBetweenPorts(customSupportDefinition, "Route", outerMostPort);
                    else
                        distOuterRouteToRoute = FLSampleSupportServices.GetHorizontalDistanceBetweenPorts(customSupportDefinition, "Route_" + index.ToString(), outerMostPort);

                    if (distStructToOuterRoute < distOuterRouteToRoute)
                    {
                        isVerticalRouteOtherside = true;
                        if (index == 1)
                            distStructToRoute = FLSampleSupportServices.GetHorizontalDistanceBetweenPorts(customSupportDefinition, "Route", "Structure");
                        else
                            distStructToRoute = FLSampleSupportServices.GetHorizontalDistanceBetweenPorts(customSupportDefinition, "Route_" + index.ToString(), "Structure");
                    }

                    if (distStructToOuterRoute > tempDistance)
                    {
                        tempDistance = distStructToRoute;
                        idxOutside = index;
                    }
                }

                if (idxOutside == 1)
                    outsideRoutePort = "Route";
                else
                    outsideRoutePort = "Route_" + idxOutside.ToString();

                object[] array = null;
                if (isVerticalRouteOtherside == false)
                    array = new object[] { 0, outsideRoutePort };
                else
                    array = new object[] { tempDistance, outsideRoutePort };

                return array;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetDistanceAndPortForVerticalRoute." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method retrieves length and outside Diameter from right and left to BBX
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <returns></returns>
        /// <code>
        /// GetLengthAndDiameterOutsideBBX(this)
        /// </code>
        public static object[] GetLengthAndDiameterOutsideBBX(CustomSupportDefinition customSupportDefinition)
        {
            try
            {
                double outSideDiaOnLeft = 0, outSideDiaOnRight = 0, lengthOutsideBBX = 0, distFromLow = 0, distFromHigh = 0, distMaxHigh = 0, distMaxLow = 0, distBBXHighToRoute = 0, distBBXLowToRoute = 0;
                BoundingBox boundingBox;
                if (customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = customSupportDefinition.BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = customSupportDefinition.BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                double boundingBoxHeight = boundingBox.Height;

                for (int index = 1; index <= customSupportDefinition.SupportHelper.SupportedObjects.Count; index++)
                {
                    if (index == 1)
                    {
                        distBBXHighToRoute = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route", "BBSR_High", PortDistanceType.Horizontal);
                        distBBXLowToRoute = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route", "BBSR_Low", PortDistanceType.Horizontal);
                    }
                    else
                    {
                        distBBXHighToRoute = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route_" + index.ToString(), "BBSR_High", PortDistanceType.Horizontal);
                        distBBXLowToRoute = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route_" + index.ToString(), "BBSR_Low", PortDistanceType.Horizontal);
                    }

                    if ((distBBXHighToRoute > boundingBoxHeight) || (distBBXLowToRoute > boundingBoxHeight))
                    {
                        if (distBBXHighToRoute < distBBXLowToRoute)
                        {
                            if (distBBXHighToRoute > distMaxHigh)
                            {
                                distMaxHigh = distBBXHighToRoute;
                                distFromHigh = distMaxHigh;
                            }
                        }
                        else
                        {
                            if (distBBXLowToRoute > distMaxLow)
                            {
                                distMaxLow = distBBXLowToRoute;
                                distFromLow = distMaxLow;
                            }
                        }
                    }
                }

                lengthOutsideBBX = distFromHigh + distFromLow;
                object[] array = { lengthOutsideBBX, distFromHigh, distFromLow, outSideDiaOnLeft, outSideDiaOnRight };
                return array;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetLengthAndDiameterOutsideBBX." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This metod returns the greatest value from the given values.
        /// <param name="firstValue"> The First Value</param>
        /// <param name="secondValue">The Second Value</param>
        /// <returns></returns>
        /// <code>
        /// GetGreaterValue(firstValue, SecondValue)
        /// </code>
        public static Double GetGreaterValue(double firstValue, double secondValue)
        {
            double tempValue = 0;

            if (firstValue > secondValue)
                tempValue = firstValue;
            else if (firstValue < secondValue)
                tempValue = secondValue;
            else
                tempValue = secondValue;

            return tempValue;
        }
        /// <summary>
        /// This method retrieves the OutDiameter of route from the side of structure
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="outerMostPort"> The outer most port</param>
        /// <param name="outerRouteDiameter"> Route Diametr</param>
        /// <returns></returns>
        /// <code>
        /// GetOutSideRouteDiameterOnOtherSideStruct(this, port, routeDia)
        /// </code>
        public static object[] GetOutSideRouteDiameterOnOtherSideStruct(CustomSupportDefinition customSupportDefinition, string outerMostPort, double outerRouteDiameter)
        {
            try
            {
                double distStructToOuterRoute = 0, distStructToRoute = 0, distOuterRouteToRoute = 0, tempDistance = 0, outsideRouteDiaNearBBX_Low = 0, outsideRouteDiaNearBBX_High = 0, pipeDiameter = 0;
                string outsideRoutePort = string.Empty;
                int idxOutside = 0;

                distStructToOuterRoute = customSupportDefinition.RefPortHelper.DistanceBetweenPorts(outerMostPort, "Structure", PortDistanceType.Horizontal);

                for (int index = 1; index <= customSupportDefinition.SupportHelper.SupportedObjects.Count; index++)
                {
                    if (index == 1)
                        distOuterRouteToRoute = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route", outerMostPort, PortDistanceType.Horizontal);
                    else
                        distOuterRouteToRoute = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route_" + index.ToString(), outerMostPort, PortDistanceType.Horizontal);

                    if (distStructToOuterRoute < distOuterRouteToRoute)
                    {
                        if (index == 1)
                            distStructToRoute = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                        else
                            distStructToRoute = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route_" + index.ToString(), "Structure", PortDistanceType.Horizontal);
                    }

                    if (distStructToRoute > tempDistance)
                    {
                        tempDistance = distStructToRoute;
                        idxOutside = index;
                    }
                }

                PipeObjectInfo pipeInfo = (PipeObjectInfo)customSupportDefinition.SupportedHelper.SupportedObjectInfo(idxOutside);
                pipeDiameter = pipeInfo.OutsideDiameter;
                if (idxOutside == 1)
                    outsideRoutePort = "Route";
                else
                    outsideRoutePort = "Route_" + idxOutside.ToString();

                if (customSupportDefinition.RefPortHelper.DistanceBetweenPorts(outsideRoutePort, "BBSR_High", PortDistanceType.Horizontal) > customSupportDefinition.RefPortHelper.DistanceBetweenPorts(outerMostPort, "BBSR_High", PortDistanceType.Horizontal))
                {
                    outsideRouteDiaNearBBX_High = outerRouteDiameter;
                    outsideRouteDiaNearBBX_Low = pipeDiameter;
                }
                else
                {
                    outsideRouteDiaNearBBX_High = pipeDiameter;
                    outsideRouteDiaNearBBX_Low = outerRouteDiameter;
                }
                object[] array = { outsideRouteDiaNearBBX_High, outsideRouteDiaNearBBX_Low };
                return array;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetOutSideRouteDiameterOnOtherSideStruct." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method retrieves the nearest distance from route
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="outerMostRoutePort1"> The OuterMost Port1</param>
        /// <param name="outerMostRoutePort2">The OuterMost Port2</param>
        /// <returns></returns>
        /// <code>
        /// GetNearestOuterMostRouteDistance(this, port1, port2)
        /// </code>
        public static Double GetNearestOuterMostRouteDistance(CustomSupportDefinition customSupportDefinition, string outerMostRoutePort1, string outerMostRoutePort2)
        {
            try
            {
                double distStructToOuterPort1 = customSupportDefinition.RefPortHelper.DistanceBetweenPorts(outerMostRoutePort1, "Structure", PortDistanceType.Horizontal);
                double distStructToOuterPort2 = customSupportDefinition.RefPortHelper.DistanceBetweenPorts(outerMostRoutePort2, "Structure", PortDistanceType.Horizontal);
                double distStructToRoutePort = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                double distRoutePortToOuterPort1 = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route", outerMostRoutePort1, PortDistanceType.Horizontal);
                double distRoutePortToOuterPort2 = customSupportDefinition.RefPortHelper.DistanceBetweenPorts("Route", outerMostRoutePort2, PortDistanceType.Horizontal);
                double pipeOrientation = (customSupportDefinition.RefPortHelper.AngleBetweenPorts(outerMostRoutePort1, PortAxisType.X, OrientationAlong.Global_Z) * 180 / Math.PI);

                if ((HgrCompareDoubleService.cmpdbl(pipeOrientation , 0) == true && pipeOrientation < 0.00001) || (pipeOrientation > 179.999999 && pipeOrientation < 180.00001))
                    distStructToOuterPort1 = GetHorizontalDistanceBetweenPorts(customSupportDefinition, outerMostRoutePort1, "Structure");

                pipeOrientation = (customSupportDefinition.RefPortHelper.AngleBetweenPorts(outerMostRoutePort2, PortAxisType.Z, "Structure", PortAxisType.Z, OrientationAlong.Global_Z) * 180 / Math.PI);
                if ((HgrCompareDoubleService.cmpdbl(pipeOrientation , 0) == true && pipeOrientation < 0.00001) || (pipeOrientation > 179.999999 && pipeOrientation < 180.00001))
                    distStructToOuterPort1 = GetHorizontalDistanceBetweenPorts(customSupportDefinition, outerMostRoutePort2, "Structure");

                pipeOrientation = (customSupportDefinition.RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Z, "Structure", PortAxisType.Z, OrientationAlong.Global_Z) * 180 / Math.PI);
                if ((HgrCompareDoubleService.cmpdbl(pipeOrientation , 0) == true && pipeOrientation < 0.00001) || (pipeOrientation > 179.999999 && pipeOrientation < 180.00001))
                    distStructToOuterPort1 = GetHorizontalDistanceBetweenPorts(customSupportDefinition, "Route", "Structure");

                pipeOrientation = (customSupportDefinition.RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Z, "Structure", PortAxisType.Z, OrientationAlong.Global_Z) * 180 / Math.PI);
                if ((HgrCompareDoubleService.cmpdbl(pipeOrientation , 0) == true && pipeOrientation < 0.00001) || (pipeOrientation > 179.999999 && pipeOrientation < 180.00001))
                    distStructToOuterPort1 = GetHorizontalDistanceBetweenPorts(customSupportDefinition, "Route", outerMostRoutePort1);

                pipeOrientation = (customSupportDefinition.RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Z, "Structure", PortAxisType.Z, OrientationAlong.Global_Z) * 180 / Math.PI);
                if ((HgrCompareDoubleService.cmpdbl(pipeOrientation , 0) == true && pipeOrientation < 0.00001) || (pipeOrientation > 179.999999 && pipeOrientation < 180.00001))
                    distStructToOuterPort1 = GetHorizontalDistanceBetweenPorts(customSupportDefinition, "Route", outerMostRoutePort2);

                if (distRoutePortToOuterPort1 > distStructToOuterPort1)
                    return distStructToOuterPort2;
                else
                    return distStructToOuterPort1;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetNearestOuterMostRouteDistance." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        /// <summary>
        /// This method checks whether a offset value needs to be explicitly specified on both end of the leg.
        /// If the input structure is either a slab or a non-parallel member, then the offset needs to be specified. Else no offset value needs to be specified
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <returns>Boolean array</returns>
        /// <code>
        ///     Boolean[] bIsOffsetApplied = GetIsLugEndOffsetApplied();  
        ///</code>
        public static Boolean[] GetIsLugEndOffsetApplied(CustomSupportDefinition customSupportDefinition)
        {
            try
            {
                Collection<BusinessObject> structureObjects = customSupportDefinition.SupportHelper.SupportingObjects;
                Boolean[] isOffsetApplied = new Boolean[2];

                //first route object is set as primary route object

                isOffsetApplied[0] = true;
                isOffsetApplied[1] = true;

                if (structureObjects != null)
                {
                    if (structureObjects.Count >= 2)
                    {
                        for (int index = 0; index <= 1; index++)
                        {
                            //check if offset is to be applied
                            const double routeStructAngle = Math.PI / 180.0;
                            Boolean varRuleApplied = true;
                            if ((customSupportDefinition.SupportHelper.SupportingObjects.Count != 0))
                            {
                                if (customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member)
                                {
                                    //if angle is within 1 degree, regard as parallel case
                                    //Also check for Sloped structure                                
                                    MemberPart memberPart = (MemberPart)customSupportDefinition.SupportHelper.SupportingObjects[index];
                                    ICurve memberCurve = memberPart.Axis;

                                    Vector supportedVector = new Vector();
                                    Vector supportingVector = new Vector();

                                    if (customSupportDefinition.SupportedHelper.SupportedObjectInfo(1).SupportedObjectType == SupportedObjectType.Pipe)
                                    {
                                        Position startLocation = new Position(customSupportDefinition.SupportedHelper.SupportedObjectInfo(1).StartLocation);
                                        Position endLocation = new Position(customSupportDefinition.SupportedHelper.SupportedObjectInfo(1).EndLocation);
                                        supportedVector = new Vector(endLocation - startLocation);
                                    }
                                    if (memberCurve is ILine)
                                    {
                                        ILine line = (ILine)memberCurve;
                                        supportingVector = line.Direction;
                                    }

                                    double angle = GetAngleBetweenVectors(supportingVector, supportedVector);
                                    double refAngle1 = customSupportDefinition.RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.Y, OrientationAlong.Global_X) - Math.PI / 2;
                                    double refAngle2 = customSupportDefinition.RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.X, OrientationAlong.Global_X);

                                    if (angle < (refAngle1 + 0.001) && angle > (refAngle1 - 0.001))
                                        angle = angle - Math.Abs(refAngle1);
                                    else if (angle < (refAngle2 + 0.001) && angle > (refAngle2 - 0.001))
                                        angle = angle - Math.Abs(refAngle2);
                                    else
                                        angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                    angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                    if (Math.Abs(angle) < routeStructAngle || Math.Abs(angle - Math.PI) < routeStructAngle)
                                        varRuleApplied = false;
                                }
                            }
                            isOffsetApplied[index] = varRuleApplied;
                        }
                    }
                }

                return isOffsetApplied;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetIsLugEndOffsetApplied." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method differentiate multiple structure input object based on elative position. For U shape support, when two structure are inputted.
        /// This method returns the port name on left and right side of the route object. (Left is on the negative Y axis of bounding box coord. sys., right is on the positive Y axis of the bounding box coord. sys.)
        /// </summary>
        /// <param name="customSupportDefinition">an abstract class, that contain the support definition</param>
        /// <param name="isOffsetApplied">IsOffsetApplied-The offset applied.(Boolean[])</param>
        /// <returns>String array</returns>
        /// <code>
        ///     string[] idxStructPort = GetIndexedStructPortName(bIsOffsetApplied);
        /// </code>
        public static String[] GetIndexedStructPortName(CustomSupportDefinition customSupportDefinition, Boolean[] isOffsetApplied)
        {
            try
            {
                String[] structurePort = new String[2];
                int structureCount = customSupportDefinition.SupportHelper.SupportingObjects.Count;
                int i;

                if (customSupportDefinition.SupportHelper.PlacementType == PlacementType.PlaceByReference)
                {
                    structurePort[0] = "Structure";
                    structurePort[1] = "Structure";
                }
                else
                {
                    structurePort[0] = "Structure";
                    structurePort[1] = "Struct_2";

                    if (structureCount > 1)
                    {
                        if (customSupportDefinition.SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                        {
                            for (i = 0; i <= 1; i++)
                            {
                                double angle = 0;
                                if ((customSupportDefinition.SupportHelper.SupportingObjects.Count != 0))
                                {
                                    if (customSupportDefinition.SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member & isOffsetApplied[i] == false)
                                        angle = customSupportDefinition.RefPortHelper.PortConfigurationAngle("Route", structurePort[i], PortAxisType.Y);
                                }
                                //the port is the right structure port
                                if (Math.Abs(angle) < Math.PI / 2.0)
                                {
                                    if (i == 0)
                                    {
                                        structurePort[0] = "Struct_2";
                                        structurePort[1] = "Structure";
                                    }
                                }
                                //the port is the left structure port
                                else
                                {
                                    if (i == 1)
                                    {
                                        structurePort[0] = "Struct_2";
                                        structurePort[1] = "Structure";
                                    }
                                }

                            }
                        }
                    }
                    else
                        structurePort[1] = "Structure";
                }
                //switch the OffsetApplied flag
                if (structurePort[0] == "Struct_2")
                {
                    Boolean flag = isOffsetApplied[0];
                    isOffsetApplied[0] = isOffsetApplied[1];
                    isOffsetApplied[1] = flag;
                }

                return structurePort;
            }
            catch (Exception e)
            {
                Type myType = customSupportDefinition.GetType();
                CmnException e1 = new CmnException("Error in GetIndexedStructPortName." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        /// <summary>
        /// This method returns the direct angle between two vectors in Radians
        /// </summary>
        /// <param name="vector1">The vector1(Vector).</param>
        /// <param name="vector2">The vector2(Vector).</param>
        /// <returns>double</returns>        
        /// <code>
        ///       ContentHelper contentHelper = new ContentHelper();
        ///       double value;
        ///       value = contentHelper. GetAngleBetweenVectors(vector1, vector2 );
        ///</code>

        public static double GetAngleBetweenVectors(Vector vector1, Vector vector2)
        {
            try
            {
                double dotProd = (vector1.Dot(vector2) / (vector1.Length * vector2.Length));
                double arcCos = 0.0;
                if (HgrCompareDoubleService.cmpdbl(Math.Abs(dotProd), 1) == false)
                {
                    arcCos = Math.PI / 2 - Math.Atan(dotProd / Math.Sqrt(1 - dotProd * dotProd));
                }
                else if (HgrCompareDoubleService.cmpdbl(dotProd, -1) == true)
                {
                    arcCos = Math.PI;
                }
                else if (HgrCompareDoubleService.cmpdbl(dotProd, 1) == true)
                {
                    arcCos = 0;
                }
                return arcCos;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        /// <summary>
        /// Gets the data on part by condition.
        /// </summary>
        /// <param name="partClassName">The data required part class name.</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="conditionInterface">The condition interface.</param>
        /// <param name="conditionProperty">The condition property.</param>
        /// <param name="referenceValue">The reference value.</param>
        /// <param name="comparisionavalue">The comparision value</param>
        /// <returns>double value</returns>
        /// <example>
        /// <code>
        ///     PropertyValueCodelist loadClassCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass");
        ///     loadClass = loadClassCodeList.PropertyInfo.DisplayName;
        ///     double dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5862","IJUAFINL_C", "C","IJUAFINL_LoadClass", "LoadClass",loadClass,0.001)   
        /// </code>
        /// </example>        
        public static double GetDataByCondition(string partClassName, string dataInterface, string dataProperty, string conditionInterface, string conditionProperty, double minimumReferencevalue, double maximumReferencevalue)
        {
            IEnumerable<BusinessObject> flSampleCmpParts = null;
            try
            {
                double distance = 0;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass flSampleCmpPartClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);

                if (flSampleCmpPartClass.PartClassType.Equals("HgrServiceClass"))
                    flSampleCmpParts = flSampleCmpPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    flSampleCmpParts = flSampleCmpPartClass.Parts;

                flSampleCmpParts = flSampleCmpParts.Where(part => (double)((PropertyValueDouble)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue > minimumReferencevalue && (double)((PropertyValueDouble)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue < maximumReferencevalue);
                if (flSampleCmpParts.Count() > 0)
                    distance = ((double)((PropertyValueDouble)flSampleCmpParts.ElementAt(0).GetPropertyValue(dataInterface, dataProperty)).PropValue);
                return distance;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get DataByConditionof FLSampleSupportServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (flSampleCmpParts is IDisposable)
                {
                    ((IDisposable)flSampleCmpParts).Dispose(); // This line will be executed
                }
            }
        }
        /// <summary>
        /// Gets the data on part by condition.
        /// </summary>
        /// <param name="partClassName">The data required part class name.</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="conditionInterface">The condition interface.</param>
        /// <param name="conditionProperty">The condition property.</param>
        /// <param name="referenceValue">The reference value.</param>
        /// <returns>double value</returns>
        /// <example>
        /// <code>
        ///     PropertyValueCodelist loadClassCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass");
        ///     loadClass = loadClassCodeList.PropertyInfo.DisplayName;
        ///     double dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5862","IJUAFINL_C", "C","IJUAFINL_LoadClass", "LoadClass",loadClass)   
        /// </code>
        /// </example>        
        public static double GetDataByCondition(string partClassName, string dataInterface, string dataProperty, string conditionInterface, string conditionProperty, string referenceValue)
        {
            IEnumerable<BusinessObject> flSampleCmpParts = null;
            try
            {
                double distance;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass flSampleCmpPartClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);

                if (flSampleCmpPartClass.PartClassType.Equals("HgrServiceClass"))
                    flSampleCmpParts = flSampleCmpPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    flSampleCmpParts = flSampleCmpPartClass.Parts;

                flSampleCmpParts = flSampleCmpParts.Where(part => (string)((PropertyValueString)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue == referenceValue);
                if (flSampleCmpParts.Count() > 0)
                    distance = ((double)((PropertyValueDouble)flSampleCmpParts.ElementAt(0).GetPropertyValue(dataInterface, dataProperty)).PropValue);
                else
                    distance = 0;
                return distance;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get DataByCondition of FLSampleSupportServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (flSampleCmpParts is IDisposable)
                {
                    ((IDisposable)flSampleCmpParts).Dispose(); // This line will be executed
                }
            }
        }
        /// <summary>
        /// Gets the data on part by condition.
        /// </summary>
        /// <param name="partClassName">The data required part class name.</param>
        /// <param name="dataInterface">Interface to which the property belongs.</param>
        /// <param name="dataProperty">Name of the property which we required.</param>
        /// <param name="conditionInterface">The condition interface.</param>
        /// <param name="conditionProperty">The condition property.</param>
        /// <param name="referenceValue">The reference value.</param>
        /// <returns>double value</returns>
        /// <example>
        /// <code>
        ///     double dis = GetStringDataByCondition("Anv09_RodET", "IJDPart", "PartNumber", "IJUAhsRodDiameter", "RodDiameter", rodSize - 0.000001, rodSize + 0.000001); 
        /// </code>
        /// </example>        
        public static string GetStringDataByCondition(string partClassName, string dataInterface, string dataProperty, string conditionInterface, string conditionProperty, double minimumReferencevalue, double maximumReferencevalue)
        {
            IEnumerable<BusinessObject> flSampleCmpParts = null;
            try
            {
                string distance = string.Empty;
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass flSampleCmpPartClass = (PartClass)catalogBaseHelper.GetPartClass(partClassName);

                if (flSampleCmpPartClass.PartClassType.Equals("HgrServiceClass"))
                    flSampleCmpParts = flSampleCmpPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                else
                    flSampleCmpParts = flSampleCmpPartClass.Parts;

                flSampleCmpParts = flSampleCmpParts.Where(part => (double)((PropertyValueDouble)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue > minimumReferencevalue && (double)((PropertyValueDouble)part.GetPropertyValue(conditionInterface, conditionProperty)).PropValue < maximumReferencevalue);
                if (flSampleCmpParts.Count() > 0)
                    distance = ((string)((PropertyValueString)flSampleCmpParts.ElementAt(0).GetPropertyValue(dataInterface, dataProperty)).PropValue);
                return distance;
            }
            catch (Exception e)
            {
                CmnException e1 = new CmnException("Error in Get GetDataByConditionString of FINLAssemblyServices class" + ". Error:" + e.Message, e);
                throw e1;
            }
            finally
            {
                if (flSampleCmpParts is IDisposable)
                {
                    ((IDisposable)flSampleCmpParts).Dispose(); // This line will be executed
                }
            }
        }
    }
}
