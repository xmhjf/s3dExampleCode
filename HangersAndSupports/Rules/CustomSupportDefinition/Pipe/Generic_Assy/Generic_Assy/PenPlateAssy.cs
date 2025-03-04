//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PenPlateAssy.cs
//   Generic_Assy,Ingr.SP3D.Content.Support.Rules.PenPlateAssy
//   Author       : Manikanth
//   Creation Date: 17-Sep-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   17-Sep-2013    Manikanth  CR-CP-224494 Convert HS_Generic_Assy to C# .Net 
//   22-Jan-2015    PVK        TR-CP-264951  Resolve coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    public class PenPlateAssy : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string PLATE = "PLATE";
        private const string HOLE = "HOLE";
        private bool routeAlligned;
        private string routePortName;
        private int hole, count,rightRouteIndex,leftRouteIndex,routeCount, routeIndex;
        private double leftOffset, rightOffset, plateThikness, widthOffset, heightOffset, filletRadius, extremePipesDistance;
        string[] holes;
        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    leftOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyPenPlate", "LeftOffset")).PropValue;
                    rightOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyPenPlate", "RightOffset")).PropValue;
                    plateThikness = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyPenPlate", "PlateThk")).PropValue;
                    widthOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyPenPlate", "WOffset")).PropValue;
                    heightOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyPenPlate", "HOffset")).PropValue;
                    filletRadius = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyPenPlate", "FilletRadius")).PropValue;

                    routeCount = SupportHelper.SupportedObjects.Count;
                    if (routeCount > 1)
                    {
                        double[] originX = new double[routeCount + 1];
                        double[] originY = new double[routeCount + 1];
                        double[] originZ = new double[routeCount + 1];
                        for (routeIndex = 1; routeIndex <= routeCount; routeIndex++)
                        {
                            if (routeIndex == 1)
                                routePortName = "Route";
                            else
                                routePortName = "Route_" + routeIndex;

                            Matrix4X4 matrix = RefPortHelper.PortLCS(routePortName);

                            originX[routeIndex] = matrix.Origin.X;
                            originY[routeIndex] = matrix.Origin.Y;
                            originZ[routeIndex] = matrix.Origin.Z;
                        }
                        Vector vector1 = new Vector(), vector2 = new Vector(), normalVector = new Vector();
                        double y=0.0;
                        if (HgrCompareDoubleService.cmpdbl(originY[1] , originY[2])==true)
                            normalVector.Set(1, 0, 0);
                        else if (HgrCompareDoubleService.cmpdbl(originX[1], originX[2]) == true)
                            normalVector.Set(0, 1, 0);
                        else if (HgrCompareDoubleService.cmpdbl(originZ[1], originZ[2]) == true)
                            normalVector.Set(0, 0, 1);

                        if (routeCount == 2)
                            routeAlligned = true;
                        else
                        {
                            for (int i = 1; i <= routeCount - 1; i += 2)
                            {
                                //Create a vector passing through first and second Routes
                                vector1.Set(originX[i + 1] - originX[i], originY[i + 1] - originY[i], originZ[i + 1] - originZ[i]);
                                //Create a vector passing through second and third Routes
                                vector2.Set(originX[i + 2] - originX[i + 1], originY[i + 2] - originY[i + 1], originZ[i + 2] - originZ[i + 1]);
                                //Calculate the angle between the two vectors w.r.to theri normals. If the angle is zero
                                //means the routes are aligned. If all the routes are aligned, then place a single
                                //Penetration plate which is of Egg kind of shape. Otherwise plates as cylinders
                                //for individual routes.

                                y = vector1.Angle(vector2, normalVector);                               
                                if ((y >= 0) && (y <= 2 * (Math3d.DistanceTolerance)))
                                    routeAlligned = true;
                                else if ((y >= (Math.PI - Math3d.DistanceTolerance)) && (y <= (Math.PI + Math3d.DistanceTolerance)))
                                    routeAlligned = true;
                                else
                                    routeAlligned = false;
                            }
                        }
                    }
                    if (routeCount == 1)
                    {
                        parts.Add(new PartInfo(PLATE, "HgrGen_PentrPlate_01"));
                        parts.Add(new PartInfo(HOLE, "Utility_USER_FIXED_CYL_1"));
                    }
                    else
                    {
                        parts.Add(new PartInfo(PLATE, "HgrGen_PentrPlate_01"));
                        count = parts.Count;
                        holes = new string[routeCount + count + 2];
                        for (int i = count + 1; i <= routeCount + count; i++)
                        {
                            holes[i] = "Hole" + i;
                            parts.Add(new PartInfo(holes[i], "Utility_USER_FIXED_CYL_1"));
                        }
                    }
                    return parts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;                
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;               
               
                //1. Load standard bounding box definition     
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;
              
                //2. Get bounding box boundary objects dimension information                
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                //====== ======
                //3. retrieve dimension of the bounding box
                //====== ======
                // Get route box geometry
                //
                //  ____________________
                // |                    |
                // |  ROUTE BOX BOUND   | BBXHeight
                // |____________________|
                //        BBXWidth
                //
                double boundingBoxWidth = boundingBox.Width, boundingBoxHeight = boundingBox.Height, planeOffset = 0.0, originOffset = 0.0;              
                
                if (!routeAlligned)
                {
                    Matrix4X4 matrix1 = RefPortHelper.PortLCS("BBR_Low"), matrix2 = RefPortHelper.PortLCS("BBR_High");
                    Position routeOrigin1 = new Position(), routeOrigin2 = new Position();
                    Vector routeXAxis1 = new Vector(), routeZAxis1 = new Vector(), routeXAxis2 = new Vector(), routeZAxis2 = new Vector(), routeVector = new Vector(), routeYAxis1 = new Vector();
                                       
                    routeOrigin1.Set(matrix1.Origin.X, matrix1.Origin.Y, matrix1.Origin.Z);
                    routeXAxis1.Set(matrix1.XAxis.X, matrix1.YAxis.Y, matrix1.ZAxis.Z);
                    routeZAxis1.Set(matrix1.ZAxis.X, matrix1.ZAxis.Y, matrix1.ZAxis.Z);

                    routeOrigin2.Set(matrix2.Origin.X, matrix2.Origin.Y, matrix2.Origin.Z);
                    routeXAxis1.Set(matrix2.XAxis.X, matrix2.YAxis.Y, matrix2.ZAxis.Z);
                    routeZAxis1.Set(matrix2.ZAxis.X, matrix2.ZAxis.Y, matrix2.ZAxis.Z);

                    routeVector.Set(routeOrigin2.X - routeOrigin1.X, routeOrigin2.Y - routeOrigin1.Y, routeOrigin2.Z - routeOrigin1.Z);
                    routeYAxis1 = routeZAxis1.Cross(routeXAxis1);
                    double angle = Math3d.ATanDeg(boundingBoxHeight / boundingBoxWidth), offset = (routeVector.Length) / 2;                  
                    planeOffset = offset * Math3d.SinDeg(angle);
                    originOffset = offset * Math3d.CosDeg(angle);
                }
                double pipeDiameter=0.0, pipeDiameter1, pipeDiameter2;

                if (routeCount == 1)
                {
                    PipeObjectInfo pipeinfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    pipeDiameter = pipeinfo.OutsideDiameter;

                    (componentDictionary[PLATE]).SetPropertyValue(pipeDiameter / 2, "IJUAHgrPntrPlate", "LeftPipeRadius");
                    (componentDictionary[PLATE]).SetPropertyValue(plateThikness, "IJUAHgrPntrPlate", "Thickness");
                    (componentDictionary[PLATE]).SetPropertyValue(leftOffset, "IJUAHgrPntrPlate", "LeftOffset");
                    (componentDictionary[PLATE]).SetPropertyValue("Round", "IJUAHgrPntrPlate", "PlateShape");
                    (componentDictionary[HOLE]).SetPropertyValue(plateThikness, "IJOAHgrUtility_USER_FIXED_CYL", "L");
                    (componentDictionary[HOLE]).SetPropertyValue(pipeDiameter / 2, "IJOAHgrUtility_USER_FIXED_CYL", "RADIUS");
                }
                else
                {
                    if (!routeAlligned)
                    {
                        (componentDictionary[PLATE]).SetPropertyValue("Rectangle", "IJUAHgrPntrPlate", "PlateShape");
                        (componentDictionary[PLATE]).SetPropertyValue(plateThikness, "IJUAHgrPntrPlate", "Thickness");
                        (componentDictionary[PLATE]).SetPropertyValue(boundingBoxWidth, "IJUAHgrPntrPlate", "PlateWidth");
                        (componentDictionary[PLATE]).SetPropertyValue(boundingBoxHeight, "IJUAHgrPntrPlate", "PlateDepth");
                        (componentDictionary[PLATE]).SetPropertyValue(filletRadius, "IJUAHgrPntrPlate", "FilletRadius");
                        (componentDictionary[PLATE]).SetPropertyValue(widthOffset, "IJUAHgrPntrPlate", "WidthOffset");
                        (componentDictionary[PLATE]).SetPropertyValue(heightOffset, "IJUAHgrPntrPlate", "DepthOffset");
                    }
                    else
                    {
                        extremePipesDistance = GetPipesIndexOfMaximumDistance(out leftRouteIndex, out rightRouteIndex);
                        PipeObjectInfo pipeinfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(leftRouteIndex);
                        pipeDiameter1 = pipeinfo.OutsideDiameter;
                        pipeinfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(rightRouteIndex);
                        pipeDiameter2 = pipeinfo.OutsideDiameter;

                        (componentDictionary[PLATE]).SetPropertyValue("Egg", "IJUAHgrPntrPlate", "PlateShape");
                        (componentDictionary[PLATE]).SetPropertyValue(extremePipesDistance, "IJUAHgrPntrPlate", "DistBetPipes");
                        (componentDictionary[PLATE]).SetPropertyValue(pipeDiameter1 / 2, "IJUAHgrPntrPlate", "LeftPipeRadius");
                        (componentDictionary[PLATE]).SetPropertyValue(pipeDiameter2 / 2, "IJUAHgrPntrPlate", "RightPipeRadius");
                        (componentDictionary[PLATE]).SetPropertyValue(plateThikness, "IJUAHgrPntrPlate", "Thickness");
                        (componentDictionary[PLATE]).SetPropertyValue(leftOffset, "IJUAHgrPntrPlate", "LeftOffset");
                        (componentDictionary[PLATE]).SetPropertyValue(rightOffset, "IJUAHgrPntrPlate", "RightOffset");
                    }
                    for (int i = 1; i <= routeCount; i++)
                    {
                        PipeObjectInfo pipeinfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(i);
                        double diameter = pipeinfo.OutsideDiameter;

                        if (i == 1)
                            hole = count + 1;
                        else
                            hole = count + i;

                        (componentDictionary[holes[hole]]).SetPropertyValue(plateThikness, "IJOAHgrUtility_USER_FIXED_CYL", "L");
                        (componentDictionary[holes[hole]]).SetPropertyValue(diameter / 2, "IJOAHgrUtility_USER_FIXED_CYL", "RADIUS");
                    }
                }
                if (routeCount == 1)
                {
                    //Add Joint Between the Penetration Plate and Route
                    JointHelper.CreateRigidJoint(PLATE, "Route", "-1", "Route", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, plateThikness / 2, 0, 0);
                    //Add Joint Between the Penetration Cyl (Hole) and Route
                    JointHelper.CreateRigidJoint(HOLE, "StartOther", "-1", "Route", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, plateThikness / 2, 0, 0);
                }
                else
                {
                    if (!routeAlligned)
                        //Add Joint Between the Penetration Plate and BBR_Low
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", PLATE, "Route", Plane.XY, Plane.ZX, Axis.Y, Axis.NegativeX, planeOffset, plateThikness / 2, originOffset);
                    else if (routeAlligned)
                    {
                        string leftRouteName = string.Empty;
                        if (leftRouteIndex == 1)
                            leftRouteName = "Route";
                        else
                            leftRouteName = "Route_" + leftRouteIndex;
                        //Add Joint Between the Penetration Plate and Route 1
                        JointHelper.CreateRevoluteJoint(PLATE, "Route", "-1", leftRouteName, Axis.X, Axis.NegativeX);
                        //Add Joint Between the Penetration Plate and Route 2
                        JointHelper.CreateSphericalJoint(PLATE, "Route_2", "-1", "Route_" + rightRouteIndex);

                        for (hole = count + 1; hole <= routeCount + count; hole++)
                        {
                            if (hole == count + 1)
                                routePortName = "Route";
                            else
                            {
                                routeIndex = hole - count;
                                routePortName = "Route_" + routeIndex;
                            }
                            JointHelper.CreateRigidJoint(holes[hole], "StartOther", "-1", routePortName, Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, plateThikness / 2, 0, 0);
                        }
                    }
                }

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                return 1;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(PLATE, 1));
                    if (routeCount > 1)
                    {
                        for (int i = count + 1; i <= routeCount + count; i++)
                        {
                            routeConnections.Add(new ConnectionInfo(holes[i], 1));
                        }
                    }
                    else
                        routeConnections.Add(new ConnectionInfo(PLATE, 1));
                    //Return the collection of Route connection information.
                    return routeConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Struct Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportingConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    //We are not connecting to any structure so we have nothing to return

                    return structConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
        #region ICustomHgrBOMDescription Members
        //---------------------------------------------------
        //BOM Description
        //-----------------------------------------------------      
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                string bomDescription = (string)((PropertyValueString)support.GetPropertyValue("IJOAHgrGenAssyBOMDesc", "BOM_DESC")).PropValue;

                if (bomDescription == "")
                    bomDescription = "Generic Penetration Plate Assembly";

                return bomDescription;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOMDescription." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
        /// <summary>
        /// Gets the Distance Between Pipes.
        /// </summary>
        /// <param name="pipeIndex1">Pipe index1</param>
        /// <param name="pipeIndex2">Pipe index2.</param>
        /// <returns>double value</returns>
        /// <example>
        /// <code>
        ///    double  extremePipesDist = GetMaxDirectDistBetPipesWithIndex(out leftRouteIndex, out rightRouteIndex);
        /// </code>
        /// </example>   
        public double GetPipesIndexOfMaximumDistance(out int pipeIndex1, out int pipeIndex2)
        {
            try
            {
            double maximumDistance = 0;
            string[] pipePortName = new string[routeCount + 1];
            for (int i = 1; i <= routeCount; i++)
            {
                if (i == 1)
                    pipePortName[1] = "Route";
                else
                    pipePortName[i] = "Route_" + i;
            }

            double[,] portsDistances = new double[21, 21];

            for (int i = 1; i <= routeCount; i++)
            {
                for (int j = i + 1; j <= routeCount; j++)
                {
                    portsDistances[i, j] = RefPortHelper.DistanceBetweenPorts(pipePortName[i], pipePortName[j], PortDistanceType.Direct);
                }
            }
            pipeIndex1 = 0;
            pipeIndex2 = 0;
            for (int i = 1; i <= routeCount; i++)
            {
                for (int j = i + 1; j <= routeCount; j++)
                {
                    if (maximumDistance < portsDistances[i, j])
                    {
                        maximumDistance = portsDistances[i, j];
                        pipeIndex1 = i;
                        pipeIndex2 = j;
                    }
                }
            }
            return maximumDistance;
        }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in GetPipesIndexOfMaximumDistance." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
    }
}