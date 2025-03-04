//----------------------------------------------------------------------------------
//      Copyright (C) 2013 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   PlateUpSide rule returns Upside of Plate    
//
//      Author:  
//
//      History:
//      July 31th, 2014   Created by Natilus-HSV
//
//-----------------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Content.Manufacturing;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Planning.Middle;
using ProcessInfo = Ingr.SP3D.Content.Manufacturing.ProcessInfo;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// PlateUpSide Rule
    /// </summary>
    public class PlateUpSide : PlateUpSideRule
    {

        /// <summary>
        /// Set PlateUpSide Rule 
        /// </summary>
        /// <param name="processInfo">The process info.</param>
        public override void Evaluate(ProcessInformation processInfo)
        {
            try
            {
                if (processInfo == null)
                    throw new ArgumentNullException("Input ProcessInfo is empty");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs
                PlatePartBase platePart = null;
                if (processInfo.ManufacturingParent != null)
                {
                    platePart = (PlatePartBase)processInfo.ManufacturingParent;
                }

                #endregion

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                int plateSide = 0; //Undefined

                //Get Arguments
                Dictionary<int, object> args = processInfo.GetArguments("ProductionControl");
                if (args.Count == 0)
                {
                    return;
                }

                string controlAttribute = Convert.ToString(processInfo.GetArguments("ProductionControl").FirstOrDefault().Value);
                TopologyPort basePort = null;
                TopologyPort offsetPort = null;
                Position centerPositionSurface = null;
                Vector normalVector = null;
                ISurface surfaceGeometry = null;

                switch (controlAttribute.ToUpper())
                {

                    case "MOLDED":
                        #region Molded

                        if (platePart != null)
                        {
                            if (processInfo.GetArguments("Hull").Count == 1 && platePart.Type == PlateType.Hull)
                            {
                                plateSide = Convert.ToInt32(processInfo.GetArguments("Hull").FirstOrDefault().Value);
                            }
                            else
                            {
                                //Get AlternateMoldedSide to consider thickness direction of Plate 
                                ContextTypes contextType = platePart.MoldedSide;
                                if (contextType == ContextTypes.Base)
                                    plateSide = (int)ManufacturingSide.Base;
                                else if (contextType == ContextTypes.Offset)
                                    plateSide = (int)ManufacturingSide.Offset;
                            }
                        }
                        #endregion
                        break;
                    case "ANTIMOLDED":
                        #region AntiModled

                        //Get AlternateMoldedSide to consider thickness direction of Plate 
                        if (platePart != null)
                        {
                            if (processInfo.GetArguments("Hull").Count == 1 && platePart.Type == PlateType.Hull)
                            {
                                plateSide = Convert.ToInt32(processInfo.GetArguments("Hull").FirstOrDefault().Value);
                            }
                            else
                            {
                                ContextTypes contextType = platePart.MoldedSide;
                                if (contextType == ContextTypes.Base)
                                    plateSide = (int)ManufacturingSide.Offset;
                                else if (contextType == ContextTypes.Offset)
                                    plateSide = (int)ManufacturingSide.Base;
                            }
                        }

                        #endregion
                        break;
                    case "ASSEMBLYORIENTATION":
                        #region AssemblyOrientation
                        if (platePart != null)
                        {
                            AssemblyBase assemblyBase = (AssemblyBase)platePart.AssemblyParent;
                            if (assemblyBase != null)
                            {
                                ContextTypes type = assemblyBase.BasePlateOrientation;
                                if (type == ContextTypes.Offset)
                                {
                                    plateSide = (int)ManufacturingSide.Offset;
                                }
                                else
                                {
                                    plateSide = (int)ManufacturingSide.Base;
                                }

                            }
                        }

                        #endregion
                        break;
                    case "MOSTNUMBEROFMARKING":
                        #region MostNumberOfMarking
                        int numMarkingOnBaseSide = 0, numMarkingOnOffSetSide = 0;

                        if (platePart != null)
                        {
                            ReadOnlyCollection<BusinessObject> connectedParts = platePart.GetConnectedObjects();
                            //Check the ports of platePart
                            ReadOnlyCollection<IPort> connectedPorts = platePart.GetConnectedPorts(PortType.Face);

                            if (connectedPorts != null)
                            {
                                foreach (IPort port in connectedPorts)
                                {
                                    ReadOnlyCollection<IConnection> connections = port.Connections;
                                    if (connections != null)
                                    {
                                        foreach (IConnection con in connections)
                                        {
                                            BusinessObject boundedObject = null;
                                            BusinessObject boundingObject = null;
                                            if (con is PhysicalConnection) //Check Physical Connection 
                                            {
                                                PhysicalConnection pc = (PhysicalConnection)con;
                                                boundedObject = pc.BoundedObject;
                                                boundingObject = pc.BoundingObject;
                                            }

                                            if (connectedParts.Contains(boundedObject) == true || connectedParts.Contains(boundingObject) == true)
                                            {
                                                if (port is TopologyPort)
                                                {
                                                    TopologyPort topoPort = (TopologyPort)port;

                                                    int result1 = 0, result2 = 0;

                                                    result1 = (int)topoPort.ContextId & (int)ContextTypes.Base;
                                                    result2 = (int)topoPort.ContextId & (int)ContextTypes.Offset;

                                                    if (result1 > 0)
                                                    {
                                                        numMarkingOnBaseSide++;
                                                    }
                                                    else if (result2 > 0)
                                                    {
                                                        numMarkingOnOffSetSide++;
                                                    }
                                                    else
                                                    {
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        if (numMarkingOnBaseSide >= numMarkingOnOffSetSide)
                        {
                            plateSide = (int)ManufacturingSide.Base;
                        }
                        else
                        {
                            plateSide = (int)ManufacturingSide.Offset;
                        }
                        #endregion
                        break;
                    case "MOSTNUMBEROFSTIFFENER":
                        #region MostNumberOfStiffener
                        //CR-CP-259243   Need 3D API to get SideWithMoreStiffeners 
                        plateSide = (int)ManufacturingSide.Base;
                        #endregion
                        break;
                    case "IN":
                        #region In
                        // Check Longitudinal Plate or Hull 

                        plateSide = 0; //UnDefinedSide

                        if (platePart != null)
                        {
                            if (platePart.Type == PlateType.LBulkheadPlate || platePart.Type == PlateType.Hull)
                            {
                                //Logic
                                //1. Get the Center and Normal Vector approximately. 
                                //2. Check the Cener location and Normal Vector and return Value 
                                basePort = platePart.GetPort(TopologyGeometryType.Face, ContextTypes.Base);
                                offsetPort = platePart.GetPort(TopologyGeometryType.Face, ContextTypes.Offset);

                                if (basePort != null)
                                {
                                    // Case 1 centerPositionSurface.y < 0 and normalVector.Y > 0
                                    // Case 2 centerPositionSurface.y > 0 and normalVector.Y < 0

                                    surfaceGeometry = (ISurface)basePort.Geometry;
                                    if (surfaceGeometry != null)
                                    {
                                        base.GetSurfaceApproxNormalAndCenter(surfaceGeometry, out centerPositionSurface, out normalVector);

                                        if (centerPositionSurface != null && normalVector != null)
                                        {
                                            if (centerPositionSurface.Y * normalVector.Y < 0)
                                            {
                                                plateSide = (int)ManufacturingSide.Base;
                                            }
                                            else
                                            {

                                            }
                                        }
                                    }

                                }

                                if (offsetPort != null)
                                {

                                    surfaceGeometry = (ISurface)offsetPort.Geometry;
                                    if (surfaceGeometry != null)
                                    {
                                        base.GetSurfaceApproxNormalAndCenter(surfaceGeometry, out centerPositionSurface, out normalVector);


                                        // Case 1 centerPositionSurface.y < 0 and normalVector.Y > 0
                                        // Case 2 centerPositionSurface.y > 0 and normalVector.Y < 0
                                        if (centerPositionSurface != null && normalVector != null)
                                        {
                                            if (centerPositionSurface.Y * normalVector.Y < 0)
                                            {
                                                plateSide = (int)ManufacturingSide.Offset;
                                            }
                                            else
                                            {

                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    //
                                }

                            }
                        }



                        #endregion
                        break;
                    case "OUT":
                        #region Out

                        // Check Longitudinal Plate or Hull 
                        plateSide = 0; //UnDefinedSide

                        if (platePart != null)
                        {
                            if (platePart.Type == PlateType.LBulkheadPlate || platePart.Type == PlateType.Hull)
                            {
                                //Logic
                                //1. Get the Center and Normal Vector approximately. 
                                //2. Check the Cener location and Normal Vector and return Value 

                                basePort = platePart.GetPort(TopologyGeometryType.Face, ContextTypes.Base);
                                offsetPort = platePart.GetPort(TopologyGeometryType.Face, ContextTypes.Offset);


                                if (basePort != null)
                                {
                                    surfaceGeometry = (ISurface)basePort.Geometry;
                                    if (surfaceGeometry != null)
                                    {
                                        base.GetSurfaceApproxNormalAndCenter(surfaceGeometry, out centerPositionSurface, out normalVector);
                                    }

                                    // Case 1 centerPositionSurface.y > 0 and normalVector.Y > 0
                                    // Case 2 centerPositionSurface.y < 0 and normalVector.Y < 0
                                    if (centerPositionSurface != null && normalVector != null)
                                    {
                                        if (centerPositionSurface.Y * normalVector.Y > 0)
                                        {
                                            plateSide = (int)ManufacturingSide.Base;
                                        }
                                        else
                                        {

                                        }
                                    }

                                }
                                if (offsetPort != null)
                                {

                                    surfaceGeometry = (ISurface)offsetPort.Geometry;
                                    if (surfaceGeometry != null)
                                    {
                                        base.GetSurfaceApproxNormalAndCenter(surfaceGeometry, out centerPositionSurface, out normalVector);
                                    }

                                    // Case 1 centerPositionSurface.y > 0 and normalVector.Y > 0
                                    // Case 2 centerPositionSurface.y < 0 and normalVector.Y < 0
                                    if (centerPositionSurface != null && normalVector != null)
                                    {
                                        if (centerPositionSurface.Y * normalVector.Y > 0)
                                        {
                                            plateSide = (int)ManufacturingSide.Offset;
                                        }
                                        else
                                        {

                                        }
                                    }
                                }
                                else
                                {
                                    //
                                }

                            }
                        }

                        #endregion
                        break;
                    case "ALONGDIRECTION":
                        #region AlongDirection

                        //Logic
                        //1. Get the Center and Normal Vector approximately. 

                        double dirX = Convert.ToDouble(processInfo.GetArguments("DirX").FirstOrDefault().Value);
                        double dirY = Convert.ToDouble(processInfo.GetArguments("DirY").FirstOrDefault().Value);
                        double dirZ = Convert.ToDouble(processInfo.GetArguments("DirZ").FirstOrDefault().Value);
                        //Radian 
                        double angleTol = Convert.ToDouble(processInfo.GetArguments("Angle").FirstOrDefault().Value);

                        Vector directionVec = new Vector(dirX, dirY, dirZ);
                        double angle = 0.0;

                        if (platePart != null)
                        {
                            basePort = platePart.GetPort(TopologyGeometryType.Face, ContextTypes.Base);
                            offsetPort = platePart.GetPort(TopologyGeometryType.Face, ContextTypes.Offset);
                        }

                        if (basePort != null)
                        {
                            surfaceGeometry = (ISurface)basePort.Geometry;
                            if (surfaceGeometry != null)
                            {
                                base.GetSurfaceApproxNormalAndCenter(surfaceGeometry, out centerPositionSurface, out normalVector);
                                normalVector.Length = 1.0;

                                //Radian 0<= angle <= PI
                                angle = Math.Acos(directionVec.Dot(normalVector));
                            }

                            if (angle <= angleTol)
                                plateSide = (int)ManufacturingSide.Base;

                        }

                        if (offsetPort != null)
                        {

                            surfaceGeometry = (ISurface)offsetPort.Geometry;
                            if (surfaceGeometry != null)
                            {
                                base.GetSurfaceApproxNormalAndCenter(surfaceGeometry, out centerPositionSurface, out normalVector);
                                normalVector.Length = 1.0;

                                //Radian 0<= angle <= PI
                                angle = Math.Acos(directionVec.Dot(normalVector));
                            }

                            if (angle <= angleTol)
                                plateSide = (int)ManufacturingSide.Offset;
                        }


                        #endregion
                        break;
                    default:
                        plateSide = 0;
                        break;

                }

                #endregion

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs
                processInfo.SetAttribute((int)ProcessInfo.ProcessValues.PlateUpSide, plateSide);
                #endregion
            }
            catch(Exception e)
            {
                LogForToDoList(1050, "Problem occurred in PlateUpSide Rule" + e.Message);
            }
        }
    }
}
