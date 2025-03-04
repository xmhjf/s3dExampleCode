//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Position rule returns the edge offset, number of templates, collection of frames and if it is a support edge of the TemplateSet.    
//
//      Author:  
//
//      History:
//      May 28th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Grids.Middle;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using TemplateProcessInfo = Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// TemplateSet Position Rule.
    /// </summary>

    public class Position : UserDefinedValuesRule
    {
        /// <summary>
        /// Gets the TemplateSet position related values.
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

                TemplateSet mfgTemplateSet = null;
                BusinessObject parentPart = null;

                PlatePartBase platePart = null;
                ProfilePart profilePart = null;

                if (processInfo.ManufacturingPart != null)
                {
                    mfgTemplateSet = (TemplateSet)processInfo.ManufacturingPart;
                }
                if (processInfo.ManufacturingParent != null)
                {
                    parentPart = (BusinessObject)processInfo.ManufacturingParent;

                    if (mfgTemplateSet != null)
                    {
                        if (mfgTemplateSet.Type == TemplateSetType.Plate)
                        {
                            platePart = (PlatePartBase)parentPart;
                        }
                        else if (mfgTemplateSet.Type == TemplateSetType.ProfileFace)
                        {
                            profilePart = (ProfilePart)parentPart;
                        }
                    }
                    else
                        throw new ArgumentNullException("TemplateSet is empty");
                }

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing                

                double edgeOffset = 0.0;
                int numberOfTemplates = 0;
                bool supportEdges = false;
                ReadOnlyCollection<GridPlaneBase> gridPlanes = null;

                // Get the edgeOffset, numberOfTemplates, supportEdges, collection of gridPlanes.
                if (mfgTemplateSet == null)
                    return;

                GetPositionData(processInfo, mfgTemplateSet, out edgeOffset, out numberOfTemplates, out supportEdges, out gridPlanes);

                Dictionary<string, object> results = new Dictionary<string, object>();

                results.Add("EdgeOffset", edgeOffset);
                results.Add("NumberOfTemplates", numberOfTemplates);
                results.Add("SupportEdges", supportEdges);
                results.Add("GridPlanes", gridPlanes);

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs
                processInfo.SetAttribute((int)TemplateProcessInfo.ProcessValues.Position, results);
                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e, 3010, "Call to TemplateSet Position Rule failed with the error" + e.Message);
            }

        }

        #region Private Methods

        /// <summary>
        /// Gets the edgeOffset, numberOfTemplates, supportEdges and the collection of GridPlanes.
        /// </summary>
        /// <param name="processInfo">The process information.</param>
        /// <param name="mfgTemplateSet">The MFG template set.</param>
        /// <param name="edgeOffset">The edge offset.</param>
        /// <param name="numberOfTemplates">The number of templates.</param>
        /// <param name="supportEdges">if set to <c>true</c> [support edges].</param>
        /// <param name="gridPlanes">The grid planes.</param>
        /// <exception cref="System.NullReferenceException">
        /// The Input ProcessInfo is null or The Input MfgTemplateSet is null
        /// </exception>
        private void GetPositionData(ProcessInformation processInfo, TemplateSet mfgTemplateSet,
                                    out double edgeOffset, out int numberOfTemplates, out bool supportEdges, out ReadOnlyCollection<GridPlaneBase> gridPlanes)
        {
            if (processInfo == null)
            {
                throw new NullReferenceException("The Input ProcessInfo is null");

            }

            if (mfgTemplateSet == null)
            {
                throw new NullReferenceException("The Input MfgTemplateSet is null");
            } 

            edgeOffset = 0.0;
            numberOfTemplates = 0;
            supportEdges = false;
            gridPlanes = null;

            //Get Arguments - PositionAttributeControl
            // 1		= Frame
            // 2		= Edge
            // 3		= Even
            Dictionary<int, object> args = processInfo.GetArguments("PositionControl");
            if (args.Count == 0)
            {
                return;
            }
               
            if( ((args.ContainsValue("1") == true) && (args.ContainsValue("2") == true)) || // "FramesAndEdges"
                (args.ContainsValue("1") == true))  // "FramesOnly"
            {
                edgeOffset = Convert.ToDouble(processInfo.GetArguments("EdgeOffset").FirstOrDefault().Value);
                numberOfTemplates = Convert.ToInt32(processInfo.GetArguments("NumberOfTemplates").FirstOrDefault().Value);
                supportEdges = Convert.ToBoolean(processInfo.GetArguments("SupportEdges").FirstOrDefault().Value);
                gridPlanes = GetGridPlanes(processInfo, mfgTemplateSet);
                      
            } 
            else if( ((args.ContainsValue("3") == true) && (args.ContainsValue("2") == true)) ||   // "EvenAndEdges"
                (args.ContainsValue("3") == true))  // "EvenOnly"
            {

                double length = 0.0;
                if (mfgTemplateSet.Type == TemplateSetType.Plate)
                {
                    PlatePartBase platePart = (PlatePartBase)mfgTemplateSet.DetailedPart;
                    length = 1.0; // TODO 1m- (double)platePart.GetPropertyValue("IJUASPSFrameFndn", "PlateLength");
                }
                else if (mfgTemplateSet.Type == TemplateSetType.ProfileFace)
                {
                    ProfilePart profilePart = (ProfilePart)mfgTemplateSet.DetailedPart;
                    length = 1.0; // TODO 1m- (double)profilePart.GetPropertyValue("IJUASPSFrameFndn", "WebLength");
                }

                double lengthTolerance1  = Convert.ToDouble(processInfo.GetArguments("LengthTolerance1").FirstOrDefault().Value);
                double lengthTolerance2 = Convert.ToDouble(processInfo.GetArguments("LengthTolerance2").FirstOrDefault().Value);

                if (length < lengthTolerance1) 
                {
                    edgeOffset = 0.1;
                    numberOfTemplates = 3;
                }
                else
                {
                    if (length < lengthTolerance2)
                    {
                        edgeOffset = 0.2;
                        numberOfTemplates = 5;
                    }
                    else
                    {
                        edgeOffset = 0.3;
                        numberOfTemplates = 7;
                    }
                }

                supportEdges = Convert.ToBoolean(processInfo.GetArguments("SupportEdges").FirstOrDefault().Value);
                gridPlanes = null;

            }
            else if (args.ContainsValue("2") == true)  // "EdgesOnly"            
            {
                edgeOffset = Convert.ToDouble(processInfo.GetArguments("EdgeOffset").FirstOrDefault().Value);
                numberOfTemplates = Convert.ToInt32(processInfo.GetArguments("NumberOfTemplates").FirstOrDefault().Value);
                supportEdges = Convert.ToBoolean(processInfo.GetArguments("SupportEdges").FirstOrDefault().Value);
                gridPlanes = null;

            }
            else
            {
                return;
            }       

        }

        /// <summary>
        /// Gets the intersecting GridPlanes according to the template direction.
        /// </summary>
        /// <param name="processInfo">The process info.</param>
        /// <param name="mfgTemplateSet">The TemplateSet.</param>
        /// <returns>The intersecting GridPlanes.</returns>
        private ReadOnlyCollection<GridPlaneBase> GetGridPlanes(ProcessInformation processInfo, TemplateSet mfgTemplateSet)
        {

            if (processInfo == null)
            {
                throw new NullReferenceException("The Input ProcessInfo is null");

            }

            if (mfgTemplateSet == null)
            {
                throw new NullReferenceException("The Input MfgTemplateSet is null");
            } 
         
            Collection<GridPlaneBase> allGridPlanes = new Collection<GridPlaneBase>();
            try
            {
                // Return 'null' if frame system is not set on the TemplateSet.
                if (mfgTemplateSet.CoordinateSystem == null)
                {
                    return null;
                }

                CodelistItem codelistItem;
                string interfaceName;
                int templateDirection;

                PlatePartBase platePart = null;
                ProfilePart profilePart = null;
                TopologyPort port = null;
                double xValue = 0.0, yValue = 0.0, zValue = 0.0;

                if (mfgTemplateSet.Type == TemplateSetType.Plate)
                {
                    platePart = (PlatePartBase)processInfo.ManufacturingParent;
                    port = platePart.GetPort(TopologyGeometryType.Face, ContextTypes.Base,-1);
                    if (port == null)
                    {
                        port = platePart.GetPort(TopologyGeometryType.Face, ContextTypes.Offset,-1);
                    }

                    interfaceName = "IJUAMfgTemplateProcessPlate";
                }
                else // (mfgTemplateSet.Type == TemplateSetType.ProfileFace)
                {
                    profilePart = (ProfilePart)processInfo.ManufacturingParent;
                    port = profilePart.GetPort(TopologyGeometryType.Face, -1, -1, ContextTypes.Lateral, (int)SectionFaceType.Web_Left, false);
                    if (port == null)
                    {
                        port = profilePart.GetPort(TopologyGeometryType.Face, -1, -1, ContextTypes.Lateral, (int)SectionFaceType.Web_Right, false);
                    }

                    interfaceName = "IJUAMfgTemplateProcessProfile";
                }

                //Get TemplateSet Direction
                codelistItem = mfgTemplateSet.GetSettingValue(SettingType.Process, interfaceName, "Direction");
                templateDirection = codelistItem.Value;

                Vector directionVector = null;

                if (templateDirection == (int)TemplateProcessInfo.Direction.Transversal)
                {
                    directionVector = mfgTemplateSet.CoordinateSystem.XAxis;
                }
                else if (templateDirection == (int)TemplateProcessInfo.Direction.Longitudinal)
                {
                    directionVector = mfgTemplateSet.CoordinateSystem.YAxis;
                }
                else if (templateDirection == (int)TemplateProcessInfo.Direction.Waterline)
                {
                    directionVector = mfgTemplateSet.CoordinateSystem.ZAxis;
                }
                else // PerpendicularToAxis or AlongAxis
                {
                    //TODO
                    // Get the Landing curve of Profile
                    // Get the vector from start point to end point of landing curve.

                    Vector landingCurveVector = null; // TODO
                    Vector vector1 = null;

                    Common.Middle.Position startPos = null;
                    Common.Middle.Position endPos = null;

                    ICurve landingCurve = EntityService.GetLandingCurve(profilePart);
                    if (landingCurve != null)
                    {
                        landingCurve.EndPoints(out startPos, out endPos);
                        landingCurveVector = endPos.Subtract(startPos);
                    }
                    else
                        return null;
                    

                    if (templateDirection == (int)TemplateProcessInfo.Direction.PerpendicularToAxis)
                    {
                        vector1 = landingCurveVector;
                    }
                    else //if (templateDirection == (int)TemplateProcessInfo.TemplateProcessDirection.AlongAxis)
                    {
                        if (port != null && landingCurveVector != null)
                        {
                            Common.Middle.Position avgRootPos;
                            Common.Middle.Vector avgNormal;
                            EntityService.GetPlatePartAvgPointAvgNormal((ISurface)port.Geometry, false, out avgRootPos, out avgNormal);
                            vector1 = avgNormal.Cross(landingCurveVector);
                        }
                    }

                    if (vector1 != null)
                    {
                        xValue = Math.Abs(vector1.Dot(mfgTemplateSet.CoordinateSystem.XAxis));
                        yValue = Math.Abs(vector1.Dot(mfgTemplateSet.CoordinateSystem.YAxis));
                        zValue = Math.Abs(vector1.Dot(mfgTemplateSet.CoordinateSystem.ZAxis));
                    }

                    if ((xValue > yValue) && (xValue > zValue))
                    {
                        directionVector = mfgTemplateSet.CoordinateSystem.XAxis;
                    }
                    else if ((yValue > xValue) && (yValue > zValue))
                    {
                        directionVector = mfgTemplateSet.CoordinateSystem.YAxis;
                    }
                    else if ((zValue > xValue) && (zValue > yValue))
                    {
                        directionVector = mfgTemplateSet.CoordinateSystem.ZAxis;
                    }
                }

                ReadOnlyCollection<GridPlaneBase> gridPlanesInPrimaryDirection = null;

                if (directionVector != null)
                    gridPlanesInPrimaryDirection = EntityService.GetPlanesInRange(port, directionVector, mfgTemplateSet.CoordinateSystem);
                else
                    return null;
 
                if (gridPlanesInPrimaryDirection != null)
                {
                    foreach (GridPlaneBase gridPlane in gridPlanesInPrimaryDirection)
                    {
                        allGridPlanes.Add(gridPlane);
                    }
                }

                string templateType = (string)processInfo.GetAttribute((int)TemplateProcessInfo.ProcessValues.Type, "TemplateType");
                if (templateType == "BOX" || templateType == "USERDEFINEDBOX" || templateType == "USERDEFINED BOX WITH EDGES")
                {                    
                    Vector secondaryDirectionVector = null;

                    if (port != null)
                    {
                        Common.Middle.Position cenetr = null;
                        Common.Middle.Vector normal = null;
                        TopologySurface surface = (TopologySurface)port.Geometry;
                        EntityService.FindApproxCenterAndNormal(surface, out cenetr, out normal);
                        xValue = Math.Abs(mfgTemplateSet.CoordinateSystem.XAxis.Dot(normal));
                        yValue = Math.Abs(mfgTemplateSet.CoordinateSystem.YAxis.Dot(normal));
                        zValue = Math.Abs(mfgTemplateSet.CoordinateSystem.ZAxis.Dot(normal));
                    }

                    if ((xValue > yValue) && (xValue > zValue))
                    {
                        secondaryDirectionVector = mfgTemplateSet.CoordinateSystem.XAxis.Cross(directionVector);
                    }
                    else if ((yValue > xValue) && (yValue > zValue))
                    {
                        secondaryDirectionVector = mfgTemplateSet.CoordinateSystem.YAxis.Cross(directionVector);
                    }
                    else if ((zValue > xValue) && (zValue > yValue))
                    {
                        secondaryDirectionVector = mfgTemplateSet.CoordinateSystem.ZAxis.Cross(directionVector);
                    }

                    ReadOnlyCollection<GridPlaneBase> gridPlanesInSecondaryDirection = null;

                    if(secondaryDirectionVector != null)
                        gridPlanesInSecondaryDirection = EntityService.GetPlanesInRange(port, secondaryDirectionVector, mfgTemplateSet.CoordinateSystem);

                    if (gridPlanesInSecondaryDirection != null)
                    {
                        foreach (GridPlaneBase gridPlane in gridPlanesInSecondaryDirection)
                        {
                            allGridPlanes.Add(gridPlane);
                        }
                    }
                }

                
            }
            catch 
            {
            }
            
            return new ReadOnlyCollection<GridPlaneBase>(allGridPlanes);
        }

        #endregion Private Methods
    }
}

