//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Determines the process settings for the Edge Template offset, which will be applied to find location of bottom curve.
//
//      Author:  
//
//      History:
//      May 28th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;
using TemplateProcessInfo = Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Offset Rule.
    /// </summary>
    public class Offset : UserDefinedValuesRule
    {
        /// <summary>
        /// Gets the TemplateSet Edge Template offset, which will be applied to find location of bottom curve.
        /// </summary>
        /// <param name="processInfo">The process info.</param>
        public override void Evaluate(ProcessInformation processInfo)
        {
            try
            {
                if (processInfo == null)
                    throw new ArgumentNullException("Input ProcessInfo is empty.");

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs

                TemplateSet mfgTemplateSet = null;
                BusinessObject parentPart = null;  
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
                        if (mfgTemplateSet.Type == TemplateSetType.ProfileEdge)
                        {
                            profilePart = (ProfilePart)parentPart;
                        }
                        else
                        {
                            return;
                        }
                    }
                    else
                    {
                        return;
                    }

                }

                // Check for valid side.
                int templateSide = Convert.ToInt32(processInfo.GetAttribute((int)TemplateProcessInfo.ProcessValues.Side, "Side"));
                if (templateSide == 0)
                {
                    return;
                }

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing
                
                double offset = 0.0;
                int sectionID = templateSide;

                if (profilePart != null)
                {
                    ReadOnlyCollection<TopologyPort> topologyPorts = profilePart.GetPorts(TopologyGeometryType.Edge, ContextTypes.Base, GeometryOperationTypes.PartFinalTrim, GraphPosition.Before);

                    foreach (TopologyPort port in topologyPorts)
                    {
                        if (port.SectionId == sectionID)
                        {
                            offset = port.Length;
                            break;
                        }
                    }
                }
                
                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                processInfo.SetAttribute((int)TemplateProcessInfo.ProcessValues.Offset, offset);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e, 3027, "Call to TemplateSet Offset Rule failed with the error" + e.Message);

            }

        }
    }
}
