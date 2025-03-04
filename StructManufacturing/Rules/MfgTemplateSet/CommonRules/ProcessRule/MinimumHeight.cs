//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   MinimumHeight rule returns the minimum height of the TemplateSet.    
//
//      Author:  
//
//      History:
//      May 28th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;
using TemplateProcessInfo = Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// MinimumHeight Rule.
    /// </summary>
    public class MinimumHeight : UserDefinedValuesRule
    {
        /// <summary>
        /// Gets the TemplateSet minimum height.
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

                PlatePartBase platePart = null;
                ProfilePart profilePart = null;
                MemberPart memberPart = null;

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
                        else if (mfgTemplateSet.Type == TemplateSetType.ProfileFace || mfgTemplateSet.Type == TemplateSetType.ProfileEdge)
                        {
                            profilePart = (ProfilePart)parentPart;
                        }
                        else if (mfgTemplateSet.Type == TemplateSetType.Tube)
                        {
                            memberPart = (MemberPart)parentPart;
                        }
                    }
                }

                #endregion GetInputs

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing
                // Set the Minimum Template Height depending on the TemplateSet Type.
                // Currently only returns a fixed value but this could be
                // in the future dependent on the workcenter assignment

                double minimumHeight = 0.0;

                if (mfgTemplateSet != null)
                {
                    switch (mfgTemplateSet.Type)
                    {
                        case TemplateSetType.Plate:
                            minimumHeight = 0.35;
                            break;
                        case TemplateSetType.ProfileFace:
                            minimumHeight = 0.35;
                            break;
                        case TemplateSetType.ProfileEdge:
                            minimumHeight = 0.35;
                            break;
                        case TemplateSetType.Tube:
                            minimumHeight = 0.5;
                            break;
                    }
                }

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                processInfo.SetAttribute((int)TemplateProcessInfo.ProcessValues.MinimumHeight, minimumHeight);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e, 3008, "Call to TemplateSet Minimum Height Rule failed with the error" + e.Message);
            }

        }
    }
}
