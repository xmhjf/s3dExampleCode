//----------------------------------------------------------------------------------
//      Copyright (C) 2014 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Extension rule returns the extension values of the TemplateSet.    
//
//      Author:  
//
//      History:
//      May 28th, 2014   Created by Natilus-India
//
//-----------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;
using TemplateProcessInfo = Ingr.SP3D.Content.Manufacturing.TemplateProcessInfo;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// TemplateSet Extension Rule.
    /// </summary>
    public class Extension : UserDefinedValues
    {
        /// <summary>
        /// Gets the TemplateSet extension values.
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
                // Get the Extension values to be used for the template beyond the seam defintion.

                TemplateSetType templateSetType = TemplateSetType.Plate;
                if (mfgTemplateSet != null)
                    templateSetType = mfgTemplateSet.Type;

                Dictionary<string, object> results = GetExtensionValues(processInfo, templateSetType);

                #endregion Processing

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs

                int key = (int)TemplateProcessInfo.ProcessValues.Extension;
                processInfo.SetAttribute(key, results);

                #endregion Set Outputs
            }
            catch (Exception e)
            {
                LogForToDoList(e,  3006, "Call to TemplateSet Extension Rule failed with the error" + e.Message);
            }
        }

        #region Private Methods

        /// <summary>
        /// Provides the Extension information.
        /// </summary>
        /// <param name="processInfo">The process info.</param>
        /// <param name="templateSetType">TemplateSet Type.</param>
        /// <returns>The Extension Values.</returns>
        private Dictionary<string, object> GetExtensionValues(ProcessInformation processInfo, TemplateSetType templateSetType)
        {
            Dictionary<string, object> results = new Dictionary<string, object>();

            switch (templateSetType)
            {
                case TemplateSetType.Plate:
                    {
                        // ***Note: Uncomment the below lines of code if the values defined in XML need to be overwritten.
                       
                        //double extension = 0.05;
                        //double offset = 0.0;
                        //double minimumExtension = 0.0;

                        //results.Add("Extension", extension);
                        //results.Add("Offset", offset);
                        //results.Add("MinimumExtension", minimumExtension);

                        break;
                    }
                case TemplateSetType.ProfileFace:
                    {
                        //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:++:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:++:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                        //  LowerSidePortion  -> Portion of Bottom Line needed on the LowerSide of BCL. 1->complete 0->Not needed 0.5->Half of the LowerSide Bottom line
                        //  LowerSideFixedExtension -> Fixed Extension to be applied on the Lower Side
                        //  LowerSideOffsetExtension -> Offset Extension to be applied on the Lower Side
                        //  LowerSideMinimumExtensionForOffset -> Minimum Extension for Offset Extension on the Lower Side
                        //
                        //  UpperSidePortion -> Portion of Bottom Line needed on the UpperSide of BCL. 1->complete 0->Not needed 0.5->Half of the UpperSide Bottom line
                        //  UpperSideFixedExtension -> Fixed Extension to be applied on the Upper Side
                        //  UpperSideOffsetExtension -> Offset Extension to be applied on the Upper Side
                        //  UpperSideMinimumExtensionForOffset ->Minimum Extension for Offset Extension on the Upper Side
                        //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:++:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:++:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+

                        // ***Note: Uncomment the below lines of code if the values defined in XML need to be overwritten.

                        //double lowerSidePortion = 1.0;
                        //double lowerSideFixedExtension = 0.05;
                        //double lowerSideOffsetExtension = 0.0;
                        //double lowerSideMinimumExtensionForOffset = 0.0;

                        //double upperSidePortion = 1.0;
                        //double upperSideFixedExtension = 0.05;
                        //double upperSideOffsetExtension = 0.0;
                        //double upperSideMinimumExtensionForOffset = 0.0;

                        //results.Add("LowerSidePortion", lowerSidePortion);
                        //results.Add("LowerSideFixedExtension", lowerSideFixedExtension);
                        //results.Add("LowerSideOffsetExtension", lowerSideOffsetExtension);
                        //results.Add("LowerSideMinimumExtensionForOffset", lowerSideMinimumExtensionForOffset);

                        //results.Add("UpperSidePortion", upperSidePortion);
                        //results.Add("UpperSideFixedExtension", upperSideFixedExtension);
                        //results.Add("UpperSideOffsetExtension", upperSideOffsetExtension);
                        //results.Add("UpperSideMinimumExtensionForOffset", upperSideMinimumExtensionForOffset);

                        break;
                    }
                case TemplateSetType.ProfileEdge:
                    {
                        //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:++:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:++:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                        //extensionBase: - The Extension value to be applied on base side
                        //extensionOffset: - The Extension value to be applied on offset side
                        //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:++:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:++:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+

                        // ***Note: Uncomment the below lines of code if the values defined in XML need to be overwritten.

                        //double extensionBase = 0.1;
                        //double extensionOffset = 0.1;

                        //results.Add("ExtensionBase", extensionBase);
                        //results.Add("ExtensionOffset", extensionOffset);

                        break;
                    }
                case TemplateSetType.Tube:
                    {
                        //+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:++:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:++:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                        //' Determines the process settings for the Template extension, which will be applied to both
                        //' the left and right side of the tube template
                        //'               Properties:
                        //'                   method: - How the extension is measured;
                        //'                           0 : Linear extension
                        //'                           1 : Perpendicular
                        //'                           2 : Along edge
                        //'                   measure: - how to measure the extension offset to the tube edge.
                        //'                           0 : Along girth
                        //'                           1 : Along tube axis
                        //'                   value1: - The offset value to be applied
                        //'                   value2: - reserved distance value for new method/offset type
                        //+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:++:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:++:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+

                        // ***Note: Uncomment the below lines of code if the values defined in XML need to be overwritten.

                        //int methodLeft = 2;
                        //int methodRight = 2;
                        //int measureLeft = 1;
                        //int measureRight = 1;

                        //double value1Left = 90.0;
                        //double value1Right = 90.0;
                        //double value2Left = 0.0;
                        //double value2Right = 0.0;

                        //results.Add("MethodLeft", methodLeft);
                        //results.Add("MethodRight", methodRight);
                        //results.Add("MeasureLeft", measureLeft);
                        //results.Add("MeasureRight", measureRight);

                        //results.Add("Value1Left", value1Left);
                        //results.Add("Value1Right", value1Right);
                        //results.Add("Value2Left", value2Left);
                        //results.Add("Value2Right", value2Right);

                        break;
                    }
            }

            return results;
        }

        #endregion Private Methods
    }
}
