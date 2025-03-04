//----------------------------------------------------------------------------------
//      Copyright (C) 2013 Intergraph Corporation.  All rights reserved.
//
//      Purpose:   Manufacturing Service Manager rule. 
//
//      Author:  
//
//      History:
//      November 14th, 2014   Created by Natilus-HSV
//
//-----------------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Content.Manufacturing;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Manufacturing
{
    /// <summary>
    /// Service Manager Rule for Custom Setting Inputs.  
    /// </summary>
    public class ServiceManager : ServiceManagerRule
    {
        /// <summary>
        /// Evaluates Service Manager Rule to set custom dependency objects
        /// </summary>
        /// <param name="customInputInfo">The Custom Input Information.</param>
        /// <exception cref="System.ArgumentNullException">Input Custom Input Information is empty</exception>
        public override void Evaluate(CustomInputInformation customInputInfo)
        {
            if (customInputInfo == null)
                throw new ArgumentNullException("Input CustomInputInformation is empty");

            try
            {
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //1. Get Inputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region GetInputs
                ManufacturingBase mfgObject = null;

                if (customInputInfo.MfgObject != null)
                {
                    mfgObject = customInputInfo.MfgObject;
                }

                #endregion

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //2. Processing 
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Processing

                #region Additional Tracking List
                //StructMfgServiceMgr_NoCustomInputs = 0,
                //StructMfgServiceMgr_CustomRule = 1,
                //StructMfgServiceMgr_PartProdRouting = 2,
                //StructMfgServiceMgr_PartParentAssembly = 4,
                //StructMfgServiceMgr_PartCommonPart = 8,
                //StructMfgServiceMgr_PartKnuckles = 16,
                //StructMfgServiceMgr_PartFeatures = 32,
                //StructMfgServiceMgr_PartOpenings = 64,
                //StructMfgServiceMgr_AllAssemblyShrinkages = 128,
                //StructMfgServiceMgr_AssemblyBasePlate = 256,
                //StructMfgServiceMgr_BasePlateStiffeners = 512,
                //StructMfgServiceMgr_BUAssemblyPCs = 1024,
                //StructMfgServiceMgr_PartPanel = 2048,
                //StructMfgServiceMgr_ConnectedParts = 4096

                Dictionary<string, long> additionalTrackingList = new Dictionary<string, long>();
                long markingLine, margin, partShrinkage, assemblyPartShrinkage, assemblyShrinkage, template, pinJig, mfgPlate, mfgProfile;
                markingLine = GetAccumulatedInputSettings(customInputInfo,"MarkingLine");
                margin = GetAccumulatedInputSettings(customInputInfo,"Margin");
                partShrinkage = GetAccumulatedInputSettings(customInputInfo,"PartShrinkage");
                assemblyPartShrinkage = GetAccumulatedInputSettings(customInputInfo,"AssemblyPartShrinkage");
                assemblyShrinkage = GetAccumulatedInputSettings(customInputInfo,"AssemblyShrinkage");
                template = GetAccumulatedInputSettings(customInputInfo,"Template");
                pinJig = GetAccumulatedInputSettings(customInputInfo,"PinJig");
                mfgPlate = GetAccumulatedInputSettings(customInputInfo,"MfgPlate");
                mfgProfile = GetAccumulatedInputSettings(customInputInfo,"MfgProfile");

                additionalTrackingList.Add("MarkingLine", markingLine);
                additionalTrackingList.Add("Margin", margin);
                additionalTrackingList.Add("Template", template);
                additionalTrackingList.Add("PinJig", pinJig);
                additionalTrackingList.Add("PartShrinkage", partShrinkage);
                additionalTrackingList.Add("AssemblyPartShrinkage", assemblyPartShrinkage);
                additionalTrackingList.Add("AssemblyShrinkage", assemblyShrinkage);
                additionalTrackingList.Add("MfgPlate", mfgPlate);
                additionalTrackingList.Add("MfgProfile", mfgProfile); 
                #endregion 

                #region DependencyObjects 
                ReadOnlyCollection<BusinessObject> dependencyObjects = null;
                #endregion 


                #endregion

                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                //3. Set Outputs
                //:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:+:++:+:+:+:+:+:+:+
                #region Set Outputs
                customInputInfo.DependencyObjects = dependencyObjects;
                customInputInfo.AdditionalTrackingList = additionalTrackingList;
                #endregion

            }
            catch(Exception e)
            {
                LogForToDoList(7, "Problem occurred in Manufacturing Service Manager Rule" + e.Message,Common.Middle.Services.ToDoMessageTypes.ToDoMessageError);
            }
        }


        /// <summary>
        /// Gets the accumulated input settings.
        /// </summary>
        /// <param name="customInputInfo">The custom input information.</param>
        /// <param name="argument">The argument.</param>
        /// <returns></returns>
        private int GetAccumulatedInputSettings(CustomInputInformation customInputInfo, string argument )
        {
            if (String.IsNullOrEmpty(argument) == true || String.IsNullOrWhiteSpace(argument) == true)
                return 0;

            try
            {
                Dictionary<int, object> customSettings = customInputInfo.GetArguments(argument);

                if (customSettings != null && customSettings.Count > 0)
                {
                    int value = 0;
                    foreach (var tempValue in customSettings)
                    {
                        value += Convert.ToInt32(tempValue.Value);
                    }
                    return value;
                }
                else
                    return 0;
            }
            catch (Exception)
            {
                return 0;
            }
        }
    }
}


