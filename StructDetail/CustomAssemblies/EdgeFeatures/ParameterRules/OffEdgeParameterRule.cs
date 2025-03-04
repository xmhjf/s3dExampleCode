//--------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  EdgeFeatureParameterRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMEdgeFeatureRules.dll
//  Original Class Name: ‘OffEdgeParm’ in VB content
//
//Abstract
//	OffEdgeParameterRule is a .NET parameter rule which is defining the custom parameter rule. 
//  This class subclasses from EdgeFeatureParameterRule. 
//---------------------------------------------------------------------------------------------------------------
using System;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Parameter rule which defining the custom parameter rule in context of the edge feature.
    /// It computes the item parameters for the edge feature and assigns them.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [DefaultLocalizer(EdgeFeaturesResourceIdentifiers.EdgeFeaturesResources, EdgeFeaturesResourceIdentifiers.EdgeFeaturesAssembly)]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get parameter value. 
    [RuleInterface(DetailingCustomAssembliesConstants.IASMEdgeFeatureRules_OffEdgeParm, DetailingCustomAssembliesConstants.IASMEdgeFeatureRules_OffEdgeParm)]
    public class OffEdgeParameterRule : EdgeFeatureParameterRule
    {
        //=======================================================================================================
        //DefinitionName/ProgID of this symbol is "EdgeFeatures,Ingr.SP3D.Content.Structure.OffEdgeParameterRule"
        //=======================================================================================================

        #region Parameters
        [ControlledParameter(1, DetailingCustomAssembliesConstants.AdjustedOffset, DetailingCustomAssembliesConstants.AdjustedOffset)]
        public ControlledParameterDouble adjustedOffset;
        #endregion Parameters

        /// <summary>
        /// Evaluates the parameter rule It validates the inputs of the edge feature.
        /// It computes the item parameters in the context of the smart occurrence and assigns them.
        /// </summary>
        public override void Evaluate()
        {
            try
            {
                //Get the edge feature from the occurrence
                Feature edgeFeature = (Feature)base.Occurrence;

                //Validating inputs
                ValidateEdgeFeatureInputs();

                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created while validating the inputs
                    return;
                }

                double bevelDepth = 0.0;

                PhysicalConnection physicalConnectionAtEdgeFeaturePort = base.PhysicalConnectionAtEdgeFeaturePort;
                //Check if the Physical connection nearest to the Edge Feature.
                if (physicalConnectionAtEdgeFeaturePort != null)
                {
                    //we have a physical connection available at the given edge feature, get its bevel depth on both sides of the plate.
                    double refSideBevelDepth = EdgeFeatureServices.GetBevelDepth(edgeFeature, physicalConnectionAtEdgeFeaturePort, true);
                    double antiRefSideBevelDepth = EdgeFeatureServices.GetBevelDepth(edgeFeature, physicalConnectionAtEdgeFeaturePort, false);

                    //use the larger depth amongst 
                    bevelDepth = (refSideBevelDepth > (antiRefSideBevelDepth - MarineSymbolConstants.CompareTolerance)) ? refSideBevelDepth : antiRefSideBevelDepth;
                }

                //default clearance from Bevel, this is the minimum clearance that should be maintained between EF and Weld
                double clearanceOffBevel = 0.005;

                //Getting the offset value of the edge feature
                double offset = StructHelper.GetDoubleProperty(edgeFeature, DetailingCustomAssembliesConstants.IJUASmartEdgeFeature, DetailingCustomAssembliesConstants.Offset);

                //check if existing edge feature offset is more than the bevel depth and clearance then we are OK, 
                //otherwise adjust the offset value with bevelDepth + clearanceOffBevel
                if ((bevelDepth + clearanceOffBevel) > (offset - MarineSymbolConstants.CompareTolerance))
                {
                    offset = bevelDepth + clearanceOffBevel;
                }

                this.adjustedOffset.Value = offset;
            }
            catch (Exception)
            {
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(EdgeFeaturesResourceIdentifiers.ToDoParameterRule,
                        "Unexpected error while evaluating Edge feature parameter rule."));
                }
            }
        }
    }
}