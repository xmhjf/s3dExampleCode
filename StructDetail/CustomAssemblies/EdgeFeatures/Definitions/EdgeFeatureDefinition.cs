//--------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  EdgeFeatureDefinition.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMEdgeFeatureRules.dll
//  Original Class Name: EdgeDef in VB content
//
//Abstract:
//  EdgeFeatureDefinition is a .NET custom assembly definition which creates the required assembly outputs. 
//  This class subclasses from EdgeFeatureCustomAssemblyDefinition.
//
//Change History:
//  dd.mmm.yyyy    who    change description
//--------------------------------------------------------------------------------------------------------------------
using System;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Implementation of Edge feature .NET custom assembly definition class.
    /// EdgeFeatureDefinition is a .NET custom assembly definition which creates a free edge treatment optionally.   
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    [DefaultLocalizer(EdgeFeaturesResourceIdentifiers.EdgeFeaturesResources, EdgeFeaturesResourceIdentifiers.EdgeFeaturesAssembly)]
    public class EdgeFeatureDefinition : EdgeFeatureCustomAssemblyDefinition
    {
        //=========================================================================================================
        //DefinitionName/ProgID of this symbol is "EdgeFeatures,Ingr.SP3D.Content.Structure.EdgeFeatureDefinition"
        //=========================================================================================================

        #region Definitions of assembly outputs

        //Creates a free edge treatment
        [AssemblyOutput(1, DetailingCustomAssembliesConstants.EdgeFeatureFET1)]
        public AssemblyOutput freeEdgeTreatment;

        #endregion Definitions of assembly outputs

        #region Public override functions and methods

        /// <summary>
        /// Validates the edge feature and creates the free edge treatment as AssemblyOutput, 
        /// if it is required, for this definition based on selector answers.
        /// </summary>
        public override void EvaluateAssembly()
        {
            try
            {
                //Get the occurrence from the edge feature
                Feature edgeFeature = (Feature)base.Occurrence;

                //Validating inputs
                ValidateEdgeFeatureInputs();

                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created while validating the inputs
                    return;
                }

                base.AddFeatureGeometry();

                Ingr.SP3D.Interop.StructDetailUtil.StructDetailHelper oHlpr = new Ingr.SP3D.Interop.StructDetailUtil.StructDetailHelper();
                if (!(oHlpr.IsPartialDetailed((Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO((Ingr.SP3D.Common.Middle.BusinessObject)(base.EdgePort.Connectable))))))
                {                                    
                    //Construct the output objects for the assembly.
                    if (base.IsFreeEdgeTreatmentNeeded(DetailingCustomAssembliesConstants.ApplyTreatment, (int)Answer.Yes))
                    {
                        //Only construct the Free edge treatment if not generated yet and add it is as output
                        if (this.freeEdgeTreatment.Output == null)
                        {
                            EdgeTreatment edgeTreatment = GetFreeEdgeTreatment(DetailingCustomAssembliesConstants.RootEdgeTreatment);
                            if (edgeTreatment != null)
                            {
                                this.freeEdgeTreatment.Output = edgeTreatment;
                            }
                            else
                            {
                                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(EdgeFeaturesResourceIdentifiers.ToDoFreeEdgeTreatment,
                                    "Free edge treatment is not created."));
                            }
                        }
                    }
                    else
                    {
                        //if assembly output is not required now, delete it if it has been previously created
                        if (this.freeEdgeTreatment.Output != null)
                        {
                            this.freeEdgeTreatment.Delete();
                        }
                    }
                }
                oHlpr = null;
            }
            catch (Exception)
            {
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(EdgeFeaturesResourceIdentifiers.ToDoEvaluateAssembly,
                        "Unexpected error while evaluating the edge feature custom assembly definition."));
                }
            }
        }
        #endregion Public override functions and methods
    }
}
