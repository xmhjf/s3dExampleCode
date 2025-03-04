//**************************************************************************************************************************/
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  CornerDefinition.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\CornerDef.dll
//  Original Class Name: ‘CornerDef’ in VB content
//
//Abstract
// CornerDefinition is a .NET custom assembly definition which creates the free EdgeTreatment and Collar CustomPlatePart.
// This class subclasses from CornerCustomAssemblyDefinition.
//
// Change History:
//  dd.mmm.yyyy    who    change description
//  12.oct.2016 svsmylav  DI-259156: Statement with 'NotifyAssemblyConnection' method call
//                             is removed (earlier it was used to update AC rules to avoid missing PC TDR).
//****************************************************************************************************************************/
using System;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Implementation of Corner Feature .NET custom assembly definition class.
    /// CornerDefinition is a .NET custom assembly definition which creates EdgeTreatment and Collar CustomPlatePart if required.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    [DefaultLocalizer(CornerFeaturesResourceIdentifiers.CornerFeaturesResources, CornerFeaturesResourceIdentifiers.CornerFeaturesAssembly)]
    public class CornerDefinition : CornerCustomAssemblyDefinition
    {
        //=====================================================================================================
        //DefinitionName/ProgID of this symbol is "CornerFeatures,Ingr.SP3D.Content.Structure.CornerDefinition"
        //=====================================================================================================

        #region Private members
        //AssemblyOutput names
        private const string FreeEdgeTreatment = "CornerFeatureFET1";
        private const string CollarCustomPlatePart = "CornerFeatureCollar";
        #endregion Private members

        #region Definitions of assembly outputs

        /// <summary>
        /// Free EdgeTreatment at Corner Feature.
        /// </summary>
        [AssemblyOutput(1, FreeEdgeTreatment)]
        public AssemblyOutput freeEdgeTreatment;

        /// <summary>
        /// Collar CustomPlatePart at Corner Feature.
        /// </summary>
        [AssemblyOutput(2, CollarCustomPlatePart)]
        public AssemblyOutput collar;

        #endregion Definitions of assembly outputs

        #region Public override properties and methods

        /// <summary>
        /// Validates the corner feature and creates the free edge treatment and collar as AssemblyOutputs,
        /// if they are required for this definition based on selector answers.
        /// </summary>
        public override void EvaluateAssembly()
        {
            try
            {
                //Validates the inputs of the Corner Feature
                base.ValidateInputs();

                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created while validating the inputs
                    return;
                }

                //It's only required to be explicitly called if the Corner Feature is not constructed using its 3D API constructor.
                base.AddFeatureToCut();

                //get the required assembly outputs for this definition           
                Dictionary<string, bool> requiredAssemblyOutputs = this.RequiredAssemblyOutputs;

                //Create the Free EdgeTreatment at corner only if needed, if not needed then delete the assembly output.
                base.CreateOrDeleteFreeEdgeTreatment(this.freeEdgeTreatment, DetailingCustomAssembliesConstants.RootEdgeTreatment, requiredAssemblyOutputs[FreeEdgeTreatment]);

                //need to pass the PlateCommonInputs when Corner Feature's face port is associated with ProfilePart else it will use the plate properties of the connectable of the Corner Feature's face port  
                PlateCommonInputs plateCommonInputs = null;
                if (base.FacePort.Connectable is ProfilePart && requiredAssemblyOutputs[CollarCustomPlatePart])
                {
                    PlateType plateType = PlateType.Standalone;
                    Material material = new CatalogStructHelper().GetMaterial("Steel - Carbon", "A");
                    int namingCategory = StructMarineHelper.GetPlateNamingCategoryCodelistValue(plateType, "TB");
                    plateCommonInputs = new PlateCommonInputs(plateType, material, 0.025, MoldedDirection.Above, Tightness.NonTight, namingCategory, MarineSymbolConstants.SpecA);
                }

                //Create the collar at corner only if needed if not needed then delete the assembly output.
                base.CreateOrDeleteCollarCustomPlatePart(this.collar, plateCommonInputs, DetailingCustomAssembliesConstants.RootCornerCollar, requiredAssemblyOutputs[CollarCustomPlatePart]);
            }
            catch (Exception)
            {
                //may be there could be specific ToDo record is created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(CornerFeaturesResourceIdentifiers.ToDoEvaluateAssembly,
                            "Unexpected error while evaluating the custom assembly of {0}"), this.ToString()));
                }
            }
        }

        #endregion Public override properties and methods

        #region Private Methods

        /// <summary>
        /// Gets all needed assembly outputs for the current configuration of the Corner Feature.
        /// </summary>
        /// <returns>This returns the needed outputs of the Corner Feature.</returns>
        private Dictionary<string, bool> RequiredAssemblyOutputs
        {
            get
            {
                //Dictionary which holds the data of the assembly output name and the Boolean which indicates
                //whether the corresponding assembly output is needed or not.
                Dictionary<string, bool> requiredAssemblyOutputs = new Dictionary<string, bool>();

                Feature cornerFeature = (Feature)base.Occurrence;

                bool isFreeEdgeTreatmentNeeded = false;

                Ingr.SP3D.Interop.StructDetailUtil.StructDetailHelper oHlpr = new Ingr.SP3D.Interop.StructDetailUtil.StructDetailHelper();
                bool isPartialDetailed = false;

                if (! (base.FacePort.Connectable is MemberPart))
                {
                    isPartialDetailed = (oHlpr.IsPartialDetailed(Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO((Ingr.SP3D.Common.Middle.BusinessObject)base.FacePort.Connectable)));
                }

                oHlpr = null;

                if (isPartialDetailed)
                {                    
                    //Default Value is False
                }
                else
                {
                    //in case of 'SMMbrACStd.EndCutCornerSel' selector rule, ApplyTreatment is not define as selector question which inturn returning the value as null. 
                    PropertyValue propertyValue = cornerFeature.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.ApplyTreatment);
                    if (propertyValue != null && ((PropertyValueCodelist)propertyValue).PropValue == (int)Answer.Yes)
                    {
                        isFreeEdgeTreatmentNeeded = true;
                    }
                }

                requiredAssemblyOutputs.Add(FreeEdgeTreatment, isFreeEdgeTreatmentNeeded);

                //If the RootCornerCollar part class is not bulk loaded then this functionality (Corner Collars) cannot be used.
                //This check is needed as the Bulkload is optional and users might or might not Bulkload some new part class as per their requirement.
                //If the part was bulk loaded then check if the collar is needed or not.
                bool isCollarCustomPlatePartNeeded = (new CatalogStructHelper().DoesPartOrPartClassExist(DetailingCustomAssembliesConstants.RootCornerCollar) && cornerFeature.PartName.Contains(DetailingCustomAssembliesConstants.WithCollar)) ? true : false;
                requiredAssemblyOutputs.Add(CollarCustomPlatePart, isCollarCustomPlatePartNeeded);

                return requiredAssemblyOutputs;
            }
        }

        #endregion Private Methods
    }
}
