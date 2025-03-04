/**************************************************************************************************************/
//
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  SlotCDefinition.cs
//
//Abstract
//	SlotCDefinition is a .NET custom assembly definition which is defining the custom assembly definition. 
//  This class subclasses from SlotCustomAssemblyDefinition.
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMSlotRules.dll
//  Original Class Name: ‘SlotCDef’ in VB content
/**************************************************************************************************************/

using System;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Content.Structure.Services;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Implementation of slot C type .NET custom assembly definition class.
    /// SlotCDefinition is a .NET custom assembly definition which creates physical connections, edge treatment and corner features if required.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    [DefaultLocalizer(PenetrationsResourceIds.PenetrationsResource, PenetrationsResourceIds.PenetrationsAssembly)]
    public class SlotCDefinition : SlotCustomAssemblyDefinition
    {
        //===================================================================================================
        //DefinitionName/ProgID of this symbol is "Penetrations,Ingr.SP3D.Content.Structure.SlotCDefinition"
        //===================================================================================================

        #region Definitions of assembly outputs
        /// <summary>
        /// Creates a Physical Connection with the bounding mapped WebLeft face port.
        /// </summary>
        [AssemblyOutput(1, "SlotCPC1")]
        public AssemblyOutput leftPhysicalConnection;

        /// <summary>
        /// Creates a free edge treatement along the slot edge.
        /// </summary>
        [AssemblyOutput(2, "SlotCFET1")]
        public AssemblyOutput freeEdgeTreatment;

        /// <summary>
        /// Creates a physical connection with the bounding mapped TopFlange face port.
        /// </summary>
        [AssemblyOutput(3, "SlotC_PC_Top")]
        public AssemblyOutput topPhysicalConnection;

        /// <summary>
        /// Creates a corner feature at bottom left slot corner.
        /// </summary>
        [AssemblyOutput(4, "SlotC_BottomLeftCF")]
        public AssemblyOutput bottomLeftCornerFeature;

        /// <summary>
        /// Creates a corner feature at top left slot corner.
        /// </summary>
        [AssemblyOutput(5, "SlotC_TopLeftCF")]
        public AssemblyOutput topLeftCornerFeature;
        #endregion Definitions of assembly outputs

        #region Public override functions and methods

        /// <summary>
        /// Creates required AssemblyOutputs for this definition.
        /// Checks whether any physical connections or edge treatment or corner features are required and if required then create them.
        /// </summary>
        public override void EvaluateAssembly()
        {
            try
            {
                Feature feature = (Feature)base.Occurrence;

                //Validating the inputs required to create the slot.
                //This is a virtual method and can be overriden to add specific validation, if specific input is missing.
                base.ValidateInputs();

                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created while validating the inputs
                    return;
                }                

                BusinessObject penetratedPart = base.Penetrated;

                //This call is only needed to support VB6 Assembly Connections and will not be needed when all content has been moved to .NET.
                base.AddFeatureGeometry(penetratedPart, feature);

                //get the required assembly outputs for this definition           
                Dictionary<string, bool> requiredAssemblyOutputs = this.RequiredAssemblyOutputs;

                //Create the left physical connection output for this definition, if not needed then delete the assembly output.
                bool isLeftPhysicalConnectionNeeded = requiredAssemblyOutputs["LeftPhysicalConnection"];
                CreateOrDeletePhysicalConnectionOutput(leftPhysicalConnection, isLeftPhysicalConnectionNeeded, (int)SectionFaceType.Web_Left, (int)SectionFaceType.Web_Left, DetailingCustomAssembliesConstants.TeeWeld);

                //Create the top physical connection output for this definition, if not needed then delete the assembly output.
                bool isTopPhysicalConnectionNeeded = requiredAssemblyOutputs["TopPhysicalConnection"];
                CreateOrDeletePhysicalConnectionOutput(topPhysicalConnection, isTopPhysicalConnectionNeeded, (int)SectionFaceType.Top, (int)SectionFaceType.Top, DetailingCustomAssembliesConstants.TeeWeld);

                //Create the free edge treatment output for this definition, if not needed then delete the assembly output.
                bool isFreeEdgeTreatmentNeeded = requiredAssemblyOutputs["FreeEdgeTreatment"];
                CreateOrDeleteFreeEdgeTreatmentOutput(freeEdgeTreatment, isFreeEdgeTreatmentNeeded, DetailingCustomAssembliesConstants.RootEdgeTreatment);

                //Create the bottom left corner feature output for this definition, if not needed then delete the assembly output.
                bool isBottomLeftCornerFeatureNeeded = requiredAssemblyOutputs["BottomLeftCornerFeature"];
                CornerEdgeOptions cornerEdgeOption = CornerEdgeOptions.EdgeAlongU;
                CreateOrDeleteCornerFeatureOutput(bottomLeftCornerFeature, isBottomLeftCornerFeatureNeeded, (int)SectionFaceType.Web_Left, cornerEdgeOption,  DetailingCustomAssembliesConstants.RootCorner);

                //Get section type of penetrating
                string sectionTypeName = GetPenetratingSectionTypeName();
                int edgeIdAlongV = PenetrationsServices.GetSectionEdgeIdAlongV(sectionTypeName);
                //Create the top left corner feature output for this definition, if not needed then delete the assembly output.
                bool isTopLeftCornerFeatureNeeded = requiredAssemblyOutputs["TopLeftCornerFeature"];
                CreateOrDeleteCornerFeatureOutput(topLeftCornerFeature, isTopLeftCornerFeatureNeeded, (int)SectionFaceType.Web_Left, edgeIdAlongV, DetailingCustomAssembliesConstants.VariableEdgeCorner);
            }
            catch (Exception)
            {
                //may be there could be specific ToDo record is created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(PenetrationsResourceIds.ToDoDefinition,
                            "Unexpected error while evaluating the custom assembly of {0}"), this.ToString()));
                }
            }
        }
        #endregion Public override functions and methods

        #region Private methods

        /// <summary>
        /// Gets all needed assembly outputs for the current configuration of the slot.
        /// </summary>
        /// <returns>This returns the needed outputs of the slot.</returns>
        private Dictionary<string, bool> RequiredAssemblyOutputs
        {
            get
            {
                //Dictionary which holds the data of the assembly output name and the bool which indicates
                //whether the corresponding assembly output is needed or not.
                Dictionary<string, bool> requiredAssemblyOutputs = new Dictionary<string, bool>();

                //Get part name
                string partName = ((Feature)base.Occurrence).PartName;               

                //By default the tolerance is considered as 0.004
                double tolerance = 0.004;

                bool isPartialDetailed = false;
                Ingr.SP3D.Interop.StructDetailUtil.StructDetailHelper oHlpr = new Ingr.SP3D.Interop.StructDetailUtil.StructDetailHelper();
                isPartialDetailed = oHlpr.IsPartialDetailed(Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO( (BusinessObject)base.Penetrated)) || oHlpr.IsPartialDetailed(Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO( (BusinessObject)base.Penetrating));

                bool isLeftPhysicalConnectionNeeded = false;

                if (!isPartialDetailed)
                {
                    //now get SectionFaceType for left physical connection if needed
                    if (PenetrationsServices.IsPCNeededOnSlotLeftEdgeOrLeftTopEdge(partName))
                    {
                        if (base.IsEdgeOverlappingWithPenetrating((int)SectionFaceType.Web_Left, tolerance))
                        {
                            isLeftPhysicalConnectionNeeded = true;
                        }
                    }
                }
                //Adding left pysical connection and top physical connection assembly output information to the dictonary
                requiredAssemblyOutputs.Add("LeftPhysicalConnection", isLeftPhysicalConnectionNeeded);

                //now add free edge treatment if it is needed
                bool isFreeEdgeTreatmentNeeded = false;

                if ( !oHlpr.IsPartialDetailed(Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO((BusinessObject)base.Penetrated)))
                {
                    isFreeEdgeTreatmentNeeded = base.IsFreeEdgeTreatmentNeeded(DetailingCustomAssembliesConstants.ApplyTreatment, (int)Answer.Yes);
                }
                oHlpr = null;

                requiredAssemblyOutputs.Add("FreeEdgeTreatment", isFreeEdgeTreatmentNeeded);

                bool isTopPhysicalConnectionNeeded = false;

                if (!isPartialDetailed)
                {
                    //now get SectionFaceType for top physical connection if needed
                    if (PenetrationsServices.IsPCNeededOnSlotTopEdgeOrLeftTopEdge(partName))
                    {
                        if (base.IsEdgeOverlappingWithPenetrating((int)SectionFaceType.Top, tolerance))
                        {
                            isTopPhysicalConnectionNeeded = true;
                        }
                    }
                }
                requiredAssemblyOutputs.Add("TopPhysicalConnection", isTopPhysicalConnectionNeeded); 

                bool isBottomLeftCornerFeatureNeeded = false; 

                //now get SectionFaceType for bottom left corner feature
                Feature slot = (Feature)base.Occurrence;
                int slotEdgeIDForBottomLeftCornerFeature = (int)PenetrationsServices.GetSlotEdgeIDForBottomLeftCornerFeature(slot, partName);
                if (slotEdgeIDForBottomLeftCornerFeature > 0)
                {
                    if (base.IsEdgeOverlappingWithPenetrating(slotEdgeIDForBottomLeftCornerFeature, tolerance))
                    {
                        isBottomLeftCornerFeatureNeeded = true;                        
                    }               
                }   
         
                //Add the BottomLeftCornerFeature assembly output information to the dictonary
                requiredAssemblyOutputs.Add("BottomLeftCornerFeature", isBottomLeftCornerFeatureNeeded);

                bool isTopLeftCornerFeatureNeeded = false; 

                //now get SectionFaceType for top left corner feature
                //PartName consistent if LT and outsideCorners is Yes
                string sectionTypeName = GetPenetratingSectionTypeName();

                int slotEdgeIDAlongVForTopLeftCornerFeature = (int)PenetrationsServices.GetSlotEdgeIDForTopLeftCornerFeature(slot, sectionTypeName, partName);
            
                if (slotEdgeIDAlongVForTopLeftCornerFeature > 0)
                {
                    if (base.IsEdgeOverlappingWithPenetrating((int)SectionFaceType.Web_Left, tolerance) && base.IsEdgeOverlappingWithPenetrating(slotEdgeIDAlongVForTopLeftCornerFeature, tolerance))
                    {
                        isTopLeftCornerFeatureNeeded = true;                       
                    }               
                }  
         
                //Add the TopLeftCornerFeature assembly output information to the dictonary
                requiredAssemblyOutputs.Add("TopLeftCornerFeature", isTopLeftCornerFeatureNeeded);

                return requiredAssemblyOutputs;
            
            }
        }
        #endregion Private methods
    }
}
