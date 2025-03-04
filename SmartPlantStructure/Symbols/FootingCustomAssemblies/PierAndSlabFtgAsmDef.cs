//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  PierAndSlabFtgAsmDef.cs
// 
//  Original Dll Name: <InstallDir>\SharedContent\Bin\SmartPlantStructure\Symbols\Release\SPSFootingMacros.dll
//  Original Class Name: ‘PierAndSlabFtgAsmDef’ in VB content
//
//Abstract
//   PierAndSlabFtgAsmDef is a .NET custom assembly definition which creates a grout pad, a pier and a slab component in the model.
//   This class subclasses from FootingCustomAssemblyDefinition.
//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
//===========================================================================================
//Namespace of this class is Ingr.SP3D.Content.Structure
//It is recommended that customers specify namespace of their symbols to be
//<CompanyName>.SP3D.Content.<Specialization>.
//It is also recommended that if customers want to change this symbol to suit their
//requirements, they should change namespace/symbol name so the identity of the modified
//symbol will be different from the one delivered by Intergraph.
//===========================================================================================
namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Pier and slab footing CustomAssemblyDefinition.
    /// It will update the IJDAttributes interface for this definition and the assembly outputs will have their attributes (IJDAttributes) modified.
    /// </summary>
    [SymbolVersion("1.1.0.0")]
    [OutputNotification(SPSSymbolConstants.IID_IJDAttributes)]
    [OutputNotification(SPSSymbolConstants.IID_IJDAttributes, true)] // The assembly outputs will have their attributes (IJDAttributes) modified.
    [OutputNotification(SPSSymbolConstants.IID_IJDGeometry, true)] // The assembly outputs will have their geometry (IJDGeometry) modified.
    [OutputNotification(SPSSymbolConstants.IID_IJStructMaterial, true)] // The assembly outputs will have their material (IJStructMaterial) modified.
    public class PierAndSlabFtgAsmDef : FootingCustomAssemblyDefinition
    {
        //==================================================================================================================
        //DefinitionName/ProgID of this symbol is "FootingCustomAssemblies,Ingr.SP3D.Content.Structure.PierAndSlabFtgAsmDef"
        //==================================================================================================================

        #region Definition of inputs
        [InputCatalogPart(1)]
        public InputCatalogPart catalogPart;
        #endregion Definition of inputs

        #region Definitions of assemblies and their outputs
        [AssemblyOutput(1, SPSSymbolConstants.Grout)]
        public AssemblyOutput groutAssemblyOutput;

        [AssemblyOutput(2, SPSSymbolConstants.Pier)]
        public AssemblyOutput pierAssemblyOutput;

        [AssemblyOutput(3, SPSSymbolConstants.Slab)]
        public AssemblyOutput slabAssemblyOutput;
        #endregion Definitions of assemblies and their outputs

        #region Public override properties and methods

        /// <summary>
        /// Construct and re-evaluate the custom assembly outputs.
        /// Here we will do the following:
        /// 1. Decide which assembly outputs are needed. 
        /// 2. Create the ones which are needed, delete which are not needed now.
        /// </summary>
        /// <remarks></remarks>
        public override void EvaluateAssembly()
        {
            try
            {
                //getting the footing from occurrence property
                Footing footing = (Footing)base.Occurrence;

                FoundationComponent groutComponent = null;
                FoundationComponent pierComponent = null;
                FoundationComponent slabComponent = null;
                string foundationComponentPartName;

                // construct the grout if required
                bool withGroutPad = FootingServices.DoesFootingHaveGrout(footing);
                if (groutAssemblyOutput.Output == null)
                {
                    if (withGroutPad)
                    {
                        foundationComponentPartName = FootingServices.GetFootingComponentPartName(footing, FootingComponentType.Grout);
                        groutComponent = base.CreateComponent(footing, foundationComponentPartName);
                        groutAssemblyOutput.Output = groutComponent;
                    }
                }
                else
                {
                    if (withGroutPad)
                    {
                        groutComponent = (FoundationComponent)groutAssemblyOutput.Output;
                    }
                    else // output was previously generated (but no longer needed) so delete it now  
                    {
                        groutAssemblyOutput.Delete();
                    }
                }

                bool isPlacedByPoint = base.IsPlacedByPoint(footing);
                //set the grout sizing rule to 'User Defined' in case placed by point
                if (isPlacedByPoint && withGroutPad)
                {
                    FootingServices.SetSizingRule(groutComponent, FootingComponentType.Grout, FootingServices.UserDefined);
                }

                // construct the pier (if not generated yet)
                if (pierAssemblyOutput.Output == null)
                {
                    foundationComponentPartName = FootingServices.GetFootingComponentPartName(footing, FootingComponentType.Pier);
                    pierComponent = base.CreateComponent(footing, foundationComponentPartName);
                    pierAssemblyOutput.Output = pierComponent;
                }
                else
                {
                    pierComponent = (FoundationComponent)pierAssemblyOutput.Output;
                }

                //set the pier sizing rule to 'User Defined' in case placed by point and without grout pad
                if (isPlacedByPoint && !withGroutPad)
                {                    
                    FootingServices.SetSizingRule(pierComponent, FootingComponentType.Pier, FootingServices.UserDefined);
                }

                // construct the slab (if not generated yet)
                if (slabAssemblyOutput.Output == null)
                {
                    foundationComponentPartName = FootingServices.GetFootingComponentPartName(footing, FootingComponentType.Slab);
                    slabComponent = base.CreateComponent(footing, foundationComponentPartName);

                    // Get the inputs from the parent occurrence
                    Collection<BusinessObject> planes = footing.SupportingObjects;
                    if (planes.Count > 0)
                    {
                        slabComponent.SupportingObjects = planes;
                    }

                    slabAssemblyOutput.Output = slabComponent;
                }
                else
                {
                    slabComponent = (FoundationComponent)slabAssemblyOutput.Output;
                }

                double sectionDepth = 0;
                double sectionWidth = 0;
                double height = 0;
                double memberAngle = 0;

                //Call the evaluate which will return the section dimension and member angle.  These values
                //will be used to compute the length and width of the footing component and the rotation angle.
                base.Evaluate(null, footing, out sectionDepth, out sectionWidth, out height, out memberAngle);
                //ToDo list is created with error type hence stop computation
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    return;
                }

                ToDoListMessage toDoListMessage = null;
                //Grout Evaluation
                if (withGroutPad)
                {
                    toDoListMessage = FootingServices.EvaluateComponent(footing, groutComponent, FootingComponentType.Grout, withGroutPad, sectionDepth, sectionWidth, memberAngle);
                    if (toDoListMessage != null)
                    {
                        base.ToDoListMessage = toDoListMessage;
                        if (toDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        {
                            return;
                        }
                    }
                }

                //Pier Evaluation
                toDoListMessage = FootingServices.EvaluateComponent(footing, pierComponent, FootingComponentType.Pier, withGroutPad, sectionDepth, sectionWidth, memberAngle);
                if (toDoListMessage != null)
                {
                    base.ToDoListMessage = toDoListMessage;
                    if (toDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        return;
                    }
                }

                //Slab Evaluation
                toDoListMessage = FootingServices.EvaluateComponent(footing, slabComponent, FootingComponentType.Slab, withGroutPad, sectionDepth, sectionWidth, memberAngle);
                if (toDoListMessage != null)
                {
                    base.ToDoListMessage = toDoListMessage;
                    if (toDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        return;
                    }
                }
            }
            catch (Exception) // General Unhandled exception 
            {
                //check if ToDo message already created or not
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError,
                        String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrEvaluatePierAndSlabFootingAssembly,
                    "Unexpected error while evaluating {0}. Check your custom code or contact S3D support."), this.ToString()));
                }
            }
        }

        /// <summary>
        /// Updates the rotation value of footing components based on the transform applied to the footing.
        /// This method needs to be overridden to handle the case when the components orientation is set to Global.
        /// In this case any transformation applied to Footing needs to be applied to the components as well.
        /// </summary>
        /// <param name="businessObject">The footing which is being transformed.</param>
        /// <param name="transformationMatrix">The transformation matrix.</param>
        public override void Transform(BusinessObject businessObject, Matrix4X4 transformationMatrix)
        {
            if (businessObject == null)
            {
                throw new ArgumentNullException("businessObject");
            }
            if (transformationMatrix == null)
            {
                throw new ArgumentNullException("transformationMatrix");
            }

            Footing footing = (Footing)businessObject;

            //if the footing is placed by point then update the rotation, but if it placed by member then update only if orientation is global
            bool isPlacedByPoint = base.IsPlacedByPoint(footing);

            //Update each component's transform matrix as that's where the geometry is
            //update grout's rotation
            bool withGroutPad = FootingServices.DoesFootingHaveGrout(footing);
            if (withGroutPad)
            {
                //get the grout component
                FoundationComponent groutComponent = FootingServices.GetComponent(footing, FootingComponentType.Grout);
                base.UpdateComponentRotationAngle(groutComponent, isPlacedByPoint, SPSSymbolConstants.IJUASPSFtgGroutPad, SPSSymbolConstants.GroutOrientation, SPSSymbolConstants.GroutRotationAngle, transformationMatrix);
            }

            //update pier's rotation
            //get the pier component
            FoundationComponent pierComponent = FootingServices.GetComponent(footing, FootingComponentType.Pier);
            base.UpdateComponentRotationAngle(pierComponent, isPlacedByPoint, SPSSymbolConstants.IJUASPSFtgPier, SPSSymbolConstants.PierOrientation, SPSSymbolConstants.PierRotationAngle, transformationMatrix);

            //update slab's rotation
            //get the slab component
            FoundationComponent slabComponent = FootingServices.GetComponent(footing, FootingComponentType.Slab);
            base.UpdateComponentRotationAngle(slabComponent, isPlacedByPoint, SPSSymbolConstants.IJUASPSFtgSlab, SPSSymbolConstants.SlabOrientation, SPSSymbolConstants.SlabRotationAngle, transformationMatrix);
        }

        /// <summary> 
        /// Gets the foul interface type from the footing.
        /// </summary> 
        /// <param name="businessObject">Business object which aggregates the symbol.</param> 
        /// <returns>Enumerated values for foul interface type. Returns Participant if the custom symbol should participant for interference check.</returns>
        public override FoulInterfaceType GetFoulInterfaceType(BusinessObject businessObject)
        {
            return FoulInterfaceType.Participant;
        }

        #endregion Public override properties and methods
    }
}