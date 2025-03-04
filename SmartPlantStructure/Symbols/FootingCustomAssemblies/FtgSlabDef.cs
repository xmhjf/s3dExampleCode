//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  FtgSlabDef.cs
// 
//  Original Dll Name: <InstallDir>\SharedContent\Bin\SmartPlantStructure\Symbols\Release\SPSFootingMacros.dll
//  Original Class Name: ‘FtgSlabDef’ and ‘FtgSlabSym’ in VB content
//
//Abstract
//   FtgSlabDef is a .NET custom assembly definition which creates graphic outputs for representing a slab footing component in the model.
//   This class subclasses from FootingCustomAssemblyDefinition.
//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.ReferenceData.Exceptions;
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
    [SymbolVersion("1.1.0.0")]
    public class FtgSlabDef : FootingCustomAssemblyDefinition
    {
        private bool isOnPreLoad = false;
        //========================================================================================================
        //DefinitionName/ProgID of this symbol is "FootingCustomAssemblies,Ingr.SP3D.Content.Structure.FtgSlabDef"
        //========================================================================================================

        #region Definition of inputs
        [InputCatalogPart(1)]
        public InputCatalogPart catalogPart;
        [InputDouble(2, "SlabShape", "Slab Shape", 1)]
        public InputDouble slabShape;
        [InputDouble(3, "SlabSizingRule", "Slab Sizing Rule", 1)]
        public InputDouble slabSizingRule;
        [InputDouble(4, "SlabOrientation", "Slab Orientation", 1)]
        public InputDouble slabOrientation;
        [InputDouble(5, "SlabRotationAngle", "Slab Rotation Angle", 0)]
        public InputDouble slabRotationAngle;
        [InputDouble(6, "SlabEdgeClearance", "Slab Edge Clearance", 24)]
        public InputDouble slabEdgeClearance;
        [InputDouble(7, "SlabSizeIncrement", "Slab Size Increment", 1)]
        public InputDouble slabSizeIncrement;
        [InputDouble(8, "SlabLength", "Slab Length", 5)]
        public InputDouble slabLength;
        [InputDouble(9, "SlabWidth", "Slab Width", 5)]
        public InputDouble slabWidth;
        [InputDouble(10, "SlabHeight", "Slab Height", 16)]
        public InputDouble slabHeight;
        [InputString(11, "SlabSPSMaterial", "Slab Material", "Concrete")]
        public InputString slabMaterialName;
        [InputString(12, "SlabSPSGrade", "Slab Material Grade", "Fc 3000")]
        public InputString slabMaterialGrade;
        #endregion Definition of inputs

        #region Definitions of aspects and their outputs
        [SymbolOutput(SPSSymbolConstants.Slab, "Slab Geometry")]
        [Aspect("SimplePhysical", "Simple Physical Aspect", AspectID.SimplePhysical)]
        public AspectDefinition simplePhysicalAspect;
        #endregion Definitions of aspects and their outputs

        #region Public override properties and methods

        /// <summary>
        /// Constructs the symbol outputs/aspects for Slab component.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                //========================================
                // Construction of Simple Physical Aspect 
                //========================================
                //checking undefined value
                this.ValidateUndefinedCodelistValue();
                //ToDo list is created with error type hence stop computation
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    return;
                }


                int shapeType = (int)this.slabShape.Value;
                double length = this.slabLength.Value;
                double width = this.slabWidth.Value;
                double height = this.slabHeight.Value;
                double clearance = this.slabEdgeClearance.Value;
                double projectionOffset = 0.0;
                int orientationType = (int)this.slabOrientation.Value;
                double rotationAngle = 0.0;
                double globalDelta = 0.0; //The rotation and member angle which is already factored into component's matrix.
                FootingComponentType componentType = FootingComponentType.Slab;

                //Construct component geometry                
                FootingServices.PlaceFootingShape(base.OccurrenceConnection, shapeType, length, width, height, clearance, rotationAngle, orientationType, globalDelta, componentType, projectionOffset, simplePhysicalAspect);

            }
            catch (Exception) // General Unhandled exception  
            {
                //check if ToDo message already created or not
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError,
                        FootingLocalizer.GetString(FootingResourceIDs.ErrSlabConstructOutputs,
                        "Error while constructing outputs for slab in Footing Slab Definition. Please check your custom code or contact S3D support."));
                }
            }
        }

        /// <summary>
        /// Construct and re-evaluate the custom assembly outputs.
        /// Here we do the following:
        /// Get and set the physical properties on the component
        /// Calculate the weight and centre of gravity (COG) of Slab FootingComponent. 
        /// Set the calculated weight and COG on the Slab FootingComponent.
        /// </summary>
        public override void EvaluateAssembly()
        {
            FoundationComponent footingSlabComponent = (FoundationComponent)base.Occurrence;
            try
            {
                //2. Get and set the physical properties on the component
                //Get physical properties - exposed surface area
                double surfaceArea = FootingServices.GetExposedSurfaceArea(footingSlabComponent, FootingComponentType.Slab);

                //3. Set the physical properties on the component - surface area, volume, and centre of gravity
                //VolumeCG is added an output and surface area property is set on the footing
                FootingServices.SetPhysicalProperties(footingSlabComponent, FootingComponentType.Slab, surfaceArea, 0.0);

                //4. Set the material 
                //Read the material info from the inputs and get the material object.
                CatalogStructHelper catalogStructHelper = FootingServices.CatalogStructHelper;
                string materialName = StructHelper.GetStringProperty(footingSlabComponent, SPSSymbolConstants.IJUASPSFtgSlab, SPSSymbolConstants.SlabMaterial);
                string materialGrade = StructHelper.GetStringProperty(footingSlabComponent, SPSSymbolConstants.IJUASPSFtgSlab, SPSSymbolConstants.SlabMaterialGrade);
                Material material = catalogStructHelper.GetMaterial(materialName, materialGrade);
                footingSlabComponent.Material = material;
            }
            catch (RefDataMaterialNotFoundException)
            {
                //If Material is not found in catalog, create todo record and continue
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, FootingLocalizer.GetString(FootingResourceIDs.ErrMaterialNotFound,
                                        "Cannot set material on the footing, as the required material is not found in catalog. Check the error log and catalog data."));
            }
            catch
            {
                //generic error. More serious stuff, create ToDoRecord and stop computation
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrEvaluatePierAndSlabFootingAssembly,
                    "Unexpected error while evaluating {0}. Check your custom code or contact S3D support."), this.ToString()));
                return;
            }
            try
            {
                //evaluating weight and COG for slab component
                FootingServices.EvaluateWeightCG(footingSlabComponent, FootingComponentType.Slab);
            }
            catch (Exception) // General Unhandled exception 
            {
                //check if ToDo message already created or not
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.Occurrence.ToString() + " " +
                        FootingLocalizer.GetString(FootingResourceIDs.ErrSlabWCOGMissingSystemAttributeData,
                        "cannot calculate weight and centre of gravity for the Slab Component in Footing Slab Definition, as some of the required user attribute value cannot be obtained from the catalog. Check the error log and catalog data."));
                }
            }
        }

        /// <summary> 
        /// OnPreLoad gets called immediately before the properties are loaded in the property page control. 
        /// Any change to the display status of properties is to be done here. 
        /// </summary> 
        /// <param name="businessObject">Delegating Business Object.</param> 
        /// <param name="CollAllDisplayedValues">Read only collection of all properties displayed in the property pages control.</param>
        public override void OnPreLoad(BusinessObject businessObject, ReadOnlyCollection<PropertyDescriptor> allDisplayedValues)
        {
            // check if any of the arguments are null and throw argument null exception. 
            // businessObject can be nothing 
            if (allDisplayedValues == null)
            {
                throw new ArgumentNullException("allDisplayedValues");
            }

            //optimization to avoid value validation in OnPropertyChange 
            this.isOnPreLoad = true;

            for (int i = 0; i < allDisplayedValues.Count; i++)
            {
                PropertyDescriptor propertyDescriptor = allDisplayedValues[i];
                PropertyValue propertyValue = propertyDescriptor.Property;
                string propertyName = propertyValue.PropertyInfo.Name;
                switch (propertyName)
                {
                    //make all these properties read-only
                    case SPSSymbolConstants.Volume:
                    case SPSSymbolConstants.SurfaceArea:
                    case SPSSymbolConstants.CenterX:
                    case SPSSymbolConstants.CenterY:
                    case SPSSymbolConstants.CenterZ:
                        propertyDescriptor.ReadOnly = true;
                        break;
                    case SPSSymbolConstants.ReportingRequirements:
                    case SPSSymbolConstants.SlabSizingRule:
                    case SPSSymbolConstants.SlabHeight:
                    case SPSSymbolConstants.SlabOrientation:
                        string errorMessage = string.Empty;
                        this.OnPropertyChange(businessObject, allDisplayedValues, propertyDescriptor, propertyValue, out errorMessage);
                        if (errorMessage.Length > 0)
                        {
                            this.isOnPreLoad = false;
                        }
                        break;
                }

                //isOnPreLoad = false, no need to execute for loop, hence break it and let the caller know about it.
                if (!this.isOnPreLoad)
                {
                    break;
                }
            }

            //setting to False so that it will do value validation in OnPropertyChange 
            this.isOnPreLoad = false;
        }

        /// <summary> 
        /// OnPropertyChange is called each time a property is modified. Any custom validation to be done here. 
        /// </summary> 
        /// <param name="businessObject">Delegating Business Object.</param> 
        /// <param name="allDisplayedValues">Read only collection of all properties displayed in the property pages control.</param> 
        /// <param name="propertyToChange">Property being modified.</param> 
        /// <param name="newPropertyValue">New value of the property.</param> 
        /// <param name="errorMessage">Custom error message returned by validation.</param> 
        /// <returns>Returns Boolean for error status. Returns false if change in property value is not valid.</returns> 
        public override bool OnPropertyChange(BusinessObject businessObject, ReadOnlyCollection<PropertyDescriptor> allDisplayedValues, PropertyDescriptor propertyToChange, PropertyValue newPropertyValue, out string errorMessage)
        {
            // check if any of the arguments are null and throw argument null exception. 
            // businessObject and allDisplayedValues can be null 
            if (propertyToChange == null)
            {
                throw new ArgumentNullException("propertyToChange");
            }
            if (newPropertyValue == null)
            {
                throw new ArgumentNullException("newPropertyValue");
            }

            bool isOnPropertyChange = true;
            //initializing with true
            errorMessage = string.Empty;
            int propValueInt = 0;
            SP3DPropType propertyType = newPropertyValue.PropertyInfo.PropertyType;
            string propertyName = propertyToChange.Property.PropertyInfo.Name;

            if (propertyType == SP3DPropType.PTInteger)
            {
                PropertyValueInt propertyValueInt = (PropertyValueInt)newPropertyValue;
                propValueInt = Convert.ToInt32(propertyValueInt.PropValue);
            }
            else if (propertyType == SP3DPropType.PTCodelist)
            {
                PropertyValueCodelist propertyValueCodelist = (PropertyValueCodelist)newPropertyValue;
                propValueInt = propertyValueCodelist.PropValue;
            }

            //If property name is ReportingRequirements 
            if (propertyName == SPSSymbolConstants.ReportingRequirements)
            {
                if (propValueInt == -1)
                {
                    //grey out the reporting type on the GOPC
                    base.SetAccessOnPropertyDescriptor(allDisplayedValues, SPSSymbolConstants.ReportingType, true);
                }
            }
            else if (propertyName == SPSSymbolConstants.SlabSizingRule) //If property name is SlabSizingRule 
            {
                //grey out the slab length and width on the GOPC
                FootingServices.SetComponentSizingPropertiesAccess(FootingComponentType.Slab, propValueInt, allDisplayedValues);
            }
            else if (propertyName == SPSSymbolConstants.SlabHeight || propertyName == SPSSymbolConstants.SlabOrientation) //for combined slab assembly make slab height read-only if placed with bottom plane
            {
                FoundationComponent foundationComponent = (FoundationComponent)businessObject;
                ISystem systemParent = foundationComponent.SystemParent;
                Footing footing = (Footing)systemParent;
                if (footing != null)
                {
                    string footingPartName = footing.PartName;
                    switch (footingPartName)
                    {
                        case SPSSymbolConstants.RECT_SLAB_COMBINED_FOOTING_ASSEMBLY:
                            if (propertyName == SPSSymbolConstants.SlabHeight)
                            {
                                Collection<BusinessObject> supportingObjects = footing.SupportingObjects;
                                if (supportingObjects.Count > 0) //means placed with plane
                                {
                                    base.SetAccessOnPropertyDescriptor(allDisplayedValues, propertyName, true);
                                }
                            }
                            else //for combined footing slabs make orientation as Global & read-only
                            {
                                base.SetValueAndAccessOnPropertyDescriptor(allDisplayedValues, propertyName, true, SPSSymbolConstants.ORIENTATION_GLOBAL);
                            }
                            break;
                        case SPSSymbolConstants.RECT_PIER_AND_SLAB_COMBINED_FOOTING_ASSEMBLY: //for combined footing slabs make orientation as Global & read-only
                            if (propertyName == SPSSymbolConstants.SlabOrientation)
                            {
                                base.SetValueAndAccessOnPropertyDescriptor(allDisplayedValues, propertyName, true, SPSSymbolConstants.ORIENTATION_GLOBAL);
                            }
                            break;
                    }
                }
            }

            //isOnPreLoad = false means this method is not called from OnPreLoad(), hence validate the input values
            if (!this.isOnPreLoad && propertyType == SP3DPropType.PTDouble)
            {
                isOnPropertyChange = this.CustomPropertyMgtValidate(propertyName, newPropertyValue, ref errorMessage);
            }

            return isOnPropertyChange;
        }

        #endregion Public override properties and methods

        #region Private methods

        /// <summary>
        /// Checks for undefined value and raise error.
        /// </summary>
        private void ValidateUndefinedCodelistValue()
        {
            int tempSlabShape = Convert.ToInt32(slabShape.Value);
            int tempSlabSizingRule = Convert.ToInt32(slabSizingRule.Value);
            int tempSlabOrientation = Convert.ToInt32(slabOrientation.Value);

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructPrismaticFootingShapes, SPSSymbolConstants.UDP, tempSlabShape))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_SLAB_SHAPE,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrSlabShapeCodeListValue,
                    "Error while validating Slab Code list value as Slab Shape value: {0} does not exist in Prismatic Footing Shapes: {1} table. Please check catalog or contact S3D support."), tempSlabShape.ToString(), SPSSymbolConstants.StructPrismaticFootingShapes));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructFootingCompSizingRule, SPSSymbolConstants.UDP, tempSlabSizingRule))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_SLAB_SIZING_RULE,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrSlabSizingRuleCodeListValue,
                    "Error while validating Slab sizing rule as SlabSizingRule value: {0} does not exist in Footing component sizing rule: {1} table. Please check catalog or contact S3D support."), tempSlabSizingRule.ToString(), SPSSymbolConstants.StructFootingCompSizingRule));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructCoordSysReference, SPSSymbolConstants.UDP, tempSlabOrientation))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_SLAB_ORIENTATION,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrSlabOrientationCodeListValue,
                    "Error while validating Slab orientation as SlabOrientation value: {0} does not exist in StructCoordSysReference: {1} table. Please check catalog or contact S3D support."), tempSlabOrientation.ToString(), SPSSymbolConstants.StructCoordSysReference));
                return;
            }
        }

        /// <summary> 
        /// This method is used to validate the footing property value. 
        /// </summary> 
        /// <param name="propertyName">The name of the Property to validate.</param> 
        /// <param name="newPropertyValue">The value of the Property to validate.</param> 
        /// <param name="errorMessage">The error message returned by reference.</param> 
        /// <returns>Returns true if valid otherwise false.</returns> 
        private bool CustomPropertyMgtValidate(string propertyName, PropertyValue newPropertyValue, ref string errorMessage)
        {
            bool isValid = true;
            //initializing with true
            PropertyValueDouble propertyValueDouble = (PropertyValueDouble)newPropertyValue;
            double value = Convert.ToDouble(propertyValueDouble.PropValue);

            switch (propertyName)
            {
                case SPSSymbolConstants.SlabLength:
                case SPSSymbolConstants.SlabWidth:
                case SPSSymbolConstants.SlabHeight:
                case SPSSymbolConstants.SlabEdgeClearance:
                    isValid = ValidationHelper.IsGreaterThanZero(value, ref errorMessage);
                    break;
                case SPSSymbolConstants.SlabSizeIncrement:
                    isValid = !ValidationHelper.IsNegative(value, ref errorMessage);
                    break;
            }

            return isValid;
        }

        #endregion Private methods
    }
}