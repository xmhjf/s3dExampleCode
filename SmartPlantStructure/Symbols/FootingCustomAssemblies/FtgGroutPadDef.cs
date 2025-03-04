//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  FtgGroutPadDef.cs
// 
//  Original Dll Name: <InstallDir>\SharedContent\Bin\SmartPlantStructure\Symbols\Release\SPSFootingMacros.dll
//  Original Class Name: ‘FtgGroutPadDef’ and ‘FtgGroutPadSym’ in VB content
//
//Abstract
//   FtgGroutPadDef is a .NET custom assembly definition which creates graphic outputs for representing a grout pad footing component in the model.
//   This class subclasses from FootingCustomAssemblyDefinition.
//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    public class FtgGroutPadDef : FootingCustomAssemblyDefinition
    {
        private bool isOnPreLoad = false;
        //============================================================================================================
        //DefinitionName/ProgID of this symbol is "FootingCustomAssemblies,Ingr.SP3D.Content.Structure.FtgGroutPadDef"
        //============================================================================================================

        #region Definition of inputs
        [InputCatalogPart(1)]
        public InputCatalogPart catalogPart;
        [InputDouble(2, "GroutShape", "Grout Shape", 1)]
        public InputDouble groutShape;
        [InputDouble(3, "GroutSizingRule", "Grout Sizing Rule", 1)]
        public InputDouble groutSizingRule;
        [InputDouble(4, "GroutOrientation", "Grout Orientation", 3)]
        public InputDouble groutOrientation;
        [InputDouble(5, "GroutRotationAngle", "Grout Rotation Angle", 0)]
        public InputDouble groutRotationAngle;
        [InputDouble(6, "GroutEdgeClearance", "Grout Edge Clearance", 0)]
        public InputDouble groutEdgeClearance;
        [InputDouble(7, "GroutLength", "Grout Length", 16)]
        public InputDouble groutLength;
        [InputDouble(8, "GroutWidth", "Grout Width", 16)]
        public InputDouble groutWidth;
        [InputDouble(9, "GroutHeight", "Grout Height", 1)]
        public InputDouble groutHeight;
        [InputString(10, "GroutSPSMaterial", "Grout Material", "Grout")]
        public InputString groutMaterialName;
        [InputString(11, "GroutSPSGrade", "Grout Material Grade", "High Strength")]
        public InputString groutMaterialGrade;
        #endregion Definition of inputs

        #region Definitions of aspects and their outputs
        [SymbolOutput(SPSSymbolConstants.Grout, "Grout Geometry")]
        [Aspect("SimplePhysical", "Simple Physical Aspect", AspectID.SimplePhysical)]
        public AspectDefinition simplePhysicalAspect;
        #endregion Definitions of aspects and their outputs

        #region Public override properties and methods

        /// <summary>
        /// Constructs the symbol outputs/aspects for Grout component.
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

                int shapeType = (int)this.groutShape.Value;
                double length = this.groutLength.Value;
                double width = this.groutWidth.Value;
                double height = this.groutHeight.Value;
                double clearance = this.groutEdgeClearance.Value;
                double projectionOffset = 0.0;
                int orientationType = (int)this.groutOrientation.Value;
                double rotationAngle = 0.0;
                double globalDelta = 0.0; //The rotation and member angle which is already factored into component's matrix.
                FootingComponentType componentType = FootingComponentType.Grout;

                //Construct component geometry                
                FootingServices.PlaceFootingShape(base.OccurrenceConnection, shapeType, length, width, height, clearance, rotationAngle, orientationType, globalDelta, componentType, projectionOffset, simplePhysicalAspect);
            }
            catch (Exception)  // General Unhandled exception 
            {
                //check if ToDo message already created or not
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError,
                        String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrEvaluatePierAndSlabFootingAssembly,
                        "Unexpected error while evaluating {0}. Check your custom code or contact S3D support."),this.ToString()));
                }
            }
        }

        /// <summary>
        /// Construct and re-evaluate the custom assembly outputs.
        /// Here we do the following:
        /// Get and set the physical properties on the component
        /// Calculate the weight and centre of gravity (COG) of grout FootingComponent. 
        /// Set the calculated weight and COG on the grout FootingComponent.
        /// </summary>
        public override void EvaluateAssembly()
        {
            FoundationComponent footingGroutComponent = (FoundationComponent)base.Occurrence;
            try
            {
                //Get and set the physical properties on the component
                //Get physical properties - exposed surface area
                double surfaceArea = FootingServices.GetExposedSurfaceArea(footingGroutComponent, FootingComponentType.Grout);

                //Set the physical properties on the footing - surface area, volume, and centre of gravity
                //VolumeCG is added an output and surface area property is set on the footing
                FootingServices.SetPhysicalProperties(footingGroutComponent, FootingComponentType.Grout, surfaceArea, 0.0);

                //Set the material
                //Read the material info from the inputs and get the material object.
                CatalogStructHelper catalogStructHelper = FootingServices.CatalogStructHelper;
                string materialName = StructHelper.GetStringProperty(footingGroutComponent, SPSSymbolConstants.IJUASPSFtgGroutPad, SPSSymbolConstants.GroutMaterial);
                string materialGrade = StructHelper.GetStringProperty(footingGroutComponent, SPSSymbolConstants.IJUASPSFtgGroutPad, SPSSymbolConstants.GroutMaterialGrade);
                Material material = catalogStructHelper.GetMaterial(materialName, materialGrade);
                footingGroutComponent.Material = material;
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
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FootingLocalizer.GetString(FootingResourceIDs.ErrEvaluateFootingGroutPadDef,
                    "Unexpected error while evaluating custom assembly outputs for Footing Grout pad definition. Check your custom code or contact S3D support."));
                return;
            }
            //evaluating weight and COG for grout component
            try
            {
                FootingServices.EvaluateWeightCG((FoundationBase)base.Occurrence, FootingComponentType.Grout);
            }
            catch (Exception) // General Unhandled exception 
            {
                //check if ToDo message already created or not
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.Occurrence.ToString() + " " +
                        FootingLocalizer.GetString(FootingResourceIDs.ErrGroutWCOGMissingSystemAttributeData,
                        "cannot calculate weight and centre of gravity for the Grout Component in Footing Grout Pad Definition, as some of the required user attribute value cannot be obtained from the catalog. Check the error log and catalog data."));
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
                        propertyDescriptor.ReadOnly = true;
                        break;
                    case SPSSymbolConstants.ReportingRequirements:
                    case SPSSymbolConstants.GroutSizingRule:
                    case SPSSymbolConstants.GroutHeight:
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
                    //gray out the reporting type on the GOPC
                    base.SetAccessOnPropertyDescriptor(allDisplayedValues, SPSSymbolConstants.ReportingType, true);
                }
            }
            else if (propertyName == SPSSymbolConstants.GroutSizingRule) //If property name is GroutSizingRule 
            {
                FootingServices.SetComponentSizingPropertiesAccess(FootingComponentType.Grout, propValueInt, allDisplayedValues);

                //gray out the grout sizing rule property field when footing is placed with point
                FoundationComponent foundationComponent = (FoundationComponent)businessObject;
                Footing footing = foundationComponent.SystemParent as Footing;
                if (footing != null)
                {
                    propertyToChange.ReadOnly = base.IsPlacedByPoint(footing) ? true : false;
                }
            }
            else if (propertyName == SPSSymbolConstants.GroutHeight) //for merged pier & combinedSlabAsm make grout height read-only
            {
                FoundationComponent foundationComponent = (FoundationComponent)businessObject;
                Footing footing = foundationComponent.SystemParent as Footing;
                if (footing != null)
                {
                    string footingPartName = footing.PartName;
                    switch (footingPartName)
                    {
                        case SPSSymbolConstants.MERGED_RECT_PIER_COMBINED_FOOTING_ASSEMBLY:
                        case SPSSymbolConstants.RECT_SLAB_COMBINED_FOOTING_ASSEMBLY:
                        case SPSSymbolConstants.OCT_SLAB_COMBINED_FOOTING_ASSEMBLY:
                            base.SetAccessOnPropertyDescriptor(allDisplayedValues, SPSSymbolConstants.GroutHeight, true);
                            break;
                        case SPSSymbolConstants.RECT_PIER_AND_SLAB_COMBINED_FOOTING_ASSEMBLY:
                        case SPSSymbolConstants.RECT_PIER_AND_OCT_SLAB_COMBINED_FOOTING_ASSEMBLY:
                            base.SetAccessOnPropertyDescriptor(allDisplayedValues, SPSSymbolConstants.GroutHeight, true);
                            break;
                    }
                }
            }

            //isOnPreLoad = false means this method is not called from OnPreLoad(), hence validate the input values
            if (!this.isOnPreLoad && (propertyType == SP3DPropType.PTDouble || propertyType == SP3DPropType.PTCodelist))
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
            int tempGroutShape = Convert.ToInt32(groutShape.Value);
            int tempGroutSizingRule = Convert.ToInt32(groutSizingRule.Value);
            int tempGroutOrientation = Convert.ToInt32(groutOrientation.Value);

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructPrismaticFootingShapes, SPSSymbolConstants.UDP, tempGroutShape))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_PIER_SHAPE,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrGroutShapeCodeListValue,
                    "Error while validating Grout Code list value as Grout Shape value: {0} does not exist in Prismatic Footing Shapes: {1} table. Please check catalog or contact S3D support."), tempGroutShape.ToString(), SPSSymbolConstants.StructPrismaticFootingShapes));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructFootingCompSizingRule, SPSSymbolConstants.UDP, tempGroutSizingRule))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_PIER_SIZING_RULE,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrGroutSizingRuleCodeListValue,
                    "Error while validating Grout sizing rule as GroutSizingRule value: {0} does not exist in Footing component sizing rule: {1} table. Please check catalog or contact S3D support."), tempGroutSizingRule.ToString(), SPSSymbolConstants.StructFootingCompSizingRule));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructCoordSysReference, SPSSymbolConstants.UDP, tempGroutOrientation))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_PIER_ORIENTATION,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrGroutOrientationCodeListValue,
                    "Error while validating Grout orientation as GroutOrientation value: {0} does not exist in StructCoordSysReference: {1} table. Please check catalog or contact S3D support."), tempGroutOrientation.ToString(), SPSSymbolConstants.StructCoordSysReference));
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
            double value = 0;
            int valueInt = 0;

            SP3DPropType propertyType = newPropertyValue.PropertyInfo.PropertyType;
            if (propertyType == SP3DPropType.PTDouble)
            {
                PropertyValueDouble propertyValueDouble = (PropertyValueDouble)newPropertyValue;
                value = Convert.ToDouble(propertyValueDouble.PropValue);
            }
            else
            {
                PropertyValueCodelist propertyValueCodelist = (PropertyValueCodelist)newPropertyValue;
                valueInt = propertyValueCodelist.PropValue;
            }

            switch (propertyName)
            {
                case SPSSymbolConstants.GroutLength:
                case SPSSymbolConstants.GroutWidth:
                case SPSSymbolConstants.GroutHeight:
                case SPSSymbolConstants.GroutEdgeClearance:
                    isValid = ValidationHelper.IsGreaterThanZero(value, ref errorMessage);
                    break;
                case SPSSymbolConstants.GroutShape:
                    if (valueInt == SPSSymbolConstants.SHAPE_OCTAGONAL)
                    {
                        isValid = false;
                        errorMessage = String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrOctagonalGroutNotSupported, "Octagonal Grout is not supported while validating footing property value in {0}. Please check your inputs."),this.ToString());
                    }
                    break;
            }

            return isValid;
        }

        #endregion Private methods
    }
}