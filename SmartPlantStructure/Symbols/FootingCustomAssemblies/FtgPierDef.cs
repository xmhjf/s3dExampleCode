//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  FtgPierDef.cs
// 
//  Original Dll Name: <InstallDir>\SharedContent\Bin\SmartPlantStructure\Symbols\Release\SPSFootingMacros.dll
//  Original Class Name: ‘FtgPierDef’ and ‘FtgPierSym’ in VB content
//
//Abstract
//   FtgPierDef is a .NET custom assembly definition which creates graphic outputs for representing a pier footing component in the model.
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
    public class FtgPierDef : FootingCustomAssemblyDefinition
    {
        private bool isOnPreLoad = false;
        //========================================================================================================
        //DefinitionName/ProgID of this symbol is "FootingCustomAssemblies,Ingr.SP3D.Content.Structure.FtgPierDef"
        //========================================================================================================

        #region Definition of inputs
        [InputCatalogPart(1)]
        public InputCatalogPart catalogPart;
        [InputDouble(2, "PierShape", "Pier Shape", 1)]
        public InputDouble pierShape;
        [InputDouble(3, "PierSizingRule", "Pier Sizing Rule", 1)]
        public InputDouble pierSizingRule;
        [InputDouble(4, "PierOrientation", "Pier Orientation", 1)]
        public InputDouble pierOrientation;
        [InputDouble(5, "PierRotationAngle", "Pier Rotation Angle", 0)]
        public InputDouble pierRotationAngle;
        [InputDouble(6, "PierEdgeClearance", "Pier Edge Clearance", 1)]
        public InputDouble pierEdgeClearance;
        [InputDouble(7, "PierSizeIncrement", "Pier Size Increment", 1)]
        public InputDouble pierSizeIncrement;
        [InputDouble(8, "PierChamfered", "Pier Chamfered", 0)]
        public InputDouble pierChamfered;
        [InputDouble(9, "PierChamferSize", "Pier Chamfer Size", 1)]
        public InputDouble pierChamferSize;
        [InputDouble(10, "PierLength", "Pier Length", 18)]
        public InputDouble pierLength;
        [InputDouble(11, "PierWidth", "Pier Width", 18)]
        public InputDouble pierWidth;
        [InputDouble(12, "PierHeight", "Pier Height", 24)]
        public InputDouble pierHeight;
        [InputString(13, "PierSPSMaterial", "Pier Material", "Concrete")]
        public InputString pierMaterialName;
        [InputString(14, "PierSPSGrade", "Pier Material Grade", "Fc 3000")]
        public InputString pierMaterialGrade;
        #endregion Definition of inputs

        #region Definitions of aspects and their outputs
        [SymbolOutput(SPSSymbolConstants.Pier, "Pier Geometry")]
        [Aspect("SimplePhysical", "Simple Physical Aspect", AspectID.SimplePhysical)]
        public AspectDefinition simplePhysicalAspect;
        #endregion Definitions of aspects and their outputs

        #region Public override properties and methods

        /// <summary>
        /// Constructs the symbol outputs/aspects for Pier component.
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

                int shapeType = (int)this.pierShape.Value;
                double length = this.pierLength.Value;
                double width = this.pierWidth.Value;
                double height = this.pierHeight.Value;
                double clearance = this.pierEdgeClearance.Value;
                double projectionOffset = 0.0;
                int orientationType = (int)this.pierOrientation.Value;
                double rotationAngle = 0.0;
                double globalDelta = 0.0; //The rotation and member angle which is already factored into component's matrix.
                FootingComponentType componentType = FootingComponentType.Pier;

                //Construct component geometry                
                FootingServices.PlaceFootingShape(base.OccurrenceConnection, shapeType, length, width, height, clearance, rotationAngle, orientationType, globalDelta, componentType, projectionOffset, simplePhysicalAspect);
            }
            catch (Exception) // General Unhandled exception 
            {
                //check if ToDo message already created or not
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError,
                        FootingLocalizer.GetString(FootingResourceIDs.ErrPierConstructOutputs,
                        "Error while constructing outputs for pier in Footing Pier Definition. Please check your custom code or contact S3D support."));
                }
            }
        }

        /// <summary>
        /// Construct and re-evaluate the custom assembly outputs.
        /// Here we will do the following:
        /// 1. Calculate the surface area.
        /// 2. Set the calculated surface area on the pier FootingComponent.
        /// 3. Calculate the weight and centre of gravity (COG) of pier FootingComponent. 
        /// 4. Set the calculated weight and COG on the pier FootingComponent.
        /// </summary>
        public override void EvaluateAssembly()
        {
            FoundationComponent footingPierComponent = (FoundationComponent)base.Occurrence;
            try
            {

                //1. Get and set the physical properties on the component
                //Get physical properties - exposed surface area
                double surfaceArea = FootingServices.GetExposedSurfaceArea(footingPierComponent, FootingComponentType.Pier);

                //2. Set the physical properties on the component - surface area, volume, and centre of gravity
                //VolumeCG is added an output and surface area property is set on the footing
                FootingServices.SetPhysicalProperties(footingPierComponent, FootingComponentType.Pier, surfaceArea, 0.0);

                //2. Set the material 
                //Read the material info from the inputs and get the material object.
                CatalogStructHelper catalogStructHelper = FootingServices.CatalogStructHelper;
                string materialName = StructHelper.GetStringProperty(footingPierComponent, SPSSymbolConstants.IJUASPSFtgPier, SPSSymbolConstants.PierMaterial);
                string materialGrade = StructHelper.GetStringProperty(footingPierComponent, SPSSymbolConstants.IJUASPSFtgPier, SPSSymbolConstants.PierMaterialGrade);
                Material material = catalogStructHelper.GetMaterial(materialName, materialGrade);
                footingPierComponent.Material = material;
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
                    "Unexpected error while evaluating {0}. Check your custom code or contact S3D support."),this.ToString()));
                return;
            }
            try
            {
                //3. evaluating weight and COG for pier component
                FootingServices.EvaluateWeightCG(footingPierComponent, FootingComponentType.Pier);
            }
            catch (Exception) // General Unhandled exception 
            {
                //check if ToDo message already created or not
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.Occurrence.ToString() + " " +
                        FootingLocalizer.GetString(FootingResourceIDs.ErrCalculatingWCOG,
                        "Error in calculating weight and centre of gravity for pier component in Footing Pier Definition, as some of the required user attribute values cannot be obtained from the catalog. Check the error log and catalog data."));
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
                    case SPSSymbolConstants.PierSizingRule:
                    case SPSSymbolConstants.PierHeight:
                    case SPSSymbolConstants.PierOrientation:
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
            else if (propertyName == SPSSymbolConstants.PierSizingRule) //If property name is PierSizingRule 
            {
                //gray out the pier length and pier width on the GOPC
                FootingServices.SetComponentSizingPropertiesAccess(FootingComponentType.Pier, propValueInt, allDisplayedValues);

                //gray out the pier sizing rule property field read-only when footing is placed with point and without grout 
                FoundationComponent foundationComponent = (FoundationComponent)businessObject;
                Footing footing = foundationComponent.SystemParent as Footing;
                if (footing != null)
                {
                    propertyToChange.ReadOnly = base.IsPlacedByPoint(footing) && !FootingServices.DoesFootingHaveGrout(footing) ? true : false;
                }
            }
            else if (propertyName == SPSSymbolConstants.PierHeight || propertyName == SPSSymbolConstants.PierOrientation)
            {
                //for combined slab & pier assembly, make pier height read-only also for merged pier & combinedSlabAsm make pier height read-only if it is placed with bottom plane
                FoundationComponent foundationComponent = (FoundationComponent)businessObject;
                ISystem systemParent = foundationComponent.SystemParent;
                Footing footing = (Footing)systemParent;
                if (footing != null)
                {
                    string footingPartName = footing.PartName;
                    switch (footingPartName)
                    {
                        case SPSSymbolConstants.RECT_PIER_AND_SLAB_COMBINED_FOOTING_ASSEMBLY:
                        case SPSSymbolConstants.RECT_PIER_AND_OCT_SLAB_COMBINED_FOOTING_ASSEMBLY:
                            if (propertyName == SPSSymbolConstants.PierHeight)
                            {
                                base.SetAccessOnPropertyDescriptor(allDisplayedValues, propertyName, true);
                            }
                            break;
                        case SPSSymbolConstants.MERGED_RECT_PIER_COMBINED_FOOTING_ASSEMBLY:
                            if (propertyName == SPSSymbolConstants.PierHeight)
                            {
                                Collection<BusinessObject> supportingObjects = footing.SupportingObjects;
                                if (supportingObjects.Count == 0)
                                {
                                    base.SetAccessOnPropertyDescriptor(allDisplayedValues, propertyName, true);
                                }
                            }
                            else //For merged pier make pier orientation as Global & read-only
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
            int tempPierShape = Convert.ToInt32(pierShape.Value);
            int tempPierSizingRule = Convert.ToInt32(pierSizingRule.Value);
            int tempPierOrientation = Convert.ToInt32(pierOrientation.Value);

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructPrismaticFootingShapes, SPSSymbolConstants.UDP, tempPierShape))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_PIER_SHAPE,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrPierShapeCodeListValue,
                    "Error while validating Pier Code list value as Pier Shape value: {0} does not exist in Prismatic Footing Shapes: {1} table. Please check catalog or contact S3D support."), tempPierShape.ToString(), SPSSymbolConstants.StructPrismaticFootingShapes));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructFootingCompSizingRule, SPSSymbolConstants.UDP, tempPierSizingRule))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_PIER_SIZING_RULE,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrPierSizingRuleCodeListValue,
                    "Error while validating Pier sizing rule as PierSizingRule value: {0} does not exist in Footing component sizing rule: {1} table. Please check catalog or contact S3D support."), tempPierSizingRule.ToString(), SPSSymbolConstants.StructFootingCompSizingRule));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructCoordSysReference, SPSSymbolConstants.UDP, tempPierOrientation))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_PIER_ORIENTATION,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrPierOrientationCodeListValue,
                    "Error while validating Pier orientation as PierOrientation value: {0} does not exist in StructCoordSysReference: {1} table. Please check catalog or contact S3D support."), tempPierOrientation.ToString(), SPSSymbolConstants.StructCoordSysReference));
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
                case SPSSymbolConstants.PierLength:
                case SPSSymbolConstants.PierWidth:
                case SPSSymbolConstants.PierHeight:
                case SPSSymbolConstants.PierEdgeClearance:
                case SPSSymbolConstants.PierChamferSize:
                    isValid = ValidationHelper.IsGreaterThanZero(value, ref errorMessage);
                    break;
                case SPSSymbolConstants.PierSizeIncrement:
                    isValid = !ValidationHelper.IsNegative(value, ref errorMessage);
                    break;
            }

            return isValid;
        }

        #endregion Private methods
    }
}