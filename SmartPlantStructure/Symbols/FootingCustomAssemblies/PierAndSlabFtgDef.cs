//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  PierAndSlabFtgDef.cs
// 
//  Original Dll Name: <InstallDir>\SharedContent\Bin\SmartPlantStructure\Symbols\Release\SPSFootingMacros.dll
//  Original Class Name: ‘PierAndSlabFtgDef’ and ‘PierAndSlabFtgSym’ in VB content
//
//Abstract
//   FtgPierDef is a .NET custom assembly definition which creates graphic outputs for representing a grout pad, a pier and a slab footing in the model.
//   This class subclasses from FootingCustomAssemblyDefinition.
//'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.ReferenceData.Exceptions;
using Ingr.SP3D.Common.Exceptions;
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
    [OutputNotification(SPSSymbolConstants.IID_IJDAttributes)]
    [OutputNotification(SPSSymbolConstants.IID_IJStructMaterial)]
    public class PierAndSlabFtgDef : FootingCustomAssemblyDefinition
    {
        private bool isOnPreLoad = false;
        //===============================================================================================================
        //DefinitionName/ProgID of this symbol is "FootingCustomAssemblies,Ingr.SP3D.Content.Structure.PierAndSlabFtgDef"
        //===============================================================================================================

        #region Definition of inputs
        [InputCatalogPart(1)]
        public InputCatalogPart catalogPartInput;
        [InputDouble(2, "SlabShape", "Slab Shape", 1)]
        public InputDouble slabShapeInput;
        [InputDouble(3, "SlabSizingRule", "Slab Sizing Rule", 1)]
        public InputDouble slabSizingRuleInput;
        [InputDouble(4, "SlabOrientation", "Slab Orientation", 1)]
        public InputDouble slabOrientationInput;
        [InputDouble(5, "SlabRotationAngle", "Slab Rotation Angle", 0)]
        public InputDouble slabRotationAngleInput;
        [InputDouble(6, "SlabEdgeClearance", "Slab Edge Clearance", 24)]
        public InputDouble slabEdgeClearanceInput;
        [InputDouble(7, "SlabSizeIncrement", "Slab Size Increment", 1)]
        public InputDouble slabSizeIncrementInput;
        [InputDouble(8, "SlabLength", "Slab Length", 5)]
        public InputDouble slabLengthInput;
        [InputDouble(9, "SlabWidth", "Slab Width", 5)]
        public InputDouble slabWidthInput;
        [InputDouble(10, "SlabHeight", "Slab Height", 16)]
        public InputDouble slabHeightInput;
        [InputDouble(11, "WithPier", "With Pier", 1)]
        public InputDouble withPierInput;
        [InputDouble(12, "WithGroutPad", "With Grout Pad", 1)]
        public InputDouble withGroutPadInput;
        [InputDouble(13, "GlobalDelta", "Global Delta", 0)]
        public InputDouble globalDeltaInput;
        [InputDouble(14, "PierShape", "Pier Shape", 3)]
        public InputDouble pierShapeInput;
        [InputDouble(15, "PierSizingRule", "Pier Sizing Rule", 0)]
        public InputDouble pierSizingRuleInput;
        [InputDouble(16, "PierOrientation", "Pier Orientation", 0)]
        public InputDouble pierOrientationInput;
        [InputDouble(17, "PierRotationAngle", "Pier Rotation Angle", 0)]
        public InputDouble pierRotationAngleInput;
        [InputDouble(18, "PierEdgeClearance", "Pier Edge Clearance", 0)]
        public InputDouble pierEdgeClearanceInput;
        [InputDouble(19, "PierSizeIncrement", "Pier Size Increment", 1)]
        public InputDouble pierSizeIncrementInput;
        [InputDouble(20, "PierChamfered", "Pier Chamfered", 0)]
        public InputDouble pierChamferedInput;
        [InputDouble(21, "PierChamferSize", "Pier Chamfer Size", 1)]
        public InputDouble pierChamferSizeInput;
        [InputDouble(22, "PierLength", "Pier Length", 20)]
        public InputDouble pierLengthInput;
        [InputDouble(23, "PierWidth", "Pier Width", 20)]
        public InputDouble pierWidthInput;
        [InputDouble(24, "PierHeight", "Pier Height", 16)]
        public InputDouble pierHeightInput;
        [InputDouble(25, "GroutShape", "Grout Shape", 1)]
        public InputDouble groutShapeInput;
        [InputDouble(26, "GroutSizingRule", "Grout Sizing Rule", 1)]
        public InputDouble groutSizingRuleInput;
        [InputDouble(27, "GroutOrientation", "Grout Orientation", 3)]
        public InputDouble groutOrientationInput;
        [InputDouble(28, "GroutRotationAngle", "Grout Rotation Angle", 0)]
        public InputDouble groutRotationAngleInput;
        [InputDouble(29, "GroutEdgeClearance", "Grout Edge Clearance", 0)]
        public InputDouble groutEdgeClearanceInput;
        [InputDouble(30, "GroutLength", "Grout Length", 16)]
        public InputDouble groutLengthInput;
        [InputDouble(31, "GroutWidth", "Grout Width", 16)]
        public InputDouble groutWidthInput;
        [InputDouble(32, "GroutHeight", "Grout Height", 1)]
        public InputDouble groutHeightInput;
        [InputString(33, "SlabSPSMaterial", "Slab Material", "Concrete")]
        public InputString slabMaterialInput;
        [InputString(34, "SlabSPSGrade", "Slab Material Grade", "Fc 3000")]
        public InputString slabMaterialGradeInput;
        [InputString(35, "PierSPSMaterial", "Pier Material", "Concrete")]
        public InputString pierMaterialInput;
        [InputString(36, "PierSPSGrade", "Pier Material Grade", "Fc 3000")]
        public InputString pierMaterialGradeInput;
        [InputString(37, "GroutSPSMaterial", "Grout Material", "Grout")]
        public InputString groutMaterialInput;
        [InputString(38, "GroutSPSGrade", "Grout Material Grade", "A")]
        public InputString groutMaterialGradeInput;
        #endregion Definition of inputs

        #region Definitions of aspects and their outputs
        [Aspect("Physical", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput(SPSSymbolConstants.Slab, "Slab Geometry")]
        [SymbolOutput(SPSSymbolConstants.Pier, "Pier Geometry")]
        [SymbolOutput(SPSSymbolConstants.Grout, "Grout Geometry")]
        public AspectDefinition simplePhysicalAspect;
        [Aspect("ReferenceGeometry", "Reference Geometry Aspect", AspectID.ReferenceGeometry)]
        [SymbolOutput("Point1", "Point1")]
        [SymbolOutput("Point2", "Point2")]
        [SymbolOutput("Point3", "Point3")]
        [SymbolOutput("Point4", "Point4")]
        public AspectDefinition referenceGeometryAspect;
        #endregion Definitions of aspects and their outputs

        #region Public override properties and methods

        /// <summary>
        /// Overridden to perform actions prior to calling the construction of the symbol outputs.
        /// Here we will set all properties on Footing which need access to Footing Occurrence.
        /// Note that for 'Shared' symbols, Occurrence will not be available in ConstructOutputs.
        /// Note that properties are set using generic property access as symbol inputs are not available in this method.
        /// </summary>
        public override void PreConstructOutputs()
        {
            double sectionWidth = 0;
            double sectionDepth = 0;
            double height = 0;
            double memberAngle = 0;
            try
            {
                Footing footing = (Footing)base.Occurrence;
                bool isPlacedByPoint = base.IsPlacedByPoint(footing);

                //Call the evaluate which will return the section dimension and member angle.  These values
                //will be used to compute the length and width of the footing component and the rotation angle.  The footing matrix
                //is set with regards to the footing origin if not attached to a member or the bottom of the member if attached to a member
                base.Evaluate(null, footing, out sectionDepth, out sectionWidth, out height, out memberAngle);
                
                //Set the member angle to global delta symbol input so that we can retrieve it in ConstructOutouts through Symbol Inputs
                footing.SetPropertyValue(memberAngle, SPSSymbolConstants.IJUASPSPierAndSlabFooting, SPSSymbolConstants.GlobalDelta);

                //set the grout sizing rule to 'User Defined' in case placed by point
                //ToDo list is created with error type hence stop computation
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    return;
                }

                bool withGroutPad = FootingServices.DoesFootingHaveGrout(footing);

                if (withGroutPad)
                {
                    //set the grout sizing rule to 'User Defined' in case placed by point
                    if (isPlacedByPoint)
                    {
                        FootingServices.SetSizingRule(footing, FootingComponentType.Grout, FootingServices.UserDefined);
                    }

                    //Set the length and width properties for Grout based on the sizing rule and angle
                    //if the sizing rule is not user defined then take into account the sizing rule and angle                    
                    int groutSizingRule = FootingServices.GetSizingRule(footing, FootingComponentType.Grout);
                    if (groutSizingRule != SPSSymbolConstants.SIZING_RULE_USERDEFINED)
                    {
                        double groutRotationAngle = FootingServices.GetRotationAngle(footing, FootingComponentType.Grout);
                        FootingServices.SetLengthAndWidthPropertiesByRule(footing, FootingComponentType.Grout, sectionDepth, sectionWidth, groutRotationAngle, withGroutPad);
                    }
                }
                else if (isPlacedByPoint)
                {
                    //set the pier sizing rule to 'User Defined' in case placed by point and without grout pad
                    FootingServices.SetSizingRule(footing, FootingComponentType.Pier, FootingServices.UserDefined);
                }

                //Set the length and width properties for Pier based on the sizing rule and angle
                //if the sizing rule is not user defined then take into account the sizing rule and angle
                int pierSizingRule = FootingServices.GetSizingRule(footing, FootingComponentType.Pier);
                if (pierSizingRule != SPSSymbolConstants.SIZING_RULE_USERDEFINED)
                {
                    double pierRotationAngle = FootingServices.GetRotationAngle(footing, FootingComponentType.Pier);
                    FootingServices.SetLengthAndWidthPropertiesByRule(footing, FootingComponentType.Pier, sectionDepth, sectionWidth, pierRotationAngle, withGroutPad);
                }
                //Set the length and width properties for Slab based on the sizing rule and angle
                //if the sizing rule is not user defined then take into account the sizing rule and angle
                int slabSizingRule = FootingServices.GetSizingRule(footing, FootingComponentType.Slab);
                if (slabSizingRule != SPSSymbolConstants.SIZING_RULE_USERDEFINED)
                {
                    double slabRotationAngle = FootingServices.GetRotationAngle(footing, FootingComponentType.Slab);
                    FootingServices.SetLengthAndWidthPropertiesByRule(footing, FootingComponentType.Slab, sectionDepth, sectionWidth, slabRotationAngle, withGroutPad);
                }
            }
            catch // General Unhandled exception  
            {
                //check if ToDo message already created or not
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError,
                        FootingLocalizer.GetString(FootingResourceIDs.ErrPreConstructOutputs,
                        "Error while evaluating Footing in PreConstructOutputs in Footing custom assembly definition. Check your custom code or contact S3D support."));
                }
            }
        }
        /// <summary>
        /// Constructs the outputs for the Footing aspects.
        /// Dimensions are set above in PreConstructOutputs. Here use these values from symbol inputs
        /// and construct outputs
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                //checking undefined value
                this.ValidateUndefinedCodelistValue();
                //ToDo list is created with error type hence stop computation
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    return;
                }

                bool withGroutPad = (int)this.withGroutPadInput.Value == 0 ? false : true;

                //Get the connection
                SP3DConnection connection = base.OccurrenceConnection;

                //Adding the created point to the ReferenceGeometry aspect  
                Point3d point3D = new Point3d(connection, 0, 0, 0);
                referenceGeometryAspect.Outputs["Point1"] = point3D;

                int shapeType = 0;
                double length = 0.0;
                double width = 0.0;
                double clearance = 0.0; 
                double projectionOffset = 0.0;
                double groutHeight = 0.0;
                double rotationAngle = 0;
                int orientationType = 0;
                int sizingRule = 0;
                double globalDelta = 0;
                FootingComponentType componentType = FootingComponentType.Grout;
                ToDoListMessage invalidCodeListItemToDoMessage = null;
                    
                //------------------------------------------------------------------------------------------------------------------------------------
                //GroutPad Geometry 
                //------------------------------------------------------------------------------------------------------------------------------------
                if (withGroutPad)
                {
                    //Place the shape, the shape is based on the shape property of the object.  The shape is a projection3d output which
                    //is transformed based on the member angle and set on the aspect as an output based on the component name.  For grout,
                    //the projection offset is 0 since we are projecting on -z to the grout height.
                    shapeType = (int)this.groutShapeInput.Value;
                    length = this.groutLengthInput.Value;
                    width = this.groutWidthInput.Value;
                    clearance = this.groutEdgeClearanceInput.Value;
                    groutHeight = this.groutHeightInput.Value;
                    orientationType = (int) this.groutOrientationInput.Value;
                    rotationAngle = this.groutRotationAngleInput.Value;
                    sizingRule = (int)this.groutSizingRuleInput.Value;
                    globalDelta = this.globalDeltaInput.Value;

                    //Validate code list values and generate ToDoRecord if needed
                    invalidCodeListItemToDoMessage = FootingServices.AreCodeListValuesValid(componentType, orientationType, sizingRule, shapeType);
                    if (invalidCodeListItemToDoMessage != null)
                    {
                        base.ToDoListMessage = invalidCodeListItemToDoMessage;
                        return;
                    }

                    //Construct component geometry                
                    FootingServices.PlaceFootingShape(base.OccurrenceConnection, shapeType, length, width, groutHeight, clearance, rotationAngle, orientationType, globalDelta, componentType, projectionOffset, simplePhysicalAspect);

                    //Add the created point to the ReferenceGeometry aspect
                    point3D = new Point3d(connection, 0, 0, -groutHeight);
                    referenceGeometryAspect.Outputs["Point2"] = point3D;
                }
                else
                {
                    groutHeight = 0.0;
                }

                //------------------------------------------------------------------------------------------------------------------------------------
                //Pier Geometry 
                //------------------------------------------------------------------------------------------------------------------------------------				
                //Place the shape, the shape is based on the shape property of the object.  The shape is a projection3d output which
                //is transformed based on the member angle and set on the aspect as an output based on the component name.  For pier,
                //the projection offset is grout height since we are projecting on -z to the pier height.  The pier is placed below the
                //the grout pad.
                componentType = FootingComponentType.Pier;
                projectionOffset = groutHeight;
                shapeType = (int) this.pierShapeInput.Value;
                length = this.pierLengthInput.Value;
                width = this.pierWidthInput.Value;
                double pierHeight = this.pierHeightInput.Value;
                clearance = this.pierEdgeClearanceInput.Value;
                orientationType = (int)this.pierOrientationInput.Value;
                rotationAngle = this.pierRotationAngleInput.Value;
                sizingRule = (int) this.groutSizingRuleInput.Value;

                //Validate code list values and generate ToDoRecord if needed
                invalidCodeListItemToDoMessage = FootingServices.AreCodeListValuesValid(componentType, orientationType, sizingRule, shapeType);
                if (invalidCodeListItemToDoMessage != null)
                {
                    base.ToDoListMessage = invalidCodeListItemToDoMessage;
                    return;
                }

                FootingServices.PlaceFootingShape(base.OccurrenceConnection, shapeType, length, width, pierHeight, clearance, rotationAngle, orientationType, globalDelta, componentType, projectionOffset, simplePhysicalAspect);

                //4. Add the created point to the ReferenceGeometry aspect
                point3D = new Point3d(connection, 0.0, 0.0, -(pierHeight + groutHeight));
                referenceGeometryAspect.Outputs["Point3"] = point3D;

                //------------------------------------------------------------------------------------------------------------------------------------
                // Slab Geometry 
                //------------------------------------------------------------------------------------------------------------------------------------
                componentType = FootingComponentType.Slab;
                shapeType = (int)this.slabShapeInput.Value;
                length = this.slabLengthInput.Value;
                width = this.slabWidthInput.Value;
                double slabHeight = this.slabHeightInput.Value;
                clearance = this.slabEdgeClearanceInput.Value;
                orientationType = (int)this.slabOrientationInput.Value;
                rotationAngle = this.slabRotationAngleInput.Value;
                sizingRule = (int)this.groutSizingRuleInput.Value;
                //Validate code list values and generate ToDoRecord if needed
                invalidCodeListItemToDoMessage = FootingServices.AreCodeListValuesValid(componentType, orientationType, sizingRule, shapeType);
                if (invalidCodeListItemToDoMessage != null)
                {
                    base.ToDoListMessage = invalidCodeListItemToDoMessage;
                    return;
                }
                //Place the shape, the shape is based on the shape property of the object.  The shape is a projection3d output which
                //is transformed based on the member angle and set on the aspect as an output based on the component name.  For slab,
                //the projection offset is grout height and pier height since we are projecting on -z to the slab height.  The slab
                //is placed below the pier.
                projectionOffset = groutHeight + pierHeight;
                FootingServices.PlaceFootingShape(base.OccurrenceConnection, shapeType, length, width, slabHeight, clearance, rotationAngle, orientationType, globalDelta, componentType, projectionOffset, simplePhysicalAspect);

                //Add the created point to the ReferenceGeometry aspect
                point3D = new Point3d(connection, 0.0, 0.0, -(slabHeight + pierHeight + groutHeight));
                referenceGeometryAspect.Outputs["Point4"] = point3D;
            }
            catch (Exception) // General Unhandled exception  
            {
                //check if ToDo message already created or not
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError,
                        FootingLocalizer.GetString(FootingResourceIDs.ErrPierAndSlabConstructOutputs,
                        "Error while constructing outputs for pier and slab in Pier and Slab Footing Definition. Please check your custom code or contact S3D support."));
                }
            }
        }

        /// <summary>
        /// Evaluate the Footing assembly. Here we will do the following:
        /// 1. Calculate the exposed surface area for each component individually and for the Footing
        /// 2. Set the component material
        /// 3. Calculate the weight and centre of gravity (COG) of Footing. 
        /// 4. Set the calculated weight and COG on the Footing.
        /// </summary>
        public override void EvaluateAssembly()
        {
            try
            {
                Footing footing = (Footing)base.Occurrence;
                bool withGroutPad = FootingServices.DoesFootingHaveGrout(footing);
                double groutSurfaceArea = 0.0;

                //Set the material 
                if (withGroutPad)
                {
                    //Get exposed surface area of the grout                
                    groutSurfaceArea = FootingServices.GetExposedSurfaceArea(footing, FootingComponentType.Grout);
                    //Set component material
                    FootingServices.SetComponentMaterial(footing, FootingComponentType.Grout);
                }
                else
                {
                    //Remove grout material relation name
                    footing.RemoveMaterial(SPSSymbolConstants.Grout);
                }

                //Get exposed surface area of the pier
                double pierSurfaceArea = FootingServices.GetExposedSurfaceArea(footing, FootingComponentType.Pier);
                //Set component material
                FootingServices.SetComponentMaterial(footing, FootingComponentType.Pier);

                //Get exposed surface area of the slab
                double slabSurfaceArea = FootingServices.GetExposedSurfaceArea(footing, FootingComponentType.Slab);
                //Set component material
                FootingServices.SetComponentMaterial(footing, FootingComponentType.Slab);

                //set the total surface area of the footing
                double totalSurfaceArea = groutSurfaceArea + pierSurfaceArea + slabSurfaceArea;
                footing.SetPropertyValue(totalSurfaceArea, SPSSymbolConstants.IJSurfaceArea, SPSSymbolConstants.SurfaceArea);
            }
            catch(RefDataMaterialNotFoundException)
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
                //Evaluate weight and COG for the footing here as evaluate is called after construction of outputs.
                this.EvaluateWeightCG();
            }
            catch (Exception) // General Unhandled exception 
            {
                //check if ToDo message already created or not
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.Occurrence.ToString() + " " +
                        FootingLocalizer.GetString(FootingResourceIDs.ErrPierAndSlabWCOGMissingSystemAttributeData,
                        "cannot calculate weight and centre of gravity for the Piier and Slab Component in PiFooting Definition, as some of the required user attribute value cannot be obtained from the catalog. Check the error log and catalog data."));
                }
            }
        }

        /// <summary>
        /// Updates the rotation value of footing components based on the transform applied to the footing.
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

            //update grout rotation
            bool isGroutNeeded = FootingServices.DoesFootingHaveGrout(footing);
            if (isGroutNeeded)
            {
                base.UpdateComponentRotationAngle(footing, isPlacedByPoint, SPSSymbolConstants.IJUASPSFtgGroutPad, SPSSymbolConstants.GroutOrientation, SPSSymbolConstants.GroutRotationAngle, transformationMatrix);
            }

            //update pier rotation
            base.UpdateComponentRotationAngle(footing, isPlacedByPoint, SPSSymbolConstants.IJUASPSFtgPier, SPSSymbolConstants.PierOrientation, SPSSymbolConstants.PierRotationAngle, transformationMatrix);

            //update slab rotation
            base.UpdateComponentRotationAngle(footing, isPlacedByPoint, SPSSymbolConstants.IJUASPSFtgSlab, SPSSymbolConstants.SlabOrientation, SPSSymbolConstants.SlabRotationAngle, transformationMatrix);
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

        /// <summary>
        /// Returns the name, material, geometry and attributes of the components of the footing.
        /// Implement this method if the footing doesn't have children business objects to represent
        /// the components but information about the component is needed to be provided during data exchange
        /// with other applications.
        /// </summary>
        /// <param name="businessObject">Footing business object.</param>
        /// <returns>Returns a collection of custom output.</returns>
        public override Collection<CustomOutput> GetComponents(BusinessObject businessObject)
        {
            if (businessObject == null)
            {
                throw new ArgumentNullException("businessObject");
            }

            Footing footing = (Footing)businessObject;

            Collection<CustomOutput> components = new Collection<CustomOutput>();
            Collection<Surface3d> groutFaces = new Collection<Surface3d>();
            Collection<Surface3d> pierFaces = new Collection<Surface3d>();
            Collection<Surface3d> slabFaces = new Collection<Surface3d>();

            CustomSurfaceOutput footingComponentSurfaceOutput = null;

            string materialGrade = null;
            string materialType = null;

            //get output geometry of each component from the physical representation and add to the respective collection
            Surface3d groutSurface = (Surface3d)SymbolHelper.GetSymbolOutput(footing, "Physical", SPSSymbolConstants.Grout);
            if (groutSurface != null)
            {
                groutFaces.Add(groutSurface);
            }

            Surface3d pierSurface = (Surface3d)SymbolHelper.GetSymbolOutput(footing, "Physical", SPSSymbolConstants.Pier);
            if (pierSurface != null)
            {
                pierFaces.Add(pierSurface);
            }

            Surface3d slabSurface = (Surface3d)SymbolHelper.GetSymbolOutput(footing, "Physical", SPSSymbolConstants.Slab);
            if (slabSurface != null)
            {
                slabFaces.Add(slabSurface);
            }

            //create output to represent grout
            if (groutFaces.Count > 0)
            {
                //get the grout material type and material grade
                materialType = StructHelper.GetStringProperty(footing, SPSSymbolConstants.IJUASPSFtgGroutPad, SPSSymbolConstants.GroutMaterial);
                materialGrade = StructHelper.GetStringProperty(footing, SPSSymbolConstants.IJUASPSFtgGroutPad, SPSSymbolConstants.GroutMaterialGrade);

                //get properties specific to the grout component from the footing and add to the respective collection
                Collection<PropertyValue> groutProperties = new Collection<PropertyValue>();

                foreach (PropertyValue propertyValue in footing.GetAllProperties())
                {
                    string interfaceName = propertyValue.PropertyInfo.InterfaceInfo.Name;
                    if (interfaceName == SPSSymbolConstants.IJUASPSFtgGroutPad || interfaceName == SPSSymbolConstants.IJUASPSFtgGroutPadDim)
                    {
                        groutProperties.Add(propertyValue);
                    }
                }

                footingComponentSurfaceOutput = new CustomSurfaceOutput(SPSSymbolConstants.Grout, materialType, materialGrade, groutFaces, groutProperties);
                components.Add(footingComponentSurfaceOutput);
            }

            //create output to represent pier
            if (pierFaces.Count > 0)
            {
                //get the pier material type and material grade
                materialType = StructHelper.GetStringProperty(footing, SPSSymbolConstants.IJUASPSFtgPier, SPSSymbolConstants.PierMaterial);
                materialGrade = StructHelper.GetStringProperty(footing, SPSSymbolConstants.IJUASPSFtgPier, SPSSymbolConstants.PierMaterialGrade);

                //get properties specific to the pier component from the footing and add to the respective collection
                Collection<PropertyValue> pierProperties = new Collection<PropertyValue>();

                foreach (PropertyValue propertyValue in footing.GetAllProperties())
                {
                    string interfaceName = propertyValue.PropertyInfo.InterfaceInfo.Name;
                    if (interfaceName == SPSSymbolConstants.IJUASPSFtgPier || interfaceName == SPSSymbolConstants.IJUASPSFtgPierDim)
                    {
                        pierProperties.Add(propertyValue);
                    }
                }

                footingComponentSurfaceOutput = new CustomSurfaceOutput(SPSSymbolConstants.Pier, materialType, materialGrade, pierFaces, pierProperties);
                components.Add(footingComponentSurfaceOutput);
            }

            //create output to represent slab
            if (slabFaces.Count > 0)
            {
                //get the slab material type and material grade
                materialType = StructHelper.GetStringProperty(footing, SPSSymbolConstants.IJUASPSFtgSlab, SPSSymbolConstants.SlabMaterial);
                materialGrade = StructHelper.GetStringProperty(footing, SPSSymbolConstants.IJUASPSFtgSlab, SPSSymbolConstants.SlabMaterialGrade);

                //get properties specific to the slab component from the footing and add to the respective collection
                Collection<PropertyValue> slabProperties = new Collection<PropertyValue>();

                foreach (PropertyValue propertyValue in footing.GetAllProperties())
                {
                    string interfaceName = propertyValue.PropertyInfo.InterfaceInfo.Name;
                    if (interfaceName == SPSSymbolConstants.IJUASPSFtgSlab || interfaceName == SPSSymbolConstants.IJUASPSFtgSlabDim)
                    {
                        slabProperties.Add(propertyValue);
                    }
                }

                footingComponentSurfaceOutput = new CustomSurfaceOutput(SPSSymbolConstants.Slab, materialType, materialGrade, slabFaces, slabProperties);
                components.Add(footingComponentSurfaceOutput);
            }

            return components;
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
                    case SPSSymbolConstants.PierSizingRule:
                    case SPSSymbolConstants.SlabSizingRule:
                    case SPSSymbolConstants.WithGroutPad:
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
            else if (propertyType == SP3DPropType.PTBool)
            {
                PropertyValueBoolean propertyValueBool = (PropertyValueBoolean)newPropertyValue;
                propValueInt = Convert.ToInt32(propertyValueBool.PropValue);
            }

            PropertyDescriptor propertyDescriptor;
            PropertyValue propertyValue;
            string propertyNameToMarkReadOnly;

            //If property name is ReportingRequirements 
            if (propertyName == SPSSymbolConstants.ReportingRequirements)
            {
                if (propValueInt == -1)
                {
                    //grey out the reporting type on the GOPC
                    base.SetAccessOnPropertyDescriptor(allDisplayedValues, SPSSymbolConstants.ReportingType, true);
                }
            }
            else if (propertyName == SPSSymbolConstants.PierSizingRule) //Read only behavior for sizing properties is based on sizing rule for that component 
            {
                FootingServices.SetComponentSizingPropertiesAccess(FootingComponentType.Pier, propValueInt, allDisplayedValues);

                //gray out the pier sizing rule property field read-only when footing is placed with point and without grout 
                propertyToChange.ReadOnly = base.IsPlacedByPoint((Foundation)businessObject) && !FootingServices.DoesFootingHaveGrout((Footing)businessObject) ? true : false;
            }
            else if (propertyName == SPSSymbolConstants.SlabSizingRule)
            {
                FootingServices.SetComponentSizingPropertiesAccess(FootingComponentType.Slab, propValueInt, allDisplayedValues);
            }
            else if (propertyName == SPSSymbolConstants.GroutSizingRule)
            {
                FootingServices.SetComponentSizingPropertiesAccess(FootingComponentType.Grout, propValueInt, allDisplayedValues);

                //gray out the grout sizing rule property field when footing is placed with point
                propertyToChange.ReadOnly = base.IsPlacedByPoint((Foundation)businessObject) ? true : false;
            }
            else if (propertyName == SPSSymbolConstants.WithGroutPad)
            {
                if (allDisplayedValues != null)
                {
                //Decide whether to grey out all grout properties or not, based on this flag
                    for (int i = 0; i < allDisplayedValues.Count; i++)
                    {
                        propertyDescriptor = allDisplayedValues[i];
                        propertyValue = propertyDescriptor.Property;
                        propertyNameToMarkReadOnly = propertyValue.PropertyInfo.Name;
                        switch (propertyNameToMarkReadOnly)
                        {
                            case SPSSymbolConstants.GroutSizingRule:
                            case SPSSymbolConstants.GroutRotationAngle:
                            case SPSSymbolConstants.GroutOrientation:
                            case SPSSymbolConstants.GroutShape:
                            case SPSSymbolConstants.GroutMaterial:
                            case SPSSymbolConstants.GroutMaterialGrade:
                            case SPSSymbolConstants.GroutHeight:
                                //Grout properties which should be read-only or editable based on grout presence only. Not dependent on sizing rule
                                propertyDescriptor.ReadOnly = (propValueInt == 0) ? true : false;
                                break;
                            case SPSSymbolConstants.GroutLength:
                            case SPSSymbolConstants.GroutWidth:
                            case SPSSymbolConstants.GroutEdgeClearance:
                                if (propValueInt == 0)
                                {
                                    propertyDescriptor.ReadOnly = true;
                                }
                                else
                                {
                                    //Need to get the value of Grout sizing rule to determine the fate of the properties
                                    int groutSizingRule = FootingServices.GetGroutSizingRule(allDisplayedValues);
                                    FootingServices.SetComponentSizingPropertiesAccess(FootingComponentType.Grout, groutSizingRule, allDisplayedValues);
                                }
                                break;
                        }
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
        /// Method to evaluate the footing weight and COG.
        /// </summary>
        private void EvaluateWeightCG()
        {
            //getting Footing from occurrence 
            Footing footing = (Footing)base.Occurrence;

            int weightCOGOrigin = StructHelper.GetIntProperty(footing, SPSSymbolConstants.IJWCGValueOrigin, SPSSymbolConstants.DryWCGOrigin);

            if (weightCOGOrigin == SPSSymbolConstants.DRY_WCOG_ORIGIN_COMPUTED)
            {
                Position localCOG = new Position();

                double totalVolume = 0.0;
                double groutWeight = 0.0;
                double pierWeight = 0.0;
                double pierCOGZ = 0.0;
                double slabWeight = 0.0;
                double slabCOGZ = 0.0;
                double groutHeight = 0.0;
                double groutCOGZ = 0.0;
                VolumeCG groutVolumeCG = null;
                bool withGroutPad = FootingServices.DoesFootingHaveGrout(footing);
                if (withGroutPad)
                {
                    //get the volumeCG for each of the components.  Pass the z offset as we are projecting to -z.  For grout there is no z offset.
                    groutHeight = FootingServices.GetComponentHeight(footing, FootingComponentType.Grout);
                    groutVolumeCG = FootingServices.GetComponentVolumeCG(footing, FootingComponentType.Grout, 0);
                    if (groutVolumeCG != null)
                    {
                        totalVolume = groutVolumeCG.Volume;
                        groutCOGZ = groutVolumeCG.COGZ;
                    }
                }

                double pierHeight = FootingServices.GetComponentHeight(footing, FootingComponentType.Pier);

                VolumeCG pierVolumeCG = FootingServices.GetComponentVolumeCG(footing, FootingComponentType.Pier, groutHeight);
                VolumeCG slabVolumeCG = FootingServices.GetComponentVolumeCG(footing, FootingComponentType.Slab, groutHeight + pierHeight);

                if (pierVolumeCG != null)
                {
                    totalVolume = totalVolume + pierVolumeCG.Volume;
                }

                if (slabVolumeCG != null)
                {
                    totalVolume = totalVolume + slabVolumeCG.Volume;
                }

                if (groutVolumeCG != null)
                {
                    groutWeight = SymbolHelper.GetWeightFromVolume(footing, SPSSymbolConstants.IJUASPSFtgGroutPad, SPSSymbolConstants.GroutMaterial, SPSSymbolConstants.GroutMaterialGrade, groutVolumeCG.Volume);
                }

                if (pierVolumeCG != null)
                {
                    pierWeight = SymbolHelper.GetWeightFromVolume(footing, SPSSymbolConstants.IJUASPSFtgPier, SPSSymbolConstants.PierMaterial, SPSSymbolConstants.PierMaterialGrade, pierVolumeCG.Volume);
                    pierCOGZ = pierVolumeCG.COGZ;
                }

                if (slabVolumeCG != null)
                {
                    slabWeight = SymbolHelper.GetWeightFromVolume(footing, SPSSymbolConstants.IJUASPSFtgSlab, SPSSymbolConstants.SlabMaterial, SPSSymbolConstants.SlabMaterialGrade, slabVolumeCG.Volume);
                    slabCOGZ = slabVolumeCG.COGZ;
                }

                //get the total weight for the cog calculation
                double totalWeight = groutWeight + pierWeight + slabWeight;
                localCOG.Z = (groutWeight * groutCOGZ + pierWeight * pierCOGZ + slabWeight * slabCOGZ) / totalWeight;

                footing.SetPropertyValue(totalVolume, SPSSymbolConstants.IJGenericVolume, SPSSymbolConstants.Volume);

                Matrix4X4 matrix = footing.Matrix;
                Position globalCOG = matrix.Transform(localCOG);

                WeightCOGServices weightCOGServices = new WeightCOGServices();
                weightCOGServices.SetWeightAndCOG(footing, totalWeight, globalCOG.X, globalCOG.Y, globalCOG.Z);
            }
        }

        /// <summary>
        /// Checks for undefined value and raise error.
        /// </summary>
        private void ValidateUndefinedCodelistValue()
        {
            int groutShape = Convert.ToInt32(groutShapeInput.Value);
            int groutSizingRule = Convert.ToInt32(groutSizingRuleInput.Value);
            int groutOrientation = Convert.ToInt32(groutOrientationInput.Value);
            int pierShape = Convert.ToInt32(pierShapeInput.Value);
            int pierSizingRule = Convert.ToInt32(pierSizingRuleInput.Value);
            int pierOrientation = Convert.ToInt32(pierOrientationInput.Value);
            int slabShape = Convert.ToInt32(slabShapeInput.Value);
            int slabSizingRule = Convert.ToInt32(slabSizingRuleInput.Value);
            int slabOrientation = Convert.ToInt32(slabOrientationInput.Value);

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructPrismaticFootingShapes, SPSSymbolConstants.UDP, groutShape))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_GROUT_SHAPE,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrGroutShapeCodeListValue,
                    "Error while validating Grout Code list value as Grout Shape value: {0} does not exist in Prismatic Footing Shapes: {1} table. Please check catalog or contact S3D support."), groutShape.ToString(), SPSSymbolConstants.StructPrismaticFootingShapes));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructFootingCompSizingRule, SPSSymbolConstants.UDP, groutSizingRule))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_GROUT_SIZING_RULE,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrGroutSizingRuleCodeListValue,
                    "Error while validating Grout sizing rule as GroutSizingRule value: {0} does not exist in Footing component sizing rule: {1} table. Please check catalog or contact S3D support."), groutSizingRule.ToString(), SPSSymbolConstants.StructFootingCompSizingRule));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructCoordSysReference, SPSSymbolConstants.UDP, groutOrientation))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_GROUT_ORIENTATION,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrGroutOrientationCodeListValue,
                    "Error while validating Grout orientation as GroutOrientation value: {0} does not exist in StructCoordSysReference: {1} table. Please check catalog or contact S3D support."), groutOrientation.ToString(), SPSSymbolConstants.StructCoordSysReference));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructPrismaticFootingShapes, SPSSymbolConstants.UDP, pierShape))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_PIER_SHAPE,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrPierShapeCodeListValue,
                    "Error while validating Pier Code list value as Pier Shape value: {0} does not exist in Prismatic Footing Shapes: {1} table. Please check catalog or contact S3D support."), pierShape.ToString(), SPSSymbolConstants.StructPrismaticFootingShapes));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructFootingCompSizingRule, SPSSymbolConstants.UDP, pierSizingRule))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_PIER_SIZING_RULE,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrPierSizingRuleCodeListValue,
                    "Error while validating Pier sizing rule as PierSizingRule value: {0} does not exist in Footing component sizing rule: {1} table. Please check catalog or contact S3D support."), pierSizingRule.ToString(), SPSSymbolConstants.StructFootingCompSizingRule));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructCoordSysReference, SPSSymbolConstants.UDP, pierOrientation))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_PIER_ORIENTATION,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrPierOrientationCodeListValue,
                    "Error while validating Pier orientation as PierOrientation value: {0} does not exist in StructCoordSysReference: {1} table. Please check catalog or contact S3D support."), pierOrientation.ToString(), SPSSymbolConstants.StructCoordSysReference));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructPrismaticFootingShapes, SPSSymbolConstants.UDP, slabShape))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_SLAB_SHAPE,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrSlabShapeCodeListValue,
                    "Error while validating Slab Code list value as Slab Shape value: {0} does not exist in Prismatic Footing Shapes: {1} table. Please check catalog or contact S3D support."), slabShape.ToString(), SPSSymbolConstants.StructPrismaticFootingShapes));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructFootingCompSizingRule, SPSSymbolConstants.UDP, slabSizingRule))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_SLAB_SIZING_RULE,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrSlabSizingRuleCodeListValue,
                    "Error while validating Slab sizing rule as SlabSizingRule value: {0} does not exist in Footing component sizing rule: {1} table. Please check catalog or contact S3D support."), slabSizingRule.ToString(), SPSSymbolConstants.StructFootingCompSizingRule));
                return;
            }

            if (!StructHelper.IsValidCodeListValue(SPSSymbolConstants.StructCoordSysReference, SPSSymbolConstants.UDP, slabOrientation))
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.StructureToDoMessageCodelist, SPSSymbolConstants.TDL_INVALID_SLAB_ORIENTATION,
                    String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrSlabOrientationCodeListValue,
                    "Error while validating Slab orientation as SlabOrientation value: {0} does not exist in StructCoordSysReference: {1} table. Please check catalog or contact S3D support."), slabOrientation.ToString(), SPSSymbolConstants.StructCoordSysReference));
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
                case SPSSymbolConstants.PierSizeIncrement:
                case SPSSymbolConstants.SlabSizeIncrement:
                    isValid = !ValidationHelper.IsNegative(value, ref errorMessage);
                    break;
                case SPSSymbolConstants.GroutLength:
                case SPSSymbolConstants.GroutWidth:
                case SPSSymbolConstants.GroutHeight:
                case SPSSymbolConstants.GroutEdgeClearance:
                case SPSSymbolConstants.PierLength:
                case SPSSymbolConstants.PierWidth:
                case SPSSymbolConstants.PierHeight:
                case SPSSymbolConstants.PierEdgeClearance:
                case SPSSymbolConstants.SlabLength:
                case SPSSymbolConstants.SlabWidth:
                case SPSSymbolConstants.SlabHeight:
                case SPSSymbolConstants.SlabEdgeClearance:
                    isValid = ValidationHelper.IsGreaterThanZero(value, ref errorMessage);
                    break;
            }

            return isValid;
        }

        #endregion Private methods
    }
}