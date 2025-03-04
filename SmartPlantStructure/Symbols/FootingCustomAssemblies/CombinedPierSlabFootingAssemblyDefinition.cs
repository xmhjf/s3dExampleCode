//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  CombinedPierSlabFootingAssemblyDefinition.cs
// 
//  Original Dll Name: <InstallDir>\SharedContent\Bin\SmartPlantStructure\Symbols\Release\SPSFootingMacros.dll
//  Original Class Name: ‘CombPierSlabFtgAsmDef’ in VB content
//
//Abstract
//   CombinedPierSlabFootingAssemblyDefinition is a .NET custom assembly definition which creates grout pads, piers, an octagonal or rectangular slab and control points in the model.
//   This class subclasses from FootingCustomAssemblyDefinition.
//''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
using System;
using System.Collections.Generic;
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
    /// Combined pier and slab footing CustomAssemblyDefinition.
    /// It will update the IJDAttributes and IJStructElevationDatum interfaces for this definition and 
    /// the assembly outputs will have their geometry (IJDGeometry) and attributes (IJDAttributes) modified.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    [OutputNotification(SPSSymbolConstants.IID_IJDAttributes)]
    [OutputNotification(SPSSymbolConstants.IID_IJStructElevationDatum)]
    [OutputNotification(SPSSymbolConstants.IID_IJDAttributes, true)] // The assembly outputs will have their attributes (IJDAttributes) modified.
    [OutputNotification(SPSSymbolConstants.IID_IJDGeometry, true)] // The assembly outputs will have their geometry (IJDGeometry) modified.
    public class CombinedPierSlabFootingAssemblyDefinition : FootingCustomAssemblyDefinition
    {
        //========================================================================================================================================
        //DefinitionName/ProgID of this symbol is "FootingCustomAssemblies,Ingr.SP3D.Content.Structure.CombinedPierSlabFootingAssemblyDefinition"
        //========================================================================================================================================
        #region Definition of inputs
        [InputCatalogPart(1)]
        public InputCatalogPart catalogPart;
        #endregion Definition of inputs

        #region Definitions of assemblies and their outputs
        //collection of grout assembly output
        //grout assembly outputs will have their geometry (IJDGeometry), attributes (IJDAttributes), material (IJStructMaterial) modified.
        [OutputNotification(SPSSymbolConstants.IID_IJStructMaterial)]
        [AssemblyOutput(1, "Combined Footing Grout")]
        public AssemblyOutputs groutAssemblyOutputs;

        //octagonal slab assembly output
        //octagonal slab assembly output will have their geometry (IJDGeometry), attributes (IJDAttributes), material (IJStructMaterial) modified.
        [OutputNotification(SPSSymbolConstants.IID_IJStructMaterial)]
        [AssemblyOutput(2, "Combined Footing Slab Oct")]
        public AssemblyOutput octagonalSlabAssemblyOutput;

        //rectangular slab assembly output
        //rectangular slab assembly output will have their geometry (IJDGeometry), attributes (IJDAttributes), material (IJStructMaterial) modified.
        [OutputNotification(SPSSymbolConstants.IID_IJStructMaterial)]
        [AssemblyOutput(3, "Combined Footing Slab Rect")]
        public AssemblyOutput rectangularSlabAssemblyOutput;

        //collection of pier assembly output
        //pier assembly outputs will have their geometry (IJDGeometry), attributes (IJDAttributes), material (IJStructMaterial) modified.                
        [OutputNotification(SPSSymbolConstants.IID_IJStructMaterial)]
        [AssemblyOutput(4, "Combined Footing Pier")]
        public AssemblyOutputs pierAssemblyOutputs;

        //collection of control point assembly output
        //control point assembly outputs will have their geometry (IJDGeometry) and attributes (IJDAttributes) modified.        
        [AssemblyOutput(5, "Control Point")]
        public AssemblyOutputs controlPointAssemblyOutputs;
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
                //getting the combined footing from occurrence property
                Footing footing = (Footing)base.Occurrence;

                double sectionDepth = 0;
                double sectionWidth = 0;
                double height = 0;
                double memberAngle = 0;
                //Call the evaluate to set the occurrence matrix on combined footing assembly
                base.Evaluate(null, footing, out sectionDepth, out sectionWidth, out height, out memberAngle);
                //ToDo list is created with error type hence stop computation
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    return;
                }

                //Gets the bottom plane method specified on the footing.
                FoundationBottomPlaneMethod bottomPlaneMethod = footing.BottomPlaneMethod;

                //Determines whether the combined footing assembly is placed by point or MemberSystem
                bool isPlacedByPoint = base.IsPlacedByPoint(footing);

                //*****************************************************
                //Construct the variable output grout pads if required
                //*****************************************************
                Collection<BusinessObject> supportedObjects = footing.SupportedObjects;
                int supportedObjectsCount = supportedObjects.Count;
                if (supportedObjectsCount < 1)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, String.Format(FootingLocalizer.GetString(FootingResourceIDs.ErrCombinedFtgInsufficientNumberOfSupportedObject,
                        "Error in evaluating custom assembly outputs of CombinedPierSlab Footing custom assembly definition as supported objects count: {0} is less than one for combined footings. Please check custom code or contact S3D support."), supportedObjectsCount.ToString()));
                    return;
                }

                //If grout pad is needed, a grout pad is needed for each supported object.
                bool withGroutPad = FootingServices.DoesFootingHaveGrout(footing);

                //Number of new grout needed or removed is difference of supportedObjectsCount - existing grout count
                int numberOfNewGroutNeededOrRemoved = (withGroutPad) ? (supportedObjectsCount - this.groutAssemblyOutputs.Count) : -(this.groutAssemblyOutputs.Count);

                //create or remove the grout assembly outputs.
                string foundationComponentPartName = FootingServices.GetFootingComponentPartName(footing, FootingComponentType.Grout);
                base.CreateOrRemoveAssemblyOutputs(footing, this.groutAssemblyOutputs, withGroutPad, numberOfNewGroutNeededOrRemoved, foundationComponentPartName);

                //set the grout sizing rule to 'User Defined' in case placed by point
                if (isPlacedByPoint && withGroutPad)
                {
                    foreach (BusinessObject groutAssemblyOutput in this.groutAssemblyOutputs)
                    {
                        FootingServices.SetSizingRule(groutAssemblyOutput, FootingComponentType.Grout, FootingServices.UserDefined);
                    }
                }

                //************************************************
                //Construct / Update the output rectangular slab
                //************************************************
                string slabComponentName = FootingServices.GetFootingComponentPartName(footing, FootingComponentType.Slab);
                bool isRectangularSlab = false; FoundationComponent slabComponent = null;
                string slabInterfaceName = string.Empty, slabHeightPropertyName = string.Empty;
                switch (slabComponentName)
                {
                    case SPSSymbolConstants.RectFootingSlab:
                        //Boolean flag which will be indicate if the slab is rectangular and octagonal.
                        isRectangularSlab = true;
                        slabComponent = base.CreateAssemblyOutput(footing, this.rectangularSlabAssemblyOutput, slabComponentName);
                        //the octagonal slab is not required now, delete it if it has been previously created
                        base.DeleteAssemblyOutput(this.octagonalSlabAssemblyOutput);
                        slabInterfaceName = SPSSymbolConstants.IJUASPSFtgSlabDim;
                        slabHeightPropertyName = SPSSymbolConstants.SlabHeight;
                        break;
                    case SPSSymbolConstants.OctFootingSlab:
                        slabComponent = base.CreateAssemblyOutput(footing, this.octagonalSlabAssemblyOutput, slabComponentName);
                        //the rectangular slab is not required now, delete it if it has been previously created
                        base.DeleteAssemblyOutput(this.rectangularSlabAssemblyOutput);
                        slabInterfaceName = SPSSymbolConstants.IJUAOctagonalSlabDim;
                        slabHeightPropertyName = SPSSymbolConstants.OctSlabHeight;
                        break;
                }

                if (slabComponent == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FootingLocalizer.GetString(FootingResourceIDs.ErrMissingSlabComponent,
                        "Unable to get slab component from CreateAssemblyOutput in CombinedPierSlab Footing custom assembly definition as Pier height is invalid, while constructing slab. Please check custom code or contact S3D support."));
                    return;
                }

                //*************************************
                //Construct the variable output piers
                //*************************************
                bool isMergePiers = StructHelper.GetBoolProperty(footing, SPSSymbolConstants.IJUACombinedPiers, SPSSymbolConstants.MergePiers);
                double maximumDistanceBetweenTwoSupportedObjects = StructHelper.GetDoubleProperty(footing, SPSSymbolConstants.IJUACombinedPiers, SPSSymbolConstants.MaxDistance);
                Dictionary<int, Collection<BusinessObject>> supportedObjectsConsideringMerge = base.GetSupportedObjectsForMerge(isMergePiers, maximumDistanceBetweenTwoSupportedObjects);
                int supportedObjectsConsideringMergeCount = supportedObjectsConsideringMerge.Count;
                //Number of new pier needed or removed is difference of supportedObjectsConsideringMergeCount - existing pier count
                int numberOfNewPierNeededOrRemoved = supportedObjectsConsideringMergeCount - this.pierAssemblyOutputs.Count;

                //Gets the lowest supported object position and its index in the supported object collection
                int indexOfLowestSupportedObject;
                Position lowestSupportedObjectPosition = base.GetLowestSupportedObjectPosition(supportedObjects, out indexOfLowestSupportedObject);

                //in case the footing bottom plane method either Associative or ElevationValue then gets the position on the datum plane by projecting the lowest supported object position.
                Position projectedPositionOnDatumPlane = null;
                if (bottomPlaneMethod != FoundationBottomPlaneMethod.None)
                {
                    projectedPositionOnDatumPlane = base.GetProjectedPositionOnDatumPlane(lowestSupportedObjectPosition);
                }

                double slabHeight = StructHelper.GetDoubleProperty(slabComponent, slabInterfaceName, slabHeightPropertyName);
                bool isValidPierHeight = FootingServices.IsValidPierHeight(this.groutAssemblyOutputs, lowestSupportedObjectPosition, indexOfLowestSupportedObject, withGroutPad, slabHeight, bottomPlaneMethod, projectedPositionOnDatumPlane);
                if (!isValidPierHeight)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, FootingLocalizer.GetString(FootingResourceIDs.ErrInvalidPierHeight,
                        "Pier height is invalid in CombinedPierSlab Footing custom assembly definition. Check the bottom plane and the defined grout thickness."));
                    return;
                }

                //create or remove the pier assembly outputs.
                foundationComponentPartName = FootingServices.GetFootingComponentPartName(footing, FootingComponentType.Pier);
                base.CreateOrRemoveAssemblyOutputs(footing, this.pierAssemblyOutputs, numberOfNewPierNeededOrRemoved, foundationComponentPartName);

                //set the pier sizing rule to 'User Defined' in case placed by point and without grout pad
                if (isPlacedByPoint && !withGroutPad)
                {                    
                    foreach (BusinessObject pierAssemblyOutput in this.pierAssemblyOutputs)
                    {
                        FootingServices.SetSizingRule(pierAssemblyOutput, FootingComponentType.Pier, FootingServices.UserDefined);
                    }
                }

                //*********************************************
                //Construct the variable output control points
                //*********************************************
                //Creates key points and control points that can be selected as layout points in drawings on combined footings.
                int numberOfPierVertices = supportedObjectsConsideringMergeCount * 4; //considering piers are rectangular and in top face it will have 4 vertices
                int numberOfSlabVertices = (isRectangularSlab) ? 8 : 16; //in case of rectangular slab we have total 8 vertices, whereas for octagonal slab it has 16 vertices considering top and bottom faces
                int numberOfPierAndSlabVertices = numberOfPierVertices + numberOfSlabVertices;
                //Number of new control points needed or removed is difference of (supportedObjectsCount + numberOfPierAndSlabVertices) - existing control point count
                int numberOfNewControlPointNeededOrRemoved = supportedObjectsCount + numberOfPierAndSlabVertices - this.controlPointAssemblyOutputs.Count;

                if (numberOfNewControlPointNeededOrRemoved > 0)
                {
                    Position location = new Position(0.0, 0.0, 0.0);
                    double diameter = 0.001;
                    bool displayInSimplePhysicalAspect = false;
                    int controlPointSubType = (int)ControlPointSubType.Foundation;

                    //New control points needed.
                    for (int i = 0; i < numberOfNewControlPointNeededOrRemoved; i++)
                    {
                        //Construct ControlPoint and add them in ControlPoint assembly outputs 
                        //Key points are placed at the vertices of slabs and piers.
                        //Control points are placed at supported object locations. 
                        int controlPointType = (i > numberOfPierAndSlabVertices - 1) ? (int)ControlPointType.ControlPoint : (int)ControlPointType.Keypoint;
                        ControlPoint controlPoint = base.CreateControlPoint(footing, location, diameter, displayInSimplePhysicalAspect, controlPointType, controlPointSubType);
                        this.controlPointAssemblyOutputs.Add(controlPoint);
                    }
                }
                else
                {
                    //Remove extra control points starting from back of the collection, numberOfNewControlPointsNeededOrRemoved is –ve here.
                    int countRemoved = 0;
                    for (int i = this.controlPointAssemblyOutputs.Count - 1; countRemoved > numberOfNewControlPointNeededOrRemoved; countRemoved--, i--)
                    {
                        this.controlPointAssemblyOutputs.RemoveAt(i);
                    }
                }

                //********************
                //Evaluate grout pads 
                //********************
                for (int i = 0; i < this.groutAssemblyOutputs.Count; i++)
                {
                    //Evaluate grout pad                     
                    BusinessObject supportedObject = supportedObjects[i];
                    FoundationComponent groutComponent = (FoundationComponent)this.groutAssemblyOutputs[i];
                    this.EvaluateGroutComponent(footing, groutComponent, supportedObject);
                }

                //***************
                //Evaluate piers 
                //***************
                for (int i = 0; i < this.pierAssemblyOutputs.Count; i++)
                {
                    //Evaluate pier                     
                    this.EvaluatePierComponent(footing, i, supportedObjectsConsideringMerge[i], isPlacedByPoint, withGroutPad, slabHeight, bottomPlaneMethod);
                    //If ToDo list is created with error type stop computation
                    if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        return;
                    }
                }

                //**************
                //Evaluate slab
                //**************
                this.EvaluateSlabComponent(footing, supportedObjects, bottomPlaneMethod, projectedPositionOnDatumPlane, slabComponent, slabHeight, withGroutPad);

                //************************
                //Evaluate control points 
                //************************
                this.EvaluateControlPoints(supportedObjects, isRectangularSlab, slabComponent, slabHeight, numberOfPierVertices, numberOfPierAndSlabVertices);

                //if the footing bottom plane method neither Associative nor ElevationValue then add it to ToDo list
                if (bottomPlaneMethod == FoundationBottomPlaneMethod.None && base.SupportingPlaneRequirementStatus == SupportingPlaneStatus.Required)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.FootingToDoMessageCodelist,
                        SPSSymbolConstants.TDL_FTG_INSUFFICIENT_INPUTS, FootingLocalizer.GetString(FootingResourceIDs.ErrFootingInsufficientInputs,
                        "Footing cannot be placed using selected inputs as bottom plane method neither Associative nor ElevationValue in CombinedPierSlab Footing custom assembly definition. Please Delete and replace inputs."));
                }
            }
            catch (Exception) // General Unhandled exception 
            {
                //check if ToDo message already created or not
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError,
                        FootingLocalizer.GetString(FootingResourceIDs.ErrEvaluateCombinedPierSlabFootingAssembly,
                        "Unexpected error while evaluating CombinedPierSlab Footing custom assembly definition. Please check custom code or contact S3D support."));
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
                Collection<FoundationComponent> grouts = FootingServices.GetGroutOrPierComponents(footing, FootingComponentType.Grout);
                for (int i = 0; i < grouts.Count; i++)
                {
                    //get the grout component orientation
                    FoundationComponent groutComponent = grouts[i];
                    base.UpdateComponentRotationAngle(groutComponent, isPlacedByPoint, SPSSymbolConstants.IJUASPSFtgGroutPad, SPSSymbolConstants.GroutOrientation, SPSSymbolConstants.GroutRotationAngle, transformationMatrix);
                }
            }

            //update pier's rotation
            Collection<FoundationComponent> piers = FootingServices.GetGroutOrPierComponents(footing, FootingComponentType.Pier);
            for (int i = 0; i < piers.Count; i++)
            {
                //get the pier component orientation
                FoundationComponent pierComponent = piers[i];
                base.UpdateComponentRotationAngle(pierComponent, isPlacedByPoint, SPSSymbolConstants.IJUASPSFtgPier, SPSSymbolConstants.PierOrientation, SPSSymbolConstants.PierRotationAngle, transformationMatrix);
            }

            //update slab rotation irrespective of place by points or slab component orientation
            //get the slab component
            FoundationComponent slabComponent = FootingServices.GetComponent(footing, FootingComponentType.Slab);
            base.UpdateComponentRotationAngle(slabComponent, true, SPSSymbolConstants.IJUASPSFtgSlab, SPSSymbolConstants.SlabOrientation, SPSSymbolConstants.SlabRotationAngle, transformationMatrix);

            //Update the datum elevation so that it is consistent with applied transform (either applied interactively or through MDR)
            base.SetElevationDatumOnTransform(footing, transformationMatrix);
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

            //Need to know what the new values of MergePiers and UseElevationDatum plane properties are in the collection of property descriptors.
            bool mergePiers = false;
            bool useElevationDatum = false;
            for (int i = 0; i < allDisplayedValues.Count; i++)
            {
                PropertyDescriptor propertyDescriptor = allDisplayedValues[i];
                PropertyValue propertyValue = propertyDescriptor.Property;
                string propertyName = propertyValue.PropertyInfo.Name;

                //make all these properties read-only
                if (propertyName == SPSSymbolConstants.MergePiers)
                {
                    PropertyValueBoolean propertyValueBool = (PropertyValueBoolean)propertyValue;
                    mergePiers = (bool)propertyValueBool.PropValue;
                }
                else if (propertyName == SPSSymbolConstants.UseElevationDatum)
                {
                    //UseElevationDatum is read-only on the property page as there is no way to select a plane if 'UseElevationDatum' can be set to 'False'
                    propertyDescriptor.ReadOnly = true;
                    //In multi select mode, if the values are different for different footings, it is empty
                    PropertyValueBoolean propertyValueBool = (PropertyValueBoolean)propertyValue;
                    bool? propValue = propertyValueBool.PropValue;
                    useElevationDatum = (propValue != null) ? (bool)propValue : false;
                }
            }

            for (int i = 0; i < allDisplayedValues.Count; i++)
            {
                PropertyDescriptor propertyDescriptor = allDisplayedValues[i];
                PropertyValue propertyValue = propertyDescriptor.Property;
                string propertyName = propertyValue.PropertyInfo.Name;
                switch (propertyName)
                {
                    //Need to make the MaxDistance property read-only if MergePiers property value is true as we don't want to allow setting maximum distance for Piers if they are already merged.
                    case SPSSymbolConstants.MaxDistance:
                        propertyDescriptor.ReadOnly = (mergePiers) ? false : true;
                        break;
                    //Need to make the BottomElevation property read-only if UseElevationDatum property is false as we don't want to allow setting bottom elevation if user is not using elevation datum.
                    case SPSSymbolConstants.BottomElevation:
                        propertyDescriptor.ReadOnly = (!useElevationDatum) ? true : false;
                        break;
                    case SPSSymbolConstants.Volume:
                        propertyDescriptor.ReadOnly = true;
                        break;
                }
            }
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

            errorMessage = string.Empty;
            bool isOnPropertyChange = true;
            string propertyName = propertyToChange.Property.PropertyInfo.Name;
            switch (propertyName)
            {
                //Need to make the MaxDistance property read-only if MergePiers property value is true as we don't want to allow setting maximum distance for Piers if they are already merged.
                case SPSSymbolConstants.MergePiers:
                    PropertyValueBoolean propertyValueBool = (PropertyValueBoolean)newPropertyValue;
                    bool mergePiers = (bool)propertyValueBool.PropValue;
                    base.SetAccessOnPropertyDescriptor(allDisplayedValues, SPSSymbolConstants.MaxDistance, !mergePiers);
                    break;
                //validate the given bottom elevation value must be lower than elevation of the lowest supported object
                case SPSSymbolConstants.BottomElevation:
                    double bottomElevationValue = ((PropertyValueDouble)newPropertyValue).PropValue.Value;
                    isOnPropertyChange = base.IsValidBottomElevationDatumValue(bottomElevationValue);
                    if (!isOnPropertyChange)
                    {
                        errorMessage = FootingLocalizer.GetString(FootingResourceIDs.ErrFootingInvalidElevation,
                            "Specify an elevation that is below the top of the Footing");
                    }
                    break;
                //validate the given MaxDistance value must be greater than zero
                case SPSSymbolConstants.MaxDistance:
                    isOnPropertyChange = ValidationHelper.IsGreaterThanZero((double)((PropertyValueDouble)newPropertyValue).PropValue, ref errorMessage);
                    break;
            }

            return isOnPropertyChange;
        }

        #endregion Public override properties and methods

        #region Private methods

        /// <summary>
        /// Evaluates the grout component.
        /// </summary>
        /// <param name="footing">The footing.</param>
        /// <param name="groutComponent">The grout component.</param>
        /// <param name="supportedObject">The supported object.</param>
        private void EvaluateGroutComponent(Footing footing, FoundationComponent groutComponent, BusinessObject supportedObject)
        {
            //1. Get the bottom position from the supported member/point to set the grout component origin and the member angle
            Position origin = null;
            double supportedObjectRotationAngle = 0.0;
            MemberSystem memberSystem = supportedObject as MemberSystem;
            Point3d point3d = supportedObject as Point3d;
            double sectionDepth = 0.0, sectionWidth = 0.0;
            if (memberSystem != null)
            {
                origin = base.GetBottomPositionFromMember(memberSystem, out supportedObjectRotationAngle);
                //get cross-section dimensions of the member part
                MemberPart memberPart = base.GetBottomMemberPart(memberSystem);
                sectionDepth = memberPart.CrossSection.Depth;
                sectionWidth = memberPart.CrossSection.Width;
            }
            else if (point3d != null)
            {
                origin = point3d.Position;
                supportedObjectRotationAngle = base.GetPlanAngleFromMatrix(footing.Matrix);
            }

            double groutRotationAngle = 0.0;
            int groutOrientation = 0, groutSizingRule = 0, groutShape = 0;

            //2. Get component rotation and orientation to set the component matrix
            FootingServices.GetDimensions(groutComponent, FootingComponentType.Grout, out groutRotationAngle, out groutOrientation, out groutSizingRule, out groutShape);

            //3. Set the grout component origin
            //Construct a identity matrix
            Matrix4X4 groutComponentMatrix = new Matrix4X4();
            if (origin != null)
            {
                groutComponentMatrix.Origin = origin;
            }

            //4. Rotate the component based on the rotation angles of the component and member angle
            //rotation vector is along global Z direction
            Vector vector = new Vector(0.0, 0.0, 1.0);
            if (groutOrientation == SPSSymbolConstants.ORIENTATION_GLOBAL)
            {
                groutComponentMatrix.Rotate(groutRotationAngle, vector);
            }
            else if (groutOrientation == SPSSymbolConstants.ORIENTATION_LOCAL)
            {
                groutRotationAngle = groutRotationAngle + supportedObjectRotationAngle;
                groutComponentMatrix.Rotate(groutRotationAngle, vector);
            }

            //5. Set the length and width properties based on the sizing rule and angle
            //if the sizing rule is not user defined then take into account the sizing rule and angle
            if (groutSizingRule != SPSSymbolConstants.SIZING_RULE_USERDEFINED)
            {
                //member angle should not be considered while setting the length and width of grout
                groutRotationAngle = Math.Abs(groutRotationAngle - supportedObjectRotationAngle);
                FootingServices.SetLengthAndWidthPropertiesByRule(groutComponent, FootingComponentType.Grout, sectionDepth, sectionWidth, groutRotationAngle, true);
            }

            //6. Finally set the grout component matrix
            groutComponent.Matrix = groutComponentMatrix;
        }

        /// <summary>
        /// Evaluates the slab component.
        /// </summary>
        /// <param name="footing">The footing.</param>
        /// <param name="supportedObjects">The supported objects.</param>
        /// <param name="bottomPlaneMethod">The footing bottom plane method.</param>
        /// <param name="projectedPositionOnDatumPlane">The projected position on datum plane.</param>
        /// <param name="slabComponent">The slab component.</param>
        /// <param name="slabHeight">Height of the slab.</param>
        /// <param name="withGroutPad">if set to <c>true</c> [with grout pad].</param>
        private void EvaluateSlabComponent(Footing footing, Collection<BusinessObject> supportedObjects, FoundationBottomPlaneMethod bottomPlaneMethod, Position projectedPositionOnDatumPlane, FoundationComponent slabComponent, double slabHeight, bool withGroutPad)
        {
            double slabRotationAngle = 0.0;
            int slabOrientation = 0, slabSizingRule = 0, slabShape = 0;

            //1. Get component rotation and orientation to set the component matrix
            FootingServices.GetDimensions(slabComponent, FootingComponentType.Slab, out slabRotationAngle, out slabOrientation, out slabSizingRule, out slabShape);

            //2. Get the cumulative range of supported objects and its mid position
            RangeBox supportedObjectsRange;
            Position supportedObjectsMidPosition;
            base.GetSupportedObjectsRangeAndMidPosition(supportedObjects, slabRotationAngle, out supportedObjectsRange, out supportedObjectsMidPosition);

            //3. set the slab component origin
            //Construct an identity matrix
            Matrix4X4 slabComponentMatrix = new Matrix4X4();
            //need to get the elevation value for slab component in case None bottom plane method.
            //if the footing does not have a supporting plane as input, 
            //then the upgraded/synchronized footing lost the computed slab height or the slab height provided by the user.             
            if (bottomPlaneMethod == FoundationBottomPlaneMethod.None)
            {
                projectedPositionOnDatumPlane = new Position();
                //Get the first grout height
                double groutHeight = 0.0;
                if (this.groutAssemblyOutputs[0] != null)
                {
                    groutHeight = StructHelper.GetDoubleProperty(this.groutAssemblyOutputs[0], SPSSymbolConstants.IJUASPSFtgGroutPadDim, SPSSymbolConstants.GroutHeight);
                }

                //Get the first pier height
                double pierHeight = 0.0;
                if (this.pierAssemblyOutputs[0] != null)
                {
                    pierHeight = StructHelper.GetDoubleProperty(this.pierAssemblyOutputs[0], SPSSymbolConstants.IJUASPSFtgPierDim, SPSSymbolConstants.PierHeight);
                }

                projectedPositionOnDatumPlane.Z = base.GetElevationForSlab(supportedObjects, groutHeight, pierHeight) - slabHeight;
            }
            if (projectedPositionOnDatumPlane != null)
            {
                supportedObjectsMidPosition.Z = projectedPositionOnDatumPlane.Z + slabHeight;
            }
            slabComponentMatrix.Origin = supportedObjectsMidPosition;

            //4. Rotate the component based on the rotation angles of the component
            //rotation vector is along global Z direction
            //In case of slab component don’t consider the supported object rotation angle
            Vector vector = new Vector(0.0, 0.0, 1.0);
            slabComponentMatrix.Rotate(slabRotationAngle, vector);

            //5. Set the length and width properties based on the sizing rule and angle
            //if the sizing rule is not user defined then take into account the sizing rule and angle
            if (slabSizingRule != SPSSymbolConstants.SIZING_RULE_USERDEFINED)
            {
                //get all piers
                bool isMergePiers = StructHelper.GetBoolProperty(footing, SPSSymbolConstants.IJUACombinedPiers, SPSSymbolConstants.MergePiers);
                double maximumDistanceBetweenTwoSupportedObjects = StructHelper.GetDoubleProperty(footing, SPSSymbolConstants.IJUACombinedPiers, SPSSymbolConstants.MaxDistance);
                Dictionary<int, Collection<BusinessObject>> supportedObjectsConsideringMerge = base.GetSupportedObjectsForMerge(isMergePiers, maximumDistanceBetweenTwoSupportedObjects);
                double pierWidth, pierLength;
                this.GetMaximumPierDimensions(supportedObjectsConsideringMerge, supportedObjectsRange, slabRotationAngle, out pierWidth, out pierLength);

                FootingServices.SetLengthAndWidthPropertiesByRule(slabComponent, FootingComponentType.Slab, pierLength, pierWidth, slabRotationAngle, withGroutPad);
            }

            //6. Set origin value on the slab component centre coordinates
            slabComponent.SetPropertyValue(supportedObjectsMidPosition.X, SPSSymbolConstants.IJUASPSFootingCenter, SPSSymbolConstants.CenterX);
            slabComponent.SetPropertyValue(supportedObjectsMidPosition.Y, SPSSymbolConstants.IJUASPSFootingCenter, SPSSymbolConstants.CenterY);
            if (projectedPositionOnDatumPlane != null)
            {
                slabComponent.SetPropertyValue(projectedPositionOnDatumPlane.Z, SPSSymbolConstants.IJUASPSFootingCenter, SPSSymbolConstants.CenterZ);
            }

            //7. Finally set the slab component matrix
            slabComponent.Matrix = slabComponentMatrix;
        }

        /// <summary>
        /// Evaluates the pier component.
        /// </summary>
        /// <param name="footing">The footing.</param>
        /// <param name="indexOfPier">Index of the pier which is evaluated.</param>
        /// <param name="supportedObjectsPerPier">The supported objects.</param>
        /// <param name="withGroutPad">if set to <c>true</c> [with grout pad].</param>
        /// <param name="slabHeight">Height of the slab.</param>     
        /// <param name="bottomPlaneMethod">The footing bottom plane method.</param>
        private void EvaluatePierComponent(Footing footing, int indexOfPier, Collection<BusinessObject> supportedObjectsPerPier, bool isPlacedByPoint, bool withGroutPad, double slabHeight, FoundationBottomPlaneMethod bottomPlaneMethod)
        {
            FoundationComponent pierComponent = (FoundationComponent)this.pierAssemblyOutputs[indexOfPier];

            double pierRotationAngle = 0.0;
            int pierOrientation = 0, pierSizingRule = 0, pierShape = 0, groutIndex = 0;
            //Get the translation part from footing but not the rotation part TR-231229
            Matrix4X4 componentMatrix = new Matrix4X4();

            // Get component rotation and orientation to set the component matrix
            FootingServices.GetDimensions(pierComponent, FootingComponentType.Pier, out pierRotationAngle, out pierOrientation, out pierSizingRule, out pierShape);

            //Get the cumulative range of supported objects and its mid position
            RangeBox range;
            Position midPosition;
            base.GetSupportedObjectsRangeAndMidPosition(supportedObjectsPerPier, pierRotationAngle, out range, out midPosition);

            //Determine if the piers are merged
            bool isMergePiers = supportedObjectsPerPier.Count == 1 ? false : true;

            double supportedObjectRotationAngle = 0.0;

            //the lowest member position is same for all supported objects in the collection. 
            //for merged case all members are on same plane and non-merged case the collection contains only one
            Position lowestSupportedObjectPosition = base.GetLowestSupportedObjectPosition(supportedObjectsPerPier, out groutIndex);
            //in case of merged piers do not consider the member angle
            if (!isMergePiers)
            {
                if (supportedObjectsPerPier[0] is MemberSystem)
                {
                    base.GetBottomPositionFromMember((MemberSystem)supportedObjectsPerPier[0], out supportedObjectRotationAngle);
                }
                groutIndex = indexOfPier;//for non merged case grout- pier will have one to one mapping. use same pier index
            }

            double groutHeight = 0.0, groutRotationAngle = 0.0;
            int groutOrientation = -1, groutSizingRule = -1, groutShape = -1;
            //If grout is needed get the grout height            
            if (withGroutPad && this.groutAssemblyOutputs[groutIndex] != null)
            {
                BusinessObject groutComponent = this.groutAssemblyOutputs[groutIndex];
                groutHeight = StructHelper.GetDoubleProperty(groutComponent, SPSSymbolConstants.IJUASPSFtgGroutPadDim, SPSSymbolConstants.GroutHeight);
                // Get grout component inputs
                FootingServices.GetDimensions((FoundationBase)groutComponent, FootingComponentType.Grout, out groutRotationAngle, out groutOrientation, out groutSizingRule, out groutShape);
            }
            else//use member angle directly
            {
                groutRotationAngle = -supportedObjectRotationAngle;
            }

            //Adjusting the slab component matrix origin to accommodate the grout height
            Position origin = new Position(midPosition); //in case of point supported object we will get read-only position so this value needs to be created using New
            origin.Z = origin.Z - groutHeight;
            componentMatrix.Origin = origin;

            //rotation vector is along global Z direction
            Vector vector = new Vector(0.0, 0.0, 1.0);

            // Rotate the component based on the rotation angles of the component and member angle
            if (pierOrientation == SPSSymbolConstants.ORIENTATION_GLOBAL)
            {
                componentMatrix.Rotate(pierRotationAngle, vector);
            }
            else if (pierOrientation == SPSSymbolConstants.ORIENTATION_LOCAL)
            {
                pierRotationAngle = (isMergePiers) ? pierRotationAngle : (pierRotationAngle + supportedObjectRotationAngle);
                componentMatrix.Rotate(pierRotationAngle, vector);
            }

            //If the footing is placed without grout pad, skip setting of length and width
            if (withGroutPad)
            {
                double componentRotationAngle = 0.0;
                // Set the length and width properties based on the sizing rule and angle
                //if the sizing rule is not user defined then take into account the sizing rule and angle
                if (pierSizingRule != SPSSymbolConstants.SIZING_RULE_USERDEFINED)
                {
                    double footingOrientationAngle = (isPlacedByPoint) ? base.GetPlanAngleFromMatrix(footing.Matrix) : 0.0;

                    double groutLengthAtXLow, groutLengthAtXHigh, groutWidthAtYLow, groutWidthAtYHigh;
                    this.GetMaximumGroutDimensions(pierComponent, supportedObjectsPerPier, range, withGroutPad, footingOrientationAngle, pierRotationAngle, out groutWidthAtYLow, out groutWidthAtYHigh, out groutLengthAtXLow, out groutLengthAtXHigh);

                    double groutWidth, groutLength;
                    if (isMergePiers)
                    {
                        //pier length and pier width calculated considering orientation angle of pier component, for circular shape no need to consider pier orientation angle
                        componentRotationAngle = (pierShape == SPSSymbolConstants.SHAPE_CIRCULAR) ? 0.0 : Math.Abs(pierRotationAngle);

                        if (componentRotationAngle > Math.PI / 2)
                        {
                            componentRotationAngle = Math.Abs(Math.PI - componentRotationAngle);
                        }

                        //Consider the range too in case of merged piers                        
                        double yLow = (range.Low.Y - groutWidthAtYLow / 2);
                        double yHigh = (range.High.Y + groutWidthAtYHigh / 2);
                        double xLow = (range.Low.X - groutLengthAtXLow / 2);
                        double xHigh = (range.High.X + groutLengthAtXHigh / 2);
                        groutLength = (xHigh - xLow) * Math.Cos(componentRotationAngle);
                        groutWidth = (yHigh - yLow) * Math.Cos(componentRotationAngle);

                        //in case of non-rotated shift the pier origin
                        if (StructHelper.AreEqual(componentRotationAngle, 0.0, Math3d.DistanceTolerance))
                        {
                            //update the merge pier origin
                            origin.X = (xLow + xHigh) / 2.0;
                            origin.Y = (yLow + yHigh) / 2.0;
                            componentMatrix.Origin = origin;
                        }
                    }
                    else
                    {
                        //footing placed by member with grout pad then, add member angle to the grout rotation angle if grout orientation is local to the member
                        if (withGroutPad && groutOrientation == SPSSymbolConstants.ORIENTATION_LOCAL)
                        {
                            groutRotationAngle = groutRotationAngle + supportedObjectRotationAngle;
                        }
                        if (pierShape == SPSSymbolConstants.SHAPE_CIRCULAR)
                        {
                            pierRotationAngle = 0;
                        }
                        if (groutShape == SPSSymbolConstants.SHAPE_CIRCULAR)
                        {
                            groutRotationAngle = 0;
                        }

                        componentRotationAngle = Math.Abs(pierRotationAngle - groutRotationAngle);
                        if (componentRotationAngle > Math.PI / 2)
                        {
                            componentRotationAngle = Math.Abs(Math.PI - componentRotationAngle);
                        }

                        //in case of single pier, there will be one to one relation between grout and pier
                        groutLength = groutLengthAtXLow;
                        groutWidth = groutWidthAtYLow;
                    }

                    FootingServices.SetLengthAndWidthPropertiesByRule(pierComponent, FootingComponentType.Pier, groutLength, groutWidth, componentRotationAngle, withGroutPad);
                }

                //in case the footing bottom plane method either Associative or ElevationValue then gets the position on the datum plane by projecting the lowest supported object position.
                if (bottomPlaneMethod != FoundationBottomPlaneMethod.None)
                {
                    Position projectedPositionOnDatumPlane = base.GetProjectedPositionOnDatumPlane(lowestSupportedObjectPosition);
                    double projectedPositionElevation = 0.0;
                    if (projectedPositionOnDatumPlane != null)
                    {
                        projectedPositionElevation = projectedPositionOnDatumPlane.Z;
                    }
                    //the projected position elevation must be less than the elevation of the lowest supported object
                    if (projectedPositionElevation > lowestSupportedObjectPosition.Z)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.FootingToDoMessageCodelist,
                            SPSSymbolConstants.TDL_FTG_BOTTOMPLANE_NOTVALID_CANNOTCOMPUTE, FootingLocalizer.GetString(FootingResourceIDs.ErrInvalidSupportingPlane,
                            "Footing support plane is not valid as projected position elevation is greater than lowest supported object position elevation in CombinedPierSlab Footing custom assembly definition. Software cannot compute the pier height. Delete and replace footing."));
                        return;
                    }

                    double bottomElevationDatum = (double)footing.BottomElevationDatum;
                    //the elevation of the lowest supported object must be greater than bottom elevation datum value
                    if (lowestSupportedObjectPosition.Z > bottomElevationDatum)
                    {
                        //calculate the pier height and set it
                        double pierHeight = Math.Abs(lowestSupportedObjectPosition.Z - projectedPositionElevation) - groutHeight - slabHeight;
                        if (!isMergePiers && pierHeight < 0.0)
                        {
                            base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.FootingToDoMessageCodelist,
                                SPSSymbolConstants.TDL_FTG_BOTTOMPLANE_NOTVALID, FootingLocalizer.GetString(FootingResourceIDs.ErrInvalidPierHeightNegativeValue,
                                "Footing support plane is not valid for some supported members. Footing pier height is an invalid negative value while evaluating pier component in CombinedPierSlab Footing custom assembly definition. Please check custom code or contact S3D support."));
                            return;
                        }

                        pierComponent.SetPropertyValue(pierHeight, SPSSymbolConstants.IJUASPSFtgPierDim, SPSSymbolConstants.PierHeight);
                    }
                    else
                    {
                        if (isMergePiers)
                        {
                            base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.FootingToDoMessageCodelist,
                                SPSSymbolConstants.TDL_FTG_BOTTOMPLANE_NOTVALID_CANNOTCOMPUTE, FootingLocalizer.GetString(FootingResourceIDs.ErrInvalidSupportingPlane,
                                "Footing support plane is not valid. Software cannot compute the pier height as its MergePiers and elevation of lowest supported object is less than bottom elevation datum value. Delete and replace footing."));
                            return;
                        }
                        else
                        {
                            base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SPSSymbolConstants.FootingToDoMessageCodelist,
                                SPSSymbolConstants.TDL_FTG_BOTTOMPLANE_NOTVALID, FootingLocalizer.GetString(FootingResourceIDs.ErrInvalidPierHeightNegativeValue,
                                "Footing support plane is not valid for some supported members. Footing pier height is an invalid negative value while evaluating pier component in CombinedPierSlab Footing custom assembly definition. Please check custom code or contact S3D support."));
                            return;
                        }
                    }
                }
            }

            // Setting the footing component matrix
            pierComponent.Matrix = componentMatrix;
        }

        /// <summary>
        /// Evaluates the ControlPoints.
        /// 1. Position the ControlPoints at the supported objects location.
        /// 2. Position the KeyPoints at the 8 vertices of the slab component.
        /// 3. Position the KeyPoints at the 4 vertices (top side) of the pier component.
        /// </summary>
        /// <param name="supportedObjects">The supported objects.</param>
        /// <param name="isRectangularSlab">if set to <c>true</c> [is rectangular slab].</param>
        /// <param name="slabComponent">The slab component.</param>
        /// <param name="slabHeight">Height of the slab component.</param>
        /// <param name="numberOfPierVertices">The number of pier vertices.</param>
        /// <param name="numberOfPierAndSlabVertices">The number of pier and slab vertices.</param>
        private void EvaluateControlPoints(Collection<BusinessObject> supportedObjects, bool isRectangularSlab, FoundationComponent slabComponent, double slabHeight, int numberOfPierVertices, int numberOfPierAndSlabVertices)
        {
            for (int i = 1; i <= this.controlPointAssemblyOutputs.Count; i++)
            {
                if (i > numberOfPierAndSlabVertices)
                {
                    //Control points at the supported objects location
                    BusinessObject supportedObject = supportedObjects[(i - 1) - numberOfPierAndSlabVertices];
                    base.SetControlPointLocationAtSupportedObject(this.controlPointAssemblyOutputs, supportedObject, i);
                }
                else
                {
                    //Evaluates the ControlPoints at slab vertices.
                    //Position the KeyPoints at the 8 vertices of the slab component.
                    if (i <= numberOfPierAndSlabVertices - numberOfPierVertices)
                    {
                        Matrix4X4 slabComponentMatrix = slabComponent.Matrix;
                        if (isRectangularSlab)
                        {
                            //in case of rectangular slab we will create Key points at top and bottom level 4 vertices of the rectangular slab (i.e. 8 vertices for rectangular slab)
                            double slabLength = StructHelper.GetDoubleProperty(slabComponent, SPSSymbolConstants.IJUASPSFtgSlabDim, SPSSymbolConstants.SlabLength);
                            double slabWidth = StructHelper.GetDoubleProperty(slabComponent, SPSSymbolConstants.IJUASPSFtgSlabDim, SPSSymbolConstants.SlabWidth);
                            base.SetControlPointLocationAtRectangularSolidVertices(this.controlPointAssemblyOutputs, slabComponentMatrix, slabLength, slabWidth, slabHeight, ref i);
                        }
                        else
                        {
                            //in case of octagonal slab we will create Key points at top and bottom level 8 vertices of the octagonal slab (i.e. 16 vertices for octagonal slab)
                            double overallDimension = StructHelper.GetDoubleProperty(slabComponent, SPSSymbolConstants.IJUAOctagonalSlabDim, SPSSymbolConstants.OctOverallDim);
                            base.SetControlPointLocationAtOctagonalSolidVertices(this.controlPointAssemblyOutputs, slabComponentMatrix, overallDimension, slabHeight, ref i);
                        }
                    }
                    else
                    {
                        //Evaluates the ControlPoints at pier top plane vertices.
                        //Position the KeyPoints at the 4 top plane vertices of the pier component.
                        //in case of pier we will create Key points at top level 4 vertices of the pier                        
                        int startIndex = numberOfPierAndSlabVertices - numberOfPierVertices;
                        double tmp = (i - startIndex) / 4.0;
                        int pierIndex = (int)Math.Round(tmp);
                        if (pierIndex >= tmp)
                        {
                            pierIndex = pierIndex - 1;
                        }
                        pierIndex = pierIndex + 1;

                        FoundationComponent pierComponent = (FoundationComponent)this.pierAssemblyOutputs[pierIndex - 1];
                        Matrix4X4 pierComponentMatrix = pierComponent.Matrix;
                        double pierLength = StructHelper.GetDoubleProperty(pierComponent, SPSSymbolConstants.IJUASPSFtgPierDim, SPSSymbolConstants.PierLength);
                        double pierWidth = StructHelper.GetDoubleProperty(pierComponent, SPSSymbolConstants.IJUASPSFtgPierDim, SPSSymbolConstants.PierWidth);
                        base.SetControlPointLocationAtRectangularPlaneVertices(this.controlPointAssemblyOutputs, pierComponentMatrix, pierLength, pierWidth, ref i);
                    }
                }
            }
        }

        /// <summary>
        /// Gets the pier dimension which is having maximum dimension located at the extreme edges of the range box of the supported objects.
        /// </summary>
        /// <param name="supportedObjectsConsideringMerge">The supported objects considering merge.</param>
        /// <param name="supportedObjectsRange">The supported objects range.</param>
        /// <param name="slabRotationAngle">The slab rotation angle.</param>
        /// <param name="pierMaximumWidth">Width of the pier.</param>
        /// <param name="pierMaximumLength">Length of the pier.</param>
        private void GetMaximumPierDimensions(Dictionary<int, Collection<BusinessObject>> supportedObjectsConsideringMerge, RangeBox supportedObjectsRange, double slabRotationAngle, out double pierMaximumWidth, out double pierMaximumLength)
        {
            pierMaximumWidth = 0.0; pierMaximumLength = 0.0;
            double tempPierMaximumWidth = 0.0, tempPierMaximumLength = 0.0;

            //process each pier to get maximum pier dimensions
            for (int i = 0; i < supportedObjectsConsideringMerge.Count; i++)
            {
                //get the each pier supported object and process each member
                Collection<BusinessObject> supportedObjectsPerPier = supportedObjectsConsideringMerge[i];

                foreach (BusinessObject supportedObjectPerPier in supportedObjectsPerPier)
                {
                    Position lowPosition; double memberAngle = 0.0;
                    if (supportedObjectPerPier is MemberSystem)
                    {
                        lowPosition = base.GetBottomPositionFromMember((MemberSystem)supportedObjectPerPier, out memberAngle);
                    }
                    else
                    {
                        lowPosition = ((Point3d)supportedObjectPerPier).Position;
                    }

                    lowPosition = base.GetTransformedPositionByAngle(lowPosition, slabRotationAngle);

                    //if the member position is at the range border, then re-evaluate pier dimensions
                    if (Math.Abs(lowPosition.X - supportedObjectsRange.Low.X) < Math3d.FitTolerance || Math.Abs(lowPosition.X - supportedObjectsRange.High.X) < Math3d.FitTolerance ||
                        Math.Abs(lowPosition.Y - supportedObjectsRange.Low.Y) < Math3d.FitTolerance || Math.Abs(lowPosition.Y - supportedObjectsRange.High.Y) < Math3d.FitTolerance)
                    {
                        double pierLength = 0.0, pierWidth = 0.0, pierHeight = 0.0, pierEdgeClearance = 0.0;
                        int pierShape = -1;
                        FoundationComponent pierComponent = (FoundationComponent)this.pierAssemblyOutputs[i];
                        FootingServices.GetDimensions(pierComponent, FootingComponentType.Pier, ref pierLength, ref pierWidth, ref pierHeight, ref pierEdgeClearance, ref pierShape);
                        double pierRotationAngle = StructHelper.GetDoubleProperty(pierComponent, SPSSymbolConstants.IJUASPSFtgPier, SPSSymbolConstants.PierRotationAngle);
                        int pierOrientation = StructHelper.GetIntProperty(pierComponent, SPSSymbolConstants.IJUASPSFtgPier, SPSSymbolConstants.PierOrientation);

                        if (supportedObjectsPerPier.Count > 1)//it is a merged pier, so calculate appropriate pier size
                        {
                            RangeBox pierRange = null;
                            Position pierMidPosition = null;
                            base.GetSupportedObjectsRangeAndMidPosition(supportedObjectsPerPier, pierRotationAngle, out pierRange, out pierMidPosition);
                            pierLength = pierLength - (pierRange.High.X - pierRange.Low.X);
                            pierWidth = pierWidth - (pierRange.High.Y - pierRange.Low.Y);
                        }

                        //Evaluate the pier rotation angle using slab and member angle
                        pierRotationAngle = FootingServices.GetPierRotationAngle(pierComponent, slabRotationAngle, memberAngle);

                        //This column is governing length of range box
                        if (Math.Abs(lowPosition.X - supportedObjectsRange.Low.X) < Math3d.FitTolerance || Math.Abs(lowPosition.X - supportedObjectsRange.High.X) < Math3d.FitTolerance)
                        {
                            tempPierMaximumLength = pierLength * Math.Cos(pierRotationAngle) + pierWidth * Math.Sin(pierRotationAngle);
                        }

                        //This column is governing width of range box
                        if (Math.Abs(lowPosition.Y - supportedObjectsRange.Low.Y) < Math3d.FitTolerance || Math.Abs(lowPosition.Y - supportedObjectsRange.High.Y) < Math3d.FitTolerance)
                        {
                            tempPierMaximumWidth = pierWidth * Math.Cos(pierRotationAngle) + pierLength * Math.Sin(pierRotationAngle);
                        }

                        // reset pier dimensions with maximum values
                        if (pierMaximumWidth < tempPierMaximumWidth)
                        {
                            pierMaximumWidth = tempPierMaximumWidth;
                        }
                        if (pierMaximumLength < tempPierMaximumLength)
                        {
                            pierMaximumLength = tempPierMaximumLength;
                        }
                        break;//no need to process further for this pier component
                    }
                }
            }

            pierMaximumWidth = (supportedObjectsRange.High.Y - supportedObjectsRange.Low.Y) + pierMaximumWidth;
            pierMaximumLength = (supportedObjectsRange.High.X - supportedObjectsRange.Low.X) + pierMaximumLength;
        }

        /// <summary>
        /// Gets the grout dimension which is having maximum dimension located at the extreme edges of the range box of the supported objects.
        /// </summary>
        /// <param name="pierComponent">The pier component.</param>
        /// <param name="supportedObjectsPerPier">The supported objects considering merge.</param>
        /// <param name="supportedObjectsRange">The supported objects range.</param>
        /// <param name="withGroutPad">Flag indicating if the footing is with GroutPad or not.</param>
        /// <param name="footingOrientationAngle">Orientation angle of the footing w.r.t X axis when rotated about Z-axis.</param>
        /// <param name="pierRotationAngle">The pier rotation angle.</param>
        /// <param name="groutWidthAtYLow">The grout width at Y-axis low.</param>
        /// <param name="groutWidthAtYHigh">The grout width at Y-axis high.</param>
        /// <param name="groutLengthAtXLow">The grout length at X-axis low.</param>
        /// <param name="groutLengthAtXHigh">The grout length at X-axis high.</param>
        private void GetMaximumGroutDimensions(FoundationComponent pierComponent, Collection<BusinessObject> supportedObjectsPerPier, RangeBox supportedObjectsRange, bool withGroutPad, double footingOrientationAngle, double pierRotationAngle, out double groutWidthAtYLow, out double groutWidthAtYHigh, out double groutLengthAtXLow, out double groutLengthAtXHigh)
        {
            groutWidthAtYLow = 0.0; groutWidthAtYHigh = 0.0; groutLengthAtXLow = 0.0; groutLengthAtXHigh = 0.0;
            double tempGroutWidthAtYLow = 0.0, tempGroutWidthAtYHigh = 0.0, tempGroutLengthAtXLow = 0.0, tempGroutLengthAtXHigh = 0.0;

            //supported objects per pier count will be more than one only if merge criteria satisfied and IsMergePiers flag is true
            bool isMergePiers = (supportedObjectsPerPier.Count > 1) ? true : false;

            //process each grout to get cumulative dimensions
            for (int i = 0; i < supportedObjectsPerPier.Count; i++)
            {
                BusinessObject supportedObject = supportedObjectsPerPier[i];
                Position lowPosition;
                if (supportedObject is MemberSystem)
                {
                    lowPosition = base.GetBottomPositionFromMember((MemberSystem)supportedObject, out footingOrientationAngle);
                }
                else
                {
                    lowPosition = ((Point3d)supportedObject).Position;
                }

                if (isMergePiers)
                {
                    lowPosition = base.GetTransformedPositionByAngle(lowPosition, pierRotationAngle);
                }

                double groutLength = 0.0, groutWidth = 0.0, groutRotationAngle = 0.0;
                if (withGroutPad)
                {
                    //get the grout component associated with the given supported object
                    //get the index of the given supported object in the supported objects collection
                    Footing footing = (Footing)base.Occurrence;
                    int index = footing.SupportedObjects.IndexOf(supportedObject);

                    //use this index to get the associated grout component from the grout assembly outputs
                    FoundationComponent groutComponent = (FoundationComponent)this.groutAssemblyOutputs[index];

                    //get the shape and size properties from the component
                    double groutHeight = 0.0, groutEdgeClearance = 0.0;
                    int groutShape = -1;
                    FootingServices.GetDimensions(groutComponent, FootingComponentType.Grout, ref groutLength, ref groutWidth, ref groutHeight, ref groutEdgeClearance, ref groutShape);
                    groutRotationAngle = StructHelper.GetDoubleProperty(groutComponent, SPSSymbolConstants.IJUASPSFtgGroutPad, SPSSymbolConstants.GroutRotationAngle);
                    int groutOrientation = StructHelper.GetIntProperty(groutComponent, SPSSymbolConstants.IJUASPSFtgGroutPad, SPSSymbolConstants.GroutOrientation);

                    if (groutOrientation == SPSSymbolConstants.ORIENTATION_LOCAL)//add the member angle if the orientation is local
                    {
                        groutRotationAngle = groutRotationAngle + footingOrientationAngle;
                    }

                    if (groutShape == SPSSymbolConstants.SHAPE_CIRCULAR)//if grout is circular then no need to consider its rotation.
                    {
                        groutRotationAngle = 0.0;
                    }
                }
                else
                {
                    if (supportedObject is MemberSystem)
                    {
                        //get cross-section dimensions of the member part
                        MemberPart memberPart = base.GetBottomMemberPart((MemberSystem)supportedObject);
                        double depth = memberPart.CrossSection.Depth;
                        double width = memberPart.CrossSection.Width;

                        if (depth > width)
                        {
                            groutLength = depth;
                            groutWidth = depth;
                        }
                        else
                        {
                            groutLength = width;
                            groutWidth = width;
                        }
                    }
                    groutRotationAngle = -footingOrientationAngle;
                }

                groutRotationAngle = (isMergePiers) ? Math.Abs(Math.Abs(groutRotationAngle) - pierRotationAngle) : Math.Abs(groutRotationAngle - pierRotationAngle);
                if (groutRotationAngle > Math.PI / 2)
                {
                    groutRotationAngle = Math.Abs(Math.PI - groutRotationAngle);
                }

                if (isMergePiers)
                {
                    //      +--------+
                    //      |---+----|
                    //      |------+-|
                    //      +--------+
                    //the below logic picks one of the corner grouts (maximum amongst) dimensions
                    //consider the dimension of grout which is at border of the range box and has maximum dimension

                    //This column is governing length of range box
                    if (Math.Abs(lowPosition.X - supportedObjectsRange.Low.X) < Math3d.FitTolerance)
                    {
                        tempGroutLengthAtXLow = groutLength * Math.Cos(groutRotationAngle) + groutWidth * Math.Sin(groutRotationAngle);
                    }

                    if (Math.Abs(lowPosition.X - supportedObjectsRange.High.X) < Math3d.FitTolerance)
                    {
                        tempGroutLengthAtXHigh = groutLength * Math.Cos(groutRotationAngle) + groutWidth * Math.Sin(groutRotationAngle);
                    }

                    //This column is governing width of range box
                    if (Math.Abs(lowPosition.Y - supportedObjectsRange.Low.Y) < Math3d.FitTolerance)
                    {
                        tempGroutWidthAtYLow = groutWidth * Math.Cos(groutRotationAngle) + groutLength * Math.Sin(groutRotationAngle);
                    }

                    if (Math.Abs(lowPosition.Y - supportedObjectsRange.High.Y) < Math3d.FitTolerance)
                    {
                        tempGroutWidthAtYHigh = groutWidth * Math.Cos(groutRotationAngle) + groutLength * Math.Sin(groutRotationAngle);
                    }
                }
                else
                {
                    tempGroutWidthAtYLow = groutWidth;
                    tempGroutLengthAtXLow = groutLength;
                }

                // reset grout dimensions with maximum values
                if (groutWidthAtYLow < tempGroutWidthAtYLow)
                {
                    groutWidthAtYLow = tempGroutWidthAtYLow;
                }
                if (groutWidthAtYHigh < tempGroutWidthAtYHigh)
                {
                    groutWidthAtYHigh = tempGroutWidthAtYHigh;
                }
                if (groutLengthAtXLow < tempGroutLengthAtXLow)
                {
                    groutLengthAtXLow = tempGroutLengthAtXLow;
                }
                if (groutLengthAtXHigh < tempGroutLengthAtXHigh)
                {
                    groutLengthAtXHigh = tempGroutLengthAtXHigh;
                }
            }
        }

        #endregion Private methods
    }
}