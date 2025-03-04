//-----------------------------------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  CopeDefinition.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\SmartPlantStructure\Symbols\Release\SPSFeatureMacros.dll
//  Original Class Name: ‘CopeDef’ in VB content
//
//Abstract:
//  CopeDefinition is a .NET custom assembly definition which creates cutback surface and cope on MemberPart after placing the fitted assembly connection. 
//  This class subclasses from FeatureCustomAssemblyDefinition.
//-----------------------------------------------------------------------------------------------------------------------------------------------------------
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Implementation of CopeDefinition .NET custom assembly definition class.
    /// CopeDefinition is a .NET custom assembly definition which creates planar cutback and copes on bounded member part at given position with respect to the bounding member part.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    [DefaultLocalizer(FeaturesResourceIds.DEFAULT_RESOURCE, FeaturesResourceIds.DEFAULT_ASSEMBLY)]
    [OutputNotification(StructureCustomAssembliesConstants.IID_IJUASPSCope)]
    public class CopeDefinition : FeatureCustomAssemblyDefinition
    {
        //===================================================================================================
        //DefinitionName/ProgID of this symbol is "Features,Ingr.SP3D.Content.Structure.CopeDefinition"
        //===================================================================================================
        #region Private members
        Feature feature = null;
        bool isOnPreload = false;
        #endregion

        #region Definitions of assembly outputs
        // Setting up output notification so that we can mark these outputs as modified on these interfaces
        // Cutback assembly output
        [OutputNotification(SPSSymbolConstants.IID_IJSurface)]
        [AssemblyOutput(1, StructureCustomAssembliesConstants.CopeCutback)]
        public AssemblyOutput cutBack;
        // Cope1 assembly output
        [OutputNotification(SPSSymbolConstants.IID_IJDModelBody)]
        [AssemblyOutput(2, StructureCustomAssembliesConstants.Cope1)]
        public AssemblyOutput cope1;
        // Cope2 assembly output
        [OutputNotification(SPSSymbolConstants.IID_IJDModelBody)]
        [AssemblyOutput(3, StructureCustomAssembliesConstants.Cope2)]
        public AssemblyOutput cope2;
        #endregion

        #region Public Override methods

        /// <summary>
        /// Construct and re-evaluate the cope feature custom assembly outputs: CopeCutback, Cope1 and Cope2.
        /// </summary>
        public override void EvaluateAssembly()
        {
            try
            {
                // Get the feature.
                this.feature = (Feature)base.Occurrence;

                // Get Top Flange Clearance 
                double sideFlangeClearance = StructHelper.GetDoubleProperty(this.feature, StructureCustomAssembliesConstants.IJUASPSCope, StructureCustomAssembliesConstants.SideFlangeClearance);
                double insideFlangeClearance = StructHelper.GetDoubleProperty(this.feature, StructureCustomAssembliesConstants.IJUASPSCope, StructureCustomAssembliesConstants.InsideFlangeClearance);

                // Get Web Clearance 
                double webClearance = StructHelper.GetDoubleProperty(this.feature, StructureCustomAssembliesConstants.IJUASPSCope, StructureCustomAssembliesConstants.WebClearance);
                double increment = StructHelper.GetDoubleProperty(this.feature, StructureCustomAssembliesConstants.IJUASPSCope, StructureCustomAssembliesConstants.Increment);
                bool isPlanar = StructHelper.GetBoolProperty(this.feature, StructureCustomAssembliesConstants.IJUASPSCope, StructureCustomAssembliesConstants.AlwaysPlanar);
                bool isSquareEnd = StructHelper.GetBoolProperty(this.feature, StructureCustomAssembliesConstants.IJUASPSCope, StructureCustomAssembliesConstants.SquaredEnd);
                // Check if cutback is required.
                bool isCutbackRequired = IsCutbackRequired(isPlanar);
                bool isTopCopeRequired = IsCopeRequired(CopePosition.WebTop);
                bool isBottomCopeRequired = IsCopeRequired(CopePosition.WebBottom);
                double radius = 0.0, webRadius = 0.0;
                double length = 0.0, depth = 0.0;
                CopeRadiusType radiusType = CopeRadiusType.None;
                CopeRadiusType webRadiusType = CopeRadiusType.None;
                MemberAxisEnd memberAxisEnd = MemberAxisEnd.Start;
                // Set the CopeProperties on Base Class.
                CopeProperties copeProperties = new CopeProperties(sideFlangeClearance, insideFlangeClearance, webClearance, radius, radiusType, webRadius, webRadiusType,
                                                                   memberAxisEnd, length, depth, increment);

                // Check if cutback already exists. If yes, modify it; else, create new one.
                // 
                TopologySurface cutbackOutput = null;
                if (this.cutBack.Output != null)
                {
                    if (isCutbackRequired)
                    {
                        // Modify the cutback surface by constructing a new Plane3d object from using CopeProperties and other information
                        // such as whether SquareEnd and/or Planar cutback is to be created.
                        // This new Plane3d will be used by ModifyCutback method.
                        Plane3d newPlane3d = CreateCutbackPlaneWithOffset(copeProperties, copeProperties.WebClearance, isSquareEnd, isPlanar);
                        
                        cutbackOutput = this.cutBack.Output as TopologySurface;
                        if(cutbackOutput == null)
                        {
                            // This means that the current cutback output is a Plane3d.
                            // Create a TopologySurface object from the cutback output and later pass it to ModufyCutback. 
                            try
                            {
                                ISurface surface = this.cutBack.Output as ISurface;
                                Collection<ISurface> surfaces = new Collection<ISurface> { surface };
                                cutbackOutput = new TopologySurface(surfaces);
                            }
                            catch(Exception)
                            {
                                //may be there could be specific ToDo record is created before, so do not override it.
                                if (base.ToDoListMessage == null)
                                {
                                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(FeaturesResourceIds.ErrInvalidCutbackSurfaceType,
                                        "The assembly output used as cutback surface must be ISurface."), this.ToString()));
                                }
                                // Stop the evaluation.
                                return;
                            }
                        }

                        // Modify the cutback with new Plane3d.
                        base.ModifyCutback(cutbackOutput, newPlane3d, copeProperties, isSquareEnd, isPlanar);
                    }
                    else
                    {
                        cutBack.Delete();
                    }
                }
                else
                {
                    // First check if cutback is required.
                    if (isCutbackRequired)
                    {
                        // Create the cutback.
                        cutbackOutput = base.CreateCutback(copeProperties, copeProperties.WebClearance, isSquareEnd, isPlanar);
                        cutBack.Output = cutbackOutput;
                    }
                }

                // If AlwaysPlanar is True, no need to create cope (as per conditional cope).
                if (!isPlanar)
                {
                    // The cope cutter created in pre v7 model may have WireBody set as cutter. Since we allow only SolidBody as cutter, on MDR of such model, we will delete
                    // the existing cutter in such case and re-create the cutter (SolidBody). 
                    // Check if the cutter is a WireBody

                    bool isCopeCutterATopologyCurve = false;
                    TopologyCurve topologyCurveCopeOutput = this.cope1.Output as TopologyCurve;
                    if(topologyCurveCopeOutput != null)
                    {
                        isCopeCutterATopologyCurve = true;
                    }

                    // If required, create cope at WebTop position. If it already exists, then modify the cutter.
                    // Do not modify the cutter if it is a TopologyCurve. 
                    if (this.HasCope(CopePosition.WebTop) && isCopeCutterATopologyCurve == false)
                    {
                        if (isTopCopeRequired)
                        {
                            // Cope1 already exists. Modify the cope. First, Get the cope contour.
                            TopologyCurve topologyCurve = CreateCopeCurve(CopePosition.WebTop, copeProperties, cutbackOutput, isPlanar, isSquareEnd);
                            if (topologyCurve != null)
                            {
                                base.ModifyCutter(CopePosition.WebTop, topologyCurve, (TopologySolid)this.cope1.Output, isSquareEnd);
                            }
                            SetCopeLengthAndDepth(copeProperties);
                        }
                        else
                        {
                            this.cope1.Delete();
                        }

                    }
                    else
                    {
                        if (isTopCopeRequired)
                        {
                            TopologyCurve topologyCurve = CreateCopeCurve(CopePosition.WebTop, copeProperties, cutbackOutput, isPlanar, isSquareEnd);
                            if (topologyCurve != null)
                            {
                                TopologySolid copeTopologySolid = base.CreateCutter(CopePosition.WebTop, topologyCurve, null, isSquareEnd);
                                if (copeTopologySolid != null)
                                {
                                    // When model prior to V7 model is migrated, some cases may have TopologyCurve as cope cutter output. 
                                    // If TopologySolid is created then only delete the existing output.
                                    if (isCopeCutterATopologyCurve)
                                    {
                                        // Replace the old output with new one.
                                        base.ReplaceOutput(this.cope1, copeTopologySolid);
                                    }
                                    else
                                    {
                                        // Set the cope output.
                                        this.cope1.Output = copeTopologySolid;
                                    }
                                }
                            }
                            SetCopeLengthAndDepth(copeProperties);
                        }
                    }

                    // The cope cutter created in pre v7 model may have WireBody set as cutter. So on MDR of such model or update of feature in that migrated model, since we allow only SolidBody as cutter, we will delete
                    // the existing cutter in such case and re-create the cutter (SolidBody). 
                    // Check if the cutter is a WireBody
                    isCopeCutterATopologyCurve = false;
                    topologyCurveCopeOutput = this.cope2.Output as TopologyCurve;
                    if (topologyCurveCopeOutput != null)
                    {
                        isCopeCutterATopologyCurve = true;
                    }

                    // If required, create cope at WebBottom position. If it already exists, then modify the cutter.
                    // Do not modify the cutter if it is a TopologyCurve.
                    if (this.HasCope(CopePosition.WebBottom) && isCopeCutterATopologyCurve == false)
                    {
                        if (isBottomCopeRequired)
                        {
                            // Cope2 already exists. Modify the cope.
                            TopologyCurve topologyCurve = CreateCopeCurve(CopePosition.WebBottom, copeProperties, cutbackOutput, isPlanar, isSquareEnd);
                            if (topologyCurve != null)
                            {
                                base.ModifyCutter(CopePosition.WebBottom, topologyCurve, (TopologySolid)this.cope2.Output, isSquareEnd);
                            }
                        }
                        else
                        {
                            this.cope2.Delete();
                        }
                    }
                    else
                    {
                        if (isBottomCopeRequired)
                        {
                            TopologyCurve topologyCurve = CreateCopeCurve(CopePosition.WebBottom, copeProperties, cutbackOutput, isPlanar, isSquareEnd);
                            if (topologyCurve != null)
                            {
                                TopologySolid copeTopologySolid = base.CreateCutter(CopePosition.WebBottom, topologyCurve, null, isSquareEnd);
                                if (copeTopologySolid != null)
                                {
                                    // When model prior to V7 model is migrated, some cases may have TopologyCurve as cope cutter output. 
                                    // If TopologySolid is created then only delete the existing output.
                                    if (isCopeCutterATopologyCurve)
                                    {
                                        // Replace the old output with new one.
                                        base.ReplaceOutput(this.cope2, copeTopologySolid);
                                    }
                                    else
                                    {
                                        // Set the cope output.
                                        this.cope2.Output = copeTopologySolid;
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    // Delete this.cope1 (that always corresponds to WebTop location) assembly outputs if present.
                    if (HasCope(CopePosition.WebTop))
                    {
                        this.cope1.Delete();
                    }
                    // Delete this.cope2 (that always corresponds to WebBottom location) assembly outputs if present.
                    if (HasCope(CopePosition.WebBottom))
                    {
                        this.cope2.Delete();
                    }
                }
            }
            catch (Exception)
            {
                //may be there could be specific ToDo record is created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(FeaturesResourceIds.ErrToDoEvaluateAssembly,
                        "Unexpected error while evaluating custom assembly of {0}. Please check your custom code or contact S3D support."), this.ToString()));
                }
            }

        }

        /// <summary>
        /// OnPreLoad gets called immediately before the properties are loaded in the property page control. 
        ///             Its default implementation calls the simpler IsPropertyReadOnly method, which allows the developer
        ///             to set a property read-only on an individual basis. However, in cases where the entire context of properties
        ///             is needed then override this method instead. By overriding this method the IsPropertyReadOnly method will
        ///             not be invoked.
        /// </summary>
        /// <param name="businessObject">Occurrence or member output business object.</param>
        /// <param name="allDisplayedValues">All displayed values.</param>
        public override void OnPreLoad(BusinessObject businessObject, ReadOnlyCollection<PropertyDescriptor> allDisplayedValues)
        {
            this.isOnPreload = true; // Optimization to avoid value validation in OnAttrChange

            for (int i = 0; i < allDisplayedValues.Count; i++)
            {
                PropertyDescriptor propertyDescriptor = allDisplayedValues[i];
                PropertyValue propertyValue = propertyDescriptor.Property;
                string propertyName = propertyValue.PropertyInfo.Name;

                // Make all these properties read-only
                if (String.Compare(propertyName, StructureCustomAssembliesConstants.CopeRadius, true) == 0)
                {
                    propertyDescriptor.ReadOnly = true;
                }
                else
                {
                    string errorMessage = string.Empty;
                    OnPropertyChange(this.feature, allDisplayedValues, propertyDescriptor, propertyDescriptor.Property, out errorMessage);

                    if (errorMessage.Length > 0)
                    {
                        // Setting to false so that the value validation will be done in OnPropertyChange.
                        this.isOnPreload = false;
                    }
                }

                // No need to execute loop if isOnPreload is false.
                if (!this.isOnPreload)
                {
                    break;
                }
            }

            this.isOnPreload = false;
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
            bool isOnPropertyChange = true;
            errorMessage = String.Empty;
            // check if any of the arguments are null and throw argument null exception. 
            // businessObject and allDisplayedValues can be null 
            if (propertyToChange == null)
            {
                ArgumentNullException exception = new ArgumentNullException("propertyToChange");
                MiddleServiceProvider.ErrorLogger.Log(exception.ToString());
                throw exception;
            }
            if (newPropertyValue == null)
            {
                ArgumentNullException exception = new ArgumentNullException("newPropertyValue");
                MiddleServiceProvider.ErrorLogger.Log(exception.ToString());
                throw exception;
            }

            // Validate the attribute new value first before any further processing
            bool isValid = true;
            if (this.isOnPreload == false)
            {
                isValid = IsValid(propertyToChange, newPropertyValue, out errorMessage);
            }

            PropertyValueCodelist propertyValueCodeList;
            PropertyDescriptor propertyDescriptor;
            if (isValid)
            {
                if (String.Compare(propertyToChange.Property.PropertyInfo.Name, StructureCustomAssembliesConstants.SizingRule, true) == 0)
                {
                    propertyValueCodeList = (PropertyValueCodelist)newPropertyValue;
                    if (propertyValueCodeList.PropValue != (int)AssemblyConnectionSizing.UserDefined)  // By rule option for sizing rule
                    {
                        if (allDisplayedValues != null)
                        {
                        // If sizing rule is by rule, then gray out cope length and depth and remove graying of side flange clearance and increment.
                        for (int i = 0; i < allDisplayedValues.Count; i++)
                        {
                            propertyDescriptor = allDisplayedValues[i];
                            string propName = propertyDescriptor.Property.PropertyInfo.Name;
                            if (String.Compare(propName, SPSSymbolConstants.Length, true) == 0 || String.Compare(propName, SPSSymbolConstants.Depth, true) == 0)
                            {
                                propertyDescriptor.ReadOnly = true;
                            }
                            if (String.Compare(propName, StructureCustomAssembliesConstants.SideFlangeClearance, true) == 0 || String.Compare(propName, StructureCustomAssembliesConstants.Increment, true) == 0)
                            {
                                propertyDescriptor.ReadOnly = false;
                            }

                            // If it's inside flange clearance, make sure always planar is not grayed out.
                            if (String.Compare(propName, StructureCustomAssembliesConstants.InsideFlangeClearance, true) == 0)
                            {
                                for (int j = 0; j < allDisplayedValues.Count; j++)
                                {
                                    PropertyDescriptor tempPropertyDescriptor = allDisplayedValues[j];
                                    if (tempPropertyDescriptor.Property.PropertyInfo.Name == StructureCustomAssembliesConstants.AlwaysPlanar)
                                    {
                                        tempPropertyDescriptor.ReadOnly = false;
                                    }
                                }
                            }
                        }
                    }
                    }
                    else
                    {
                        if (allDisplayedValues != null)
                        {
                        // If its user defined,  make sure cope length and width is not grayed out.
                        for (int i = 0; i < allDisplayedValues.Count; i++)
                        {
                            propertyDescriptor = allDisplayedValues[i];
                            string propName = propertyDescriptor.Property.PropertyInfo.Name;
                            if (String.Compare(propName, SPSSymbolConstants.Length, true) == 0 || String.Compare(propName, SPSSymbolConstants.Depth, true) == 0)
                            {
                                propertyDescriptor.ReadOnly = false;
                            }
                            if (String.Compare(propName, StructureCustomAssembliesConstants.SideFlangeClearance, true) == 0 ||
                                String.Compare(propName, StructureCustomAssembliesConstants.SideFlangeClearance, true) == 0 ||
                                String.Compare(propName, StructureCustomAssembliesConstants.Increment, true) == 0)
                            {
                                propertyDescriptor.ReadOnly = true;
                            }
                        }
                    }
                }
                }

                // If the always planar bit is set to true, then the flange inside clearance
                // and the web clearance are to be grayed out; they should be editable otherwise

                string propToChangeName = propertyToChange.Property.PropertyInfo.Name;
                if (String.Compare(propToChangeName, StructureCustomAssembliesConstants.AlwaysPlanar, true) == 0)
                {
                    if (allDisplayedValues != null)
                    {
                    PropertyValueBoolean propValueBoolean = (PropertyValueBoolean)newPropertyValue;
                    if (propValueBoolean.PropValue == true)
                    {
                        // Gray out the flange inside clearance and web clearance
                        for (int i = 0; i < allDisplayedValues.Count; i++)
                        {
                            propertyDescriptor = allDisplayedValues[i];
                            string propName = propertyDescriptor.Property.PropertyInfo.Name;
                            if (String.Compare(propName, StructureCustomAssembliesConstants.InsideFlangeClearance, true) == 0 || String.Compare(propName, StructureCustomAssembliesConstants.WebClearance, true) == 0)
                            {
                                if (!propertyDescriptor.ReadOnly)
                                {
                                    propertyDescriptor.ReadOnly = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < allDisplayedValues.Count; i++)
                        {
                            propertyDescriptor = allDisplayedValues[i];
                            string propName = propertyDescriptor.Property.PropertyInfo.Name;
                            if (String.Compare(propName, StructureCustomAssembliesConstants.InsideFlangeClearance, true) == 0 || String.Compare(propName, StructureCustomAssembliesConstants.WebClearance, true) == 0)
                            {
                                if (propertyDescriptor.ReadOnly)
                                {
                                    propertyDescriptor.ReadOnly = false;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                isOnPropertyChange = false;
            }

            // No error, so set it to empty string.
            return isOnPropertyChange;
        }

        #endregion

        #region Private Properties and Methods

        /// <summary>
        /// Returns topology curve based on given cope location, cope properties and given cutback surface and existing cope output.
        /// </summary>
        /// <param name="copePosition">The cope position.</param>
        /// <param name="copeProperties">The cope properties.</param>
        /// <param name="cutbackOutput">The cutback output.</param>
        /// <param name="isPlanar">Boolean indicating whether the end of the member has copes or not.</param>
        /// <param name="isSquareEnd">If this is given as true, the member part is cut such that its end is square (normal to the axis of the member part); else if it is false, it creates a skewed end (parallel to the surface).</param>
        /// <returns>
        /// Cope curve contour
        /// </returns>
        private TopologyCurve CreateCopeCurve(CopePosition copePosition, CopeProperties copeProperties, TopologySurface cutbackOutput, bool isPlanar, bool isSquareEnd)
        {
            TopologyCurve topologyCurve = null;
            if (isSquareEnd)
            {
                topologyCurve = base.CreateSquaredCopeCutter(cutbackOutput, copePosition, copeProperties, isPlanar);
            }
            else
            {
                topologyCurve = base.CreateCutterContour(cutbackOutput, copePosition, copeProperties);
            }
            return topologyCurve;
        }

        /// <summary>
        /// Sets the length and depth on cope feature from its properties.
        /// </summary>
        /// <param name="copeProperties">The cope properties.</param>
        private void SetCopeLengthAndDepth(CopeProperties copeProperties)
        {
            // Set length and depth of the cope
            this.feature.SetPropertyValue(copeProperties.Length, StructureCustomAssembliesConstants.IJUASPSCope, SPSSymbolConstants.Length);
            this.feature.SetPropertyValue(copeProperties.Depth, StructureCustomAssembliesConstants.IJUASPSCope, SPSSymbolConstants.Depth);
        }

        /// <summary>
        /// Determines whether assembly output contains cope corresponding to the given cope position.
        /// </summary>
        /// <param name="copePosition">The cope position.</param>
        /// <returns>Boolean indicating whether assembly output contains cope at given position.</returns>
        private bool HasCope(CopePosition copePosition)
        {
            bool hasCope = false;
            if ((copePosition == CopePosition.WebTop && this.cope1.Output != null) || (copePosition == CopePosition.WebBottom && this.cope2.Output != null))
            {
                hasCope = true;
            }
            return hasCope;
        }

        /// <summary>
        /// Checks if given property value is valid for the property to changed and returns appropriate error message, if any.
        /// </summary>
        /// <param name="propertyToChange">The property to change.</param>
        /// <param name="newPropertyValue">The new property value.</param>
        /// <param name="errorMessage">The error message.</param>
        /// <returns>Boolean indicating if the given newPropertyValue is valid for the propertyToChange.</returns>
        private bool IsValid(PropertyDescriptor propertyToChange, PropertyValue newPropertyValue, out string errorMessage)
        {
            errorMessage = String.Empty;
            bool valid = true;
            PropertyValueDouble propValue = null;
            bool isPropertyValueDouble = false;
            if (newPropertyValue.PropertyInfo.PropertyType == SP3DPropType.PTDouble)
            {
                propValue = (PropertyValueDouble)newPropertyValue;
                isPropertyValueDouble = true;
                ValidationHelper.IsGreaterThanZero((double)((PropertyValueDouble)newPropertyValue).PropValue, ref errorMessage);
                if (errorMessage.Length > 0)
                {
                    valid = false;
                }
            }

            if (valid)
            {
                // If still the value is valid, if the attribute is InsideFlangeClearance, make sure the value is less than half of the bounding member's width
                string propertyName = propertyToChange.Property.PropertyInfo.Name;
                if (String.Compare(propertyName, StructureCustomAssembliesConstants.InsideFlangeClearance, true) == 0)
                {
                    if (isPropertyValueDouble)
                    {
                        double value = (double)propValue.PropValue;
                        valid = base.IsInsideFlangeClearanceValid(value);
                        if (!valid)
                        {
                            errorMessage = base.GetString(FeaturesResourceIds.ErrInvalidInsideFlangeClearance, "FlangeInsideClearance > half of Member depth");
                            if (errorMessage.Length > 0)
                            {
                                valid = false;
                            }
                        }
                    }
                }
            }

            return valid;
        }
        #endregion
    }
}