//=================================================================================================================
//
//Copyright 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  SurfaceTrimDefinition.cs 
//  Original Dll Name: <InstallDir>\SharedContent\Bin\SmartPlantStructure\Symbols\Release\SPSFeatureMacros.dll
//  Original Class Name: ‘SurfaceTrimDef’ in VB content
//
//Abstract
//	SurfaceTrimDefinition is a .NET custom assembly definition which creates graphic outputs for representing a Surface Trim Feature in the model.
//   This class subclasses from FeatureCustomAssemblyDefinition.
//=================================================================================================================
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Implementation of SurfaceTrimDefinition .NET custom assembly definition class.
    /// SurfaceTrimDefinition is a .NET custom assembly definition which trim the member with respect to the given surface    
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    [DefaultLocalizer(FeaturesResourceIds.DEFAULT_RESOURCE, FeaturesResourceIds.DEFAULT_ASSEMBLY)]
    [OutputNotification(StructureCustomAssembliesConstants.IID_IJUASPSSurfaceTrim)]
    public class SurfaceTrimDefinition : FeatureCustomAssemblyDefinition
    {
        //=====================================================================================================
        //DefinitionName/ProgID of this symbol is "Features,Ingr.SP3D.Content.Structure.SurfaceTrimDefinition"
        //=====================================================================================================

        #region Definitions of assembly outputs
        /// <summary>
        /// The surface trim cutback output
        /// </summary>
        [OutputNotification(SPSSymbolConstants.IID_IJSurface)]
        [AssemblyOutput(1, StructureCustomAssembliesConstants.SurfaceTrimCutback)]
        public AssemblyOutput surfaceTrimCutback;

        #endregion

        #region Public override functions and methods

        /// <summary>
        /// Construct and re-evaluate the custom assembly outputs.
        /// Here we will do the following:
        /// 1. Decide which assembly outputs are needed. 
        /// 2. Create the ones which are needed, delete which are not needed in the current context.
        /// 3. Sets definition properties on assembly outputs.
        /// </summary>     
        public override void EvaluateAssembly()
        {
            try
            {
                // Get the feature.
                Feature feature = (Feature)base.Occurrence;

                // Get the bounded port.
                BusinessObject boundedPort, boundingPort;
                feature.GetInputs(out boundedPort, out boundingPort);
                MemberPartAxisPort memberPartAxisPort = (MemberPartAxisPort)boundedPort;

                //Check whether surface trim required or not.
                if (base.IsSurfaceTrimRequired() == true)
                {
                    //Update the trim end value on feature
                    PropertyValueCodelist trimEndCodeList = (PropertyValueCodelist)feature.GetPropertyValue(StructureCustomAssembliesConstants.IJUASPSSurfaceTrim, StructureCustomAssembliesConstants.TrimEnd);
                    CodelistItem codeListItem = trimEndCodeList.PropertyInfo.CodeListInfo.GetCodelistItem((int)memberPartAxisPort.AxisPortType);
                    feature.SetPropertyValue(codeListItem, StructureCustomAssembliesConstants.IJUASPSSurfaceTrim, StructureCustomAssembliesConstants.TrimEnd);

                    // Construction/Update the surface trim
                    bool isSquareEnd = false;
                    double webAngle, flangeAngle;
                    double clearance = StructHelper.GetDoubleProperty(feature, StructureCustomAssembliesConstants.IJUASPSSurfaceTrim, StructureCustomAssembliesConstants.Offset);
                        
                    //ReadOnlyCollection<PropertyValue> properties = feature.GetAllProperties(); 
                    PropertyValue propertyValue = feature.GetPropertyValue(StructureCustomAssembliesConstants.IJUASPSSurfaceTrim, StructureCustomAssembliesConstants.SquaredEnd);
                    PropertyValueBoolean propertyValueBoolean = (PropertyValueBoolean)propertyValue;
                    if (propertyValueBoolean.PropValue != null)
                    {
                        isSquareEnd = (bool)propertyValueBoolean.PropValue;
                    }

                    TopologySurface surfaceCutback = null;

                    if (this.surfaceTrimCutback.Output == null)
                    {
                        //Construction Of Surface Trim                    
                        surfaceCutback = base.CreateSurfaceTrim(isSquareEnd, clearance, out webAngle, out flangeAngle);

                        //Set the cut back surface as assembly output.
                        if (surfaceCutback != null)
                        {
                            this.surfaceTrimCutback.Output = surfaceCutback;
                        }
                    }
                    else
                    {
                        //Update the Surface Trim   
                        surfaceCutback = (TopologySurface)this.surfaceTrimCutback.Output;
                        base.ModifySurfaceTrim(surfaceCutback, clearance, isSquareEnd, out webAngle, out flangeAngle);
                    }

                    //Set the Surface Trim part angles               
                    feature.SetPropertyValue(webAngle, StructureCustomAssembliesConstants.IJUASPSSurfaceTrim, StructureCustomAssembliesConstants.WebAngle);
                    feature.SetPropertyValue(flangeAngle, StructureCustomAssembliesConstants.IJUASPSSurfaceTrim, StructureCustomAssembliesConstants.FlangeAngle);
                }
            }
            catch (Exception)
            {
                //may be there could be specific to-do record is created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(FeaturesResourceIds.ErrToDoEvaluateAssembly,
                        "Unexpected error while evaluating custom assembly of {0}. Please check your custom code or contact S3D support."), this.ToString()));
                }
            }
        }

        /// <summary>
        /// Modifies read-only state of surface trim feature properties.
        /// </summary>
        /// <param name="assemblyOutputName">Name of the assembly output.</param>
        /// <param name="interfaceName">Interface name of the property.</param>
        /// <param name="propertyName">Name of the property.</param>
        /// <returns>
        /// A Boolean defining whether read-only.
        /// </returns>
        public override bool IsPropertyReadOnly(string assemblyOutputName, string interfaceName, string propertyName)
        {
            bool result = false;
            switch (propertyName)
            {
                case StructureCustomAssembliesConstants.WebAngle:
                case StructureCustomAssembliesConstants.FlangeAngle:
                case StructureCustomAssembliesConstants.TrimEnd:
                    result = true;
                    break;
            }
            return result;
        }

        /// <summary>
        /// Validates the values of the surface trim feature Properties.
        /// </summary>
        /// <param name="assemblyOutputName">Name of the assembly output or null string for the parent assembly.</param>
        /// <param name="interfaceName">Interface name for the property being validated.</param>
        /// <param name="propertyName">Name of the property being validated.</param>
        /// <param name="propertyValue">New property value being proposed.</param>
        /// <param name="errorMessage">Returned error message for the user to indicate why the property is not valid.</param>
        /// <returns>
        /// Boolean indicating whether the property is valid or not.
        /// </returns>
        public override bool IsPropertyValid(string assemblyOutputName, string interfaceName, string propertyName, object propertyValue, out string errorMessage)
        {
            // by default set the property value as valid. Override the value later for known checks
            bool isValidPropertyValue = true;
            errorMessage = string.Empty;
            if (propertyValue != null)
            {
                switch (propertyName)
                {
                    // following property values must be greater than 0
                    case StructureCustomAssembliesConstants.WebAngle:
                    case StructureCustomAssembliesConstants.FlangeAngle:
                    case StructureCustomAssembliesConstants.Offset:
                        isValidPropertyValue = ValidationHelper.IsGreaterThanZero(Convert.ToDouble(propertyValue), ref errorMessage);
                        break;

                }
            }
            return isValidPropertyValue;
        }

        #endregion
    }
}