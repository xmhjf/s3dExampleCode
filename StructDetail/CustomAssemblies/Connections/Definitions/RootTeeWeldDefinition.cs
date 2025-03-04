//------------------------------------------------------------------------------------------------------------
//Copyright 1998-2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  RootTeeWeldDefinition.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMPhysConnRules.dll
//  Original Class Name: ‘RootTeeWeldDef’ in VB content
//
//Abstract  
//  RootTeeWeldDefinition is a .NET definition rule which is defining the custom definition rule. 
//  This class derives from PhysicalConnectionCustomAssemblyDefinition. 
//--------------------------------------------------------------------------------------------------------------

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Implementation of RootTeeWeld .NET custom assembly definition class. 
    /// RootTeeWeldDefinition is a .NET custom assembly definition which creates physical connections.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    [DefaultLocalizer(ConnectionsResourceIdentifiers.ConnectionsResource, ConnectionsResourceIdentifiers.ConnectionsAssembly)]
    public class RootTeeWeldDefinition : PhysicalConnectionCustomAssemblyDefinition
    {
        //===================================================================================================================
        //DefinitionName/ProgID of this symbol is "Connections,Ingr.SP3D.Content.Structure.RootTeeWeldDefinition"
        //===================================================================================================================
        #region Definitions of assembly outputs
        /// <summary>
        /// Places AutomaticSplitter on  Physical Connection .
        /// </summary>
        [AssemblyOutput(1, "AutomaticSplitter")]
        public AssemblyOutputs automaticSplitters;
        #endregion Definitions of assembly outputs

        #region Public override functions and methods
        /// <summary>
        /// Construct and re-evaluate the custom assembly outputs.
        /// Here we will do the following:
        /// 1. Decide which assembly outputs are needed. 
        /// 2. Create the ones which are needed, delete which are not needed now.
        /// </summary>
        public override void EvaluateAssembly()
        {
            try
            {
                //Validating the inputs required to create  the Physical Connection.
                //This is a virtual method and can be overriden to add specific validation, if specific input is missing.
                ValidateInputs();
                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created while validating the inputs
                    return;
                }
                PhysicalConnection physicalConnection = (PhysicalConnection)base.Occurrence;
                //Get the weld category
                string category = ((PropertyValue)(physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.Category))).ToString();
                // If the physical connection angle is not constant (e.g., because of twist), 
                // then it needs to be split to allow different weld properties on different segments.  The list of threshold split angles 
                // are used to determine if a physical connection needs to be split by checking if the angles between the plates are greater than
                // the given threshold angles.
                IEnumerable<double> thresholdAngles = ConnectionServices.ComputeSplitAngles(category, physicalConnection.BoundedObject, physicalConnection.BoundingObject);
                ReadOnlyCollection<Position> splitLocations = physicalConnection.GetAutomaticSplitterLocations(thresholdAngles);
                //Update the split points if exists any otherwise create a new split point 
                for (int i = 0; i < splitLocations.Count; i++)
                {
                    Position newReferencePoint = splitLocations[i];
                    if (i < this.automaticSplitters.Count)
                    {
                        Point3d existingReferencePoint = (Point3d)automaticSplitters[i];

                        if (!StructHelper.AreEqual(newReferencePoint, existingReferencePoint.Position))
                        {
                            existingReferencePoint.Set(new Point3d(newReferencePoint.X, newReferencePoint.Y, newReferencePoint.Z));
                        }
                    }
                    else
                    {
                        Point3d point = base.AddAutomaticSpliter(newReferencePoint);
                        automaticSplitters.Add(point);
                    }
                }

                //Delete the existing split points which are not required
                for (int i = splitLocations.Count; i < automaticSplitters.Count; i++)
                {
                    automaticSplitters[i].Delete();
                }
            }
            catch
            {
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError,
                        base.GetString(ConnectionsResourceIdentifiers.ErrEvaluateAssembly,
                        "Unexpected error while evaluating the PhysicalConnectionCustomAssemblyDefinition."));
                }
            }
        }
        #endregion Public override functions and methods


    }
}
