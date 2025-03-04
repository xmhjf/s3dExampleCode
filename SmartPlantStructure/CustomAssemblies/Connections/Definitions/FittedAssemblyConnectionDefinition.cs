//-----------------------------------------------------------------------------------------------------------------------------------------
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  FittedAssemblyConnectionDefinition.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\SmartPlantStructure\Symbols\Release\SPSACMacros.dll
//  Original Class Name: ‘FittedAsmConnDef’ in VB content
//
//Abstract
//	FittedAssemblyConnectionDefinition is a .NET custom assembly definition, which creates an assembly connection along with cope feature.
//
//Change History:
//  5th Sept 2013    srkombat    -Created
//-----------------------------------------------------------------------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Implementation of FittedAssemblyConnectionDefinition.
    /// Evaluates and creates assembly connection outputs for fitted ACs.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    public class FittedAssemblyConnectionDefinition : AssemblyConnectionCustomAssemblyDefinition
    {
        //===================================================================================================================================
        //DefinitionName/ProgID of this symbol is "MemberAssemblyConnections,Ingr.SP3D.Content.Structure.FittedAssemblyConnectionDefinition"
        //===================================================================================================================================

        #region Private members

        AssemblyConnection assemblyConnection = null;
        MemberPartAxisPort boundedPort = null;
        MemberPartAxisPort boundingPort = null;
        #endregion

        #region Definitions of assembly outputs

        [AssemblyOutput(1, StructureCustomAssembliesConstants.FittedAsmConnCope)]
        public AssemblyOutput fittedAssemblyConnectionCope;
        #endregion Definitions of assembly outputs

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
                this.assemblyConnection = (AssemblyConnection)base.Occurrence;

                //get the bounded port
                boundedPort = (MemberPartAxisPort)this.BoundedPort;
                if (boundedPort == null)
                {
                    //stop evaluating, no bounded port available.
                    return;
                }
                boundingPort = (MemberPartAxisPort)this.BoundingPort;
                if (boundingPort == null)
                {
                    //stop evaluating, no bounding port available.
                    return;
                }

                MemberPart boundedMemberPart = (MemberPart)boundedPort.Connectable;
                MemberPart boundingMemberPart = (MemberPart)boundingPort.Connectable;

                //check if the cope feature is needed
                Feature copeFeature = null;
                if (IsCopeFeatureNeeded(boundedMemberPart.SectionType, boundingMemberPart.SectionType))
                {
                    if (this.fittedAssemblyConnectionCope.Output == null)
                    {
                        //call the feature constructor to create the assembly output (cope)
                        copeFeature = new Feature(assemblyConnection, boundedPort, boundingPort, FeatureType.Trim, StructureCustomAssembliesConstants.CopeFeature);

                        //add this feature as assembly output for the AC
                        this.fittedAssemblyConnectionCope.Output = copeFeature;
                    }
                    else
                    {
                        //already created one
                        copeFeature = (Feature)this.fittedAssemblyConnectionCope.Output;
                    }
                    //now copy the parent attributes on child feature occurrence
                    base.SetParentAttributesOnChild(copeFeature, StructureCustomAssembliesConstants.IJUASPSCope);

                }
            }
            catch (Exception)
            {
                //may be there could be specific ToDo record is created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, String.Format(base.GetString(MemberAssemblyConnectionsResourceIds.ToDoEvaluateAssembly,
                        "Unexpected error while evaluating {0}. Check your custom code or contact S3D support."),this.ToString()));
                }
            }
        }

        /// <summary>
        /// Gets the valid input objects.
        /// </summary>
        /// <param name="inputObjects">The input objects.</param>
        /// <param name="validInputObjects">The valid input objects.</param>
        /// <returns></returns>
        public override AssemblyConnectionInputObjectsStatus GetValidInputObjects(Collection<BusinessObject> inputObjects, out Collection<BusinessObject> validInputObjects)
        {
            validInputObjects = null;
            //initialize the status enum to ok
            AssemblyConnectionInputObjectsStatus inputObjectsStatus = AssemblyConnectionInputObjectsStatus.Ok;
            //get the actual ports in the input objects.
            List<BusinessObject> inputPorts = inputObjects.Where(inputObject => inputObject is IPort).ToList();

            //we expect only two ports in the collection, if not return bad no of object status
            if (inputPorts.Count != 2)
            {
                inputObjectsStatus = AssemblyConnectionInputObjectsStatus.BadNumberOfObjects;
            }
            else
            {
                //we have two ports, now check if these ports are MemberPartAxisPort, if so, assign along port to bounding and start/end to bounded
                MemberPartAxisPort boundedPort = null;
                MemberPartAxisPort boundingPort = null;
                foreach (BusinessObject inputPort in inputPorts)
                {
                    MemberPartAxisPort tempPort = inputPort as MemberPartAxisPort;
                    if (tempPort != null)
                    {
                        if (tempPort.AxisPortType == MemberAxisPortType.Along)
                        {
                            boundingPort = tempPort;
                        }
                        else
                        {
                            boundedPort = tempPort;
                        }
                    }
                }

                //both should not be null
                if (boundedPort == null || boundingPort == null)
                {
                    inputObjectsStatus = AssemblyConnectionInputObjectsStatus.InvalidTypeOfObject;
                }
                else //add the ports back to valid collections
                {
                    validInputObjects = new Collection<BusinessObject>();
                    validInputObjects.Add(boundedPort);
                    validInputObjects.Add(boundingPort);
                }
            }
            return inputObjectsStatus;
        }

        #endregion Public override functions and methods

        #region Private methods

        /// <summary>
        /// Determines whether [is cope feature needed] [the specified bounded section type].
        /// </summary>
        /// <param name="boundedSectionType">Type of the bounded section.</param>
        /// <param name="boundingSectionType">Type of the bounding section.</param>
        /// <returns></returns>
        private bool IsCopeFeatureNeeded(string boundedSectionType, string boundingSectionType)
        {
            bool isCopeFeatureNeeded = true;
            string message;
            if (!IsValidSectionType(boundedSectionType))
            {
                message = String.Format(base.GetString(MemberAssemblyConnectionsResourceIds.InValidBoundedSection,
                        "Bounded MemberPart section type is invalid, while determining whether cope feature is needed or not, in {0}. Please check your inputs or contact S3D support."),this.ToString());

                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, StructureCustomAssembliesConstants.StructACToDoMessages,
                                                            StructureCustomAssembliesConstants.UnknownSectionType, message, assemblyConnection);
                throw new CmnException(message);
            }
            if (!IsValidSectionType(boundingSectionType))
            {
                message = String.Format(base.GetString(MemberAssemblyConnectionsResourceIds.InValidBoundingSection,
                          "Bounding MemberPart section type is invalid, for determining whether cope feature needed or not, in {0}. Please check your inputs or contact S3D support."),this.ToString());
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, StructureCustomAssembliesConstants.StructACToDoMessages,
                                                            StructureCustomAssembliesConstants.UnknownSectionType, message, assemblyConnection);
                throw new CmnException(message);

            }

            if (boundedPort.AxisPortType != MemberAxisPortType.Along)//Along port can be connected to more than one AC
            {
                if (base.IsPortConnectedToOtherAC(boundedPort))
                {
                    message = String.Format(base.GetString(MemberAssemblyConnectionsResourceIds.InvalidInputForBounded,
                            "Bounded port is already in relation with other assembly connection, for determining whether cope feature is needed or not, in {0}. Please check your inputs or contact S3D support."),this.ToString());

                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, StructureCustomAssembliesConstants.StructACToDoMessages,
                                                            StructureCustomAssembliesConstants.MultipleACsExist, message, assemblyConnection);
                    throw new CmnException(message);

                }
            }

            if (boundingPort.AxisPortType != MemberAxisPortType.Along)//Along port can be connected to more than one AC
            {
                if (base.IsPortConnectedToOtherAC(boundingPort))
                {
                    message = String.Format(base.GetString(MemberAssemblyConnectionsResourceIds.InvalidInputForBounding,
                            "Bounding port is already in relation with other assembly connection, for determining whether cope feature is needed or not, in {0}. Please check your inputs or contact S3D support."),this.ToString());

                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, StructureCustomAssembliesConstants.StructACToDoMessages,
                                                            StructureCustomAssembliesConstants.MultipleACsExist, message, assemblyConnection);
                    throw new CmnException(message);

                }
            }


            return isCopeFeatureNeeded;
        }

        /// <summary>
        /// Determines whether [is valid section type] [the specified bounded section type].
        /// </summary>
        /// <param name="boundedSectionType">Type of the bounded section.</param>
        /// <returns></returns>
        private bool IsValidSectionType(string boundedSectionType)
        {
            bool isValid = false;
            switch (boundedSectionType)
            {
                case StructureCustomAssembliesConstants.W:
                case StructureCustomAssembliesConstants.S:
                case StructureCustomAssembliesConstants.HP:
                case StructureCustomAssembliesConstants.M:
                case StructureCustomAssembliesConstants.HSSC:
                case StructureCustomAssembliesConstants.CS:
                case StructureCustomAssembliesConstants.PIPE:
                case StructureCustomAssembliesConstants.L:
                case StructureCustomAssembliesConstants.C:
                case StructureCustomAssembliesConstants.MC:
                case StructureCustomAssembliesConstants.WT:
                case StructureCustomAssembliesConstants.MT:
                case StructureCustomAssembliesConstants.ST:
                case StructureCustomAssembliesConstants.Double_L:
                case StructureCustomAssembliesConstants.RS:
                case StructureCustomAssembliesConstants.HSSR:
                    isValid = true;
                    break;
                default:
                    break;
            }
            return isValid;
        }


        /// <summary>
        /// Bounded port for the assembly connection
        /// </summary>
        private StructPortBase BoundedPort
        {
            get
            {
                AssemblyConnection assemblyConnection = (AssemblyConnection)base.Occurrence;

                Collection<IPort> boundedPorts = assemblyConnection.BoundedPorts;
                //Get the BoundingPorts and BoundedPorts from AssemblyConnection
                if (boundedPorts.Count == 0)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(MemberAssemblyConnectionsResourceIds.ErrBoundedPortCount,
                        "No bounded ports associated with the AssemblyConnection while retrieving bounded port for the assembly connection. Please check your inputs or contact S3D support."));

                    //ToDo list is created with error type hence stop computation
                    return null;
                }
                return (StructPortBase)boundedPorts[0];
            }
        }

        /// <summary>
        /// Bounding port for the assembly connection
        /// </summary>
        private StructPortBase BoundingPort
        {
            get
            {
                AssemblyConnection assemblyConnection = (AssemblyConnection)base.Occurrence;

                Collection<IPort> boundingPorts = assemblyConnection.BoundingPorts;
                //Get the BoundingPorts and BoundedPorts from AssemblyConnection
                if (boundingPorts.Count == 0)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(MemberAssemblyConnectionsResourceIds.ErrBoundingPortCount,
                        "No bounding ports associated with the AssemblyConnection while retrieving bounding port for the assembly connection. Please check your inputs or contact S3D support."));

                    //ToDo list is created with error type hence stop computation
                    return null;
                }
                return (StructPortBase)boundingPorts[0];
            }
        }
        #endregion Private methods
    }
}