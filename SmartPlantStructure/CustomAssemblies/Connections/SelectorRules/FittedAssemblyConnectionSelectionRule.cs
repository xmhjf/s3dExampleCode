//-----------------------------------------------------------------------------------------------------------------------------------------
//Copyright 1992 - 2013 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  FittedAssemblyConnectionSelectionRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\SmartPlantStructure\Symbols\Release\SPSACMacros.dll
//  Original Class Name: ‘FittedAsmConnSel’ in VB content
//
//Abstract
//	FittedAssemblyConnectionSelectionRule is a .NET selection rule for fitted assembly connections. 
//      
//Change History:
//   5th Sept 2013    srkombat    -Created
//-----------------------------------------------------------------------------------------------------------------------------------------
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Selector class for Fitted AssemblyConnection, which selects the list of available items in the context of the Fitted AssemblyConnection.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [RuleInterface(StructureCustomAssembliesConstants.IASPSACMacros_FittedAsmConnSel, StructureCustomAssembliesConstants.IASPSACMacros_FittedAsmConnSel, StructureCustomAssembliesConstants.CUSPSFittedAsmConnOV)]
    public class FittedAssemblyConnectionSelectionRule : SelectorRule
    {
        //======================================================================================================================================
        //DefinitionName/ProgID of this symbol is "MemberAssemblyConnections,Ingr.SP3D.Content.Structure.FittedAssemblyConnectionSelectionRule"
        //======================================================================================================================================

        #region Public override properties and methods

        /// <summary>
        /// Returns different selections based on the connection configuration
        /// </summary>
        public override ReadOnlyCollection<string> Selections
        {
            get
            {
                Collection<string> choices = new Collection<string>();
                try
                {
                    AssemblyConnection assemblyConnection = (AssemblyConnection)base.Occurrence;

                    //Get the BoundingPorts and BoundedPorts from AssemblyConnection
                    Collection<IPort> boundingPorts = assemblyConnection.BoundingPorts;
                    if (boundingPorts.Count == 0)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(MemberAssemblyConnectionsResourceIds.ErrBoundingPortCount,
                            "No bounding ports associated with the AssemblyConnection while retrieving the bounding port for the assembly connection. Please check your inputs or contact S3D support."));

                        //ToDo list is created with error type hence stop computation
                        return new ReadOnlyCollection<string>(choices);
                    }

                    Collection<IPort> boundedPorts = assemblyConnection.BoundedPorts;
                    if (boundedPorts.Count == 0)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(MemberAssemblyConnectionsResourceIds.ErrBoundedPortCount,
                            "No bounded ports associated with the AssemblyConnection while retrieving the bounded port for the assembly connection. Please check your inputs or contact S3D support."));

                        //ToDo list is created with error type hence stop computation
                        return new ReadOnlyCollection<string>(choices);
                    }

                    MemberPartAxisPort boundedPort = boundedPorts[0] as MemberPartAxisPort;
                    MemberPartAxisPort boundingPort = boundingPorts[0] as MemberPartAxisPort;
                    if (boundingPort == null || boundedPort == null)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(MemberAssemblyConnectionsResourceIds.ErrInvalidPortType,
                            "Either bounded port or bounding port is not a MemberPartAxisPort. Please check your inputs or contact S3D support."));

                        //ToDo list is created with error type hence stop computation
                        return new ReadOnlyCollection<string>(choices);
                    }

                    MemberPart boundedMemberPart = (MemberPart)boundedPort.Connectable;
                    MemberPart boundingMemberPart = (MemberPart)boundingPort.Connectable;
                    if (boundedMemberPart.SectionType == StructureCustomAssembliesConstants.L && boundingMemberPart.SectionType == StructureCustomAssembliesConstants.L)
                    {
                        //for angle sections, select fitted with no clearance
                        string partName = StructureCustomAssembliesConstants.FittedAsmConn_No_Clearance;
                        CatalogStructHelper catalogHelper = new CatalogStructHelper();
                        if (catalogHelper.DoesPartOrPartClassExist(partName))
                        {
                            choices.Add(partName);
                        }
                        else//old catalog may not have this item,add regular part
                        {
                            choices.Add(StructureCustomAssembliesConstants.FittedAsmConn_1);
                        }
                    }
                    else
                    {
                        choices.Add(StructureCustomAssembliesConstants.FittedAsmConn_1);
                    }
                }
                catch (Exception)
                {
                    if (base.ToDoListMessage == null)//if it is not null, which means ToDo record with specific reason has been already created. so don't over write with following generic ToDo message.
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(MemberAssemblyConnectionsResourceIds.ErrUnExpectedInSelectionRule,
                        "Unexpected error while getting the selections from Fitted AssemblyConnection selector rule. Please check your custom code or contact S3D support."));
                    }
                }

                return new ReadOnlyCollection<string>(choices);
            }
        }

        #endregion Public override properties and methods
    }
}