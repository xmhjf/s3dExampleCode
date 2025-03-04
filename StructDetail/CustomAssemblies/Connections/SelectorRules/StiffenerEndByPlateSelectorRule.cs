//-----------------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  StiffenerEndByPlateSelectorRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMAssyConRul.dll
//  Original Class Name: ‘StiffEndByPlateSel’ in VB content
//
//Abstract
//	StiffenerEndByPlateSelectorRule is a .NET selection rule for stifener bounded by plate. 
//      
//Change History:
//  dd.mmm.yyyy    who    change description
//-----------------------------------------------------------------------------------------------------------------------------------------
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using System.Collections.Generic;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Selector class to make various selections for stiffener bounded by plate
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [RuleInterface(DetailingCustomAssembliesConstants.IASMAssyConRul_StiffEndByPlateSel, DetailingCustomAssembliesConstants.IASMAssyConRul_StiffEndByPlateSel)]
    [DefaultLocalizer(ConnectionsResourceIdentifiers.ConnectionsResource, ConnectionsResourceIdentifiers.ConnectionsAssembly)]
    public class StiffenerEndByPlateSelectorRule : SelectorRule
    {
        //=================================================================================================================
        //DefinitionName/ProgID of this symbol is "Connections,Ingr.SP3D.Content.Structure.StiffenerEndByPlateSelectorRule"
        //=================================================================================================================
        #region Selector questions
        [SelectorQuestionCodelist(100, DetailingCustomAssembliesConstants.EndCutType, DetailingCustomAssembliesConstants.EndCutType, DetailingCustomAssembliesConstants.EndCutTypeCodeList, (int)EndCutTypes.Welded)]
        public SelectorQuestionCodelist endcutTypeAnswer;
        [SelectorQuestionCodelist(101, DetailingCustomAssembliesConstants.PlaceBracket, DetailingCustomAssembliesConstants.PlaceBracket, DetailingCustomAssembliesConstants.BooleanCol, (int)Answer.No)]
        public SelectorQuestionCodelist placeBracketAnswer;
        #endregion Selector questions

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
                    if (boundingPorts.Count != 1)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ConnectionsResourceIdentifiers.ErrBoundingPortCount,
                            "AssemblyConnection has more than one bounding port."));

                        //ToDo list is created with error type hence stop computation
                        return new ReadOnlyCollection<string>(choices);
                    }

                    Collection<IPort> boundedPorts = assemblyConnection.BoundedPorts;
                    if (boundedPorts.Count != 1)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ConnectionsResourceIdentifiers.ErrBoundedPortCount,
                            "No bounded ports associated with the AssemblyConnection."));

                        //ToDo list is created with error type hence stop computation
                        return new ReadOnlyCollection<string>(choices);
                    }

                    TopologyPort boundedPort = boundedPorts[0] as TopologyPort;
                    TopologyPort boundingPort = boundingPorts[0] as TopologyPort;
                    if (boundingPort == null || boundedPort == null)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ConnectionsResourceIdentifiers.ErrACInputs,
                            "One of the input ports is not a TopologyPort."));

                        //ToDo list is created with error type hence stop computation
                        return new ReadOnlyCollection<string>(choices);
                    }

                    //Get the context ID of bounding Port
                    ContextTypes boundingPortContextId = boundingPort.ContextId;

                    //case where the stiffener is bounded by plate face (base or offset)
                    if (boundingPortContextId.HasFlag(ContextTypes.Base) || boundingPortContextId.HasFlag(ContextTypes.Offset))
                    {
                        choices.Add(DetailingCustomAssembliesConstants.StiffenerEndToPlateFace);
                    }
                    else if (boundingPortContextId.HasFlag(ContextTypes.Lateral))
                    {
                        //case to indentify plate edge chamfered by stiffener edge
                        if (this.IsChamferNeeded)
                        {
                            CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                            //Verify that the required/expected AsseemblyConnection part class/part have been bulkloaded                            
                            if (catalogStructHelper.DoesPartOrPartClassExist(DetailingCustomAssembliesConstants.ChamferEdgeToStiffenerEdge))
                            {
                                if (catalogStructHelper.DoesPartOrPartClassExist(DetailingCustomAssembliesConstants.ChamferEdgeBase))
                                {
                                    choices.Add(DetailingCustomAssembliesConstants.StiffenerEndToChamferEdge);
                                }
                            }
                        }

                        //if stiffener bounded by lateral face
                        choices.Add(DetailingCustomAssembliesConstants.StiffenerEndToPlateEdge);
                    }
                    else
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ConnectionsResourceIdentifiers.ErrInvalidBoundingPort,
                            "Invalid bounding port, ContextId should be Base or Offset or Lateral."));
                    }
                }
                catch (Exception)
                {
                    //There could be specific ToDo records created before, so do not override it.
                    if (base.ToDoListMessage == null)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(ConnectionsResourceIdentifiers.ToDoSelections,
                            "Unexpected error while evaluating the {0}"), this.ToString()));
                    }                   
                }

                return new ReadOnlyCollection<string>(choices);
            }
        }

        /// <summary>
        /// Sets the value of the selector question when it is controlled by the system. 
        /// The default answer will be the value provided via the SelectorQuestion attribute and
        /// will be invoked for each “system controlled” question prior to invoking the 'Selections' method.
        /// </summary>
        /// <param name="selectorQuestion">The selector question whose default answer is being overridden by the user through code.</param>
        public override void OverrideDefaultAnswer(SelectorQuestion selectorQuestion)
        {
            try
            {
                //get the selector question name
                string selectorQuestionName = selectorQuestion.Name;

                if (selectorQuestionName == DetailingCustomAssembliesConstants.EndCutType)
                {
                    //set 'Weleded' for scenarios except for the following condition
                    EndCutTypes endCutType = EndCutTypes.Welded;

                    AssemblyConnection assemblyConnection = (AssemblyConnection)base.Occurrence;
                    TopologyPort boundedPort = (TopologyPort)assemblyConnection.BoundedPorts[0];
                    StiffenerPartBase stiffenerPartBase = boundedPort.Connectable as StiffenerPartBase;
                    if (stiffenerPartBase != null)
                    {
                        //get stiffener connection method                        
                        int connectionMethod = boundedPort.ContextId.HasFlag(ContextTypes.Base) ? stiffenerPartBase.StartConnectionMethod : stiffenerPartBase.EndConnectionMethod;
                        if (connectionMethod > (int)StiffenerConnectionMethod.ByRule)
                        {
                            switch (connectionMethod)
                            {
                                case (int)StiffenerConnectionMethod.Bracketed:
                                    endCutType = EndCutTypes.Bracketed;
                                    break;
                                case (int)StiffenerConnectionMethod.Cutback:
                                    endCutType = EndCutTypes.Cutback;
                                    break;
                                case (int)StiffenerConnectionMethod.Snipe:
                                    endCutType = EndCutTypes.Snip;
                                    break;
                            }
                        }
                        else
                        {
                            //Endcut type should be snip if the bounded object is a stiffener on bracket plate system.
                            StiffenerSystemBase rootStiffenerSystem = stiffenerPartBase.RootStiffenerSystem;
                            if (rootStiffenerSystem != null)
                            {
                                BracketPlateSystem plateToStiffen = rootStiffenerSystem.PlateToStiffen as BracketPlateSystem;
                                if (plateToStiffen != null)
                                {
                                    endCutType = EndCutTypes.Snip;
                                }
                            }
                        }
                    }

                    this.endcutTypeAnswer.Value = (int)endCutType;
                }
                else if (selectorQuestionName == DetailingCustomAssembliesConstants.PlaceBracket)
                {
                    //place bracket only when the end cut type is bracket.
                    this.placeBracketAnswer.Value = (this.endcutTypeAnswer.Value == (int)EndCutTypes.Bracketed) ? (int)Answer.Yes : (int)Answer.No;
                }
            }
            catch (Exception)
            {
                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ConnectionsResourceIdentifiers.ToDoDefaultAnswer,
                    "Unexpected error while setting the default answer of StiffenerEndByPlate AssemblyConnection selector rule."));
            }
        }

        #endregion Public override properties and methods

        #region Private methods

        /// <summary>
        /// Determines whether a chamfer can be placed between plate or stiffener.
        /// </summary>
        /// <param name="assemblyConnection">The AssemblyConnection.</param>
        /// <returns>True if a chamfer can be placed between plate or stiffener, otherwise false.</returns>
        private bool IsChamferNeeded
        {
            get
            {
                bool isChamferNeeded = false;
                AssemblyConnection assemblyConnection = (AssemblyConnection)base.Occurrence;
                IPort boundedPort = assemblyConnection.BoundedPorts[0];
                IPort boundingPort = assemblyConnection.BoundingPorts[0];

                TopologyPort platePartBasePort = null, stiffenerPartBasePort = null;

                //chamfer can only be possible between plate and stiffener or stiffener and plate so check for its connectables
                if (boundedPort.Connectable is PlatePartBase && boundingPort.Connectable is StiffenerPartBase)
                {
                    platePartBasePort = boundedPort as TopologyPort;
                    stiffenerPartBasePort = boundingPort as TopologyPort;
                }
                else if (boundedPort.Connectable is StiffenerPartBase && boundingPort.Connectable is PlatePartBase)
                {
                    platePartBasePort = boundingPort as TopologyPort;
                    stiffenerPartBasePort = boundedPort as TopologyPort;
                }

                if (platePartBasePort != null && stiffenerPartBasePort != null)
                {
                    //connectables are valid.
                    //now, check for its context id
                    //for plate, only lateral ports are valid
                    //for stifeners, base/offset/lateral ports are valid
                    ContextTypes platePortContextId = platePartBasePort.ContextId;
                    ContextTypes stiffenerPortContextId = stiffenerPartBasePort.ContextId;

                    if (platePortContextId.HasFlag(ContextTypes.Lateral) &&
                        (stiffenerPortContextId.HasFlag(ContextTypes.Base) ||
                        stiffenerPortContextId.HasFlag(ContextTypes.Offset) ||
                        stiffenerPortContextId.HasFlag(ContextTypes.Lateral)))
                    {
                        StiffenerPartBase stiffenerPartBase = (StiffenerPartBase)stiffenerPartBasePort.Connectable;
                        //currently not handling other type of cross sections
                        if (stiffenerPartBase.SectionType == MarineSymbolConstants.FB)
                        {
                            //now check the mounting face name is either web left or web right
                            int mountingFaceName = stiffenerPartBase.MountingFaceName;
							if (mountingFaceName == (int)StiffenerMountingFaceName.LeftWeb ||
                                mountingFaceName == (int)StiffenerMountingFaceName.RightWeb)
                            {
                                isChamferNeeded = true;
                            }
                        }
                    }
                }

                return isChamferNeeded;
            }
        }

        #endregion Private methods
    }
}
