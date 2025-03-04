//------------------------------------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  LeafTeeWeldSelectorRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMPhysConnRules.dll
//  Original Class Name: ‘LeafTeeWeldSel’ in VB content
//
//Abstract
//	 LeafTeeWeldSelectorRule is a .NET selection rule for Physicalconnections which selects the list of available items in the context of the LeafTeeWeld.
//   This class subclasses from PhysicalConnectionSelectorRule.    
//
//History:
//    2-Mar-2016    svsmylav     TR-288522: Removed unnecessary range check 'if AC is within the CF rangebox, and if so do not select PC item': after the fix,
//                                 'DeletedPCPortToCA' relationship will be added incase of no overlapping geometry between the bounded and the bounding parts. 
//-------------------------------------------------------------------------------------------------------------------------------------------------------------
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Selector for LeafTeeWeldSel, which selects the list of available items in the context of the LeafTeeWeldSelection.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    //New content does not need to specify the interface name, but this is specified here to support code written against old rules which directly uses interface name to get selector question. 
    [RuleInterface(DetailingCustomAssembliesConstants.IASMPhysConnRules_LeafTeeWeldSel, DetailingCustomAssembliesConstants.IASMPhysConnRules_LeafTeeWeldSel)]
    [DefaultLocalizer(ConnectionsResourceIdentifiers.ConnectionsResource, ConnectionsResourceIdentifiers.ConnectionsAssembly)]
    public class LeafTeeWeldSelectorRule : PhysicalConnectionSelectorRule
    {
        //=========================================================================================================
        //DefinitionName/ProgID of this symbol is "Connections,Ingr.SP3D.Content.Structure.LeafTeeWeldSelectorRule"
        //=========================================================================================================

        #region Public methods
        /// <summary>
        ///  Selector for LeafTeeWelds, which selects the list of available items in the context of the Physical Connections.
        /// </summary>
        public override ReadOnlyCollection<string> Selections
        {
            get
            {
                Collection<string> choices = new Collection<string>();
                try
                {
                    // Get Class Arguments
                    PhysicalConnection physicalConnection = (PhysicalConnection)base.Occurrence;

                    //Validating the inputs required to create the slot.
                    //This is a virtual method and can be overriden to add specific validation, if specific input is missing.
                    ValidateInputs();
                    //Check if any ToDo message of type error has been created during the validation of inputs
                    if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                    {
                        //ToDo list is created while validating the inputs
                        return new ReadOnlyCollection<string>(choices);
                    }
                    // Use filter in place of selection logic, if set
                    choices = ConnectionServices.GetFilteredSelections(physicalConnection);
                    if (choices.Count == 0)
                    {
                        double mountingAngle = physicalConnection.GetConnectionAngle();
                        if (Math.Abs((mountingAngle - StructHelper.DISTTOL)) >= Math.PI / 2)
                        {
                            mountingAngle = Math.PI - mountingAngle;
                        }
                        //this is the theta angle in the requirements
                        double mountingAngleCompliment = (Math.PI / 2) - mountingAngle;

                        //Check if this Physical Connection is "Square Trim"
                        //Currently, it is NOT expected that "Square Trim" cases are mixed
                        //ie: One side "Square Trim" but the other is NOT "Square Trim"                       
                        bool hasSquareTrim = ConnectionServices.HasSquareTrim(base.GetRootLogicalConnection(physicalConnection));

                        //Get the Question answers
                        string category = physicalConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.Category).ToString();
                        double boundedObjectThickness = 0.0, boundingObjectThickness = 0.0;
                        ConnectionServices.GetPhysicalConnectionPartsThickness(physicalConnection, out boundedObjectThickness, out boundingObjectThickness);
                        //If the returned Mounting Angle is greater than 90 degrees, subtract from 180 to get the smaller angle
                        //Start with the Plate by Plate cases
                        if (base.BoundedObject is PlatePartBase && base.BoundingObject is PlatePartBase)
                        {
                            choices = ConnectionServices.GetSelectionsForPlateToPlate(base.BoundedObject, category, boundedObjectThickness, hasSquareTrim);
                        }
                        else if (base.BoundedObject is PlatePartBase && base.BoundingObject is StiffenerPartBase) //this is a plate to profile case
                        {
                            choices = ConnectionServices.GetSelectionsForPlateToStiffener(base.BoundedObject, category, boundedObjectThickness, hasSquareTrim, physicalConnection.GetConnectionAngle());
                        }
                        else if (base.BoundedObject is StiffenerPartBase && base.BoundingObject is StiffenerPartBase)  //this is a profile to profile case
                        {
                            choices = ConnectionServices.GetSelectionsForStiffenerToStiffener(base.BoundedObject, base.BoundedPort, base.BoundingPort, physicalConnection.GetConnectionAngle());
                        }
                        else if (base.BoundedObject is StiffenerPartBase && base.BoundingObject is PlatePartBase) // Possibly need to consider long point / short point issue (stiffener attachment)
                        {
                            try
                            {
                                choices = ConnectionServices.GetSelectionsForStiffenerToPlate(base.BoundedObject, base.BoundingObject, category, boundedObjectThickness, physicalConnection.GetConnectionAngle(), hasSquareTrim);
                            }
                            catch (CmnException ex)
                            {
                                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ConnectionsResourceIdentifiers.ErrInvalidCategory,
                              ex.Message + this.ToString() + "."));
                            }
                        }
                        else
                        {
                            switch (category)
                            {
                                case DetailingCustomAssembliesConstants.Normal:
                                    {
                                        double opening = boundedObjectThickness * Math.Abs(Math.Tan(mountingAngleCompliment));
                                        if (opening < Math3d.FitTolerance * 3)
                                        {
                                            choices.Add(DetailingCustomAssembliesConstants.FilletWeld1);
                                        }
                                        else if (opening >= Math3d.FitTolerance * 3)
                                        {
                                            choices.Add(DetailingCustomAssembliesConstants.FilletWeld2);
                                        }
                                    }
                                    break;
                                case DetailingCustomAssembliesConstants.Full:
                                    {
                                        if (boundedObjectThickness < Math3d.FitTolerance * 4)
                                        {
                                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldChill);
                                        }
                                        else
                                        {
                                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldV);
                                        }
                                    }
                                    break;
                                case DetailingCustomAssembliesConstants.Deep:
                                    {
                                        if (boundedObjectThickness > StructHelper.DISTTOL * 25)
                                        {
                                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldK);
                                        }
                                        else
                                        {
                                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                                        }
                                    }
                                    break;
                                case DetailingCustomAssembliesConstants.Chain:
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.ChainWeld);
                                    }
                                    break;
                                case DetailingCustomAssembliesConstants.Staggered:
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.StaggeredWeld);
                                    }
                                    break;
                                case DetailingCustomAssembliesConstants.OneSidedBevel:
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldY);
                                        if (boundedObjectThickness <= Math3d.FitTolerance * 25)
                                        {
                                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldV);
                                        }
                                    }
                                    break;
                                case DetailingCustomAssembliesConstants.TwoSidedBevel:
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldK);
                                        if (boundedObjectThickness > Math3d.FitTolerance * 25)
                                        {
                                            choices.Add(DetailingCustomAssembliesConstants.TeeWeldX);
                                        }
                                    }
                                    break;
                                case DetailingCustomAssembliesConstants.Chill:
                                    {
                                        choices.Add(DetailingCustomAssembliesConstants.TeeWeldChill);
                                    }
                                    break;
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    //may be there could be specific ToDo record is created before, so do not override it.
                    if (base.ToDoListMessage == null)
                    {                       
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(ConnectionsResourceIdentifiers.ToDoPhysicalConnectionSelections,
                        "Unexpected error while evaluating the {0}"), this.ToString()));
                    }
                }
                return new ReadOnlyCollection<string>(choices);
            }
        }
        #endregion
    }
}
