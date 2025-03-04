//-------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  FreeEndCutSelectorRule.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMFreeEndCutRules.dll
//  Original Class Name: ‘FreeEndCutSel’ in VB content
//
//Abstract
//	FreeEndCutSelectorRule is a .NET selector rule which selects the list of available items in the context of the FreeEndCut. 
//  This class subclasses from SelectorRule.
//
//Change History:
//  dd.mmm.yyyy    who    change description
//-------------------------------------------------------------------------------------------------------------------------------
using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Structure.Middle;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Selector for FreeEndCut, which selects the list of available items in the context of the FreeEndCut.
    /// </summary>
    [RuleVersion("1.0.0.0")]
    [DefaultLocalizer(EndCutsResourceIds.EndCutsResource, EndCutsResourceIds.EndCutsAssembly)]
    [RuleInterface(DetailingCustomAssembliesConstants.IASMFreeEndCutRules_FreeEndCutSel, DetailingCustomAssembliesConstants.IASMFreeEndCutRules_FreeEndCutSel)]
    public class FreeEndCutSelectorRule : SelectorRule
    {
        //=====================================================================================================
        //DefinitionName/ProgID of this symbol is "EndCuts,Ingr.SP3D.Content.Structure.FreeEndCutSelectorRule"
        //=====================================================================================================

        #region Selector questions
        [SelectorQuestionCodelist(100, DetailingCustomAssembliesConstants.EndCutType, DetailingCustomAssembliesConstants.EndCutType, DetailingCustomAssembliesConstants.EndCutTypeCodeList, (int)EndCutTypes.Welded)]
        public SelectorQuestionCodelist endCutTypeAnswer;
        [SelectorQuestionString(101, DetailingCustomAssembliesConstants.ChamferType, DetailingCustomAssembliesConstants.ChamferType, DetailingCustomAssembliesConstants.ChamferTypeNone)]
        public SelectorQuestionString chamferTypeAnswer;
        #endregion Selector questions

        #region Public override properties and methods

        /// <summary>
        /// Gets the list of named part items that are possible choices based on provided inputs.
        /// </summary>
        public override ReadOnlyCollection<string> Selections
        {
            get
            {
                Collection<string> choices = new Collection<string>();

                try
                {
                    //get the free end port on which this FreeEndCut will be created
                    FreeEndCut freeEndCut = (FreeEndCut)base.Occurrence;
                    IPort endPort = freeEndCut.EndPort;

                    //if it is member part use following part because this handles FreeEndCut migration if the member splits
                    if (endPort.Connectable is MemberPart)
                    {
                        choices.Add(DetailingCustomAssembliesConstants.MemberFreeEndCutPart);
                    }
                    else
                    {
                        choices.Add(DetailingCustomAssembliesConstants.StiffenerPartBaseFreeEndCutPart);
                    }
                }
                catch (Exception)
                {
                    //There could be specific ToDo records created before, so do not override it.
                    if (base.ToDoListMessage == null)
                    {
                        base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(EndCutsResourceIds.ToDoSelections,
                            "Unexpected error while getting the selections from {0}"), this.ToString()));
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
                    //if the FreeEndCut is because of convex manual profile knuckle then it returns the EndCutType based on the selection otherwise returns the default value "Snip"
                    EndCutTypes endCutType = EndCutTypes.Snip;

                    //retrieve the end port connectable from the FreeEndCut
                    FreeEndCut freeEndCut = (FreeEndCut)base.Occurrence;
                    StiffenerPartBase stiffenerPartBase = freeEndCut.EndPort.Connectable as StiffenerPartBase;
                    if (stiffenerPartBase != null)
                    {
                        //get stiffener connection method
                        bool atStart = freeEndCut.AtStart;
                        int connectionMethod = atStart ? stiffenerPartBase.StartConnectionMethod : stiffenerPartBase.EndConnectionMethod;
                        if (connectionMethod > (int)StiffenerConnectionMethod.ByRule)
                        {
                            switch (connectionMethod)
                            {
                                case (int)StiffenerConnectionMethod.Welded:
                                    endCutType = EndCutTypes.Welded;
                                    break;
                                case (int)StiffenerConnectionMethod.Bracketed:
                                    endCutType = EndCutTypes.Bracketed;
                                    break;
                                case (int)StiffenerConnectionMethod.Cutback:
                                    endCutType = EndCutTypes.Cutback;
                                    break;
                            }
                        }
                        else if (this.IsFromExtendedConvexKnuckle(stiffenerPartBase, atStart))
                        {
                            //Needs welded end cut for stiffener part height less than or equal swage height,
                            //otherwise welded up to swage height and from there a sniped web cut.
                            CrossSection crossSection = stiffenerPartBase.CrossSection;
                            double height = crossSection.Depth;
                            if (height < StructHelper.DISTTOL)
                            {
                                height = DetailingCustomAssembliesServices.GetWebLength(crossSection);
                            }

                            //Assumption - Swage is considered to be high enough to have a completely welded web cut
                            double swageHeight = height; //to get SwageHeight need an API from detailing. Refer CR-213420 and TR-215133 
                            if (StructHelper.AreEqual(height, swageHeight, StructHelper.MEDIUMDISTTOL))
                            {
                                //end cut needs PhysicalConnection, to be welded to plate.
                                endCutType = EndCutTypes.Welded;
                            }
                        }
                    }

                    this.endCutTypeAnswer.Value = (int)endCutType;
                }
                else if (selectorQuestionName == DetailingCustomAssembliesConstants.ChamferType)
                {
                    this.chamferTypeAnswer.Value = DetailingCustomAssembliesConstants.ChamferTypeNone;
                }
            }
            catch (Exception)
            {
                //There could be specific ToDo records created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(EndCutsResourceIds.ToDoDefaultAnswer,
                        "Unexpected error while setting the default answer in {0}"), this.ToString()));
                }
            }
        }

        #endregion Public override properties and methods

        #region Private methods

        /// <summary>
        /// Gets a value indicating whether the FreeEndCut is because of convex manual profile knuckle or not.
        /// </summary>        
        /// <param name="stiffenerPartBase">The FreeEndCut's end port connectable.</param>
        /// <param name="atStart">The Boolean to indicate whether the free end cut is placed at the start or end of a profile part.</param>
        /// <returns>True if the FreeEndCut is because of convex manual profile knuckle; otherwise, false.</returns>
        private bool IsFromExtendedConvexKnuckle(StiffenerPartBase stiffenerPartBase, bool atStart)
        {
            bool isFromExtendedConvexKnuckle = false;
            ContextTypes contextType = atStart ? ContextTypes.Base : ContextTypes.Offset;
            ProfileKnuckle profileKnuckle = stiffenerPartBase.GetKnuckleRelatedToExtendedStiffenerEnd(contextType);
            if (profileKnuckle != null)
            {
                //Check the knuckle is convex manual or not
                if (profileKnuckle.IsManual && profileKnuckle.IsConvex)
                {
                    isFromExtendedConvexKnuckle = true;
                }
            }

            return isFromExtendedConvexKnuckle;
        }

        #endregion Private methods
    }
}