//-----------------------------------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  FreeEndCutDefinition.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMFreeEndCutRules.dll
//  Original Class Name: ‘FreeEndCutDef’ and ‘MbrFreeEndCutDef’ in VB content
//
//Abstract:
//  FreeEndCutDefinition is a .NET custom assembly definition which creates a web cut and optionally a top and bottom flange cut at the end of a ProfilePart 
//  (i.e., MemberPart, StiffenerPart, StandAloneStiffenerPart or EdgeReinforcementPart). 
//  This class subclasses from FreeEndCutCustomAssemblyDefinition.
//
//Change History:
//  dd.mmm.yyyy    who    change description
//-----------------------------------------------------------------------------------------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Implementation of FreeEndCut .NET custom assembly definition class.
    /// FreeEndCutDefinition is a .NET custom assembly definition which creates a web cut and optionally a top and bottom flange cut at the end of a ProfilePart 
    /// (i.e., MemberPart, StiffenerPart, StandAloneStiffenerPart or EdgeReinforcementPart).
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    [DefaultLocalizer(EndCutsResourceIds.EndCutsResource, EndCutsResourceIds.EndCutsAssembly)]
    public class FreeEndCutDefinition : FreeEndCutCustomAssemblyDefinition
    {
        //===================================================================================================
        //DefinitionName/ProgID of this symbol is "EndCuts,Ingr.SP3D.Content.Structure.FreeEndCutDefinition"
        //===================================================================================================
        #region Private members
        private const string FreeEndWebCut = "FreeEndWebCut";
        private const string FreeEndTopFlangeCut = "FreeEndTopFlangeCut";
        private const string FreeEndBottomFlangeCut = "FreeEndBottomFlangeCut";
        #endregion Private members

        #region Definitions of assembly outputs
        [AssemblyOutput(1, FreeEndWebCut)]
        public AssemblyOutput webCutAssemblyOutput;
        [AssemblyOutput(2, FreeEndTopFlangeCut)]
        public AssemblyOutput topFlangeCutAssemblyOutput;
        [AssemblyOutput(3, FreeEndBottomFlangeCut)]
        public AssemblyOutput bottomFlangeCutAssemblyOutput;
        #endregion Definitions of assembly outputs

        #region Public override properties and methods

        /// <summary>
        /// Construct and re-evaluate the custom assembly outputs.
        /// Here we will do the following:
        /// 1. Decide which assembly outputs are needed. 
        /// 2. Create the ones which are needed, delete which are not needed now.
        /// 3. Sets the answer of the question on assembly outputs.
        /// </summary>
        public override void EvaluateAssembly()
        {
            try
            {
                //Construct the assembly output objects
                //First the WebCut feature                
                FreeEndCut freeEndCut = (FreeEndCut)base.Occurrence;
                Feature webCutFeature = null, topFlangeCutFeature = null, bottomFlangeCutFeature = null;

                //get the required assembly outputs for this definition
                Dictionary<string, bool> requiredAssemblyOutputs = this.RequiredAssemblyOutputs;
                if (requiredAssemblyOutputs[FreeEndWebCut])
                {
                    //flag to know that the WebCut is in the modification stage
                    bool isWebCutInModificationStage = true;

                    //Only construct the WebCut if not generated yet and add it is as output
                    if (this.webCutAssemblyOutput.Output == null)
                    {
                        isWebCutInModificationStage = false; //we are in the initial creation stage

                        //get the WebCut feature root selector 
                        string webCutFeatureRootSelector = this.GetWebCutFeatureRootSelector();

                        //create the WebCut feature
                        webCutFeature = base.CreateWebCut(webCutFeatureRootSelector);
                        if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        {
                            //ToDo list is created with error type hence stop computation
                            return;
                        }
                        this.webCutAssemblyOutput.Output = webCutFeature;
                    }
                    else
                    {
                        webCutFeature = (Feature)this.webCutAssemblyOutput.Output;
                    }

                    //Set the parent's (FreeEndCut) answer of EndCutType question on the WebCut feature.
                    base.SetAnswer(webCutFeature, DetailingCustomAssembliesConstants.EndCutType);

                    if (isWebCutInModificationStage)
                    {
                        //Updates Webcut so that it is re-evaluate. 
                        //Re-evaluate the WebCut and its related FlangeCut when a question-answer on FreeEndCut is changed.
                        base.UpdateFeature(webCutFeature);
                    }
                }
                else
                {
                    //if WebCut is not required now, delete it if it has been previously created
                    if (this.webCutAssemblyOutput.Output != null)
                    {
                        this.webCutAssemblyOutput.Delete();
                    }
                }

                //get the bottom flange question depending upon the FreeEndCut is on MemberPart or not 
                string bottomFlangeQuestion = (base.IsMemberFreeEndCut) ? DetailingCustomAssembliesConstants.BottomFlange : DetailingCustomAssembliesConstants.TheBottomFlange;

                //Now construct the top flange cut if required.                
                if (requiredAssemblyOutputs[FreeEndTopFlangeCut])
                {
                    //Only construct the top flange cut if not generated yet and add it as output
                    if (this.topFlangeCutAssemblyOutput.Output == null)
                    {
                        //get the FlangeCut feature root selector 
                        string flangeCutFeatureRootSelector = this.GetFlangeCutFeatureRootSelector();

                        //create the top FlangeCut feature
                        topFlangeCutFeature = base.CreateFlangeCut(webCutFeature, EndCutType.FlangeCutTop, flangeCutFeatureRootSelector);
                        if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        {
                            //ToDo list is created with error type hence stop computation
                            return;
                        }
                        this.topFlangeCutAssemblyOutput.Output = topFlangeCutFeature;
                    }
                    else
                    {
                        topFlangeCutFeature = (Feature)this.topFlangeCutAssemblyOutput.Output;
                    }

                    //Set the parent's (FreeEndCut) of EndCutType question on the top flange cut feature.
                    base.SetAnswer(topFlangeCutFeature, DetailingCustomAssembliesConstants.EndCutType);

                    //Set the answer(AnswerNo) of BottomFlange/TheBottomFlange question on the top flange cut feature.
                    base.SetAnswer(topFlangeCutFeature, bottomFlangeQuestion, Answer.No);
                }
                else
                {
                    //if top flange cut is not required now, delete it if it has been previously created
                    if (this.topFlangeCutAssemblyOutput.Output != null)
                    {
                        this.topFlangeCutAssemblyOutput.Delete();
                    }
                }

                //Now construct the bottom flange cut if required.
                if (requiredAssemblyOutputs[FreeEndBottomFlangeCut])
                {
                    //Only construct the bottom flange cut if not generated yet and add it as output
                    if (this.bottomFlangeCutAssemblyOutput.Output == null)
                    {
                        //get the FlangeCut feature root selector 
                        string flangeCutFeatureRootSelector = this.GetFlangeCutFeatureRootSelector();

                        //create the bottom FlangeCut feature
                        bottomFlangeCutFeature = base.CreateFlangeCut(webCutFeature, EndCutType.FlangeCutBottom, flangeCutFeatureRootSelector);
                        if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        {
                            //ToDo list is created with error type hence stop computation
                            return;
                        }
                        this.bottomFlangeCutAssemblyOutput.Output = bottomFlangeCutFeature;
                    }
                    else
                    {
                        bottomFlangeCutFeature = (Feature)this.bottomFlangeCutAssemblyOutput.Output;
                    }

                    //Set the parent's (FreeEndCut) of EndCutType question on the bottom flange cut feature.
                    base.SetAnswer(bottomFlangeCutFeature, DetailingCustomAssembliesConstants.EndCutType);

                    //Set the answer(AnswerYes) of BottomFlange/TheBottomFlange question on the bottom flange cut feature.
                    base.SetAnswer(bottomFlangeCutFeature, bottomFlangeQuestion, Answer.Yes);
                }
                else
                {
                    //if bottom flange cut is not required now, delete it if it has been previously created
                    if (this.bottomFlangeCutAssemblyOutput.Output != null)
                    {
                        this.bottomFlangeCutAssemblyOutput.Delete();
                    }
                }
            }
            catch (Exception)
            {
                //may be there could be specific ToDo record is created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(EndCutsResourceIds.ToDoEvaluateAssembly,
                            "Unexpected error while evaluating the custom assembly of {0}"), this.ToString()));
                }
            }
        }

        /// <summary>
        /// Re-evaluate the FreeEndCut due to split of member.
        /// The port which created the FreeEndCut will be replaced by the closest port of, one of the replacing members.
        /// </summary>
        public override void InputsReplaced()
        {
            try
            {
                //select the closest one from the replacing list and replace input relation with the selected one.
                FreeEndCut freeEndCut = (FreeEndCut)base.Occurrence;
                IPortGeometry endPort = (IPortGeometry)freeEndCut.EndPort;
                IPoint originalEndPortPointGeometry = endPort.Geometry as IPoint;
                if (originalEndPortPointGeometry == null)
                {
                    string message = base.GetString(EndCutsResourceIds.ToDoPortGeometryNotPoint,
                        "The existing port geometry must be point for replacing the inputs.");
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, message);
                    return;
                }

                IPort replacingEndPort = base.GetClosestReplacingPort((IPortGeometry)freeEndCut.EndPort, originalEndPortPointGeometry);
                if (replacingEndPort != null)
                {
                    //TR-218172 FreeEndCuts on a standard member recreated in split migration 
                    new FreeEndCut((ISystem)replacingEndPort.Connectable, replacingEndPort, DetailingCustomAssembliesConstants.MbrRootFreeEndCuts);
                }
            }
            catch (Exception)
            {
                //may be there could be specific ToDo record is created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(EndCutsResourceIds.ToDoInputsReplaced,
                        "Unexpected error while inputs replaced the custom assembly of {0}"), this.ToString()));                    
                }                
            }
        }

        #endregion Public override properties and methods

        #region Private methods

        /// <summary>
        /// Gets all needed assembly outputs for the current configuration of the FreeEndCut.
        /// </summary>
        /// <returns>This returns the needed outputs of the FreeEndCut.</returns>
        private Dictionary<string, bool> RequiredAssemblyOutputs
        {
            get
            {
                //Dictionary which holds the data of the assembly output name and the Boolean which indicates
                //whether the corresponding assembly output is needed or not.
                Dictionary<string, bool> requiredAssemblyOutputs = new Dictionary<string, bool>();

                //for this definition WebCut feature always be created, so add to the needed outputs directly
                requiredAssemblyOutputs.Add(FreeEndWebCut, true);

                //add the top and bottom flange as not required initially and later update it accordingly
                requiredAssemblyOutputs.Add(FreeEndTopFlangeCut, false);
                requiredAssemblyOutputs.Add(FreeEndBottomFlangeCut, false);

                FreeEndCut freeEndCut = (FreeEndCut)base.Occurrence;

                //if it is a MemberPart FreeEndCut and the answer of EndCutType question is Welded, flange cuts are not required
                PropertyValue endCutTypePropertyValue = freeEndCut.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.EndCutType);
                int endCutType = (endCutTypePropertyValue != null) ? ((PropertyValueCodelist)endCutTypePropertyValue).PropValue : 0;

                if (base.IsMemberFreeEndCut && endCutType == (int)EndCutTypes.Welded)
                {
                    return requiredAssemblyOutputs;
                }

                //get the section type of the profile part in which FreeEndCut being created
                string sectionType = ((ProfilePart)freeEndCut.EndPort.Connectable).SectionType;

                //top flange is required if the section type has top flange
                if (DetailingCustomAssembliesServices.HasTopFlange(sectionType))
                {
                    requiredAssemblyOutputs[FreeEndTopFlangeCut] = true;
                }

                //bottom flange is required if the section type has bottom flange
                if (DetailingCustomAssembliesServices.HasBottomFlange(sectionType))
                {
                    requiredAssemblyOutputs[FreeEndBottomFlangeCut] = true;
                }

                return requiredAssemblyOutputs;
            }
        }

        /// <summary>
        /// Gets the WebCut feature root selector.
        /// </summary>
        /// <returns>The root selector for WebCut feature.</returns>
        private string GetWebCutFeatureRootSelector()
        {
            string rootSelector;
            if (base.IsMemberFreeEndCut)
            {
                //Check if newly added part class exist in catalog, 
                //thereby we can ensure that user has updated the catalog
                //with new part class needed for new functionality, otherwise point to old rules
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                rootSelector = catalogStructHelper.DoesPartOrPartClassExist(DetailingCustomAssembliesConstants.MbrFreeWebSel) ?
                    DetailingCustomAssembliesConstants.MbrFreeWebSel :
                    DetailingCustomAssembliesConstants.MBR_FreeWebCut;
            }
            else //FreeEndCut on StiffenerPart or StandAloneStiffenerPart or EdgeReinforcementPart
            {
                //get EndCutType answer
                FreeEndCut freeEndCut = (FreeEndCut)base.Occurrence;

                PropertyValue endCutTypePropertyValue = freeEndCut.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.EndCutType);
                int endCutType = (endCutTypePropertyValue != null) ? ((PropertyValueCodelist)endCutTypePropertyValue).PropValue : 0;

                rootSelector = (base.IsFromExtendedConvexKnuckle && endCutType == (int)EndCutTypes.Welded) ?
                    DetailingCustomAssembliesConstants.WebCutsLongBox : //WebCut is because of 'Extend' Mfg. option and convex knuckle, also this is welded endcut so use long box end cuts selection
                    DetailingCustomAssembliesConstants.WebCuts; //regular FreeEndCut                    
            }

            return rootSelector;
        }

        /// <summary>
        /// Gets the FlangeCut feature root selector.
        /// </summary>
        /// <returns>The root selector for FlangeCut feature.</returns>
        private string GetFlangeCutFeatureRootSelector()
        {
            string rootSelector;
            if (base.IsMemberFreeEndCut)
            {
                //Check if newly added part class exist in catalog, 
                //thereby we can ensure that user has updated the catalog
                //with new part class needed for new functionality, otherwise point to old rules
                CatalogStructHelper catalogStructHelper = new CatalogStructHelper();
                rootSelector = catalogStructHelper.DoesPartOrPartClassExist(DetailingCustomAssembliesConstants.MbrEndFlangeSel) ?
                    DetailingCustomAssembliesConstants.MbrEndFlangeSel :
                    DetailingCustomAssembliesConstants.MBR_EndFlangeSel;
            }
            else //FreeEndCut on StiffenerPart or StandAloneStiffenerPart or EdgeReinforcementPart
            {
                rootSelector = DetailingCustomAssembliesConstants.FlangeCuts;
            }

            return rootSelector;
        }

        #endregion Private methods
    }
}