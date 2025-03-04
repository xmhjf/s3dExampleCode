//-----------------------------------------------------------------------------------------------------------------------------------------
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  StiffenerEndToPlateFaceDefinition.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMAssyConRul.dll
//  Original Class Name: ‘StiffEndToPlateFaceDef’ in VB content
//
//Abstract
//	StiffenerEndToPlateFaceDefinition is a .NET custom assembly definition, which creates an assembly connection with four possible assembly outputs 
//  (web cut, top, bottom flange cut and CustomPlatePart) for stiffener bounded by plate.
//
//Change History:
//  dd.mmm.yyyy    who    change description
//-----------------------------------------------------------------------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Implementation of StiffenerEndToPlateFaceDefinition.
    /// Evaluates and creates AssemblyConnection outputs for stiffener bounded by plate case.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    [DefaultLocalizer(ConnectionsResourceIdentifiers.ConnectionsResource, ConnectionsResourceIdentifiers.ConnectionsAssembly)]
    public class StiffenerEndToPlateFaceDefinition : StiffenerAssemblyConnectionCustomAssemblyDefinition
    {
        //====================================================================================================================
        //DefinitionName/ProgID of this symbol is "Connections,Ingr.SP3D.Content.Structure.StiffenerEndToPlateFaceDefinition"
        //====================================================================================================================        
        #region Private members
        AssemblyConnection assemblyConnection = null;
        IPort boundedPort = null;
        TopologyPort boundingPort = null;
        private const string StiffenerEndToPlateFaceWebCut = "StiffEndToPlateFaceWebCut1";
        private const string StiffenerEndToPlateFaceTopFlangeCut = "StiffEndToPlateFaceFlangeCut1";
        private const string StiffenerEndToPlateFaceBottomFlangeCut = "StiffEndToPlateFaceFlangeCut2";
        private const string StiffenerEndToStiffenerFlangeBracket = "StiffEndToStiffFlangeBracket1";
        #endregion Private members

        #region Definitions of assembly outputs
        [AssemblyOutput(1, StiffenerEndToPlateFaceWebCut)]
        public AssemblyOutput webCutAssemblyOutput;

        [AssemblyOutput(2, StiffenerEndToPlateFaceTopFlangeCut)]
        public AssemblyOutput topFlangeCutAssemblyOutput;

        [AssemblyOutput(3, StiffenerEndToPlateFaceBottomFlangeCut)]
        public AssemblyOutput bottomFlangeCutAssemblyOutput;

        [AssemblyOutput(4, StiffenerEndToStiffenerFlangeBracket)]
        public AssemblyOutput customPlatePartAssemblyOutput;
        #endregion Definitions of assembly outputs

        #region Public override functions and methods

        /// <summary>
        /// Creates required AssemblyOutputs for this definition based on profile section and selector answers.
        /// This method is expected to be overridden by the inheriting class to construct and re-evaluate the custom assembly outputs.
        /// </summary>
        public override void EvaluateAssembly()
        {
            try
            {
                this.assemblyConnection = (AssemblyConnection)base.Occurrence;

                //Get the BoundingPorts and BoundedPorts from AssemblyConnection
                if (this.assemblyConnection.BoundedPorts.Count == 0)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ConnectionsResourceIdentifiers.ErrBoundedPortCount,
                        "No bounded ports associated with the AssemblyConnection."));

                    //ToDo list is created with error type hence stop computation
                    return;
                }

                Collection<IPort> boundingPorts = this.assemblyConnection.BoundingPorts;
                if (boundingPorts.Count != 1)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ConnectionsResourceIdentifiers.ErrBoundingPortCount,
                        "AssemblyConnection has more than one bounding port."));

                    //ToDo list is created with error type hence stop computation
                    return;
                }

                this.boundedPort = this.assemblyConnection.BoundedPorts[0];
                this.boundingPort = (TopologyPort)boundingPorts[0];

                //get the required assembly outputs for this definition
                Dictionary<string, bool> requiredAssemblyOutputs = this.RequiredAssemblyOutputs;
                Feature webCutFeature = null, topFlangeCutFeature = null, bottomFlangeCutFeature = null;

                //Construct the assembly output objects
                //First the WebCut feature
                if (requiredAssemblyOutputs[StiffenerEndToPlateFaceWebCut])
                {
                    //Only construct the WebCut if not generated yet and add it is as output
                    if (this.webCutAssemblyOutput.Output == null)
                    {
                        //get the WebCut feature root selector 
                        string webCutFeatureRootSelector = (base.IsLappedConnection) ? DetailingCustomAssembliesConstants.WebCutsLapped : //root selector for lapped connection
                                                                                       DetailingCustomAssembliesConstants.WebCuts; //default root selector
                        //create the WebCut feature
                        webCutFeature = base.CreateWebCut(webCutFeatureRootSelector);
                        if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        {
                            //ToDo list is created with error type hence stop computation
                            return;
                        }

                        //On Seam deletion, updates the corner features on the connectable of the
                        //AssemblyConnection's bounded port, which are nearer to the Seam.
                        this.UpdateCornerFeatureOnSeamDeletion();

                        this.webCutAssemblyOutput.Output = webCutFeature;
                    }
                    else
                    {
                        webCutFeature = (Feature)this.webCutAssemblyOutput.Output;
                    }

                    //Set the parent's (AssemblyConnection) answer of EndCutType question on the WebCut feature.
                    base.SetAnswer(webCutFeature, DetailingCustomAssembliesConstants.EndCutType);
                }
                else
                {
                    //if WebCut is not required now, delete it if it has been previously created
                    if (this.webCutAssemblyOutput.Output != null)
                    {
                        this.webCutAssemblyOutput.Delete();
                    }
                }

                //Now construct the top flange cut if required.
                if (requiredAssemblyOutputs[StiffenerEndToPlateFaceTopFlangeCut])
                {
                    //Only construct the top flange cut if not generated yet and add it as output
                    if (this.topFlangeCutAssemblyOutput.Output == null)
                    {
                        //get the FlangeCut feature root selector 
                        string flangeCutFeatureRootSelector = DetailingCustomAssembliesConstants.FlangeCuts;

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

                    //Set the parent's (AssemblyConnection) of EndCutType question on the top flange cut feature.
                    base.SetAnswer(topFlangeCutFeature, DetailingCustomAssembliesConstants.EndCutType);

                    //Set the answer(AnswerNo) of TheBottomFlange question on the top flange cut feature.
                    base.SetAnswer(topFlangeCutFeature, DetailingCustomAssembliesConstants.TheBottomFlange, Answer.No);
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
                if (requiredAssemblyOutputs[StiffenerEndToPlateFaceBottomFlangeCut])
                {
                    //Only construct the bottom flange cut if not generated yet and add it as output
                    if (this.bottomFlangeCutAssemblyOutput.Output == null)
                    {
                        //get the FlangeCut feature root selector 
                        string flangeCutFeatureRootSelector = DetailingCustomAssembliesConstants.FlangeCuts;

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

                    //Set the parent's (AssemblyConnection) of EndCutType question on the bottom flange cut feature.
                    base.SetAnswer(bottomFlangeCutFeature, DetailingCustomAssembliesConstants.EndCutType);

                    //Set the answer(AnswerYes) of TheBottomFlange question on the bottom flange cut feature.
                    base.SetAnswer(bottomFlangeCutFeature, DetailingCustomAssembliesConstants.TheBottomFlange, Answer.Yes);
                }
                else
                {
                    //if bottom flange cut is not required now, delete it if it has been previously created
                    if (this.bottomFlangeCutAssemblyOutput.Output != null)
                    {
                        this.bottomFlangeCutAssemblyOutput.Delete();
                    }
                }

                //Now construct the CustomPlatePart if required.                
                if (requiredAssemblyOutputs[StiffenerEndToStiffenerFlangeBracket])
                {
                    //Only construct the CustomPlatePart if not generated yet and add it as output
                    if (this.customPlatePartAssemblyOutput.Output == null)
                    {
                        //Get plate common inputs
                        Material material = new CatalogStructHelper().GetMaterial("Steel - Carbon", "A");
                        PlateCommonInputs plateCommonInputs = new PlateCommonInputs(PlateType.BracketPlate, material, 0.01, Continuity.Continuous, MoldedDirection.Above, Tightness.NonTight, 0, 0);

                        //create a custom plate part of type bracket
                        CustomPlatePart customPlatePart = null;
                        try
                        {
                            customPlatePart = base.CreateCustomPlatePart(plateCommonInputs);
                        }
                        catch (CmnException ex)
                        {
                            if (base.ToDoListMessage == null)
                            {
                                base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, ex.Message);

                                //ToDo list is created with error type hence stop computation
                                return;
                            }
                        }

                        if (customPlatePart != null)
                        {
                            //Set ThicknessOffset value based on the bounded stiffener cross-section
                            string sectionType = ((ICrossSection)this.boundedPort.Connectable).SectionType;
                            switch (sectionType)
                            {
                                case MarineSymbolConstants.FB:
                                case MarineSymbolConstants.SB:
                                case MarineSymbolConstants.T_XType:
                                case MarineSymbolConstants.TSType:
                                case MarineSymbolConstants.I:
                                case MarineSymbolConstants.ISType:
                                case MarineSymbolConstants.H:
                                case MarineSymbolConstants.BUT:
                                case MarineSymbolConstants.BUTL2:
                                case MarineSymbolConstants.BUTL3:
                                    //default ThicknessOffset is 0.0 on the CustomPlatePart
                                    break;
                                default:
                                    customPlatePart.ThicknessOffset = 0.005;
                                    break;
                            }

                            this.customPlatePartAssemblyOutput.Output = customPlatePart;
                        }
                        else
                        {
                            base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, base.GetString(ConnectionsResourceIdentifiers.ErrCustomPlatePart,
                                "Unexpected error while creating the CustomPlatePart."));
                        }
                    }
                }
                else
                {
                    //if CustomPlatePart is not required now, delete it if it has been previously created
                    if (this.customPlatePartAssemblyOutput.Output != null)
                    {
                        this.customPlatePartAssemblyOutput.Delete();
                    }
                }
            }
            catch (Exception)
            {
                //may be there could be specific ToDo record is created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(ConnectionsResourceIdentifiers.ToDoEvaluateAssembly,
                        "Unexpected error while evaluating the {0}"), this.ToString()));
                }
            }
        }

        #endregion Public override functions and methods

        #region Private methods

        /// <summary>
        /// Gets all needed assembly outputs for the current configuration of the AssemblyConnection.
        /// </summary>
        /// <returns>This returns the needed outputs of the AssemblyConnection.</returns>
        private Dictionary<string, bool> RequiredAssemblyOutputs
        {
            get
            {
                //Dictionary which holds the data of the assembly output name and the Boolean which indicates
                //whether the corresponding assembly output is needed or not.
                Dictionary<string, bool> requiredAssemblyOutputs = new Dictionary<string, bool>();

                //for this definition WebCut feature always be created, so add to the needed outputs directly
                requiredAssemblyOutputs.Add(StiffenerEndToPlateFaceWebCut, true);

                //add the top and bottom flange as not required initially and later update it accordingly
                requiredAssemblyOutputs.Add(StiffenerEndToPlateFaceTopFlangeCut, false);
                requiredAssemblyOutputs.Add(StiffenerEndToPlateFaceBottomFlangeCut, false);

                //get the information about top/bottom FlangeCut is required or not
                //FlangeCuts are not required for Split/Angled SplitEndToEndCase
                if (!this.IsSplitEndToEndCase)
                {
                    //get the section type of the bounded port connectable, in case it is StiffenerPartBase 
                    StiffenerPartBase boundedStiffener = this.boundedPort.Connectable as StiffenerPartBase;
                    if (boundedStiffener != null)
                    {
                        string sectionType = boundedStiffener.SectionType;

                        //top flange is required if the section type has top flange
                        if (DetailingCustomAssembliesServices.HasTopFlange(sectionType))
                        {
                            requiredAssemblyOutputs[StiffenerEndToPlateFaceTopFlangeCut] = true;
                        }

                        //bottom flange is required if the section type has bottom flange
                        if (DetailingCustomAssembliesServices.HasBottomFlange(sectionType))
                        {
                            requiredAssemblyOutputs[StiffenerEndToPlateFaceBottomFlangeCut] = true;
                        }
                    }
                }

                //get the information about CustomPlatePart is required or not
                requiredAssemblyOutputs.Add(StiffenerEndToStiffenerFlangeBracket, this.IsCustomPlatePartNeeded);

                return requiredAssemblyOutputs;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this assemblyConnection is SplitEndToEndCase or not.
        /// </summary>
        /// <value>True if this assemblyConnection is SplitEndToEndCase; otherwise, false.</value>
        private bool IsSplitEndToEndCase
        {
            get
            {
                bool isSplitEndToEndCase = false;

                //may be this assembly connection won't have SplitEndToEndCase question
                PropertyValue splitEndToEndCasePropertyValue = this.assemblyConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.SplitEndToEndCase);
                if (splitEndToEndCasePropertyValue != null)
                {
                    int splitEndToEndCase = ((PropertyValueCodelist)splitEndToEndCasePropertyValue).PropValue;
                    switch (splitEndToEndCase)
                    {
                        case (int)SplitEndCutTypes.AngleWebSquareFlange:
                        case (int)SplitEndCutTypes.AngleWebBevelFlange:
                        case (int)SplitEndCutTypes.AngleWebAngleFlange:
                        case (int)SplitEndCutTypes.DistanceWebDistanceFlange:
                        case (int)SplitEndCutTypes.OffsetWebOffsetFlange:
                            isSplitEndToEndCase = true;
                            break;
                        default:
                            break;
                    }
                }

                return isSplitEndToEndCase;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the CustomPlatePart is needed or not.
        /// </summary>
        /// <value>True if the CustomPlatePart is needed; otherwise, false.</value>
        private bool IsCustomPlatePartNeeded
        {
            get
            {
                //get the answer of PlaceBracket question from the occurence
                PropertyValue placeBracketPropertyValue = this.assemblyConnection.GetSelectionRuleAnswer(DetailingCustomAssembliesConstants.PlaceBracket);
                int placeBracket = (placeBracketPropertyValue != null) ? ((PropertyValueCodelist)placeBracketPropertyValue).PropValue : 0;
                return (placeBracket == (int)Answer.Yes) ? true : false;
            }
        }

        /// <summary>
        /// On Seam deletion, updates the corner features on the connectable of the AssemblyConnection's bounded port, which are nearer to the Seam.
        /// </summary>        
        private void UpdateCornerFeatureOnSeamDeletion()
        {
            //get the bounded connectable
            AssemblyConnection assemblyConnection = (AssemblyConnection)base.Occurrence;
            IPort boundedPort = assemblyConnection.BoundedPorts[0];
            IDetailable boundedConnectable = boundedPort.Connectable as IDetailable;

            //do nothing if the bounded object is not a detailable object
            if (boundedConnectable != null)
            {
                //get all corner features on the bounded port connectable
                ReadOnlyCollection<Feature> cornerFeatures = boundedConnectable.GetFeatures(FeatureType.Corner);

                //update the corner features with part name : LongScallopWithSeam
                foreach (Feature cornerFeature in cornerFeatures)
                {
                    string partName = cornerFeature.PartName;
                    if (partName == DetailingCustomAssembliesConstants.LongScallopWithSeam)
                    {
                        //Updates Corner feature so that it is re-evaluate. 
                        //Re-evaluate Corner feature when seam is deleted.
                        base.UpdateFeature(cornerFeature);
                    }
                }
            }
        }

        #endregion Private methods
    }
}
