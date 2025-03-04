//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SlidePlate.cs
//   HSSmartPart,Ingr.SP3D.Content.Support.Symbols.SlidePlate
//   Author       :  Hema
//   Creation Date:  9-Feb-2013
//   Description:    Converted SlidePlate Smartpart VB Project to C# .Net 

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   9-Feb-2013      Hema    CR-CP-222490 Converted SlidePlate Smartpart VB Project to C# .Net 
//   25/Mar/2013     Hema    DI-CP-228142  Modify the error handling for delivered H&S symbols
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    [CacheOption(CacheOptionType.Cached)]

    public class SlidePlate : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.SlidePlate"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputString(2, "GuideShape", "GuideShape", "No Value")]
        public InputString m_oGuideShape;

        #endregion

        #region "Definition of Additional Inputs"

        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddSlidePlateInputs(3, out endIndex, additionalInputs);
                return additionalInputs;
            }
        }

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Definition of Additional Outputs"

        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddSlidePlateOutputs(additionalOutputs);
                AddGuideOutputs(additionalOutputs);
            }
            return additionalOutputs;
        }

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;

                //Loading the SlidePlate Additional Inputs
                SlidePlateInputs slidePlate = LoadSlidePlateData(3);
                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;
                GuideInputs guide = new GuideInputs();

                string guideShape = m_oGuideShape.Value;
                double portOffset = 0;
                int j = 0;
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                if (guideShape != "0" && guideShape != " " && guideShape != "No Value" && guideShape.Trim() != "")
                {
                    guide = LoadGuideDataByQuery(guideShape);
                }

                double[] ZP1 = { slidePlate.ZPl1, slidePlate.ZPl2, slidePlate.ZPl3, slidePlate.ZPl4, slidePlate.ZPl5, slidePlate.ZPl6 };
                double[] thicknessP1 = { slidePlate.Thickness1, slidePlate.Thickness2, slidePlate.Thickness3, slidePlate.Thickness4, slidePlate.Thickness5, slidePlate.Thickness6 };

                for (int i = 0; i < ZP1.Length; i++)
                {
                    if (ZP1[i] > portOffset)
                    {
                        portOffset = ZP1[i];
                        j = i;
                    }
                }

                portOffset = portOffset + thicknessP1[j];

                if (guideShape == "0" || guideShape == " " || guideShape == "No Value" || guideShape.Trim() == "")
                {
                    portOffset = portOffset + slidePlate.ShoeHeight + slidePlate.PipeDia / 2;
                }
                else
                {
                    portOffset = portOffset + slidePlate.ShoeHeight + slidePlate.PipeDia / 2 + guide.Thickness1;
                }

                //Ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, portOffset), new Vector(0, 1, 0), new Vector(0, 0, -1));
                m_PhysicalAspect.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(0, 1, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port2"] = port2;

                //Add SlidePlate
                if (guideShape == "0" || guideShape == " " || guideShape == "No Value" || guideShape.Trim() == "")
                {
                    AddSlidePlate(slidePlate, matrix, m_PhysicalAspect.Outputs, "SlidePlate");
                }
                else
                {
                    matrix.Translate(new Vector(0, 0, guide.Thickness1));

                    AddSlidePlate(slidePlate, matrix, m_PhysicalAspect.Outputs, "SlidePlate");
                }
                // //Add the Guide
                if (guideShape != "0" && guideShape != " " && guideShape != "No Value" && guideShape.Trim() != "")
                {
                    matrix.SetIdentity();
                    matrix.Rotate(Math.PI / 2, new Vector(0, 0, 1));

                    AddGuide(guide, matrix, m_PhysicalAspect.Outputs, "Guide");
                }
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SlidePlate"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                ////System WCG Attributes

                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = 0;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
                }
                catch
                {
                    cogZ = 0;
                }

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of SlidePlate"));
                }
            }
        }
        #endregion
    }
}
