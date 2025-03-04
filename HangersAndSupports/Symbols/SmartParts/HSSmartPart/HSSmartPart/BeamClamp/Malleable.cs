//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Malleable.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Malleable
//   Author       :   Vijay
//   Creation Date:  14.Feb.2013 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   14.Feb.2013    Vijay     CR-CP-222474 Initial Creation    
//	 25.Mar.2013	Vijay 	  DI-CP-228142  Modify the error handling for delivered H&S symbols
//   10.Jun.2015    PVK	      TR-CP-274155	SmartPart TDL Errors should be corrected.
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

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Malleable : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.Malleable"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;

        #region "Definition of Additional Inputs"

        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                //Add Properties
                additionalInputs.Add(new InputString(2, "Size", "Size", "", false));
                additionalInputs.Add(new InputDouble(3, "RodDiameter", "RodDiameter", 0, false));

                //Beam Clamp Inputs
                AddMalleableBeamClampInputs(4, out endIndex, additionalInputs);

                //Struct Dimensions
                additionalInputs.Add(new InputDouble(++endIndex, "ClampThickness", "ClampThickness", 0, false));
                additionalInputs.Add(new InputDouble(++endIndex, "StructWidth", "StructWidth", 0, false));
                return additionalInputs;
            }
        }
        #endregion

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Structure", "Structure")]
        [SymbolOutput("RodEnd", "RodEnd")]
        [SymbolOutput("Pin", "Pin")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Definition of Additional Outputs"

        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddMalleableBeamClampOutputs(additionalOutputs);
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
                int endIndex, startIndex;
                Matrix4X4 matrix = new Matrix4X4();
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //Load attributes
                String size = GetStringInputValue(2);
                Double rodDiameter = GetDoubleInputValue(3);

                //Load Malleable Beam Clamp Attributes
                MalleableBeamClampInputs malleableBeamClampInputs = LoadMaleableBCData(4, out endIndex);
                startIndex = endIndex;

                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                //Load Struct Attributes
                Double flangeThickness = GetDoubleInputValue(++startIndex);
                Double flangeWidth = GetDoubleInputValue(++startIndex);

                NutInputs nut = new NutInputs();
                EyeNutInputs eyeNut = new EyeNutInputs();
                //Load Nut Attributes from Query
                if (malleableBeamClampInputs.NutShape != "")
                    nut = LoadNutDataByQuery(malleableBeamClampInputs.NutShape, 1);

                //Load Eyenut Attributes from Query
                if (malleableBeamClampInputs.EyeNutShape != "")
                    eyeNut = LoadEyeNutDataByQuery(malleableBeamClampInputs.EyeNutShape);

                //Error and Warnings
                if (malleableBeamClampInputs.Depth <= 0)
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrMalleableDepth, "Depth should not be less than or equal to zero");
                    return;
                }

                if (malleableBeamClampInputs.TopWidth <= 0)
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrMalleableTopWidth, "TopWidth should not be less than or equal to zero");
                    return;
                }

                if (malleableBeamClampInputs.Thickness <= 0)
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrMalleableThickness, "Thickness should not be less than or equal to zero");
                    return;
                }

                if (flangeThickness <= 0)
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrMalleableClamp, "Clamp Thickness should not be less than or equal to zero");
                    return;
                }

                if (flangeWidth <= 0)
                {
                    RaiseSmartPartTDLError(SmartPartSymbolResourceIDs.ErrMalleableStructureWidth, "Structure Width should not be less than or equal to zero");
                    return;
                }

                //For Pin Raise the Warning
                if ((malleableBeamClampInputs.Pin1Diameter != 0) && (malleableBeamClampInputs.Pin1Length != 0))
                {
                    if (malleableBeamClampInputs.Pin1Length < flangeWidth + 2 * malleableBeamClampInputs.Thickness)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrMalleablePin1Length, "Pin1Length should be more than FlangeWidth + 2* Thickness. Resetting value to FlangeWidth + 4* Thickness"));
                        malleableBeamClampInputs.Pin1Length = flangeWidth + 4 * malleableBeamClampInputs.Thickness;
                    }
                }

                if ((malleableBeamClampInputs.Pin1Diameter != 0) && (malleableBeamClampInputs.Pin1Length != 0))
                {
                    if (malleableBeamClampInputs.Pin2Diameter > malleableBeamClampInputs.Thickness)
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrMalleablePin2Diameter, "Pin2 Diameter should not be more than the Thickness of the Malleable Beam Clamp"));

                    if (malleableBeamClampInputs.Pin2Length < malleableBeamClampInputs.Depth)
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrMalleablePin2Length, "Pin2 Length should be more than the Depth of the Malleable Beam Clamp"));
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Double zOffset;
                if (malleableBeamClampInputs.MalleableConfig == 3)
                    zOffset = eyeNut.InnerLength2;
                else
                    zOffset = 0;

                Double tempL1, tempL2;
                tempL1 = malleableBeamClampInputs.RodTakeOut - eyeNut.InnerLength2 + malleableBeamClampInputs.Pin2Diameter / 2;
                tempL2 = eyeNut.InnerLength2 - malleableBeamClampInputs.Pin2Diameter / 2 - malleableBeamClampInputs.OverLength;

                Port port1 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_PhysicalAspect.Outputs["Structure"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "RodEnd", new Position(0, tempL2 * Math.Sin(malleableBeamClampInputs.Angle), -(tempL1 + tempL2 * Math.Cos(malleableBeamClampInputs.Angle))), new Vector(1, 0, 0), new Vector(0, Math.Sin(malleableBeamClampInputs.Angle), -Math.Cos(malleableBeamClampInputs.Angle)));
                m_PhysicalAspect.Outputs["RodEnd"] = port2;

                Port port3 = new Port(OccurrenceConnection, part, "Pin", new Position(0, 0, -malleableBeamClampInputs.RodTakeOut - malleableBeamClampInputs.Pin2Diameter / 2 + zOffset), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_PhysicalAspect.Outputs["Pin"] = port3;

                //Add Malleable Beam Clamp

                matrix = new Matrix4X4();
                matrix.Origin = new Position(0, 0, 0);
                AddMalleablBeamClamp(malleableBeamClampInputs, flangeThickness, flangeWidth, matrix, m_PhysicalAspect.Outputs, "MalleableBC");

            }
            catch (SmartPartSymbolException hgrEx)
            {
                throw hgrEx;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Malleable"));
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

                Part CatalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];

                Double weight, cogX, cogY, cogZ;
                try
                {
                    weight = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryWeight")).PropValue;
                }
                catch
                {
                    weight = 0;
                }
                //Center of Gravity
                try
                {
                    cogX = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogX")).PropValue;
                }
                catch
                {
                    cogX = 0;
                }
                try
                {
                    cogY = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogY")).PropValue;
                }
                catch
                {
                    cogY = 0;
                }
                try
                {
                    cogZ = (double)((PropertyValueDouble)CatalogPart.GetPropertyValue("IJCatalogWtAndCG", "DryCogZ")).PropValue;
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of Malleable"));
                    return;
                }
            }
        }
        #endregion
    }

}
