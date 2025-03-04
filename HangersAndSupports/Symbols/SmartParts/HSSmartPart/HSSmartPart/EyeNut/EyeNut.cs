//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   EyeNut.cs
//    HSSmartPart,Ingr.SP3D.Content.Support.Symbols.EyeNut
//   Author       :  Vijay  
//   Creation Date:  16-01-2013
//   Description: CR-CP-222483 .Net EyeNut project creation   

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   16-01-2013     Vijay   CR-CP-222483 .Net EyeNut project creation   
//   25/Mar/2013    Vijay   DI-CP-228142  Modify the error handling for delivered H&S symbols
//   09/Feb/2015    PVK     TR-CP-257909	JIMC Eye Nut toggle to incorrect placement point
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

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
    public class EyeNut : SmartPartComponentDefinition, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HSSmartPart,Ingr.SP3D.Content.Support.Symbols.EyeNut"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "RodDiameter", "RodDiameter", 0)]
        public InputDouble m_dRodDiameter;
        [InputDouble(3, "PinDiameter", "PinDiameter", 0)]
        public InputDouble m_dPinDiameter;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Definition of Additional Inputs"
        public override IEnumerable<Input> AdditionalInputs
        {
            get
            {
                int endIndex;
                List<Input> additionalInputs = new List<Input>();
                AddEyeNutInputs(4, out endIndex, additionalInputs);
                return additionalInputs;
            }
        }
        #endregion

        #region "Definition of Additional Outputs"
        public override IEnumerable<OutputDefinition> AdditionalOutputs(string aspectName)
        {
            List<OutputDefinition> additionalOutputs = new List<OutputDefinition>();
            if (aspectName == "Symbolic")
            {
                AddEyeNutOutputs(additionalOutputs);
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
                Double RodDiameter = m_dRodDiameter.Value;
                Double PinDiameter = m_dPinDiameter.Value;

                EyeNutInputs eyenut = LoadEyeNutData(4);

                if (base.ToDoListMessage != null)
                    if (base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                        return;

                Port port1 = new Port(OccurrenceConnection, part, "RodEnd", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, eyenut.OverLength1 - eyenut.InnerLength2 + PinDiameter / 2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port2"] = port2;

                Port port3 = new Port(OccurrenceConnection, part, "Surface", new Position(0, 0, eyenut.OverLength1 + eyenut.Nut.ShapeLength), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port3"] = port3;

                if (PinDiameter > eyenut.InnerWidth2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrEyeNutPinDiameter, "Pin diameter too large: clashes with eye nut dimensions"));

                ArrayList eyeNutObjectCollection;
                Geometry3d geometry3d;
                Matrix4X4 matrix = new Matrix4X4();
                AddEyeNut(eyenut, m_PhysicalAspect.Outputs, "EyeNut", out eyeNutObjectCollection);
                matrix.Origin = new Position(0, 0, 0);
                foreach (string item in eyeNutObjectCollection)
                {
                    geometry3d = (Geometry3d)m_PhysicalAspect.Outputs[item];
                    geometry3d.Transform(matrix);
                }  
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of EyeNut"));
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, SmartPartLocalizer.GetString(SmartPartSymbolResourceIDs.ErrWeightCG, "Error in Weight CG of EyeNut"));
                    return;
                }
            }
        }
        #endregion
    }

}
