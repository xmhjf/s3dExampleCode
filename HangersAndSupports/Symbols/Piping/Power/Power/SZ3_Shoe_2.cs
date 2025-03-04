//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SZ3_Shoe_2.cs
//    Power,Ingr.SP3D.Content.Support.Symbols.SZ3_Shoe_2
//   Author       :  Rajeswari
//   Creation Date:  12-Dec-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   12-Dec-2012   Rajeswari CR222282 .Net HS_Power project creation
//	 25-Mar-2013    Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
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
    [VariableOutputs]
    [CacheOption(CacheOptionType.NonCached)]
    [SymbolVersion("1.0.0.0")]
    public class SZ3_Shoe_2 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Power,Ingr.SP3D.Content.Support.Symbols.SZ3_Shoe_2"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(3, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(4, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(5, "T", "T", 0.999999)]
        public InputDouble m_dT;
        [InputDouble(6, "Dw", "Dw", 0.999999)]
        public InputDouble m_dDw;
        [InputString(7, "InputBomDesc", "InputBomDesc", "No Value")]
        public InputString m_sInputBomDesc;
        [InputDouble(8, "Angle", "Angle", 0.999999)]
        public InputDouble m_dAngle;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("RIGHT", "RIGHT")]
        public AspectDefinition m_Symbolic;

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
             
                Double B = m_dB.Value;
                Double A = m_dA.Value;
                Double H = m_dH.Value;
                Double T = m_dT.Value;
                Double Dw = m_dDw.Value;
                Double angle = m_dAngle.Value;

                if (Dw == 0 && H == 0  && T == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidDwInvalidHAndT, "Outer Diameter,H and Thickness cannot be zero"));
                    return;
                }
                if (A <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidAGTZero, "A should be greater than zero"));
                    return;
                }
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidB, "B should be greater than zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                string[] outputString = new string[3];

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -H);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d top = (Projection3d)symbolGeometryHelper.CreateBox(null, T, A, B);
                m_Symbolic.Outputs["TOP"] = top;
                outputString[0] = "TOP";

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -(B / 2 - T / 2), -H + T);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d left = (Projection3d)symbolGeometryHelper.CreateBox(null, H - Dw / 2 - T, A, T);
                m_Symbolic.Outputs["LEFT"] = left;
                outputString[1] = "LEFT";

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, B / 2 - T / 2, -H + T);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d right = (Projection3d)symbolGeometryHelper.CreateBox(null, H - Dw / 2 - T, A, T);
                m_Symbolic.Outputs["RIGHT"] = right;
                outputString[2] = "RIGHT";

                if (angle > 0)
                {
                    Matrix4X4 matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(angle, new Vector(0, 1, 0));

                    for (int i = 0; i < outputString.Length; i++)
                        if (outputString[i] != null)
                        {
                            Geometry3d transformObject = (Geometry3d)m_Symbolic.Outputs[outputString[i]];
                            transformObject.Transform(matrix);
                        }
                }
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SZ3_Shoe_2"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                string inputBomDescripttion = (string)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJUAHgrPowerBomDesc", "InputBomDesc")).PropValue;

                if (inputBomDescripttion == "None")
                {
                    bomDescription = "";
                }
                else
                {
                    if (inputBomDescripttion == null)
                    {
                        bomDescription = catalogPart.PartDescription;
                    }
                    else
                    {
                        bomDescription = inputBomDescripttion.Trim();
                    }
                }
                return bomDescription;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of SZ3_Shoe_2"));
                }
                return "";
            }
        }
        #endregion
    }
}
