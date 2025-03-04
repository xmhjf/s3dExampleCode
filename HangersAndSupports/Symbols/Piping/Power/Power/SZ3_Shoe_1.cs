//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SZ3_Shoe_1.cs
//    Power,Ingr.SP3D.Content.Support.Symbols.SZ3_Shoe_1
//   Author       :  Rajeswari
//   Creation Date: 12-Dec-2012 
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   12-Dec-2012   Rajeswari CR222282 .Net HS_Power project creation
//	 25-Mar-2013    Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
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
    public class SZ3_Shoe_1 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Power,Ingr.SP3D.Content.Support.Symbols.SZ3_Shoe_1"
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
        [InputString(6, "InputBomDesc", "InputBomDesc", "No Value")]
        public InputString m_sInputBomDesc;
        [InputDouble(7, "Dw", "Dw", 0.999999)]
        public InputDouble m_dDw;
        [InputDouble(8, "Angle", "Angle", 0.999999)]
        public InputDouble m_dAngle;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("BODY", "BODY")]
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

                if (A == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidA, "A cannot be zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(A / 2, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(0, -B / 2, -H));
                pointCollection.Add(new Position(0, B / 2, -H));
                pointCollection.Add(new Position(0, B / 2, -H + T));
                pointCollection.Add(new Position(0, T / 2, -H + T));
                pointCollection.Add(new Position(0, T / 2, -Dw / 2));
                pointCollection.Add(new Position(0, -T / 2, -Dw / 2));
                pointCollection.Add(new Position(0, -T / 2, -H + T));
                pointCollection.Add(new Position(0, -B / 2, -H + T));
                pointCollection.Add(new Position(0, -B / 2, -H));

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                curveCollection.Add(new LineString3d(pointCollection));

                Vector projectionVector = new Vector(A, 0, 0);
                Projection3d body = new Projection3d((ICurve)(new ComplexString3d(curveCollection)), projectionVector, projectionVector.Length, true);
                m_Symbolic.Outputs["BODY"] = body;

                if (angle > 0)
                {
                    Matrix4X4 matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate(angle, new Vector(0, 1, 0));
                    body.Transform(matrix);
                }
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SZ3_Shoe_1"));
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of SZ3_Shoe_1"));
                }
                return "";
            }
        }
        #endregion
    }
}
