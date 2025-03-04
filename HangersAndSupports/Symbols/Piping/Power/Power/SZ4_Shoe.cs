//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SZ4_Shoe.cs
//    Power,Ingr.SP3D.Content.Support.Symbols.SZ4_Shoe
//   Author       :  Rajeswari
//   Creation Date:  13-Dec-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   13-Dec-2012   Rajeswari CR222282 .Net HS_Power project creation
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
    public class SZ4_Shoe : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Power,Ingr.SP3D.Content.Support.Symbols.SZ4_Shoe"
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
        [InputDouble(6, "B1", "B1", 0.999999)]
        public InputDouble m_dB1;
        [InputDouble(7, "Dw", "Dw", 0.999999)]
        public InputDouble m_dDw;
        [InputString (8, "InputBomDesc", "InputBomDesc","No Value")]
        public InputString m_sInputBomDesc;
        [InputDouble(9, "Angle", "Angle", 0.999999)]
        public InputDouble m_dAngle;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("TOP", "TOP")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("RIGHT", "RIGHT")]
        [SymbolOutput("CURVEDPLATE", "CURVEDPLATE")]
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
                Double B1 = m_dB1.Value;
                Double Dw = m_dDw.Value;
                Double angle = m_dAngle.Value;

                Double calc1, calc2, calc3, calc4, angle1;

                if (Dw == 0 && H == 0 && T == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidDwInvalidHAndT, "Outer Diameter,H and Thickness cannot be zero"));
                    return;
                }
                if (B <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidB, "B should be greater than zero"));
                    return;
                }
                if (A == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrInvalidA, "A cannot be zero"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                string[] outputString = new string[4];

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0 , -H);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d top = (Projection3d)symbolGeometryHelper.CreateBox(null, T, A, B);
                m_Symbolic.Outputs["TOP"] = top;
                outputString[0] = "TOP";

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0,-(B/2-T/2), -H+T);
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

                angle1 = 2 * Math.Asin(B1 / Dw);
                calc1 = Math.Sin(angle1 / 2) * Dw / 2;
                calc2 = Math.Sqrt((Dw / 2 * Dw / 2) - (calc1 * calc1));
                calc3 = Math.Sin(angle1 / 2) * (Dw / 2 + T);
                calc4 = Math.Sqrt(((Dw / 2 + T) * (Dw / 2 + T)) - (calc3 * calc3));

                if (angle < 180)
                {
                    calc2 = -calc2;
                    calc4 = -calc4;
                }

                Line3d line1;
                Collection<ICurve> curveCollection = new Collection<ICurve>();

                Matrix4X4 matrix = new Matrix4X4();

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                matrix.SetIdentity();
                matrix.Rotate(Math.PI / 2 - angle1 / 2.0, new Vector(0, 0, 1));

                Arc3d outerArc = symbolGeometryHelper.CreateArc(null, Dw/2+T, angle1);
                outerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                outerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, -A / 2.0, 0));
                outerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                outerArc.Transform(matrix);
                curveCollection.Add(outerArc);
                
                line1 = new Line3d(new Position(-A / 2, calc3, calc4), new Position(-A / 2, calc1, calc2));
                curveCollection.Add(line1);

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(Math.PI / 2 - angle1 / 2.0, new Vector(0, 0, 1));

                Arc3d innerArc = symbolGeometryHelper.CreateArc(null, Dw/2, angle1);
                innerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                innerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, -A / 2.0, 0));
                innerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                innerArc.Transform(matrix);
                curveCollection.Add(innerArc);
              
                line1 = new Line3d(new Position(-A / 2, -calc1, calc2), new Position(-A / 2, -calc3, calc4));
                curveCollection.Add(line1);
                ComplexString3d lineString = new ComplexString3d(curveCollection);
                Vector lineVector = new Vector(1, 0, 0);
                Projection3d curvedPlate = new Projection3d(lineString, lineVector, A, true);
                m_Symbolic.Outputs["CURVEDPLATE"] = curvedPlate;
                outputString[3] = "CURVEDPLATE";

                if (angle > 0)
                {
                    matrix = new Matrix4X4();
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of SZ4_Shoe"));
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
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PowerLocalizer.GetString(PowerSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of SZ4_Shoe"));
                }
                return ToDoListMessage.ToString();
            }
        }
        #endregion
    }
}
