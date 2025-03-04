//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Gen_UBolt_E.cs
//    Generic,Ingr.SP3D.Content.Support.Symbols.Gen_UBolt_E
//   Author       :  Hema
//   Creation Date:  21-11-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-11-2012     Hema     CR-CP-222274  Converted HS_Generic VB Project to C# .Net 
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   6/11/2013     Vijaya   CR-CP-242533  Provide the ability to store GType outputs as a single blob for H&S .net symbols  
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

    [CacheOption(CacheOptionType.NonCached)] 
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Gen_UBolt_E : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Generic,Ingr.SP3D.Content.Support.Symbols.Gen_UBolt_E"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "d", "d", 0.999999)]
        public InputDouble m_dd;
        [InputDouble(3, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(4, "R", "R", 0.999999)]
        public InputDouble m_dR;
        [InputString(5, "BOM_DESC1", "BOM_DESC1", "No Value")]
        public InputString m_oBOM_DESC1;
        [InputDouble(6, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(7, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(8, "FlgThick", "FlgThick", 0.999999)]
        public InputDouble m_dFlgThick;
        [InputDouble(9, "Material", "Material", 1)]
        public InputDouble m_oMaterial;
        [InputDouble(10, "PadL", "PadL", 0.999999)]
        public InputDouble m_dPadL;
        [InputDouble(11, "PipeRadius", "PipeRadius", 0.999999)]
        public InputDouble m_dPipeRadius;
        [InputDouble(12, "P1", "P1", 0.999999)]
        public InputDouble m_dP1;
        [InputDouble(13, "Angle", "Angle", 0.999999)]
        public InputDouble m_dAngle;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("BEND", "BEND")]
        [SymbolOutput("RIGHT", "RIGHT")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("RIGHTNUT1", "RIGHTNUT1")]
        [SymbolOutput("LEFTNUT1", "LEFTNUT1")]
        [SymbolOutput("RIGHTNUT2", "RIGHTNUT2")]
        [SymbolOutput("LEFTNUT2", "LEFTNUT2")]
        [SymbolOutput("PAD", "PAD")]
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

                Double diameter = m_dd.Value;
                
                Double H = m_dH.Value;
                Double R = m_dR.Value;                
                Double E = m_dE.Value;
                Double F = m_dF.Value;
                Double T = m_dFlgThick.Value;
                Double padL = m_dPadL.Value;                
                Double pipeRadius = m_dPipeRadius.Value;
                Double P1 = m_dP1.Value;
                Double angle = m_dAngle.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                if (diameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidDiameter, "Diameter should be greater than zero"));
                    return;
                }
                if (H == 0 && R == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidHandR, "H: Height and R: Radius of Bend cannot be zero"));
                    return;
                }
                if (padL <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidPadL, "Pad Len should be greater than zero"));
                    return;
                }
                if (T > F)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidTGTF, "Flange Thickness should not be greater than Flange Depth"));
                    return;
                }

                string[] Outputstring = new string[8];

                Revolution3d bend = new Revolution3d((new Circle3d(new Position(0, R + diameter / 2, 0), new Vector(0, 0, 1), diameter / 2)), new Vector(1, 0, 0), new Position(0, 0, 0), Math.PI, true);
                m_Symbolic.Outputs["BEND"] = bend;
                Outputstring[0] = "BEND";

                symbolGeometryHelper.ActivePosition = new Position(0, P1 / 2, -H + R);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d right = (Projection3d)symbolGeometryHelper.CreateCylinder(null, diameter / 2, H - R);
                m_Symbolic.Outputs["RIGHT"] = right;
                Outputstring[1] = "RIGHT";

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -P1 / 2, -H + R);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d left = (Projection3d)symbolGeometryHelper.CreateCylinder(null, diameter / 2, H - R);
                m_Symbolic.Outputs["LEFT"] = left;
                Outputstring[2] = "LEFT";

                symbolGeometryHelper.ActivePosition = new Position(0, 0, -(H - 2 * diameter - pipeRadius + T / 2));
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d pad = (Projection3d)symbolGeometryHelper.CreateBox(null, 0.005, padL, 1.4 * diameter);
                m_Symbolic.Outputs["PAD"] = pad;
                Outputstring[3] = "PAD";

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, P1 / 2, -(H - diameter / 2 - pipeRadius - T / 4));
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d rightNut1 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, diameter, 0.6 * diameter);
                m_Symbolic.Outputs["RIGHTNUT1"] = rightNut1;
                Outputstring[4] = "RIGHTNUT1";

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -P1 / 2, -(H - diameter / 2 - pipeRadius - T / 4));
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d leftNut1 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, diameter, 0.6 * diameter);
                m_Symbolic.Outputs["LEFTNUT1"] = leftNut1;
                Outputstring[5] = "LEFTNUT1";

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, P1 / 2, -(H - 2 * diameter - pipeRadius + T / 2));
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d rightNut2 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, diameter, 0.6 * diameter);
                m_Symbolic.Outputs["RIGHTNUT2"] = rightNut2;
                Outputstring[6] = "RIGHTNUT2";

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -P1 / 2, -(H - 2 * diameter - pipeRadius + T / 2));
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d leftNut2 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, diameter, 0.6 * diameter);
                m_Symbolic.Outputs["LEFTNUT2"] = leftNut2;
                Outputstring[7] = "LEFTNUT2";

                if (angle > 0)
                {
                    Matrix4X4 matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    Vector vec = new Vector(0, 1, 0);
                    matrix.Rotate(angle, vec);

                    for (int i = 0; i < Outputstring.Length; i++)
                        if (Outputstring[i] != null)
                        {
                            Geometry3d obj = (Geometry3d)m_Symbolic.Outputs[Outputstring[i]];
                            obj.Transform(matrix);
                        }
                }
            }
            catch    //General Unhandled exception
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrConstructOutputs, "Error in ConstructOutputs of Gen_UBolt_E"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                String material = "";
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                PropertyValueCodelist materialCodeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrGenericUBoltMat", "Material");
                CodelistItem codeList = materialCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(materialCodeList.PropValue);

                if (codeList != null)
                {
                    material = codeList.ShortDisplayName;
                }

                String bomDescription = (String)((PropertyValueString)part.GetPropertyValue("IJOAHgrGenPartBOMDesc", "BOM_DESC1")).PropValue;

                if ((bomDescription).ToUpper() == null)
                    bomString = "";
                else
                    bomString = part.PartDescription + ", " + material;
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrBOMDescription, "Error in BOMDescription of Gen_UBolt_E"));
                }
                return "";
            }
        }
        #endregion
    }
}
