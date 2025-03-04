//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Gen_ClampPlate.cs
//   HVAC,Ingr.SP3D.Content.Support.Symbols.Gen_ClampPlate
//   Author       :  Hema
//   Creation Date:  30-11-2012
//   Description:    Converted HS_HVAC VB Project to C# .Net

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   30-11-2012     Hema     CR-CP-222294 Converted HS_HVAC VB Project to C# .Net
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   03/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
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

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Gen_ClampPlate : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HVAC,Ingr.SP3D.Content.Support.Symbols.Gen_ClampPlate"
        //----------------------------------------------------------------------------------

        double angle1, angle2;
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Width", "Width", 0.999999)]
        public InputDouble m_dWidth;
        [InputDouble(3, "Depth", "Depth", 0.999999)]
        public InputDouble m_dDepth;
        [InputDouble(4, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        [InputDouble(5, "Radius", "Radius", 0.999999)]
        public InputDouble m_dRadius;
        [InputDouble(6, "Angle", "Angle", 0.999999)]
        public InputDouble m_dAngle;
        [InputString(7, "InputBomDesc1", "InputBomDesc1", "No Value")]
        public InputString m_oInputBomDesc1;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("DuctPlate", "DuctPlate")]
        [SymbolOutput("Port1", "Port1")]
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
               
                double width = m_dWidth.Value;
                double depth = m_dDepth.Value;
                double thickness = m_dThickness.Value;
                double radius = m_dRadius.Value;
                double angle = m_dAngle.Value;
                double angle1Value = ((width / 2) / (radius));
                double angle2Value = ((width / 2) / (radius + thickness));
                angle1 = Math.Asin(angle1Value);
                angle2 = Math.Asin(angle2Value);

                if (width == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrInvalidWidth, "Width cannot be zero"));
                    return;
                }
                if (radius == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrInvalidRadius, "Radius cannot be zero"));
                    return;
                }
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Collection<ICurve> curveCollection = new Collection<ICurve>();
                Matrix4X4 matrix = new Matrix4X4();

                curveCollection.Add(new Line3d( new Position(-depth / 2, -width / 2, (radius + thickness)), new Position(-depth / 2, -width / 2, radius * Math.Cos(angle1))));

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                matrix.SetIdentity();
                matrix.Rotate(Math.PI / 2 - angle1, new Vector(0, 0, 1));

                Arc3d outerArc = symbolGeometryHelper.CreateArc(null, radius, 2 * angle1);
                outerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                outerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Translate(new Vector(0, -depth / 2, 0));
                outerArc.Transform(matrix);

                matrix = new Matrix4X4();
                matrix.SetIdentity();
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1));
                outerArc.Transform(matrix);
                curveCollection.Add(outerArc);

                curveCollection.Add(new Line3d( new Position(-depth / 2, width / 2, radius * Math.Cos(angle1)), new Position(-depth / 2, width / 2, (radius + thickness))));
                curveCollection.Add(new Line3d( new Position(-depth / 2, width / 2, (radius + thickness)), new Position(-depth / 2, -width / 2, (radius + thickness))));

                Vector lineVector = new Vector(width, 0, 0);
                Projection3d ductPlate = new Projection3d(new ComplexString3d(curveCollection), new Vector(1, 0, 0), width, true);
                m_Symbolic.Outputs["DuctPlate"] = ductPlate;
                matrix = new Matrix4X4();
                matrix.Rotate(angle, new Vector(1, 0, 0));
                ductPlate.Transform(matrix);
            }
            catch    //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Gen_ClampPlate"));
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
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                string bomDescription = (String)((PropertyValueString)oSupportOrComponent.GetPropertyValue("IJOAHgrHVACBomDesc", "InputBomDesc1")).PropValue;

                if (bomDescription.Trim() == "None")
                    bomString = " ";
                else if (bomDescription.Trim() == " ")
                    bomString = part.PartDescription;
                else
                    bomString = bomDescription.Trim();

                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HVACLocalizer.GetString(HVACSymbolResourceIDs.ErrConstructOutputs, "Error in BOMDescription of Gen_ClampPlate"));
                }
                return "";
            }
        }
        #endregion
    }
}
