//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Gen_UStrap.cs
//    Generic,Ingr.SP3D.Content.Support.Symbols.Gen_UStrap
//   Author       :  Hema
//   Creation Date:  15-11-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   15-11-2012     Hema     CR-CP-222274  Converted HS_Generic VB Project to C# .Net 
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   6/11/2013     Vijaya   CR-CP-242533  Provide the ability to store GType outputs as a single blob for H&S .net symbols  
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

    [CacheOption(CacheOptionType.NonCached)] 
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Gen_UStrap : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Generic,Ingr.SP3D.Content.Support.Symbols.Gen_UStrap"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(3, "t", "t", 0.999999)]
        public InputDouble m_dt;
        [InputDouble(4, "H1", "H1", 0.999999)]
        public InputDouble m_dH1;
        [InputDouble(5, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(6, "P", "P", 0.999999)]
        public InputDouble m_dP;
        [InputDouble(7, "R", "R", 0.999999)]
        public InputDouble m_dR;
        [InputDouble(8, "P1", "P1", 0.999999)]
        public InputDouble m_dP1;
        [InputString(9, "BOM_DESC1", "BOM_DESC1","No Value")]
        public InputString m_oBOM_DESC1;
        [InputDouble(10, "Material", "Material", 1)]
        public InputDouble m_oMaterial;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BODY1", "BODY1")]
        [SymbolOutput("BODY2", "BODY2")]
        [SymbolOutput("NUT1", "NUT1")]
        [SymbolOutput("NUT2", "NUT2")]
        [SymbolOutput("BOLTHEAD1", "BOLTHEAD1")]
        [SymbolOutput("BOLTHEAD2", "BOLTHEAD2")]
        [SymbolOutput("BOLT1", "BOLT1")]
        [SymbolOutput("BOLT2", "BOLT2")]
        [SymbolOutput("arc", "arc")]
        [SymbolOutput("arc1", "arc1")]
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
                Double T = m_dt.Value;               
                Double H1 = m_dH1.Value;
                Double H = m_dH.Value;
                Double P = m_dP.Value;
                Double R = m_dR.Value;
                Double P1 = m_dP1.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, -R - H1), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                if (B == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidB, "Strap Length cannot be zero"));
                    return;
                }
                if (T == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidT, "t: Strap Thickness cannot be zero"));
                    return;
                }

                Collection<ICurve> curveCollection;
                curveCollection = new Collection<ICurve>();

                curveCollection.Add(new Line3d( new Position(-B / 2, -R - T, 0), new Position(-B / 2, -R - T, -(R - T))));

                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d outerArc = symbolGeometryHelper.CreateArc(null, R + T, Math.PI);
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -B / 2, 0));
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1), new Position(0, 0, 0));
                outerArc.Transform(matrix);
                curveCollection.Add(outerArc);

                curveCollection.Add(new Line3d(new Position(-B / 2, R + T, 0), new Position(-B / 2, R + T, -(R - T))));
                curveCollection.Add(new Line3d(new Position(-B / 2, (P / 2 - T), -(R - T)), new Position(-B / 2, R + T, -(R - T))));
                curveCollection.Add(new Line3d(new Position(-B / 2, (P / 2 - T), -(R)), new Position(-B / 2, (P / 2 - T), -(R - T))));
                curveCollection.Add(new Line3d(new Position(-B / 2, R, -(R)), new Position(-B / 2, (P / 2 - T), -(R))));
                curveCollection.Add(new Line3d(new Position(-B / 2, R, 0), new Position(-B / 2, R, -(R))));

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Arc3d innerArc = symbolGeometryHelper.CreateArc(null, R , Math.PI);
                matrix = new Matrix4X4();
                matrix.Rotate(Math.PI, new Vector(0, 0, 1));
                matrix.Rotate(-(Math.PI / 2), new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, -B / 2, 0));
                matrix.Rotate(3 * (Math.PI / 2), new Vector(0, 0, 1), new Position(0, 0, 0));
                innerArc.Transform(matrix);
                curveCollection.Add(innerArc);

                curveCollection.Add(new Line3d( new Position(-B / 2, -R, 0), new Position(-B / 2, -R, -(R))));
                curveCollection.Add(new Line3d( new Position(-B / 2, -R, -(R)), new Position(-B / 2, -(P / 2 - T), -(R))));
                curveCollection.Add(new Line3d( new Position(-B / 2, -(P / 2 - T), -(R)), new Position(-B / 2, -(P / 2 - T), -(R - T))));
                curveCollection.Add(new Line3d( new Position(-B / 2, -(P / 2 - T), -(R - T)), new Position(-B / 2, -R - T, -(R - T))));

                Vector lineVector = new Vector(B, 0, 0);
                Projection3d BODY1 = new Projection3d( new ComplexString3d(curveCollection), lineVector, lineVector.Length, true);
                m_Symbolic.Outputs["BODY1"] = BODY1;

                curveCollection = new Collection<ICurve>();

                curveCollection.Add(new Line3d( new Position(-B / 2, P / 2 + T, -R - H1), new Position(-B / 2, P / 2 + T, -R)));
                curveCollection.Add(new Line3d( new Position(-B / 2, P / 2 + T, -R), new Position(-B / 2, -P / 2 - T, -R)));
                curveCollection.Add(new Line3d( new Position(-B / 2, -P / 2 - T, -R), new Position(-B / 2, -P / 2 - T, -R - H1)));
                curveCollection.Add(new Line3d( new Position(-B / 2, -P / 2 - T, -R - H1), new Position(-B / 2, -P / 2, -R - H1)));
                curveCollection.Add(new Line3d( new Position(-B / 2, -P / 2, -R - H1), new Position(-B / 2, -P / 2, -R - T)));
                curveCollection.Add(new Line3d( new Position(-B / 2, -P / 2, -R - T), new Position(-B / 2, P / 2, -R - T)));
                curveCollection.Add(new Line3d( new Position(-B / 2, P / 2, -R - T), new Position(-B / 2, P / 2, -R - H1)));
                curveCollection.Add(new Line3d( new Position(-B / 2, P / 2, -R - H1), new Position(-B / 2, P / 2 + T, -R - H1)));

                
                Vector lineVector1 = new Vector(B, 0, 0);
                Projection3d BODY2 = new Projection3d( new ComplexString3d(curveCollection), lineVector1, lineVector1.Length, true);
                m_Symbolic.Outputs["BODY2"] = BODY2;

                symbolGeometryHelper.ActivePosition = new Position(0, P1 / 2, -R - T - 0.6 * T);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d nut1 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, T, 0.6 * T);
                m_Symbolic.Outputs["NUT1"] = nut1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -P1 / 2, -R - T - 0.6 * T);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d nut2 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, T, 0.6 * T);
                m_Symbolic.Outputs["NUT2"] = nut2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, P1 / 2, -R + T);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d boltHead1 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, T, 0.6 * T);
                m_Symbolic.Outputs["BOLTHEAD1"] = boltHead1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -P1 / 2, -R + T);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d boltHead2 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, T, 0.6 * T);
                m_Symbolic.Outputs["BOLTHEAD2"] = boltHead2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -P1 / 2, -R - 2 * T);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d bolt1 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, 0.6 * T, 3 * T);
                m_Symbolic.Outputs["BOLT1"] = bolt1;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, P1 / 2, -R - 2 * T);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d bolt2 = (Projection3d)symbolGeometryHelper.CreateCylinder(null, 0.6 * T, 3 * T);
                m_Symbolic.Outputs["BOLT2"] = bolt2;
            }
            catch    //General Unhandled exception
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrConstructOutputs, "Error in ConstructOutputs of Gen_UStrap"));
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

                PropertyValueCodelist materialCodeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrGenericUStrapMat", "Material");
                CodelistItem codeList = materialCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(materialCodeList.PropValue);

                if (codeList != null)
                {
                    material = codeList.ShortDisplayName;
                }
                Double H1Value = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrGenericUStrap", "H1")).PropValue;
                String H1 = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, H1Value, UnitName.DISTANCE_MILLIMETER);
                String bomDescription = (String)((PropertyValueString)part.GetPropertyValue("IJOAHgrGenPartBOMDesc", "BOM_DESC1")).PropValue;

                if ((bomDescription).ToUpper() == null)
                    bomString = "";
                else
                    bomString = part.PartDescription + ", H1=" + H1 + ", " + material;
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrBOMDescription, "Error in BOMDescription of Gen_UStrap"));
                }
                return "";
            }
        }
        #endregion
    }
}

