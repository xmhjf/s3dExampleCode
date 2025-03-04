//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG63A.cs
//   Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG63A
//   Author       :  Hema
//   Creation Date:  05-05-2013
//   Description:    Converted Anvil_Parts VB Project to C#.Net Project

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   05-05-2013     Hema     CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

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
    [CacheOption(CacheOptionType.Cached)]

    public class Anvil_FIG63A : HangerComponentSymbolDefinition
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG63A"
        //----------------------------------------------------------------------------------


        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Length", "Length", 0.999999)]
        public InputDouble m_dLength;
        [InputDouble(3, "STANCHION_NOM_DIA", "STANCHION_NOM_DIA", 1)]
        public InputDouble m_oSTANCHION_NOM_DIA;
        [InputDouble(4, "Elbow_Radius", "Elbow_Radius", 0.999999)]
        public InputDouble m_dElbow_Radius;
        [InputDouble(5, "N", "N", 0.999999)]
        public InputDouble m_dN;
        [InputDouble(6, "A", "A", 0.999999)]
        public InputDouble m_dA;
        [InputDouble(7, "B", "B", 0.999999)]
        public InputDouble m_dB;
        [InputDouble(8, "C", "C", 0.999999)]
        public InputDouble m_dC;
        [InputDouble(9, "D", "D", 0.999999)]
        public InputDouble m_dD;
        [InputDouble(10, "E", "E", 0.999999)]
        public InputDouble m_dE;
        [InputDouble(11, "F", "F", 0.999999)]
        public InputDouble m_dF;
        [InputDouble(12, "G", "G", 0.999999)]
        public InputDouble m_dG;
        [InputDouble(13, "H", "H", 0.999999)]
        public InputDouble m_dH;
        [InputDouble(14, "I", "I", 0.999999)]
        public InputDouble m_dI;
        [InputDouble(15, "J", "J", 0.999999)]
        public InputDouble m_dJ;
        [InputDouble(16, "K", "K", 0.999999)]
        public InputDouble m_dK;
        [InputDouble(17, "L", "L", 0.999999)]
        public InputDouble m_dL;
        [InputDouble(18, "M", "M", 0.999999)]
        public InputDouble m_dM;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BASE", "BASE")]
        [SymbolOutput("PIPE", "PIPE")]
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

                Double x = 0, baseThickness = 0, baseWidth = 0, baseDepth = 0, E, stanchionDia = 0;
                Double length = m_dLength.Value;
                Double stanchionNomDia = m_oSTANCHION_NOM_DIA.Value;
                Double elbowRadius = m_dElbow_Radius.Value;
                Double[] array = new double[15];

                array[0] = m_dA.Value;
                array[1] = m_dB.Value;
                array[2] = m_dC.Value;
                array[3] = m_dD.Value;
                array[4] = m_dE.Value;
                array[5] = m_dF.Value;
                array[6] = m_dG.Value;
                array[7] = m_dH.Value;
                array[8] = m_dI.Value;
                array[9] = m_dJ.Value;
                array[10] = m_dK.Value;
                array[11] = m_dL.Value;
                array[12] = m_dM.Value;
                array[13] = m_dN.Value;
               
                string actualSnd = string.Empty, columnHeading = string.Empty;

                //Iniatizing SymbolGeometryHelper
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();
                
                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                    actualSnd = metadataManager.GetCodelistInfo("Anvil_Stanchion_Dia", "UDP").GetCodelistItem((int)stanchionNomDia).ShortDisplayName.Trim();
                else
                    actualSnd = "2";

                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass anvilfig63AClass = (PartClass)catalogBaseHelper.GetPartClass("Anvil_FIG63_DIA");
                ReadOnlyCollection<BusinessObject> fig63AClass = anvilfig63AClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItems in fig63AClass)
                {
                    if (((string)((PropertyValueString)classItems.GetPropertyValue("IJUAHgrAnvil_FIG63_DIA", "STANCHION_NOM_DIA")).PropValue == actualSnd))
                    {
                        stanchionDia = (double)((PropertyValueDouble)classItems.GetPropertyValue("IJUAHgrAnvil_FIG63_DIA", "Stanchion_Dia")).PropValue;
                        columnHeading = (string)((PropertyValueString)classItems.GetPropertyValue("IJUAHgrAnvil_FIG63_DIA", "X")).PropValue;
                        baseThickness = (double)((PropertyValueDouble)classItems.GetPropertyValue("IJUAHgrAnvil_FIG63_DIA", "BASE_THICKNESS")).PropValue;
                        baseWidth = (double)((PropertyValueDouble)classItems.GetPropertyValue("IJUAHgrAnvil_FIG63_DIA", "BASE_WIDTH")).PropValue;
                        baseDepth = (double)((PropertyValueDouble)classItems.GetPropertyValue("IJUAHgrAnvil_FIG63_DIA", "BASE_DEPTH")).PropValue;
                        break;
                    }
                }
                string[] array1 = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N" };
                for (int i = 0; i <= array1.Length; i++)
                {
                    if (columnHeading == array1[i])
                    {
                            x = array[i];
                            break;
                    }
                }

                E = length - baseThickness + x;

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (stanchionNomDia < 1 && stanchionNomDia > 14)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidstanchionNomDiaCodelist, "StanchionNom diameter should be between 1 to 14"));
                    return;
                }
                if (stanchionDia <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidstanchionDiaGTZero, "Stanchion diameter should be greater than zero"));
                    return;
                }
                if (baseWidth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidbaseWidthGTZero, "Base width should be greater than zero"));
                    return;
                }
                if (baseDepth <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidbaseDepthGTZero, "Base depth should be greater than zero"));
                    return;
                }
                if (baseThickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidbaseThicknessGTZero, "Base thickness should be greater than zero"));
                    return;
                }
                if (E == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidENZero, "E value cannot be zero"));
                    return;
                }

                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                Projection3d pipe = symbolGeometryHelper.CreateCylinder(null, stanchionDia / 2, E);
                matrix.Rotate(Math.PI, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, 0, length - baseThickness));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                pipe.Transform(matrix);
                m_Symbolic.Outputs["PIPE"] = pipe;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-baseWidth / 2, -baseDepth / 2, length - baseThickness);
                Projection3d Base = symbolGeometryHelper.CreateBox(null, baseWidth, baseDepth, baseThickness, 9);
                matrix = new Matrix4X4();
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Base.Transform(matrix);
                m_Symbolic.Outputs["BASE"] = Base;
            }
            catch       //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG63A"));
                    return;
                }
            }
        }
        #endregion

    }

}
