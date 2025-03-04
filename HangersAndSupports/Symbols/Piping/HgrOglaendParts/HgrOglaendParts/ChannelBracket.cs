//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   ChannelBracket.cs
//   HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.ChannelBracket
//   Author       :  Pavan
//   Creation Date:  19.June.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   19.June.2012    Pavan     Initial Creation
//   23-01-2015      Chethan  CR-CP-241198  Convert all Oglaend Content to .NET 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.ObjectModel;

using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.Support.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [VariableOutputs]
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class ChannelBracket : CustomSymbolDefinition, ICustomHgrWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "HgrOglaendParts,Ingr.SP3D.Content.Support.Symbols.ChannelBracket"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Height1", "Height of Base", 0)]
        public InputDouble m_Height1;
        [InputDouble(3, "Width1", "Width of Base", 0)]
        public InputDouble m_Width1;
        [InputDouble(4, "Height2", "Height of Channel", 0)]
        public InputDouble m_Height2;
        [InputDouble(5, "Width2", "Width of Channel", 0)]
        public InputDouble m_Width2;
        [InputDouble(6, "Thickness", "Thickness of Channel Bracket", 0)]
        public InputDouble m_Thickness;
        [InputDouble(7, "Gap", "Gap of Channel Bracket", 0)]
        public InputDouble m_Gap;
        [InputDouble(8, "ChamferX", "ChamferX", 0)]
        public InputDouble m_ChamferX;
        [InputDouble(9, "ChamferY", "ChamferY", 0)]
        public InputDouble m_ChamferY;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Plate1", "Plate1")]
        [SymbolOutput("Plate2", "Plate2")]
        [SymbolOutput("Plate3", "Plate3")]
        [SymbolOutput("Surface1", "Surface1")]
        [SymbolOutput("Surface2", "Surface2")]
        [SymbolOutput("Surface3", "Surface3")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = null;
                SymbolGeometryHelper symbolGeomHlpr = new SymbolGeometryHelper();
                SP3DConnection connection = default(SP3DConnection);
                connection = OccurrenceConnection;

                part = m_PartInput.Value as Part;
                double height1 = 0;
                double width1 = 0;
                double height2 = 0;
                double width2 = 0;
                double thickness = 0;
                double gap = 0;
                double chamferX = m_ChamferX.Value;
                double chamferY = m_ChamferY.Value;

                height1 = m_Height1.Value;
                if (height1 < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Height1 value should be greater than 0"));
                    return;
                }

                width1 = m_Width1.Value;
                if (width1 < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Width1 value should be greater than 0"));
                    return;
                }

                height2 = m_Height2.Value;
                if (height2 < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Height2 value should be greater than 0"));
                    return;
                }

                width2 = m_Width2.Value;
                if (width2 < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Width2 value should be greater than 0"));
                    return;
                }

                thickness = m_Thickness.Value;
                if (thickness < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Thickness value should be greater than 0"));
                    return;
                }

                gap = m_Gap.Value;
                if (gap < 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrInvalidArguments, "Gap value should be greater than 0"));
                    return;
                }

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //Initialize SymbolGeometryHelper. Set the active position and orientation 
                symbolGeomHlpr.ActivePosition = new Position(0, 0, 0);
                symbolGeomHlpr.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));

                //Plate1 (Left Side Plate of Channel Bracket)
                Line3d plate1_1, plate1_2, plate1_3, plate1_4, plate1_5, plate1_6, plate1_7;
                Collection<ICurve> plate1Coll = new Collection<ICurve>();
                plate1_1 = new Line3d(new Position(0, gap / 2, 0), new Position(width1, gap / 2, 0));
                plate1_2 = new Line3d(new Position(width1, gap / 2, 0), new Position(width1 + width2, gap / 2, chamferY));
                plate1_3 = new Line3d(new Position(width1 + width2, gap / 2, chamferY), new Position(width1 + width2, gap / 2, height2));
                plate1_4 = new Line3d(new Position(width1 + width2, gap / 2, height2), new Position(width1, gap / 2, height2));
                plate1_5 = new Line3d(new Position(width1, gap / 2, height2), new Position(width1, gap / 2, height1));
                plate1_6 = new Line3d(new Position(width1, gap / 2, height1), new Position(0, gap / 2, height1));
                plate1_7 = new Line3d(new Position(0, gap / 2, height1), new Position(0, gap / 2, 0));

                ////Add the curves into ICurves Collection
                plate1Coll.Add(plate1_1);
                plate1Coll.Add(plate1_2);
                plate1Coll.Add(plate1_3);
                plate1Coll.Add(plate1_4);
                plate1Coll.Add(plate1_5);
                plate1Coll.Add(plate1_6);
                plate1Coll.Add(plate1_7);

                //Plate2 (Right Side Plate of Channel Bracket)
                Line3d plate2_1, plate2_2, plate2_3, plate2_4, plate2_5, plate2_6, plate2_7;
                Collection<ICurve> plate2Coll = new Collection<ICurve>();

                plate2_1 = new Line3d(new Position(0, -gap / 2, 0), new Position(width1, -gap / 2, 0));
                plate2_2 = new Line3d(new Position(width1, -gap / 2, 0), new Position(width1 + width2, -gap / 2, chamferY));
                plate2_3 = new Line3d(new Position(width1 + width2, -gap / 2, chamferY), new Position(width1 + width2, -gap / 2, height2));
                plate2_4 = new Line3d(new Position(width1 + width2, -gap / 2, height2), new Position(width1, -gap / 2, height2));
                plate2_5 = new Line3d(new Position(width1, -gap / 2, height2), new Position(width1, -gap / 2, height1));
                plate2_6 = new Line3d(new Position(width1, -gap / 2, height1), new Position(0, -gap / 2, height1));
                plate2_7 = new Line3d(new Position(0, -gap / 2, height1), new Position(0, -gap / 2, 0));

                ////Add the curves into ICurves Collection
                plate2Coll.Add(plate2_1);
                plate2Coll.Add(plate2_2);
                plate2Coll.Add(plate2_3);
                plate2Coll.Add(plate2_4);
                plate2Coll.Add(plate2_5);
                plate2Coll.Add(plate2_6);
                plate2Coll.Add(plate2_7);

                //Plate3 (Right Side Plate of Channel Bracket)
                Line3d plate3_1, plate3_2, plate3_3, plate3_4;
                Collection<ICurve> plate3Coll = new Collection<ICurve>();

                plate3_1 = new Line3d(new Position(width1, -(gap / 2) - thickness, height2), new Position(width1 + width2, -(gap / 2) - thickness, height2));
                plate3_2 = new Line3d(new Position(width1 + width2, -(gap / 2) - thickness, height2), new Position(width1 + width2, (gap / 2) + thickness, height2));
                plate3_3 = new Line3d(new Position(width1 + width2, (gap / 2) + thickness, height2), new Position(width1, (gap / 2) + thickness, height2));
                plate3_4 = new Line3d(new Position(width1, (gap / 2) + thickness, height2), new Position(width1, -(gap / 2) - thickness, height2));

                ////Add the curves into ICurves Collection
                plate3Coll.Add(plate3_1);
                plate3Coll.Add(plate3_2);
                plate3Coll.Add(plate3_3);
                plate3Coll.Add(plate3_4);

                ComplexString3d plate1ComplxString = new ComplexString3d(plate1Coll);
                Vector plate1Vector = new Vector(0, 1, 0);
                Projection3d plate1Projection = new Projection3d(connection, plate1ComplxString, plate1Vector, thickness, true);

                ComplexString3d plate2ComplxString = new ComplexString3d(plate2Coll);
                Vector plate2Vector = new Vector(0, -1, 0);
                Projection3d plate2Projection = new Projection3d(connection, plate2ComplxString, plate2Vector, thickness, true);

                ComplexString3d plate3ComplxString = new ComplexString3d(plate3Coll);
                Vector plate3Vector = new Vector(0, 0, -1);
                Projection3d plate3Projection = new Projection3d(connection, plate3ComplxString, plate3Vector, thickness, true);

                m_PhysicalAspect.Outputs["Plate1"] = plate1Projection;
                m_PhysicalAspect.Outputs["Plate2"] = plate2Projection;
                m_PhysicalAspect.Outputs["Plate3"] = plate3Projection;

                //Add the ports
                Port surface1Port = new Port(connection, part, "Surface1", new Position(width1 + (width2 / 2), 0, height2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Surface1"] = surface1Port;

                Port surface2Port = new Port(connection, part, "Surface2", new Position(0, -gap / 2, height2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Surface2"] = surface2Port;

                Port surface3Port = new Port(connection, part, "Surface3", new Position(0, gap / 2, height2), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Surface3"] = surface3Port;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    Type myType = this.GetType();
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, OglaendPartLocalizer.GetString(OglaendPartSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of ChannelBracket.cs " + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name));
                }
            }

        #endregion
        }
#region ICustomHgrWeightCG Members
        void ICustomHgrWeightCG.WeightCG(SupportComponent supportComponent, ref double weight, ref double cogX, ref double cogY, ref double cogZ)
        {

            ////System WCG Attributes

            Part catalogPart = (Part)supportComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

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

        }
        #endregion
    }
}