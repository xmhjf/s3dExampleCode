
//-----------------------------------------------------------------------------
//      Copyright (C) 2011-15 Intergraph Corporation.  All rights reserved.
//
//      Component:  SectionFBPortMap to be implemented by all classes
//                  that are returning a EmulatedPort map according to the sectionAlias.
//
//      Author:  
//
//      History:
//      November 13, 2011       BSLee                  Created
//      Nov 19, 2015       mchandak               DI-275514 and DI-275249 Updated Slot Mapping Rule to avoid interop calls for GetNormalFromPostion() and IsExternalWire()
//      December 11, 2015       hgajula                DI-CP-284051  Fix coverity defects stated in November 6, 2015 report
//      December 11, 2015       knukala                DI-CP-275525  Replace PlaceIntersectionObject method in SlotMapng rule with .Net API 
//-----------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Structure.Middle;




namespace Ingr.SP3D.Content.Structure.EmulatedPortMappings
{
    /// <summary>
    /// The SectionFBPortMap class is designed to return emulated port maps based on FB SectionAlias
    /// </summary>
    internal class SectionFBPortMap : IEmulatedPortMap
    {
        #region IEmulatedPortMap Members


        /// <summary>
        /// Get the Ports on the penetrating parts depends on SectionAlias with SectionFaceType (ex) Web_Left..)
        /// </summary>
        /// <param name="penetratingPart">The Penetrating Object.</param>
        /// <param name="penetratedPart">The Penetrated Object.</param>
        /// <param name="basePltPort">BasePlate Port.</param>
        /// <param name="sectionAlias">SectionAlias .</param>
        /// <param name="web">The Web Object.</param>
        /// <param name="flange">The Flange Object.</param>
        /// <param name="secondWeb">The 2ndWeb Object.</param>
        /// <param name="secondFlange">The 2ndFlange Object.</param>
        /// <param name="basePlatePort">Base Plate Port.</param>
        /// <returns>The Ports on Penetrating parts. Ports that don't apply are omitted </returns>

        public Dictionary<int, IPort> GetEmulatedPortsMap(object penetratingPart, object penetratedPart, IPort basePltPort, string sectionAlias, object web, object flange, object secondWeb, object secondFlange, out IPort basePlatePort)
        {


            if (penetratingPart == null)
            {
                throw new ArgumentNullException("Input PenetratingPart");
            }

            if (penetratedPart == null)
            {
                throw new ArgumentNullException("Input PenetratedPart");
            }

            if (basePltPort == null)
            {
                throw new ArgumentNullException("Input PenetratedLateralEdgePort");
            }

            if (web == null)
            {
                throw new ArgumentNullException("Input Web");
            }


            basePlatePort = null;
            Dictionary<int, IPort> mappedPorts = null;

            mappedPorts = new Dictionary<int, IPort>();
            IPort basePortOfWeb = null;
            IPort offsetPortOfWeb = null;
            IPort bottomPortOfWeb = null;
            IPort topPortOfWeb = null;
            object rootBasePlateSystem = null;

            // Get Web and Base plate ports
            CommonFuncs.GetWebAndBasePlatePorts(penetratingPart, penetratedPart, basePltPort, web, flange, secondWeb, secondFlange, out basePlatePort,
                out basePortOfWeb, out offsetPortOfWeb, out bottomPortOfWeb, out topPortOfWeb, out rootBasePlateSystem);



            //                      Top
            //                     ******    
            //                     *    *
            //                     *    *  
            //             WebLeft *    *  WebRight
            //             (Base)  *    *  (Offset)
            //                     *    *
            //                     ****** 
            //                     Bottom
            // 
            // In this case, Base port of Penetrating plate is Web Left 

            mappedPorts.Add((int)SectionFaceType.Top, topPortOfWeb);
            mappedPorts.Add((int)SectionFaceType.Bottom, bottomPortOfWeb);
            mappedPorts.Add((int)SectionFaceType.Web_Left, basePortOfWeb);
            mappedPorts.Add((int)SectionFaceType.Web_Right, offsetPortOfWeb);

            return mappedPorts;
        }

        #endregion

    }
}