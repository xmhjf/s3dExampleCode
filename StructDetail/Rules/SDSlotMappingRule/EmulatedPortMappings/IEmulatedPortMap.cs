//-----------------------------------------------------------------------------
//      Copyright (C) 2011-15 Intergraph Corporation.  All rights reserved.
//
//      Component:  IEmulatedPortMap interface to be implemented by all classes
//                  that are returning a EmulatedPort map according to the sectionAlias.
//
//      Author:  
//
//      History:
//      November 13, 2011       BSLee                  Created
//      Nov 19, 2015       mchandak               DI-275514 and DI-275249 Updated Slot Mapping Rule to avoid interop calls for GetNormalFromPostion() and IsExternalWire()
//-----------------------------------------------------------------------------

using System.Collections.Generic;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Common.Middle;


namespace Ingr.SP3D.Content.Structure.EmulatedPortMappings
{
    /// <summary>
    /// Implemented by the SectionAliasPort Map classes.
    /// </summary>
    interface IEmulatedPortMap
    {
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
        Dictionary<int, IPort> GetEmulatedPortsMap(object penetratingPart, object penetratedPart, IPort basePltPort, string sectionAlias, object web, object flange, object secondWeb, object secondFlange, out IPort basePlatePort);
    }
}