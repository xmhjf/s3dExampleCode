//-----------------------------------------------------------------------------
//      Copyright (C) 2011 Intergraph Corporation.  All rights reserved.
//
//      Component:  SlotMappingRuleException class handles exceptions for  
//                  Slot Mapping Rule  services.  
//
//      Author:  
//
//      History:
//      November 16, 2011       BS Lee                  Created
//
//-----------------------------------------------------------------------------

using System;
using System.Runtime.Serialization;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Exception class to handle exceptions for the SlotMappingService 
    /// </summary>
    [Serializable]
    public class SlotMappingException : CmnException
    {

        /// <summary>
        /// Initializes a new instance of the <see cref="SlotMappingException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        public SlotMappingException(string message)
            : base(message)
        {
        }

    }
}
