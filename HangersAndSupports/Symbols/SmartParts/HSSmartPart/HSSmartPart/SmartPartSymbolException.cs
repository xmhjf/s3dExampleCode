//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SmartPartSymbolException.cs
//   Author       :  PVK
//   Creation Date:  13-05-2015
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Support.Symbols
{
    public class SmartPartSymbolException : CmnException
    {
        /// <summary>
        /// Raised when the SmartPart is not valid.
        /// </summary>
        /// <param name="iMessageID">The error message number that explains the reason for the exception.</param>
        /// <param name="sDefaultMessage"> Default error message. </param>
        public SmartPartSymbolException(int iMessageID, string sDefaultMessage)
            : base(iMessageID, sDefaultMessage, SmartPartSymbolResourceIDs.DEFAULT_RESOURCE, SmartPartSymbolResourceIDs.DEFAULT_ASSEMBLY)
        {
            //No Implementation Required.
        }
    }
}
