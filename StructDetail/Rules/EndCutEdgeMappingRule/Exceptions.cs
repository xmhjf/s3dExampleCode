//-----------------------------------------------------------------------------
//      Copyright (C) 2010 Intergraph Corporation.  All rights reserved.
//
//      Component:  S3DEndCutMappingException class handles exceptions for  
//                  cross section mapping services.  
//
//      Author:  3XCalibur
//
//      History:
//      November 05, 2010       WR                  Created
//
//-----------------------------------------------------------------------------

using System;
using System.Runtime.Serialization;
using Ingr.SP3D.Common.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Exception class to handle exceptions for the cross section mapping services.
    /// </summary>
    [Serializable]
    public class EdgeMappingException : CmnException
    {
        /// <summary>
        /// Raise exception when unable to map the cross section because of an invalid quadrant.
        /// </summary>
        /// <param name="quadrant">The current mapping quadrant.</param>
        public EdgeMappingException(int quadrant)
            : base("Unable to map cross section.  Quadrant " + quadrant + " is invalid.")
        {
        }
        
        /// <summary>
        /// Initializes a new instance of the <see cref="EdgeMappingException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        public EdgeMappingException(string message)
            : base(message)
        {
        }
        
        /// <summary>
        /// Initializes a new instance of the <see cref="EdgeMappingException"/> class.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="innerException">The inner exception.</param>
        public EdgeMappingException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
        
        /// <summary>
        /// Initializes a new instance of the <see cref="EdgeMappingException"/> class.
        /// </summary>
        /// <param name="info">The <see cref="T:System.Runtime.Serialization.SerializationInfo"/> that holds the serialized object data about the exception being thrown.</param>
        /// <param name="context">The <see cref="T:System.Runtime.Serialization.StreamingContext"/> that contains contextual information about the source or destination.</param>
        /// <exception cref="T:System.ArgumentNullException">
        /// The <paramref name="info"/> parameter is null.
        /// </exception>
        /// <exception cref="T:System.Runtime.Serialization.SerializationException">
        /// The class name is null or <see cref="P:System.Exception.HResult"/> is zero (0).
        /// </exception>
        protected EdgeMappingException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }
}
