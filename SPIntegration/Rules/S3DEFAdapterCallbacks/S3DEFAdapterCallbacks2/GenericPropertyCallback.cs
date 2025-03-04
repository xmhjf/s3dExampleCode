using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Route.Middle;
using System.Collections;
using System.Collections.ObjectModel;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.EFAdapterCallbacks
{
	/// <summary>
	/// This code is delivered to end user so that he can modify the same to change the default behavior of the EF Adapter.
	/// PropertyCallbacks class inherits from EFPropertyCallbacks class and overrides methods to provide custom behavior for the rules.
	/// </summary>
	public class GenericPropertyCallback : PropertyCallback
	{
		/// <summary>
		/// ProcessProperty.
		/// <para>Returns: 
		///     Process a property using an Info map.</para>
		/// </summary>     
		/// <param name="oInfo">The information data for the callback.</param>
		/// <returns>returns input value for base implementation.</returns>
		public override bool ProcessProperty(PropertyCallbackInformation oInfo)
		{
			return true;
		}
	}
}