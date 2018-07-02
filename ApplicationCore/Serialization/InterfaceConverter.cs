// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using Newtonsoft.Json.Converters;

namespace ApplicationCore.Serialization
{
    /// <summary>
    /// Handles resolving interfaces to the correct concrete class during serialization/deserialization.
    /// </summary>
    /// <typeparam name="T">The concrete instance type.</typeparam>
    public class InterfaceConverter<T> : CustomCreationConverter<T>
        where T : new()
    {
        public override T Create(Type objectType)
        {
            return new T();
        }
    }
}
