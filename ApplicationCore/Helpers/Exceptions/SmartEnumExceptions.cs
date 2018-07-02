// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;

namespace ApplicationCore.Helpers.Exceptions
{
    public class SmartEnumNotFoundException : Exception
    {
        public SmartEnumNotFoundException() : base()
        {
        }

        protected SmartEnumNotFoundException(System.Runtime.Serialization.SerializationInfo info, System.Runtime.Serialization.StreamingContext context) : base(info, context)
        {
        }

        public SmartEnumNotFoundException(string message) : base(message)
        {
        }

        public SmartEnumNotFoundException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
