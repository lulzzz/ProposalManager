// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information

using System;
using System.Collections.Generic;
using System.Text;

namespace ApplicationCore.Helpers
{
    /// <summary>
    /// A key value pair object.
    /// </summary>
    public abstract class Option
    {
        /// <summary>
        /// Create a new option.
        /// </summary>
        /// <param name="name">The name of the option.</param>
        /// <param name="value">The value of the option.</param>
        protected Option(string name, string value)
        {
            this.Name = name;
            this.Value = value;
        }

        /// <summary>
        /// The name, or key, of the option.
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// The value of the option.
        /// </summary>
        public string Value { get; private set; }
    }
}
