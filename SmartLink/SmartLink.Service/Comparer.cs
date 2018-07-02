// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartLink.Service
{
    public class Comparer<T> : IEqualityComparer<T>
    {
        private readonly Func<T, T, bool> _equalityComparer;

        public Comparer(Func<T, T, bool> equalityComparer)
        {
            _equalityComparer = equalityComparer;
        }

        public bool Equals(T first, T second)
        {
            return _equalityComparer(first, second);
        }

        public int GetHashCode(T value)
        {
            return value.GetHashCode();
        }
    }
}
