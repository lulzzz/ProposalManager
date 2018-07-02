// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using ApplicationCore.Helpers.Exceptions;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;

namespace ApplicationCore.Helpers
{
    /// <summary>
    /// A collection of common guard clauses, implented as extensions.
    /// </summary>
    /// <example>
    /// Guard.Against.Null(input, nameof(input));
    /// </example>
    public static class GuardClauseExtensions
    {
        /// <summary>
        /// Throws an <see cref="ArgumentNullException" /> if <see cref="input" /> is null.
        /// </summary>
        /// <param name="guardClause"></param>
        /// <param name="input"></param>
        /// <param name="parameterName"></param>
        /// <exception cref="ArgumentNullException"></exception>
        public static void Null(this IGuardClause guardClause, object input, string parameterName, string requestId = "")
        {
            if (null == input)
            {
                throw new ArgumentNullException($"RequestId: {requestId} - Null: {parameterName}");
            }
        }

        /// <summary>
        /// Throws an <see cref="ArgumentNullException" /> if <see cref="input" /> is null.
        /// Throws an <see cref="ArgumentException" /> if <see cref="input" /> is an empty string.
        /// </summary>
        /// <param name="guardClause"></param>
        /// <param name="input"></param>
        /// <param name="parameterName"></param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public static void NullOrEmpty(this IGuardClause guardClause, string input, string parameterName, string requestId = "")
        {
            Guard.Against.Null(input, parameterName, requestId);
            if (input == String.Empty)
            {
                throw new ArgumentException($"RequestId: {requestId} - Required input {parameterName} was empty.", parameterName);
            }
        }

        /// <summary>
        /// Throws an <see cref="ArgumentNullException" /> if <see cref="input" /> is null.
        /// Throws an <see cref="ArgumentException" /> if <see cref="input" /> is an empty or white space string.
        /// </summary>
        /// <param name="guardClause"></param>
        /// <param name="input"></param>
        /// <param name="parameterName"></param>
        /// <exception cref="ArgumentNullException"></exception>
        /// <exception cref="ArgumentException"></exception>
        public static void NullOrWhiteSpace(this IGuardClause guardClause, string input, string parameterName, string requestId = "")
        {
            Guard.Against.NullOrEmpty(input, parameterName, requestId);
            if (String.IsNullOrWhiteSpace(input))
            {
                throw new ArgumentException($"RequestId: {requestId} - Required input {parameterName} was empty.", parameterName);
            }
        }

        /// <summary>
        /// Throws an <see cref="ArgumentOutOfRangeException" /> if <see cref="input" /> is less than <see cref="rangeFrom" /> or greater than <see cref="rangeTo" />.
        /// </summary>
        /// <param name="guardClause"></param>
        /// <param name="input"></param>
        /// <param name="parameterName"></param>
        /// <param name="rangeFrom"></param>
        /// <param name="rangeTo"></param>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public static void OutOfRange(this IGuardClause guardClause, int input, string parameterName, int rangeFrom, int rangeTo, string requestId = "")
        {
            OutOfRange<int>(guardClause, input, parameterName, rangeFrom, rangeTo, requestId);
        }

        /// <summary>
        /// Throws an <see cref="ArgumentOutOfRangeException" /> if <see cref="input" /> is less than <see cref="rangeFrom" /> or greater than <see cref="rangeTo" />.
        /// </summary>
        /// <param name="guardClause"></param>
        /// <param name="input"></param>
        /// <param name="parameterName"></param>
        /// <param name="rangeFrom"></param>
        /// <param name="rangeTo"></param>
        /// <exception cref="ArgumentException"></exception>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public static void OutOfRange(this IGuardClause guardClause, DateTime input, string parameterName, DateTime rangeFrom, DateTime rangeTo, string requestId = "")
        {
            OutOfRange<DateTime>(guardClause, input, parameterName, rangeFrom, rangeTo, requestId);
        }

        /// <summary>
        /// Throws an <see cref="ArgumentOutOfRangeException" /> if <see cref="input" /> is not in the range of valid <see cref="SqlDateTIme" /> values.
        /// </summary>
        /// <param name="guardClause"></param>
        /// <param name="input"></param>
        /// <param name="parameterName"></param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public static void OutOfSQLDateRange(this IGuardClause guardClause, DateTime input, string parameterName, string requestId = "")
        {
            // System.Data is unavailable in .NET Standard so we can't use SqlDateTime.
            const long sqlMinDateTicks = 552877920000000000;
            const long sqlMaxDateTicks = 3155378975999970000;

            OutOfRange<DateTime>(guardClause, input, parameterName, new DateTime(sqlMinDateTicks), new DateTime(sqlMaxDateTicks), requestId);
        }

        private static void OutOfRange<T>(this IGuardClause guardClause, T input, string parameterName, T rangeFrom, T rangeTo, string requestId = "")
        {
            Comparer<T> comparer = Comparer<T>.Default;

            if (comparer.Compare(rangeFrom, rangeTo) > 0)
            {
                throw new ArgumentException($"RequestId: {requestId} - {nameof(rangeFrom)} should be less or equal than {nameof(rangeTo)}");
            }

            if (comparer.Compare(input, rangeFrom) < 0 || comparer.Compare(input, rangeTo) > 0)
            {
                throw new ArgumentOutOfRangeException($"RequestId: {requestId} - Input {parameterName} was out of range", parameterName);
            }
        }

        public static void NotStatus200OK(this IGuardClause guardClause, HttpStatusCode responseCode, string methodName, string requestId ="")
        {
            if (responseCode != HttpStatusCode.OK)
                throw new ResponseException($"RequestId: {requestId} - Method name: {methodName} - InvalidHttpResponse code: {responseCode}");
        }

        public static void NotStatus200OK(this IGuardClause guardClause, StatusCodes responseCode, string methodName, string requestId = "")
        {
            if (responseCode != StatusCodes.Status200OK)
                throw new ResponseException($"RequestId: {requestId} - Method name: {methodName} - InvalidHttpResponse code: {responseCode.Name}");
        }

        public static void NotStatus201Created(this IGuardClause guardClause, HttpStatusCode responseCode, string methodName, string requestId = "")
        {
            if (responseCode != HttpStatusCode.Created)
                throw new ResponseException($"RequestId: {requestId} - Method name: {methodName} - InvalidHttpResponse code: {responseCode}");
        }

        public static void NotStatus201Created(this IGuardClause guardClause, StatusCodes responseCode, string methodName, string requestId = "")
        {
            if (responseCode != StatusCodes.Status201Created)
                throw new ResponseException($"RequestId: {requestId} - Method name: {methodName} - InvalidHttpResponse code: {responseCode.Name}");
        }

        public static void NotStatus202Accepted(this IGuardClause guardClause, HttpStatusCode responseCode, string methodName, string requestId = "")
        {
            if (responseCode != HttpStatusCode.Accepted)
                throw new ResponseException($"RequestId: {requestId} - Method name: {methodName} - InvalidHttpResponse code: {responseCode}");
        }

        public static void NotStatus202Accepted(this IGuardClause guardClause, StatusCodes responseCode, string methodName, string requestId = "")
        {
            if (responseCode != StatusCodes.Status202Accepted)
                throw new ResponseException($"RequestId: {requestId} - Method name: {methodName} - InvalidHttpResponse code: {responseCode.Name}");
        }

        public static void NotStatus204NoContent(this IGuardClause guardClause, HttpStatusCode responseCode, string methodName, string requestId = "")
        {
            if (responseCode != HttpStatusCode.NoContent)
                throw new ResponseException($"RequestId: {requestId} - Method name: {methodName} - InvalidHttpResponse code: {responseCode}");
        }

        public static void NotStatus204NoContent(this IGuardClause guardClause, StatusCodes responseCode, string methodName, string requestId = "")
        {
            if (responseCode != StatusCodes.Status204NoContent)
                throw new ResponseException($"RequestId: {requestId} - Method name: {methodName} - InvalidHttpResponse code: {responseCode.Name}");
        }

        public static void NullOrUndefined(this IGuardClause guardClause, string input, string parameterName, string requestId = "")
        {
            Guard.Against.Null(input, parameterName, requestId);
            if (input.ToLower() == "undefined")
            {
                throw new ArgumentException($"RequestId: {requestId} - Required input {parameterName} is undefined.", parameterName);
            }
        }
    }
}
