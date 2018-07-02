// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using ApplicationCore.Helpers;
using ApplicationCore.Entities;
using ApplicationCore.Helpers.Exceptions;
using ApplicationCore.Artifacts;
using WebReact.ViewModels;

namespace WebReact.Serialization
{
    /// <summary>
    /// Converts a <see cref="SmartEnum"/> to and from a string.
    /// </summary>
    public class OpportunityStateModelConverter : JsonConverter
    {
        /// <summary>
        /// Writes the JSON representation of the object.
        /// </summary>
        /// <param name="writer">The <see cref="JsonWriter"/> to write to.</param>
        /// <param name="value">The value.</param>
        /// <param name="serializer">The calling serializer.</param>
        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            if (value == null)
            {
                writer.WriteNull();
            }
            else if (value is OpportunityStateModel)
            {
                var opportunityStateModel = (OpportunityStateModel)value;

                writer.WriteValue(opportunityStateModel);

                //writer.WriteStartObject();
                //writer.WritePropertyName("name");
                //serializer.Serialize(writer, value.name, reflectionObject.GetType(KeyName));
                //writer.WritePropertyName((resolver != null) ? resolver.GetResolvedPropertyName(ValueName) : ValueName);
                //serializer.Serialize(writer, reflectionObject.GetValue(value, ValueName), reflectionObject.GetType(ValueName));
                //writer.WriteEndObject();
            }
            else
            {
                throw new JsonSerializationException("Expected OpportunityStateModel object value");
            }
        }

        /// <summary>
        /// Reads the JSON representation of the object.
        /// </summary>
        /// <param name="reader">The <see cref="JsonReader"/> to read from.</param>
        /// <param name="objectType">Type of the object.</param>
        /// <param name="existingValue">The existing property value of the JSON that is being converted.</param>
        /// <param name="serializer">The calling serializer.</param>
        /// <returns>The object value.</returns>
        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            if (reader.TokenType == JsonToken.Null)
            {
                return null;
            }
            else
            {
                if (reader.TokenType == JsonToken.Integer)
                {
                    try
                    {
                        var v = OpportunityStateModel.FromValue(Convert.ToInt32(reader.Value));
                        return v;
                    }
                    catch (Exception ex)
                    {
                        //throw JsonSerializationException(reader, "Error parsing version string: {0}".FormatWith(CultureInfo.InvariantCulture, reader.Value), ex);
                        throw new JsonSerializationException($"Error parsing version string: {ex.Message}");
                    }
                }
                else
                {
                    //throw JsonSerializationException.Create(reader, "Unexpected token or value when parsing version. Token: {0}, Value: {1}".FormatWith(CultureInfo.InvariantCulture, reader.TokenType, reader.Value));
                    throw new JsonSerializationException($"Unexpected token or value when parsing version. Token: {reader.TokenType}, Value: {reader.Value}");
                }
            }
        }

        /// <summary>
        /// Determines whether this instance can convert the specified object type.
        /// </summary>
        /// <param name="objectType">Type of the object.</param>
        /// <returns>
        /// 	<c>true</c> if this instance can convert the specified object type; otherwise, <c>false</c>.
        /// </returns>
        public override bool CanConvert(Type objectType)
        {
            //return objectType == typeof(Version);

            return true;
        }
    }
}
