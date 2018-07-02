// Copyright(c) Microsoft Corporation. 
// All rights reserved.
//
// Licensed under the MIT license. See LICENSE file in the solution root folder for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using System.IO;
using Microsoft.Azure;

namespace SmartLink.Service
{
    public class AES
    {
        private static byte[] _key;
        private UTF8Encoding _encoder = new UTF8Encoding();
        private RijndaelManaged _rm = new RijndaelManaged();

        static AES()
        {
            var sourceKey = CloudConfigurationManager.GetSetting("Key");
            _key = sourceKey.Split(',').Select(o => byte.Parse(o)).ToArray();
        }

        public string Encrypt(string unencrypted)
        {
            return Convert.ToBase64String(Encrypt(_encoder.GetBytes(unencrypted)));
        }

        public string Decrypt(string encrypted)
        {
            return _encoder.GetString(Decrypt(Convert.FromBase64String(encrypted)));
        }

        public byte[] Encrypt(byte[] buffer)
        {
            byte[] retValue;
            using (var rm = new RijndaelManaged())
            {
                rm.GenerateIV();
                var iv = rm.IV;

                using (var encryptor = rm.CreateEncryptor(_key, iv))
                using (var cipherStream = new MemoryStream())
                {
                    cipherStream.Write(iv, 0, 16);

                    using (CryptoStream cs = new CryptoStream(cipherStream, encryptor, CryptoStreamMode.Write))
                    {
                        cs.Write(buffer, 0, buffer.Length);
                    }
                    retValue = cipherStream.ToArray();
                }
            }
            return retValue;
        }

        public byte[] Decrypt(byte[] buffer)
        {
            byte[] iv = new byte[16];
            Array.Copy(buffer, iv, 16);

            byte[] retValue;
            using (var rm = new RijndaelManaged())
            {
                using (var decryptor = rm.CreateDecryptor(_key, iv))
                using (var ms = new MemoryStream())
                {
                    using (CryptoStream cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Write))
                    {
                        cs.Write(buffer, 16, buffer.Length - 16);
                    }
                    retValue = ms.ToArray();
                }
            }
            return retValue;
        }
    }
}
