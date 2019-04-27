﻿using System;
using System.IO;
using System.Runtime.Serialization;
using System.Xml;

namespace Rubberduck.SettingsProvider
{
    internal class XmlContractPersistenceService<T> : XmlPersistenceServiceBase<T> where T : class, IEquatable<T>, new()
    {
        private const string DefaultConfigFile = "rubberduck.references";

        // ReSharper disable once StaticMemberInGenericType
        private static readonly DataContractSerializerSettings SerializerSettings = new DataContractSerializerSettings
        {
            RootNamespace = XmlDictionaryString.Empty            
        };

        public XmlContractPersistenceService(IPersistencePathProvider pathProvider) : base(pathProvider) { }

        protected override string FilePath => Path.Combine(RootPath, DefaultConfigFile);

        protected override T Read(string filePath)
        {
            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                using (var reader = XmlReader.Create(stream))
                {
                    var serializer = new DataContractSerializer(typeof(T));
                    return (T)serializer.ReadObject(reader);
                }
            }
            catch (Exception ex)
            {
                return default;
            }
        }

        protected override void Write(T toSerialize, string filePath)
        {
            // overwriting on write is intentional, we only expect this to be used for References settings
            using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            using (var writer = XmlWriter.Create(stream, OutputXmlSettings))
            {
                var serializer = new DataContractSerializer(typeof(T), SerializerSettings);
                serializer.WriteObject(writer, toSerialize);
            }
        }
    }
}
