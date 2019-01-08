//
// NamedProperties.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com> and Travis Semple
//
// Copyright (c) 2015-2018 Magic-Sessions. (www.magic-sessions.com)
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
//

using System;
using System.Collections.Generic;
using System.Linq;
using MsgKit.Enums;
using MsgKit.Structures;
using OpenMcdf;
using System.Text;
using MsgKit.Helpers;

namespace MsgKit.Streams
{
    internal sealed class NamedProperties : List<NamedProperty>
    {
        #region Fields
        /// <summary>
        ///     <see cref="TopLevelProperties" />
        /// </summary>
        private readonly TopLevelProperties _topLevelProperties;

        /// <summary>
        ///     The offset index for a <see cref="NamedProperty"/>
        /// </summary>
        private ushort _namedPropertyIndex;
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object
        /// </summary>
        /// <param name="topLevelProperties">
        ///     <see cref="TopLevelProperties" />
        /// </param>
        public NamedProperties(TopLevelProperties topLevelProperties)
        {
            _topLevelProperties = topLevelProperties;
        }
        #endregion

        #region AddProperty
        /// <summary>
        ///     Adds a <see cref="NamedPropertyTag" />
        /// </summary>
        /// <remarks>
        ///     Only support for properties by ID for now.
        /// </remarks>
        /// <param name="mapiTag"></param>
        /// <param name="obj"></param>
        internal void AddProperty(NamedPropertyTag mapiTag, object obj)
        {
            // Named property field 0000. 0x8000 + property offset
            //_topLevelProperties.AddProperty(new PropertyTag((ushort)(0x8000 + _namedPropertyIndex++), mapiTag.Type), obj);

            var propertyIndex = (ushort)(0x8000 + this.Count);
            var kind = mapiTag.Name.StartsWith("PidName") ? PropertyKind.Name : PropertyKind.Lid;
            var namedProperty = new NamedProperty
            {
                NameIdentifier = kind == PropertyKind.Lid ? mapiTag.Id : propertyIndex,
                Guid = mapiTag.Guid,
                Kind = kind,
                Name = mapiTag.Name.Replace("PidName", "").Replace("PidLid",""),
                NameSize = (uint)(kind == PropertyKind.Name ? mapiTag.Name.Length : 0)
            };

            if(mapiTag.Guid != PropertySets.PS_MAPI 
                && mapiTag.Guid != PropertySets.PS_PUBLIC_STRINGS
                && !Guids.Contains(mapiTag.Guid))
            {
                Guids.Add(mapiTag.Guid);
            }

            _topLevelProperties.AddProperty(new PropertyTag(propertyIndex, mapiTag.Type), obj);

            Add(namedProperty);
        }
        #endregion

        private ushort GetGuidIndex(NamedProperty namedProperty)
        {
            return (ushort)(namedProperty.Guid == PropertySets.PS_MAPI ? 1
                            : namedProperty.Guid == PropertySets.PS_PUBLIC_STRINGS ? 2
                            : (Guids.IndexOf(namedProperty.Guid) + 3));
        }

        private IList<Guid> Guids = new List<Guid>();

        #region WriteProperties
        /// <summary>
        ///     Writes the properties to the <see cref="CFStorage" />
        /// </summary>
        /// <param name="storage"></param>
        /// <param name="messageSize"></param>
        /// <remarks>
        ///     Unfortunately this is going to have to be used after we already written the top level properties.
        /// </remarks>
        internal void WriteProperties(CFStorage storage, long messageSize = 0)
        {
            // Grab the nameIdStorage, 3.1 on the SPEC
            storage = storage.GetStorage(PropertyTags.NameIdStorage);

            var entryStream = new EntryStream(storage);
            var stringStream = new StringStream(storage);
            var guidStream = new GuidStream(storage);
            var nameIdMappingStream = new EntryStream(storage);

            ushort propertyIndex = 0;

            foreach (var guid in Guids.Where(g => g!= PropertySets.PS_MAPI && g != PropertySets.PS_PUBLIC_STRINGS))
                guidStream.Add(guid);

            foreach (var namedProperty in this)
            {
                var guidIndex = GetGuidIndex(namedProperty);

                var indexAndKind = new IndexAndKindInformation(propertyIndex, guidIndex, namedProperty.Kind);

                if (namedProperty.Kind == PropertyKind.Name)
                {
                    var stringStreamItem = new StringStreamItem(namedProperty.Name);
                    stringStream.Add(stringStreamItem);
                    entryStream.Add(new EntryStreamItem(stringStream.GetItemByteOffset(stringStreamItem), indexAndKind));
                }
                else
                {
                    entryStream.Add(new EntryStreamItem(namedProperty.NameIdentifier, indexAndKind));
                }

                nameIdMappingStream.Add(new EntryStreamItem(GenerateNameIdentifier(namedProperty), indexAndKind));
                nameIdMappingStream.Write(storage, GenerateStreamName(namedProperty));

                // Dependign on the property type. This is doing name. 
                //entryStream.Add(new EntryStreamItem(namedProperty.NameIdentifier, new IndexAndKindInformation(propertyIndex, guidIndex, PropertyKind.Lid))); //+3 as per spec.
                //entryStream2.Add(new EntryStreamItem(namedProperty.NameIdentifier, new IndexAndKindInformation(propertyIndex, guidIndex, PropertyKind.Lid)));



                // 3.2.2 of the SPEC Needs to be written, because the stream changes as per named object.
                nameIdMappingStream.Clear();
                propertyIndex++;
            }

            guidStream.Write(storage);
            entryStream.Write(storage);
            stringStream.Write(storage);
        }

        #region GenerateStreamString
        /// <summary>
        ///     Generates the stream id of the named properties
        /// </summary>
        /// <param name="namedProperty"></param>
        /// <returns></returns>
        internal string GenerateStreamName(NamedProperty namedProperty)
        {
            var guidTarget = GetGuidIndex(namedProperty);
            var identifier = GenerateNameIdentifier(namedProperty);
            switch (namedProperty.Kind)
            {
                case PropertyKind.Lid:
                    return "__substg1.0_" +
                           (((4096 + (identifier ^ (guidTarget << 1)) % 0x1F) << 16) | 0x00000102).ToString("X")
                           .PadLeft(8, '0');
                case PropertyKind.Name:
                    return "__substg1.0_" +
                           (((0x1000 + ((identifier ^ (guidTarget << 1 | 1)) % 0x1F)) << 16) | 0x00000102).ToString("X8")
                           .PadLeft(8, '0');
                default:
                    throw new NotImplementedException();
            }
        }

        internal uint GenerateNameIdentifier(NamedProperty namedProperty)
        {
            switch (namedProperty.Kind)
            {
                case PropertyKind.Lid:
                    return namedProperty.NameIdentifier;
                case PropertyKind.Name:
                    return Crc32Calculator.CalculateCrc32(Encoding.Unicode.GetBytes(namedProperty.Name));
                default:
                    throw new NotImplementedException();
            }
        }
        #endregion

        #endregion
    }
}