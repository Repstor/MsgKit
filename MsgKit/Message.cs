//
// Message.cs
//
// Author: Kees van Spelde <sicos2002@hotmail.com>
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
using System.Linq;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using MsgKit.Enums;
using MsgKit.Helpers;
using MsgKit.Streams;
using OpenMcdf;

// ReSharper disable InconsistentNaming

namespace MsgKit
{
    /// <summary>
    ///     The base class for all the different types of Outlook MSG files
    /// </summary>
    public class Message : IDisposable
    {
        # region Fields

        /// <summary>
        ///     The subject of the E-mail
        /// </summary>
        private string _subject;

        /// <summary>
        ///     The <see cref="Regex" /> to find the prefix in a subject
        /// </summary>
        private static readonly Regex SubjectPrefixRegex = new Regex(@"^(\D{1,3}:\s)(.*)$");

        /// <summary>
        ///     The E-mail <see cref="Attachments" />
        /// </summary>
        private Attachments _attachments;

        /// <summary>
        ///     The <see cref="MessageFlags" /> 
        /// </summary>
        protected MessageFlags _messageFlags;

        #endregion

        #region Properties
        /// <summary>
        ///     The <see cref="CompoundFile" />
        /// </summary>
        internal CompoundFile CompoundFile { get; }

        /// <summary>
        ///     The <see cref="MessageClass"/>
        /// </summary>
        internal MessageClass Class;

        /// <summary>
        ///     The E-mail <see cref="Attachments" />
        /// </summary>
        public Attachments Attachments
        {
            get { return _attachments ?? (_attachments = new Attachments()); }
        }

        /// <summary>
        ///     Returns <see cref="Class"/> as a string that is written into the MSG file
        /// </summary>
        internal string ClassAsString
        {
            get
            {
                switch (Class)
                {
                    case MessageClass.Unknown:
                        throw new ArgumentException("Class field is not set");
                    case MessageClass.IPM_Note:
                        return "IPM.Note";
                    case MessageClass.IPM_Note_SMIME:
                        return "IPM.Note.SMIME";
                    case MessageClass.IPM_Note_SMIME_MultipartSigned:
                        return "IPM.Note.SMIME.MultipartSigned";
                    case MessageClass.IPM_Note_Receipt_SMIME:
                        return "IPM.Note.Receipt.SMIME";
                    case MessageClass.IPM_Post:
                        return "IPM.Post";
                    case MessageClass.IPM_Octel_Voice:
                        return "IPM.Octel.Voice";
                    case MessageClass.IPM_Voicenotes:
                        return "IPM.Voicenotes";
                    case MessageClass.IPM_Sharing:
                        return "IPM.Sharing";
                    case MessageClass.REPORT_IPM_NOTE_NDR:
                        return "REPORT.IPM.NOTE.NDR";
                    case MessageClass.REPORT_IPM_NOTE_DR:
                        return "REPORT.IPM.NOTE.DR";
                    case MessageClass.REPORT_IPM_NOTE_DELAYED:
                        return "REPORT.IPM.NOTE.DELAYED";
                    case MessageClass.REPORT_IPM_NOTE_IPNRN:
                        return "*REPORT.IPM.NOTE.IPNRN";
                    case MessageClass.REPORT_IPM_NOTE_IPNNRN:
                        return "*REPORT.IPM.NOTE.IPNNRN";
                    case MessageClass.REPORT_IPM_SCHEDULE_MEETING_REQUEST_NDR:
                        return "REPORT.IPM.SCHEDULE. MEETING.REQUEST.NDR";
                    case MessageClass.REPORT_IPM_SCHEDULE_MEETING_RESP_POS_NDR:
                        return "REPORT.IPM.SCHEDULE.MEETING.RESP.POS.NDR";
                    case MessageClass.REPORT_IPM_SCHEDULE_MEETING_RESP_TENT_NDR:
                        return "REPORT.IPM.SCHEDULE.MEETING.RESP.TENT.NDR";
                    case MessageClass.REPORT_IPM_SCHEDULE_MEETING_CANCELED_NDR:
                        return "REPORT.IPM.SCHEDULE.MEETING.CANCELED.NDR";
                    case MessageClass.REPORT_IPM_NOTE_SMIME_NDR:
                        return "REPORT.IPM.NOTE.SMIME.NDR";
                    case MessageClass.REPORT_IPM_NOTE_SMIME_DR:
                        return "*REPORT.IPM.NOTE.SMIME.DR";
                    case MessageClass.REPORT_IPM_NOTE_SMIME_MULTIPARTSIGNED_NDR:
                        return "*REPORT.IPM.NOTE.SMIME.MULTIPARTSIGNED.NDR";
                    case MessageClass.REPORT_IPM_NOTE_SMIME_MULTIPARTSIGNED_DR:
                        return "*REPORT.IPM.NOTE.SMIME.MULTIPARTSIGNED.DR";
                    case MessageClass.IPM_Appointment:
                        return "IPM.Appointment";
                    case MessageClass.IPM_Task:
                        return "IPM.Task";
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
        }

        /// <summary>
        ///     Contains a number that indicates which icon to use when you display a group
        ///     of e-mail objects. Default set to <see cref="MessageIconIndex.NewMail" />
        /// </summary>
        /// <remarks>
        ///     This property, if it exists, is a hint to the client. The client may ignore the
        ///     value of this property.
        /// </remarks>
        public MessageIconIndex IconIndex { get; set; }

        /// <summary>
        ///     The size of the message
        /// </summary>
        public long MessageSize { get; internal set; }

        /// <summary>
        ///     The <see cref="TopLevelProperties"/>
        /// </summary>
        internal TopLevelProperties TopLevelProperties;
        
        /// <summary>
        ///     The <see cref="NamedProperties"/>
        /// </summary>
        internal NamedProperties NamedProperties;

        /// <summary>
        ///     Returns or sets the UTC date and time the <see cref="Sender"/> has submitted the 
        ///     <see cref="Message"/>
        /// </summary>
        /// <remarks>
        ///     This property has to be set to UTC datetime. When not set then the current date 
        ///     and time is used
        /// </remarks>
        public DateTime? SentOn { get; set; }

        /// <summary>
        ///     Returns or sets the UTC date and time the <see cref="Sender"/> has created the
        ///     <see cref="Message"/>
        /// </summary>
        /// <remarks>
        ///     This property has to be set to UTC datetime. When not set then the current date 
        ///     and time is used
        /// </remarks>
        public DateTime? CreatedOn { get; set; }

        /// <summary>
        ///     Returns or sets the UTC date and time, when the message was last modified
        /// </summary>
        /// <remarks>
        ///     This property has to be set to UTC datetime. When not set then the current date 
        ///     and time is used
        /// </remarks>
        public DateTime? LastModifiedOn { get; set; }

        /// <summary>
        /// Name of the last user to modify the message
        /// </summary>
        public string LastModifiedBy { get; set; }

        /// <summary>
        ///     Returns or sets the text body of the E-mail
        /// </summary>
        public string BodyText { get; set; }

        /// <summary>
        ///     Returns or sets the html body of the E-mail
        /// </summary>
        public string BodyHtml { get; set; }

        /// <summary>
        ///     The compressed RTF body part
        /// </summary>
        /// <remarks>
        ///     When not set then the RTF is generated from <see cref="BodyHtml"/> (when this property is set)
        /// </remarks>
        public string BodyRtf { get; set; }

        /// <summary>
        ///     Returns or set to <c>true</c> when <see cref="BodyRtf"/> is compressed
        /// </summary>
        public bool BodyRtfCompressed { get; set; }

        /// <summary>
        ///     Returns the subject prefix of the E-mail
        /// </summary>
        public string SubjectPrefix { get; private set; }

        /// <summary>
        ///     Returns or sets the subject of the E-mail
        /// </summary>
        public string Subject
        {
            get { return _subject; }
            set
            {
                _subject = value;
                SetSubject();
            }
        }

        /// <summary>
        ///     Returns the normalized subject of the E-mail
        /// </summary>
        public string SubjectNormalized { get; private set; }

        /// <summary>
        ///     Returns or sets the  the depth of the reply in a hierarchical representation of Post objects in one conversation
        /// </summary>
        public byte[] ConversationIndex { get; set; }

        /// <summary>
        ///     contains an unchanging copy of the original subject.
        /// </summary>
        public string ConversationTopic { get; set; }

        /// <summary>
        ///     Gets or sets a valud indicating the message sender's opinion of the sensitivity of a message
        /// </summary>
        public long Sensitiviy { get; set; }

        /// <summary>
        ///     Returns or sets the <see cref="Enums.MessagePriority"/>
        /// </summary>
        public MessagePriority Priority { get; set; }

        /// <summary>
        ///     Returns or sets the <see cref="Enums.MessageImportance"/>
        /// </summary>
        public MessageImportance Importance { get; set; }

        /// <summary>
        ///     Returns or sets keywords or categories for the Message object
        /// </summary>
        public string[] Keywords { get; set; }

        /// <summary>
        ///     Returns or sets the text labels assigned to this Message object
        /// </summary>
        public string[] Categories { get; set; }
        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all it's properties
        /// </summary>
        internal Message()
        {
            CompoundFile = new CompoundFile();

            // In the preceding figure, the "__nameid_version1.0" named property mapping storage contains the 
            // three streams  used to provide a mapping from property ID to property name 
            // ("__substg1.0_00020102", "__substg1.0_00030102", and "__substg1.0_00040102") and various other 
            // streams that provide a mapping from property names to property IDs.
            var nameIdStorage = CompoundFile.RootStorage.TryGetStorage(PropertyTags.NameIdStorage) ??
                                CompoundFile.RootStorage.AddStorage(PropertyTags.NameIdStorage);

            var entryStream = nameIdStorage.AddStream(PropertyTags.EntryStream);
            entryStream.SetData(new byte[0]);
            var stringStream = nameIdStorage.AddStream(PropertyTags.StringStream);
            stringStream.SetData(new byte[0]);
            var guidStream = nameIdStorage.AddStream(PropertyTags.GuidStream);
            guidStream.SetData(new byte[0]);

            TopLevelProperties = new TopLevelProperties();
            NamedProperties = new NamedProperties(TopLevelProperties);

            Importance = MessageImportance.IMPORTANCE_NORMAL;
        }

        internal void WriteToStorage()
        {
            var rootStorage = CompoundFile.RootStorage;
            MessageSize += Attachments.WriteToStorage(rootStorage);
            var attachmentCount = Attachments.Count;

            TopLevelProperties.NextAttachmentId = attachmentCount;
            TopLevelProperties.AttachmentCount = attachmentCount;
            TopLevelProperties.AddProperty(PropertyTags.PR_HASATTACH, attachmentCount > 0);

            var messageFlags = MessageFlags.MSGFLAG_UNMODIFIED;

            if (attachmentCount > 0)
                messageFlags |= MessageFlags.MSGFLAG_HASATTACH;

            if (!SentOn.HasValue)
                SentOn = DateTime.UtcNow;

            if (!CreatedOn.HasValue)
                CreatedOn = DateTime.UtcNow;

            if (!LastModifiedOn.HasValue)
                LastModifiedOn = DateTime.UtcNow;

            TopLevelProperties.AddProperty(PropertyTags.PR_CLIENT_SUBMIT_TIME, SentOn.Value.ToUniversalTime());
            TopLevelProperties.AddProperty(PropertyTags.PR_CREATION_TIME, CreatedOn.Value.ToUniversalTime());
            TopLevelProperties.AddProperty(PropertyTags.PR_LAST_MODIFICATION_TIME, LastModifiedOn.Value.ToUniversalTime());
            TopLevelProperties.AddProperty(PropertyTags.PR_BODY_W, BodyText);

            if (string.IsNullOrWhiteSpace(BodyRtf) && !string.IsNullOrWhiteSpace(BodyHtml))
            {
                BodyRtf = Strings.GetEscapedRtf(BodyHtml);
                BodyRtfCompressed = true;
            }

            if (!string.IsNullOrWhiteSpace(BodyRtf))
            {
                TopLevelProperties.AddProperty(PropertyTags.PR_RTF_COMPRESSED, new RtfCompressor().Compress(Encoding.ASCII.GetBytes(BodyRtf)));
                TopLevelProperties.AddProperty(PropertyTags.PR_RTF_IN_SYNC, BodyRtfCompressed);
            }

            SetSubject();
            TopLevelProperties.AddProperty(PropertyTags.PR_SUBJECT_W, Subject);
            TopLevelProperties.AddProperty(PropertyTags.PR_NORMALIZED_SUBJECT_W, SubjectNormalized);
            TopLevelProperties.AddProperty(PropertyTags.PR_SUBJECT_PREFIX_W, SubjectPrefix);
            TopLevelProperties.AddProperty(PropertyTags.PR_CONVERSATION_TOPIC_W, ConversationTopic);
            TopLevelProperties.AddProperty(PropertyTags.PR_CONVERSATION_INDEX, ConversationIndex);
            TopLevelProperties.AddProperty(PropertyTags.PR_LAST_MODIFIER_NAME_W, LastModifiedBy);
            TopLevelProperties.AddProperty(PropertyTags.PR_SENSITIVITY, Sensitiviy);
            TopLevelProperties.AddProperty(PropertyTags.PR_PRIORITY, Priority);
            TopLevelProperties.AddProperty(PropertyTags.PR_IMPORTANCE, Importance);
            TopLevelProperties.AddProperty(PropertyTags.PR_ICON_INDEX, IconIndex);

            if (Keywords != null && Keywords.Any())
                NamedProperties.AddProperty(NamedPropertyTags.PidNameKeywords, Keywords);

            if (Categories != null && Categories.Any())
                NamedProperties.AddProperty(NamedPropertyTags.PidLidCategories, Categories);

        }
        #endregion

        #region SetSubject
        /// <summary>
        ///     These properties are computed by message store or transport providers from the PR_SUBJECT (PidTagSubject) 
        ///     and PR_SUBJECT_PREFIX (PidTagSubjectPrefix) properties in the following manner. If the PR_SUBJECT_PREFIX 
        ///     is present and is an initial substring of PR_SUBJECT, PR_NORMALIZED_SUBJECT and associated properties are 
        ///     set to the contents of PR_SUBJECT with the prefix removed. If PR_SUBJECT_PREFIX is present, but it is not 
        ///     an initial substring of PR_SUBJECT, PR_SUBJECT_PREFIX is deleted and recalculated from PR_SUBJECT using 
        ///     the following rule: If the string contained in PR_SUBJECT begins with one to three non-numeric characters 
        ///     followed by a colon and a space, then the string together with the colon and the blank becomes the prefix.
        ///     Numbers, blanks, and punctuation characters are not valid prefix characters. If PR_SUBJECT_PREFIX is not 
        ///     present, it is calculated from PR_SUBJECT using the rule outlined in the previous step.This property then 
        ///     is set to the contents of PR_SUBJECT with the prefix removed.
        /// </summary>
        /// <remarks>
        ///     When PR_SUBJECT_PREFIX is an empty string, PR_SUBJECT and PR_NORMALIZED_SUBJECT are the same. Ultimately, 
        ///     this property should be the part of PR_SUBJECT following the prefix. If there is no prefix, this property 
        ///     becomes the same as PR_SUBJECT.
        /// </remarks>
        protected void SetSubject()
        {
            if (!string.IsNullOrEmpty(SubjectPrefix) && !string.IsNullOrEmpty(Subject))
            {
                if (Subject.StartsWith(SubjectPrefix))
                {
                    SubjectNormalized = Subject.Substring(SubjectPrefix.Length);
                }
                else
                {
                    var matches = SubjectPrefixRegex.Matches(Subject);
                    if (matches.Count > 0)
                    {
                        SubjectPrefix = matches.OfType<Match>().First().Groups[1].Value;
                        SubjectNormalized = matches.OfType<Match>().First().Groups[2].Value;
                    }
                }
            }
            else if (!string.IsNullOrEmpty(Subject))
            {
                var matches = SubjectPrefixRegex.Matches(Subject);
                if (matches.Count > 0)
                {
                    SubjectPrefix = matches.OfType<Match>().First().Groups[1].Value;
                    SubjectNormalized = matches.OfType<Match>().First().Groups[2].Value;
                }
                else
                    SubjectNormalized = Subject;
            }
            else
                SubjectNormalized = Subject;

            if (SubjectPrefix == null) SubjectPrefix = string.Empty;
        }
        #endregion

        #region Save
        internal void Save()
        {
            TopLevelProperties.AddProperty(PropertyTags.PR_MESSAGE_CLASS_W, ClassAsString);
            TopLevelProperties.WriteProperties(CompoundFile.RootStorage, MessageSize);
            NamedProperties.WriteProperties(CompoundFile.RootStorage);
        }

        /// <summary>
        ///     Saves the message to the given <paramref name="fileName" />
        /// </summary>
        /// <param name="fileName"></param>
        internal void Save(string fileName)
        {
            Save();
            CompoundFile.Save(fileName);
        }

        /// <summary>
        ///     Saves the message to the given <paramref name="stream" />
        /// </summary>
        /// <param name="stream"></param>
        internal void Save(Stream stream)
        {
            Save();
            CompoundFile.Save(stream);
        }
        #endregion

        #region Dispose
        /// <summary>
        ///     Disposes this object and all its resources
        /// </summary>
        public void Dispose()
        {
            foreach (var attachment in Attachments)
                attachment.Stream.Dispose();

            CompoundFile?.Close();
        }
        #endregion
    }
}