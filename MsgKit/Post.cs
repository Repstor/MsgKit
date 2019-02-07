//
// Email.cs
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
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using MsgKit.Enums;
using MsgKit.Helpers;
using MsgKit.Mime.Header;
using OpenMcdf;
using MessageImportance = MsgKit.Enums.MessageImportance;
using MessagePriority = MsgKit.Enums.MessagePriority;
using Stream = System.IO.Stream;

namespace MsgKit
{
    /// <summary>
    ///     A class used to make a new Outlook E-mail MSG file
    /// </summary>
    /// <remarks>
    ///     See https://msdn.microsoft.com/en-us/library/office/cc979231.aspx
    /// </remarks>
    public class Post : Message, IDisposable
    {
        #region Fields
        /// <summary>
        ///     The <see cref="Regex" /> to find the prefix in a subject
        /// </summary>
        private static readonly Regex SubjectPrefixRegex = new Regex(@"^(\D{1,3}:\s)(.*)$");

        /// <summary>
        ///     The E-mail <see cref="Attachments" />
        /// </summary>
        private Attachments _attachments;

        /// <summary>
        ///     The subject of the E-mail
        /// </summary>
        private string _subject;
        #endregion

        #region Properties
        /// <summary>
        ///     Returns the sender of the E-mail from the <see cref="Recipients" />
        /// </summary>
        public Sender Sender { get; }

        /// <summary>
        ///     Contains the e-mail address for the messaging user represented by the <see cref="Sender"/>.
        /// </summary>
        /// <remarks>
        ///     These properties are examples of the address properties for the messaging user who is being represented by the
        ///     <see cref="Receiving" /> user. They must be set by the incoming transport provider, which is also responsible for
        ///     authorization or verification of the delegate. If no messaging user is being represented, these properties should
        ///     be set to the e-mail address contained in the PR_RECEIVED_BY_EMAIL_ADDRESS (PidTagReceivedByEmailAddress) property.
        /// </remarks>
        public Representing Representing { get; }      


        /// <summary>
        /// Sets of Returns the name of the parent folder
        /// </summary>
        public string PostedTo { get; set; }


        /// <summary>
        ///     Returns or sets the Internet Message Id
        /// </summary>
        /// <remarks>
        ///     Corresponds to the message ID field as specified in [RFC2822].<br/><br/>
        ///     If set then this value will be used, when not set the value will be read from the
        ///     <see cref="Message.TransportMessageHeaders"/> when this property is set
        /// </remarks>
        public string InternetMessageId { get; set; }

        /// <summary>
        ///     Returns or set the the value of a Multipurpose Internet Mail Extensions (MIME) message's References header field
        /// </summary>
        /// <remarks>
        ///     If set then this value will be used, when not set the value will be read from the
        ///     <see cref="Message.TransportMessageHeaders"/> when this property is set
        /// </remarks>
        public string InternetReferences { get; set; }

        /// <summary>
        ///     Returns or sets the original message's PR_INTERNET_MESSAGE_ID (PidTagInternetMessageId) property value
        /// </summary>
        /// <remarks>
        ///     If set then this value will be used, when not set the value will be read from the
        ///     <see cref="Message.TransportMessageHeaders"/> when this property is set
        /// </remarks>
        public string InReplyToId { get; set; }       

        /// <summary>
        ///     Returns <c>true</c> when the message is set as a draft message
        /// </summary>
        public bool Draft { get; }

        /// <summary>
        ///     Specifies the format for an editor to use to display a message.   
        /// </summary>
        public MessageEditorFormat MessageEditorFormat { get; set; }


        #endregion

        #region Constructor
        /// <summary>
        ///     Creates this object and sets all the needed properties
        /// </summary>
        /// <param name="sender">The <see cref="Sender"/> of the E-mail</param>
        /// <param name="subject">The subject of the E-mail</param>
        /// <param name="draft">Set to <c>true</c> to save the E-mail as a draft message</param>
        public Post(Sender sender,
                     string subject,
                     bool draft = false)
        {
            Sender = sender;
            Subject = subject;
            Importance = MessageImportance.IMPORTANCE_NORMAL;
            IconIndex = MessageIconIndex.NewMail;
            Draft = draft;
        }
        
        #endregion


        #region WriteToStorage
        /// <summary>
        ///     Writes all the properties that are part of the <see cref="Email"/> object either as <see cref="CFStorage"/>'s
        ///     or <see cref="CFStream"/>'s to the <see cref="CompoundFile.RootStorage"/>
        /// </summary>
        internal new void WriteToStorage()
        {
            var rootStorage = CompoundFile.RootStorage;

            Class = MessageClass.IPM_Post;

            base.WriteToStorage();

            var attachmentCount = Attachments.Count;
            TopLevelProperties.AttachmentCount = attachmentCount;
            TopLevelProperties.NextAttachmentId = attachmentCount;

            //TopLevelProperties.AddProperty(PropertyTags.PR_ENTRYID, Mapi.GenerateEntryId());
            //TopLevelProperties.AddProperty(PropertyTags.PR_INSTANCE_KEY, Mapi.GenerateInstanceKey());
            TopLevelProperties.AddProperty(PropertyTags.PR_STORE_SUPPORT_MASK, StoreSupportMaskConst.StoreSupportMask, PropertyFlags.PROPATTR_READABLE);
            TopLevelProperties.AddProperty(PropertyTags.PR_STORE_UNICODE_MASK, StoreSupportMaskConst.StoreSupportMask, PropertyFlags.PROPATTR_READABLE);
            TopLevelProperties.AddProperty(PropertyTags.PR_ALTERNATE_RECIPIENT_ALLOWED, true, PropertyFlags.PROPATTR_READABLE);
            TopLevelProperties.AddProperty(PropertyTags.PR_HASATTACH, attachmentCount > 0);
           
            TopLevelProperties.AddProperty(PropertyTags.PR_PARENT_DISPLAY_W, PostedTo);


            if (!string.IsNullOrWhiteSpace(InternetMessageId))
                TopLevelProperties.AddOrReplaceProperty(PropertyTags.PR_INTERNET_MESSAGE_ID_W, InternetMessageId);

            if (!string.IsNullOrWhiteSpace(InternetReferences))
                TopLevelProperties.AddOrReplaceProperty(PropertyTags.PR_INTERNET_REFERENCES_W, InternetReferences);

            if (!string.IsNullOrWhiteSpace(InReplyToId))
                TopLevelProperties.AddOrReplaceProperty(PropertyTags.PR_IN_REPLY_TO_ID_W, InReplyToId);

            var messageFlags = MessageFlags.MSGFLAG_UNMODIFIED;

            if (attachmentCount > 0)
                messageFlags |= MessageFlags.MSGFLAG_HASATTACH;

            TopLevelProperties.AddProperty(PropertyTags.PR_INTERNET_CPID, Encoding.UTF8.CodePage);            

            if (MessageEditorFormat != MessageEditorFormat.EDITOR_FORMAT_DONTKNOW)
                TopLevelProperties.AddProperty(PropertyTags.PR_MSG_EDITOR_FORMAT, MessageEditorFormat);

            if (!SentOn.HasValue)
                SentOn = DateTime.UtcNow;

            TopLevelProperties.AddProperty(PropertyTags.PR_ACCESS, MapiAccess.MAPI_ACCESS_DELETE | MapiAccess.MAPI_ACCESS_MODIFY | MapiAccess.MAPI_ACCESS_READ);
            TopLevelProperties.AddProperty(PropertyTags.PR_ACCESS_LEVEL, MapiAccess.MAPI_ACCESS_MODIFY);
            TopLevelProperties.AddProperty(PropertyTags.PR_OBJECT_TYPE, MapiObjectType.MAPI_MESSAGE);


            // http://www.meridiandiscovery.com/how-to/e-mail-conversation-index-metadata-computer-forensics/
            // http://stackoverflow.com/questions/11860540/does-outlook-embed-a-messageid-or-equivalent-in-its-email-elements
            //propertiesStream.AddProperty(PropertyTags.PR_CONVERSATION_INDEX, Subject);

            // TODO: Change modification time when this message is opened and only modified
            
            TopLevelProperties.AddProperty(PropertyTags.PR_PRIORITY, Priority);
           
            if (Draft)
            {
                messageFlags |= MessageFlags.MSGFLAG_UNSENT;
            }

            IconIndex = MessageIconIndex.Post;

            TopLevelProperties.AddProperty(PropertyTags.PR_MESSAGE_FLAGS, messageFlags);

            Sender?.WriteProperties(TopLevelProperties);
            Representing?.WriteProperties(TopLevelProperties);
            
        }
        #endregion

        #region Save
        /// <summary>
        ///     Saves the message to the given <paramref name="stream" />
        /// </summary>
        /// <param name="stream"></param>
        public new void Save(Stream stream)
        {
            WriteToStorage();
            base.Save(stream);
        }

        /// <summary>
        ///     Saves the message to the given <paramref name="fileName" />
        /// </summary>
        /// <param name="fileName"></param>
        public new void Save(string fileName)
        {
            WriteToStorage();
            base.Save(fileName);
        }
        #endregion

        #region Dispose
        /// <summary>
        ///     Disposes all the attachment streams
        /// </summary>
        public new void Dispose()
        {
            foreach (var attachment in _attachments)
                attachment.Stream.Dispose();

            base.Dispose();
        }
        #endregion
    }
}